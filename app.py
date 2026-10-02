
# ================== УТИЛИТЫ И КОНФИГ ==================
import re
from io import BytesIO
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st
import yaml
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment

CFG = yaml.safe_load("""
columns:
  doctor:       ["врач", "фио врача", "доктор"]
  spec:         ["специализация доктора", "специализация", "специальность"]
  hours_plan:   ["часов по табелю", "рабочих часов по графику"]
  hours_patient: ["часов по записи", "часов по дошедшим", "время с пациентом"]
  visits:       ["количество услуг", "кол-во посещений", "посещений"]
  sum_rub:      ["сумма фактич", "общая сумма", "сумма"]
  patient:      ["пациентов", "кол-во пациентов", "фио пациента"]
  card:         ["комп. номер", "номер карты"]
  to_pay:       ["выставление в оплату"]
  service:      ["название услуги", "услуга"]
  first_visit:  ["первая услуга", "первое обращение", "первичн"]
  date:         ["дата услуги", "дата первой услуги", "дата"]
  pot_hours:    ["рабочих часов всего"]
  pot_revenue:  ["потенциал"]
service_keep: ["допплер", "узи", "прием", "приём", "эхокг", "эхо кг", "эхо-кг"]
to_pay_keep_value: "есть"
stavka_norm: 148.8
overload_threshold: 80
""")


def norm_text(s) -> str:
    s = str(s).lower().strip()
    s = s.replace("ё", "е")
    return re.sub(r"\s+", " ", s)


def doctor_key(name: str) -> str:
    k = norm_text(name).replace(".", " ")
    return re.sub(r"\s+", " ", k).strip()


TOTAL_MARKERS = ("итого", "всего", "сумма", "отображаемые")


def _is_total(val: str) -> bool:
    v = norm_text(val)
    return any(m in v for m in TOTAL_MARKERS)


def find_header_row(df, must_have, max_scan=40):
    best, best_score = None, -1
    for i in range(min(max_scan, len(df))):
        row = df.iloc[i]
        cells = [norm_text(c) for c in row.tolist()]
        if all(any(m in cell for cell in cells) for m in must_have):
            score = row.notna().sum()
            if score > best_score:
                best, best_score = i, score
    if best is not None:
        return best
    raise ValueError(f"Не нашел строку заголовка (нужны: {must_have}).")


def build_table(df, header_row):
    t = df.iloc[header_row + 1:].copy()
    t.columns = [norm_text(c) for c in df.iloc[header_row].tolist()]
    # дубли названий колонок (пустые хвосты, повторы в выгрузке) схлопываем
    t = t.loc[:, ~t.columns.duplicated()]
    return t.dropna(how="all")


def find_col(cols, key, required=True):
    # 1) точное совпадение с синонимом (по приоритету синонимов),
    # 2) затем подстрока (по приоритету синонимов). Точное важнее:
    # иначе «Фактическая загрузка врача» перебивает «Доктор»,
    # а «Специализация доктора» — «Врач».
    norms = {col: norm_text(col) for col in cols}
    for syn in CFG["columns"][key]:
        for col, n in norms.items():
            if n == syn:
                return col
    for syn in CFG["columns"][key]:
        for col, n in norms.items():
            if syn in n:
                return col
    if required:
        raise ValueError(f"Колонка '{key}' не найдена. Есть: {list(cols)}")
    return None


def read_any(uploaded):
    name = uploaded.name.lower()
    engine = "xlrd" if name.endswith(".xls") else "openpyxl"
    return pd.read_excel(BytesIO(uploaded.read()), sheet_name=0,
                         header=None, engine=engine)


def detect_month_label(df_raw, default=""):
    """Месяц отчета = месяц, в котором больше всего дней периода.
    Период берем из шапки ('С: 05.08.2026' / 'ПО: 05.09.2026'): даже если
    границы идут '5-го по 5-е', выигрывает месяц с большим числом дней.
    """
    start = end = None
    for i in range(min(8, len(df_raw))):
        for v in df_raw.iloc[i].tolist():
            t = norm_text(v)
            m1 = re.search(r"с:\s*(\d{2})\.(\d{2})\.(\d{4})", t)
            m2 = re.search(r"по:\s*(\d{2})\.(\d{2})\.(\d{4})", t)
            if m1:
                start = date(int(m1.group(3)), int(m1.group(2)),
                             int(m1.group(1)))
            if m2:
                end = date(int(m2.group(3)), int(m2.group(2)),
                           int(m2.group(1)))
    if end is None:
        return default
    if start is None or start > end:
        return f"{end.strftime('%m.%Y')}"
    # пересечение периода с каждым месяцем
    from calendar import monthrange
    best, best_days = None, -1
    y, m = start.year, start.month
    while (y, m) <= (end.year, end.month):
        m_start = date(y, m, 1)
        m_end = date(y, m, monthrange(y, m)[1])
        days = (min(end, m_end) - max(start, m_start)).days + 1
        if days > best_days:
            best, best_days = f"{m:02d}.{y}", days
        m += 1
        if m == 13:
            y, m = y + 1, 1
    return best or default

# ================== ПАРСЕРЫ ОТЧЕТОВ ==================

DATE_RE = re.compile(r"^(\d{2})\.(\d{2})\.(\d{4})")


def parse_zagruzka(df_raw, month=None):
    """month=None -> сводные строки врачей за весь период.
    month='MM.YYYY' -> суммы по дням выбранного месяца."""
    hr = find_header_row(df_raw, must_have=("доктор", "специализация"))
    t = build_table(df_raw, hr)
    c_doc = find_col(t.columns, "doctor")
    c_spec = find_col(t.columns, "spec")

    def num_series(key):
        col = find_col(t.columns, key, required=False)
        if col is None:
            return pd.Series(0.0, index=t.index)
        return pd.to_numeric(t[col], errors="coerce").fillna(0)

    norm_rate = float(CFG.get("stavka_norm", 148.8))

    if month is not None:
        hs = num_series("hours_plan")
        hp = num_series("hours_patient")
        vs = num_series("visits")
        acc = {}   # key -> [fio, hours, hours_patient, visits]
        cur_key, cur_fio = None, ""
        for idx, r in t.iterrows():
            doc = str(r[c_doc]).strip()
            m = DATE_RE.match(doc)
            if m:
                if cur_key and f"{m.group(2)}.{m.group(3)}" == month:
                    a = acc.setdefault(
                        cur_key, [cur_fio, 0.0, 0.0, 0.0])
                    a[1] += float(hs[idx])
                    a[2] += float(hp[idx])
                    a[3] += float(vs[idx])
            else:
                spec = str(r[c_spec]).strip()
                if spec.lower() in ("", "nan", "none") and not _is_total(doc) \
                        and doc.lower() not in ("nan", "none"):
                    cur_key, cur_fio = doctor_key(doc), doc
        rows = [{"key": k, "ФИО врача": v[0], "Специализация": "",
                 "Рабочих часов по графику": v[1],
                 "Время с пациентом": v[2],
                 "Кол-во посещений": v[3]} for k, v in acc.items()]
        return pd.DataFrame(rows), norm_rate

    def is_doctor_row(r):
        doc = str(r[c_doc]).strip()
        spec = str(r[c_spec]).strip()
        if spec.lower() not in ("", "nan", "none"):
            return False
        if DATE_RE.match(doc):
            return False
        return not _is_total(doc)

    t = t[t.apply(is_doctor_row, axis=1)]

    def num(key):
        col = find_col(t.columns, key, required=False)
        if col is None:
            return 0.0
        return pd.to_numeric(t[col], errors="coerce").fillna(0)

    df = pd.DataFrame({
        "key": t[c_doc].map(doctor_key),
        "ФИО врача": t[c_doc].astype(str).str.strip(),
        "Специализация": "",
        "Рабочих часов по графику": num("hours_plan"),
        "Время с пациентом": num("hours_patient"),
        "Кол-во посещений": num("visits"),
    })
    return df, norm_rate


def parse_obschaya_summa(df_raw):
    hr = find_header_row(df_raw, must_have=("врач", "сумма"))
    t = build_table(df_raw, hr)
    c_doc = find_col(t.columns, "doctor")
    c_val = find_col(t.columns, "sum_rub")
    c_spec = find_col(t.columns, "spec", required=False)
    t = t[t[c_doc].notna()]
    t = t[~t[c_doc].apply(_is_total)]
    t = t[t[c_doc].map(lambda x: not re.match(r"^\d{2}\.\d{2}\.\d{4}",
                                              str(x).strip()))]
    key = t[c_doc].map(doctor_key)
    val = pd.to_numeric(t[c_val], errors="coerce").fillna(0)
    g = pd.DataFrame({"key": key, "v": val}).groupby("key")["v"].sum()
    spec = None
    if c_spec is not None:
        spec = pd.DataFrame({"key": key, "s": t[c_spec].astype(str).str.strip()}
                            ).drop_duplicates("key").set_index("key")["s"]
    return g, spec


def parse_unic_patients(df_raw):
    hr = find_header_row(df_raw, must_have=("доктор", "пациентов"))
    t = build_table(df_raw, hr)
    c_doc = find_col(t.columns, "doctor")
    c_val = find_col(t.columns, "patient")
    t = t[t[c_doc].notna() & t[c_val].notna()]
    t = t[~t[c_doc].apply(_is_total)]
    key = t[c_doc].map(doctor_key)
    val = pd.to_numeric(t[c_val], errors="coerce").fillna(0)
    return pd.DataFrame({"key": key, "v": val}).groupby("key")["v"].sum()


def parse_pervoe_obr(df_raw, month=None):
    hr = find_header_row(df_raw, must_have=("доктор", "номер карты"))
    t = build_table(df_raw, hr)
    c_doc = find_col(t.columns, "doctor")
    t = t[t[c_doc].notna()]
    t = t[~t[c_doc].apply(_is_total)]
    if month is not None:
        c_date = find_col(t.columns, "date", required=False)
        if c_date is not None:
            d = pd.to_datetime(t[c_date], errors="coerce", format="mixed")
            t = t[d.dt.strftime("%m.%Y") == month]
    return t[c_doc].map(doctor_key).value_counts()


def parse_svodnyj_patients(df_raw, doctor_keys, month=None):
    hr = find_header_row(df_raw, must_have=("врач", "пациент"))
    t = build_table(df_raw, hr)
    c_doc = find_col(t.columns, "doctor")
    c_pat = find_col(t.columns, "patient")
    c_pay = find_col(t.columns, "to_pay")
    c_srv = find_col(t.columns, "service", required=False)
    c_id = find_col(t.columns, "card", required=False) or c_pat
    t = t[t[c_doc].notna() & t[c_pat].notna()]
    if month is not None:
        c_date = find_col(t.columns, "date", required=False)
        if c_date is not None:
            d = pd.to_datetime(t[c_date], errors="coerce", format="mixed")
            t = t[d.dt.strftime("%m.%Y") == month]
    t = t[t[c_doc].map(doctor_key).isin(doctor_keys)]
    t = t[t[c_pay].map(norm_text) == norm_text(CFG["to_pay_keep_value"])]
    if c_srv is not None:
        keep = [norm_text(s) for s in CFG["service_keep"]]
        t = t[t[c_srv].map(lambda v: any(k in norm_text(v) for k in keep))]
    t = t.drop_duplicates(subset=[c_doc, c_id])
    return t[c_doc].map(doctor_key).value_counts(), int(t[c_id].nunique())



# ================== ОСНОВНОЙ ОТЧЕТ (2 листа) ==================

MAIN_COLUMNS = ["ФИО врача", "Специализация", "Рабочих часов по графику",
                "Время с пациентом", "Стоимость оказанных услуг(руб)",
                "Кол-во посещений", "Кол-во пациентов", "Кол-во первичных"]
CALC_COLUMNS = ["Загрузка", "Стоимость фактического часа",
                "Стоимость посещения", "Выручка на физическое лицо",
                "Посещений 1-м физическим лицом", "% первичных к пациентам",
                "Состояние врача"]
ALL_COLUMNS = MAIN_COLUMNS + CALC_COLUMNS


def _div(a, b):
    return pd.to_numeric(a, errors="coerce") / pd.to_numeric(
        b, errors="coerce").replace(0, pd.NA)


def add_calc_columns(df):
    th = CFG["overload_threshold"]
    df = df.copy()
    df["Загрузка"] = _div(df["Время с пациентом"],
                          df["Рабочих часов по графику"]) * 100
    df["Стоимость фактического часа"] = _div(
        df["Стоимость оказанных услуг(руб)"], df["Время с пациентом"])
    df["Стоимость посещения"] = _div(df["Стоимость оказанных услуг(руб)"],
                                     df["Кол-во посещений"])
    df["Выручка на физическое лицо"] = _div(
        df["Стоимость оказанных услуг(руб)"], df["Кол-во пациентов"])
    df["Посещений 1-м физическим лицом"] = _div(df["Кол-во посещений"],
                                                df["Кол-во пациентов"])
    df["% первичных к пациентам"] = _div(df["Кол-во первичных"],
                                         df["Кол-во пациентов"]) * 100
    df["Состояние врача"] = df["Загрузка"].map(
        lambda z: "Перегруз" if pd.notna(z) and z > th else "Недогруз")
    return df


def build_main_table(zag, sum_by_doc, spec_by_doc, unic_by_doc, perv_by_doc,
                     svod_by_doc, svod_clinic_total):
    df = zag.copy()
    df["Стоимость оказанных услуг(руб)"] = df["key"].map(sum_by_doc).fillna(0)
    df["Специализация"] = df["key"].map(
        spec_by_doc if spec_by_doc is not None else {}).fillna(df["Специализация"])
    # по ТЗ: по врачам — из "Уникальных пациентов за период"
    # (выгружать за нужный месяц), Итого по клинике — из сводного
    df["Кол-во пациентов"] = df["key"].map(unic_by_doc).fillna(0).astype(int)
    df["Кол-во первичных"] = df["key"].map(perv_by_doc).fillna(0).astype(int)
    df = add_calc_columns(df.drop(columns=["key"]))
    tot = {c: df[c].sum() for c in MAIN_COLUMNS[2:]}
    tot["Кол-во пациентов"] = svod_clinic_total
    total_row = pd.DataFrame([{MAIN_COLUMNS[0]: "Итого по клинике",
                               MAIN_COLUMNS[1]: "", **tot}])
    total_row = add_calc_columns(total_row)
    return pd.concat([df, total_row], ignore_index=True)[ALL_COLUMNS]


def totals_for_dinamika(main_df):
    row = main_df[main_df["ФИО врача"] == "Итого по клинике"].iloc[0]
    return {"main_total_sum": row["Стоимость оказанных услуг(руб)"],
            "main_total_hours_plan": row["Рабочих часов по графику"],
            "main_total_hours_patient": row["Время с пациентом"],
            "main_total_patients": row["Кол-во пациентов"],
            "main_total_visits": row["Кол-во посещений"],
            "main_total_first": row["Кол-во первичных"]}


POT_COLUMNS = ["Специализация", "Норма ставки", "Количество ставок",
               "Рабочих часов всего", "Стоимость часа",
               "Целевая загрузка", "Потенциал"]


def parse_potential(df_raw):
    """Таблица потенциала (эталон, загружается файлом).
    Возвращает (часы_итого, потенциал_итого, сырой df для копии листа)."""
    hr = find_header_row(df_raw, must_have=("потенциал", "рабочих часов"))
    t = build_table(df_raw, hr)
    c_h = find_col(t.columns, "pot_hours")
    c_r = find_col(t.columns, "pot_revenue")
    total = t[t.apply(lambda r: any(_is_total(v) for v in r.tolist()),
                      axis=1)]
    if total.empty:
        raise ValueError("В таблице потенциала не найдена строка 'Итого'.")
    row = total.iloc[0]
    hours = float(pd.to_numeric(pd.Series([row[c_h]]),
                                errors="coerce").fillna(0).iloc[0])
    revenue = float(pd.to_numeric(pd.Series([row[c_r]]),
                                  errors="coerce").fillna(0).iloc[0])
    return hours, revenue, df_raw


def write_main_report(month_label, main_df, potential_raw, pot_hours,
                      pot_revenue):
    """Excel с листом 'Потенциал' и листом 'Отчет по докторам'.
    Расчетные колонки — формулы Excel (как в референсе)."""
    th = CFG["overload_threshold"]
    wb = Workbook()
    # --- лист 1: Потенциал (копия загруженного эталона) ---
    ws0 = wb.active
    ws0.title = "Потенциал"
    if potential_raw is not None:
        for i, row in potential_raw.iterrows():
            for j, v in enumerate(row.tolist()):
                if pd.notna(v):
                    if isinstance(v, str) and v.strip() == "Фреболог":
                        v = "Флеболог"  # опечатка в исходном эталоне
                    ws0.cell(i + 1, j + 1, v)
    for j in range(1, 10):
        ws0.column_dimensions[get_column_letter(j)].width = 16
    bold = Font(bold=True)

    # --- лист 2: Отчет по докторам ---
    EXTRA = ["Главный вывод по врачу", "План действий по врачу"]
    ws = wb.create_sheet("Отчет по докторам")
    ws.cell(1, 1, f"Отчет по докторам, {month_label}").font = bold
    ws.cell(2, 1, f"Отчет по докторам, {month_label}").font = bold
    headers = ["№"] + MAIN_COLUMNS + CALC_COLUMNS + EXTRA
    for j, col in enumerate(headers, start=1):
        c = ws.cell(3, j, col)
        c.font = bold
        c.alignment = Alignment(wrap_text=True, vertical="center")
    first_r = 4
    last_r = first_r + len(main_df) - 1          # строка Итого
    for i, (_, row) in enumerate(main_df.iterrows()):
        r = first_r + i
        ws.cell(r, 1, i + 1 if i < len(main_df) - 1 else "")
        ws.cell(r, 2, row["ФИО врача"])
        ws.cell(r, 3, row["Специализация"])
        for j, col in enumerate(MAIN_COLUMNS[2:], start=4):
            v = row[col]
            if not pd.isna(v):
                ws.cell(r, j, round(float(v), 4)
                        if isinstance(v, (int, float)) else v)
        # расчетные колонки — формулы
        ws.cell(r, 10, f'=IFERROR(E{r}/D{r}*100,"")')          # Загрузка
        ws.cell(r, 11, f'=IFERROR(F{r}/E{r},"")')              # факт. час
        ws.cell(r, 12, f'=IFERROR(F{r}/G{r},"")')              # стоим. посещения
        ws.cell(r, 13, f'=IFERROR(F{r}/H{r},"")')              # выручка на ФЛ
        ws.cell(r, 14, f'=IFERROR(G{r}/H{r},"")')              # посещений 1-м ФЛ
        ws.cell(r, 15, f'=IFERROR(I{r}/H{r}*100,"")')          # % первичных
        ws.cell(r, 16, f'=IF(J{r}="","",IF(J{r}>{th},'
                       f'"Перегруз","Недогруз"))')             # Состояние
        for j in range(10, 16):
            ws.cell(r, j).number_format = "#,##0.00"
    # Итого: суммы по врачам в базовых колонках
    for j in range(4, 10):
        L = get_column_letter(j)
        ws.cell(last_r, j, f"=SUM({L}{first_r}:{L}{last_r - 1})")
    for j in range(1, len(headers) + 1):
        ws.cell(last_r, j).font = bold

    # --- блок под таблицей ---
    tot = main_df.iloc[-1]
    r0 = last_r + 2
    ws.cell(r0, 4, round(float(tot["Рабочих часов по графику"]) / 1.488, 3))
    ws.cell(r0, 6, round(float(pd.to_numeric(
        main_df["Стоимость фактического часа"], errors="coerce")
        .iloc[:-1].mean()), 3))
    block = [("Исп мощности",
              round(float(tot["Рабочих часов по графику"]) / pot_hours * 100, 2)
              if pot_hours else None),
             ("Дост потенциала",
              round(float(tot["Стоимость оказанных услуг(руб)"]) / pot_revenue * 100, 2)
              if pot_revenue else None),
             ("Альтернативный потенциал", None),
             ("Разница с классическим", None),
             ("Вывод", None)]
    for k, (label, val) in enumerate(block):
        ws.cell(r0 + 1 + k, 5, label).font = bold
        if val is not None:
            ws.cell(r0 + 1 + k, 7, val)
    for j, col in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(j)].width = max(12, len(col) // 2 + 4)
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ================== ОБНОВЛЕНИЕ ФАЙЛА «ДИНАМИКА» ==================

RU_MONTHS = {1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель", 5: "Май",
             6: "Июнь", 7: "Июль", 8: "Август", 9: "Сентябрь",
             10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"}


def _fmt_month(v):
    if isinstance(v, pd.Timestamp):
        return v.strftime("%m.%Y")
    if isinstance(v, (date, datetime)):
        return v.strftime("%m.%Y")
    if isinstance(v, (int, float)):
        if 20000 < v < 80000:
            return (date(1899, 12, 30) + timedelta(days=int(v))).strftime("%m.%Y")
        if 2000 < v < 2100:  # хранится как MM.YYYY, напр. 8.2025
            m = int(v)
            return f"{m:02d}.{int(round((v - m) * 10000))}"
    return str(v)


def _ru_month(label):
    if isinstance(label, (pd.Timestamp, datetime, date)):
        return RU_MONTHS[int(label.strftime("%m"))]
    m = re.match(r"(\d{2})\.(\d{4})", str(label))
    if m:
        return RU_MONTHS[int(m.group(1))]
    return str(label)


def _short_key(name):
    """'Гук Наталья Владимировна' / 'Гук Н.В.' -> 'гук нв'"""
    toks = norm_text(name).replace(".", " ").split()
    if not toks:
        return ""
    return toks[0] + " " + "".join(t[0] for t in toks[1:])


DERIVED = {
    "коэффициент использования мощности":
        '=IFERROR({c}{r_hours}/{c}{r_power}*100,"")',
    "% достижения потенциала":
        '=IFERROR({c}{r_sum}/{c}{r_pot}*100,"")',
    "% загруженности клиники":
        '=IFERROR({c}{r_hp}/{c}{r_hours}*100,"")',
    "стоимость помощи на фл": '=IFERROR({c}{r_sum}/{c}{r_fl},"")',
    "посещений на фл": '=IFERROR({c}{r_vis}/{c}{r_fl},"")',
    "стоимость посещения": '=IFERROR({c}{r_sum}/{c}{r_vis},"")',
}
MANUAL = {"мощность в часах": "potential_hours_total",
          "потенциал по выручке": "potential_revenue_total",
          "стоимость помощи": "main_total_sum",
          "рабочие часы врачей": "main_total_hours_plan",
          "часы с пациентом": "main_total_hours_patient",
          "фл": "main_total_patients",
          "посещений": "main_total_visits",
          "первичных": "main_total_first"}


def _sheet1_extend(ws, ref, month_label, metrics):
    """Лист ДИНАМИКА: новый столбец после последнего месяца."""
    hdr = next(i for i in range(10) if ref.iloc[i].notna().sum() >= 3)
    header = ref.iloc[hdr].tolist()
    param_col = next(i for i, v in enumerate(header)
                     if norm_text(v) == "параметр")
    izm_col = next(i for i, v in enumerate(header) if "изм" in norm_text(v))
    month_cols = [i for i in range(param_col + 1, izm_col)
                  if pd.notna(header[i])]
    body = ref.iloc[hdr + 1:].dropna(how="all")
    labels = body[param_col].tolist()

    # вставляем новый столбец в ws (после последнего месяца, перед Изм)
    insert_at = 3 + len(month_cols)  # 1-based: новый столбец на месте Изм
    ws.insert_cols(insert_at)
    new_col = insert_at
    prev_col = new_col - 1
    izm_new = new_col + 1

    # заголовки
    ws.cell(1, new_col, month_label).font = Font(bold=True)
    prev_label = ws.cell(1, prev_col).value
    ws.cell(1, izm_new,
            f"Изм {_ru_month(month_label)} к {_ru_month(prev_label)} в %"
            ).font = Font(bold=True)

    pos = {norm_text(l).replace("потеницал", "потенциал"): r_off + 2
           for r_off, l in enumerate(labels)}

    def find_row(key):
        key = key.replace("потеницал", "потенциал")
        for k, r in pos.items():
            if key in k:
                return r
        return None

    for r_off, label in enumerate(labels):
        r = r_off + 2
        k = norm_text(label).replace("потеницал", "потенциал")
        L = get_column_letter(new_col)
        if k in MANUAL and MANUAL[k] in metrics:
            c = ws.cell(r, new_col, metrics[MANUAL[k]])
            c.number_format = "#,##0.00"
        elif k in DERIVED:
            ws.cell(r, new_col, DERIVED[k].format(
                c=L, r_power=find_row("мощность в часах"),
                r_pot=find_row("потенциал по выручке"),
                r_sum=find_row("стоимость помощи"),
                r_hours=find_row("рабочие часы врачей"),
                r_hp=find_row("часы с пациентом"),
                r_fl=find_row("фл"), r_vis=find_row("посещений")))
        Lp = get_column_letter(prev_col)
        ws.cell(r, izm_new,
                f'=IFERROR({L}{r}/{Lp}{r}-1,"")').number_format = "0.0%"
    ws.column_dimensions[get_column_letter(new_col)].width = 12
    ws.column_dimensions[get_column_letter(izm_new)].width = 20


# метрики листа «Динамика по врачам» (порядок групп столбцов)
S2_METRICS = ["Рабочих часов по графику", "Стоимость оказанных услуг(руб)",
              "Загрузка", "Стоимость фактического часа",
              "Посещений 1-м физическим лицом", "% первичных к пациентам"]


def _month_like(v):
    if isinstance(v, (pd.Timestamp, datetime, date)):
        return True
    v = str(v)
    return v in RU_MONTHS.values() or re.match(r"^\d{2}\.\d{4}$", v)


def _sheet2_extend(ws, ref, month_label, main_df):
    """Лист 'Динамика по врачам': добавляем столбец нового месяца в каждую
    группу метрик и пересчитываем Итог."""
    df = main_df[main_df["ФИО врача"] != "Итого по клинике"].copy()
    by_doc = {_short_key(n): r for n, r in
              zip(df["ФИО врача"], df.to_dict("records"))}
    by_spec = {norm_text(k): g for k, g in df.groupby("Специализация")}

    # --- детекция структуры ---
    hdr_r = first_mc = None
    best = -1
    for r in range(1, 12):
        cells = [ws.cell(r, c).value for c in range(1, 12)]
        score = sum(1 for v in cells if _month_like(v))
        if score > best:
            best, hdr_r = score, r
    if hdr_r is None or best < 6:
        return
    for c in range(1, 12):
        if _month_like(ws.cell(hdr_r, c).value):
            first_mc = c
            break
    n_total = 0
    c = first_mc
    while c <= ws.max_column and ws.cell(hdr_r, c).value not in (None, ""):
        n_total += 1
        c += 1
    n_groups = 6
    if n_total % n_groups:
        return
    gm = n_total // n_groups  # месяцев в группе
    label_col = first_mc - 1
    ru_label = _ru_month(month_label)

    def block_values(label):
        k = norm_text(label)
        if k in by_spec:
            g = by_spec[k]
            return [float(pd.to_numeric(g[mm], errors="coerce").mean())
                    for mm in S2_METRICS]
        rec = by_doc.get(_short_key(label))
        if rec:
            return [float(rec[mm]) if pd.notna(rec[mm]) else None
                    for mm in S2_METRICS]
        return [None] * 6

    # --- вставка столбца в каждую группу (справа налево) ---
    for g in range(n_groups - 1, -1, -1):
        pos = first_mc + (g + 1) * gm
        if g == n_groups - 1:
            pos = first_mc + g * gm + gm  # после последнего месяца группы
        ws.insert_cols(pos)
    # новые позиции: группа g, новый месяц = first_mc + g*(gm+1) + gm
    for g in range(n_groups):
        col = first_mc + g * (gm + 1) + gm
        ws.cell(hdr_r, col, ru_label).font = Font(bold=True)

    # --- заполнение значений ---
    for r in range(hdr_r + 1, ws.max_row + 1):
        label = ws.cell(r, label_col).value
        if label in (None, ""):
            continue
        vals = block_values(label)
        for g in range(n_groups):
            v = vals[g]
            if v is not None:
                col = first_mc + g * (gm + 1) + gm
                ws.cell(r, col, round(v, 4))

    # --- пересчет Итог (последние 6 столбцов) ---
    itog_start = ws.max_column - 5
    group_ranges = []
    cur_group, cur_docs = None, []
    for r in range(hdr_r + 1, ws.max_row + 1):
        label = ws.cell(r, label_col).value
        if label in (None, ""):
            continue
        if norm_text(label) in by_spec:
            if cur_group:
                group_ranges.append((cur_group, cur_docs))
            cur_group, cur_docs = r, []
        elif cur_group is not None:
            cur_docs.append(r)
    if cur_group:
        group_ranges.append((cur_group, cur_docs))

    for group_r, doc_rows in group_ranges:
        for row in [group_r] + doc_rows:
            for g in range(n_groups):
                base = first_mc + g * (gm + 1)
                cells = [ws.cell(row, base + m).value
                         for m in range(gm + 1)]
                nums = [v for v in cells if isinstance(v, (int, float))]
                if not nums:
                    continue
                if row == group_r:
                    # взвешенное по числу врачей с данными за месяц
                    num = den = 0.0
                    for m, v in enumerate(cells):
                        if isinstance(v, (int, float)):
                            cnt = sum(
                                1 for dr in doc_rows
                                if isinstance(
                                    ws.cell(dr, base + m).value,
                                    (int, float)))
                            cnt = max(cnt, 1)
                            num += v * cnt
                            den += cnt
                    if den:
                        ws.cell(row, itog_start + g,
                                round(num / den, 4))
                else:
                    ws.cell(row, itog_start + g,
                            round(sum(nums) / len(nums), 4))


BASE_COLUMNS = ["Месяц", "ФИО врача", "Специализация",
                "Рабочих часов по графику", "Время с пациентом",
                "Стоимость оказанных услуг(руб)", "Кол-во посещений",
                "Кол-во пациентов", "Кол-во первичных", "Загрузка",
                "Стоимость фактического часа", "Стоимость посещения",
                "Выручка на физическое лицо",
                "Посещений 1-м физическим лицом",
                "% первичных к пациентам", "Состояние врача"]


def _to_initials(name):
    toks = str(name).replace(".", " ").split()
    if len(toks) < 2:
        return str(name)
    return f"{toks[0]} {''.join(t[0] + '.' for t in toks[1:])}"


def _base_extend(ws, ref, month_label, main_df):
    """Лист 'База': дописываем строки нового месяца."""
    df = main_df[main_df["ФИО врача"] != "Итого по клинике"].copy()
    ru = _ru_month(month_label)
    start = ws.max_row + 1
    for i, (_, row) in enumerate(df.iterrows()):
        r = start + i
        ws.cell(r, 1, ru)
        ws.cell(r, 2, _to_initials(row["ФИО врача"]))
        for j, col in enumerate(BASE_COLUMNS[2:], start=3):
            v = row[col]
            if pd.isna(v):
                continue
            c = ws.cell(r, j, round(float(v), 4)
                        if isinstance(v, (int, float)) else v)
            if isinstance(v, (int, float)):
                c.number_format = "#,##0.00"


def update_dinamika_file(reference_bytes, month_label, metrics, main_df):
    """Загруженный файл «Динамика» -> тот же файл + новый месяц на всех
    листах. Возвращает bytes обновленного xlsx."""
    sheets = pd.read_excel(BytesIO(reference_bytes), sheet_name=None,
                           header=None)
    wb = Workbook()
    wb.remove(wb.active)
    for name, ref in sheets.items():
        n = norm_text(name)
        ws = wb.create_sheet(title=str(name)[:31])
        for i, row in ref.iterrows():
            for j, v in enumerate(row.tolist()):
                if pd.notna(v):
                    ws.cell(i + 1, j + 1, v)
        if "динамика" in n and "врачам" not in n:
            _sheet1_extend(ws, ref, month_label, metrics)
        elif "врачам" in n:
            _sheet2_extend(ws, ref, month_label, main_df)
        elif "база" in n:
            _base_extend(ws, ref, month_label, main_df)
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ================== ИНТЕРФЕЙС ==================

st.set_page_config(page_title="Сборка месячных отчетов по врачам",
                   layout="wide")
st.title("Сборка месячных отчетов по врачам")
st.caption("Загрузите таблицу потенциала и 5 отчетов (названия файлов не "
           "важны — важна структура) + опционально референс листа «Динамика».")

with st.sidebar:
    st.header("Файлы")
    f_pot = st.file_uploader("Таблица потенциала (эталон)",
                             type=["xls", "xlsx"])
    f_zag = st.file_uploader("1. Загрузка врачей (без кабинетов)",
                             type=["xls", "xlsx"])
    f_sum = st.file_uploader("2. Общая сумма", type=["xls", "xlsx"])
    f_unic = st.file_uploader("3. Уникальных пациентов за период",
                              type=["xls", "xlsx"])
    f_perv = st.file_uploader("4. Первое обращение пациента",
                              type=["xls", "xlsx"])
    f_svod = st.file_uploader("5. Сводный по врачам с услугами и пациентами",
                              type=["xls", "xlsx"])
    f_dyn = st.file_uploader("Файл «Динамика» предыдущего месяца",
                             type=["xls", "xlsx"])

files = [f_zag, f_sum, f_unic, f_perv, f_svod]
if not all(files):
    st.info("Загрузите все 5 отчетов.")
    st.stop()

try:
    with st.spinner("Разбираю отчеты..."):
        pot_raw, pot_hours, pot_revenue = None, None, None
        if f_pot is not None:
            pot_raw = read_any(f_pot)
            pot_hours, pot_revenue, _ = parse_potential(pot_raw)
        zag_raw = read_any(f_zag)
        month_label = detect_month_label(zag_raw, default="")
        zag, norm_rate = parse_zagruzka(zag_raw)
        doctor_keys = set(zag["key"])
        sum_by_doc, spec_by_doc = parse_obschaya_summa(read_any(f_sum))
        unic_by_doc = parse_unic_patients(read_any(f_unic))
        perv_by_doc = parse_pervoe_obr(read_any(f_perv))
        svod_by_doc, svod_clinic_total = parse_svodnyj_patients(
            read_any(f_svod), doctor_keys)
        main_df = build_main_table(zag, sum_by_doc, spec_by_doc, unic_by_doc,
                                   perv_by_doc, svod_by_doc, svod_clinic_total)

    month_name = st.text_input("Месяц нового столбца «Динамики» "
                               "(подставлен из отчетов)",
                               value=month_label or "")

    st.subheader("Отчет по докторам")
    st.dataframe(main_df, use_container_width=True, hide_index=True)
    main_bytes = write_main_report(month_name, main_df, pot_raw,
                                   pot_hours or 0, pot_revenue or 0)
    st.download_button("Скачать отчет (xlsx, 2 листа)", main_bytes,
                       "Отчет_по_докторам.xlsx")

    metrics = totals_for_dinamika(main_df)
    metrics["potential_hours_total"] = pot_hours if pot_hours else None
    metrics["potential_revenue_total"] = pot_revenue if pot_revenue else None
    if f_pot is None:
        st.warning("Таблица потенциала не загружена — лист «Потенциал» "
                   "не будет в отчете, а строки «Мощность в часах» и "
                   "«Потенциал по выручке» в «Динамике» останутся пустыми.")

    st.subheader("Файл «Динамика»")
    if f_dyn:
        dyn_bytes = update_dinamika_file(BytesIO(f_dyn.read()).getvalue(),
                                         month_name, metrics, main_df)
        st.download_button("Скачать обновленную «Динамику» (xlsx)",
                           dyn_bytes, "Динамика.xlsx")
    else:
        st.info("Загрузите файл «Динамики» предыдущего месяца — получите "
                "его же с добавленным столбцом нового месяца.")

    with st.expander("Диагностика: как сматчились врачи"):
        c1, c2 = st.columns(2)
        c1.metric("Врачей в отчете", len(doctor_keys))
        c2.metric("Уникальных пациентов по клинике (из сводного)",
                  svod_clinic_total)
        if pot_hours:
            st.write(f"Потенциал: {pot_hours:.1f} ч, "
                     f"{pot_revenue:,.0f} руб.")
        missed = doctor_keys - set(svod_by_doc.index)
        if missed:
            st.write("Нет в сводном отчете:", sorted(missed))
except ValueError as e:
    st.error(str(e))
    st.stop()
