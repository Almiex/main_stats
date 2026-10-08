
# ================== УТИЛИТЫ И КОНФИГ ==================
import re
from io import BytesIO
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st
import yaml
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

CFG = yaml.safe_load("""
columns:
  doctor:       ["врач", "фио врача", "доктор"]
  spec:         ["специализация доктора", "специализация", "специальность"]
  hours_plan:   ["часов по табелю", "рабочих часов по графику"]
  hours_patient: ["часов по дошедшим", "время с пациентом"]
  visits:       ["дошедших пациентов", "кол-во посещений", "посещений"]
  visits_alt:   ["количество услуг"]
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
# подстроки в ФИО: такие строки не врачи (комиссии и пр.)
exclude_name_keywords: ["комиссия"]
# у этих специализаций посещения = "Количество услуг" (повторные
# обращения одного пациента видны только там), остальным = "Дошедших"
visits_uzi_specs: ["врач ультразвуковой диагностики", "врач функциональной диагностики"]
# принудительная специализация для конкретных врачей
# (ключ: фамилия + инициалы, значение: специальность для отображения)
doctor_spec_override:
  "якушева е": "Терапевт"
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

    def is_excluded_name(doc):
        d = norm_text(doc)
        return any(k in d for k in CFG.get("exclude_name_keywords", []))

    def is_doctor_row(r):
        doc = str(r[c_doc]).strip()
        spec = str(r[c_spec]).strip()
        if spec.lower() not in ("", "nan", "none"):
            return False
        if DATE_RE.match(doc):
            return False
        if is_excluded_name(doc):
            return False
        return not _is_total(doc)

    # специализация врача — берем из строк по дням (там она заполнена)
    spec_map = {}
    cur = None
    for _, r in t.iterrows():
        doc = str(r[c_doc]).strip()
        if DATE_RE.match(doc):
            sp = str(r[c_spec]).strip()
            if sp.lower() not in ("", "nan", "none") and cur:
                spec_map[cur] = sp
        else:
            if is_doctor_row(r):
                cur = doctor_key(doc)

    t = t[t.apply(is_doctor_row, axis=1)]

    def num(key):
        col = find_col(t.columns, key, required=False)
        if col is None:
            return 0.0
        return pd.to_numeric(t[col], errors="coerce").fillna(0)

    uzi_specs = [norm_text(x) for x in CFG.get("visits_uzi_specs", [])]
    is_uzi = t[c_doc].map(doctor_key).map(
        lambda k: any(u in norm_text(spec_map.get(k, "")) for u in uzi_specs))
    visits = num("visits")
    if is_uzi.any():
        visits = visits.copy()
        visits[is_uzi] = num("visits_alt")[is_uzi]

    df = pd.DataFrame({
        "key": t[c_doc].map(doctor_key),
        "ФИО врача": t[c_doc].astype(str).str.strip(),
        "Специализация": "",
        "Рабочих часов по графику": num("hours_plan"),
        "Время с пациентом": num("hours_patient"),
        "Кол-во посещений": visits,
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
    # одна строка на пару (врач, специализация): врач с несколькими
    # специализациями даст несколько строк
    key = t[c_doc].map(doctor_key)
    val = pd.to_numeric(t[c_val], errors="coerce").fillna(0)
    spec = (t[c_spec].astype(str).str.strip()
            if c_spec is not None else pd.Series("", index=t.index))
    spec = spec.mask(spec.str.lower().isin(["nan", "none"]), "")
    out = pd.DataFrame({"key": key, "spec": spec, "v": val})
    return out.groupby(["key", "spec"], as_index=False)["v"].sum()


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
    c_card = find_col(t.columns, "card", required=False)
    clinic_total = int(t[c_card].nunique()) if c_card is not None else int(len(t))
    return t[c_doc].map(doctor_key).value_counts(), clinic_total


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
    # 1) "Выставление в оплату" = "Есть"
    if CFG.get("to_pay_keep_value"):
        t = t[t[c_pay].map(norm_text) == norm_text(CFG["to_pay_keep_value"])]
    # 2) только нужные врачи (полное совпадение ФИО со списком отчета загрузки)
    t = t[t[c_doc].map(doctor_key).isin(doctor_keys)]
    # 3) только услуги: допплерография / УЗИ / приемы / эхоКГ
    if c_srv is not None:
        keep = [norm_text(s) for s in CFG["service_keep"]]
        t = t[t[c_srv].map(lambda v: any(k in norm_text(v) for k in keep))]
    # 4) удаление дубликатов по "Комп. номер" (карта пациента), а не по ФИО:
    # один человек может числиться под разными написаниями ФИО у разных
    # врачей, и тогда дедуп по ФИО теряет пересечения
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


def build_main_table(zag, obs, unic_by_doc, perv_by_doc,
                     svod_by_doc, svod_clinic_total, perv_clinic_total=None):
    """obs_by_spec: DataFrame(key, spec, v) — суммы по парам врач+специализация.
    Врач с несколькими специализациями -> отдельная строка на каждую;
    часы/посещения/пациенты только на первой строке (Итого не задваивается)."""
    obs = obs.copy()
    # переопределение специальностей отдельных врачей (см. CFG)
    ov = {_short_key(k): v for k, v in
          (CFG.get("doctor_spec_override") or {}).items()}
    if ov:
        obs["spec"] = [ov.get(_short_key(k), sp)
                       for k, sp in zip(obs["key"], obs["spec"])]
        obs = obs.groupby(["key", "spec"], as_index=False)["v"].sum()
    by_key = {k: list(zip(g["spec"], g["v"]))
              for k, g in obs.groupby("key")} if len(obs) else {}
    rows = []
    for _, r in zag.iterrows():
        specs = by_key.get(r["key"]) or [("", 0.0)]
        for i, (spec, val) in enumerate(specs):
            d = r.to_dict()
            d["Специализация"] = spec
            d["Стоимость оказанных услуг(руб)"] = val
            d["_extra"] = i > 0
            if i > 0:  # доп. специализация: без часов/посещений
                for c in ("Рабочих часов по графику", "Время с пациентом",
                          "Кол-во посещений"):
                    d[c] = 0.0
            rows.append(d)
    df = pd.DataFrame(rows)
    # по ТЗ: по врачам — из "Уникальных пациентов за период"
    # (выгружать за нужный месяц), Итого по клинике — из сводного
    # Кол-во пациентов: уникальные по каждому врачу из отчета 20
    # (после фильтров; один пациент у нескольких врачей НЕ дублируется,
    # так как сводный расчет идет по врачам)
    df["Кол-во пациентов"] = df["key"].map(svod_by_doc).fillna(0).astype(int)
    df["Кол-во первичных"] = df["key"].map(perv_by_doc).fillna(0).astype(int)
    if df["_extra"].any():  # пациенты/первичные только на основной строке
        df.loc[df["_extra"], ["Кол-во пациентов", "Кол-во первичных"]] = 0
    df = add_calc_columns(df.drop(columns=["key", "_extra"]))
    tot = {c: df[c].sum() for c in MAIN_COLUMNS[2:]}
    # Итого по клинике: пациенты — из отчета 20 (сводный, уникальные
    # по клинике: один пациент мог быть у нескольких врачей);
    # первичные — как все прочие: Итого = сумма по врачам из отчета
    tot["Кол-во пациентов"] = svod_clinic_total
    total_row = pd.DataFrame([{MAIN_COLUMNS[0]: "Итого по клинике",
                               MAIN_COLUMNS[1]: "", **tot}])
    total_row = add_calc_columns(total_row)
    return pd.concat([df, total_row], ignore_index=True)[ALL_COLUMNS]


def totals_for_dinamika(main_df, pot_hours=None, pot_revenue=None):
    row = main_df[main_df["ФИО врача"] == "Итого по клинике"].iloc[0]
    out = {"main_total_sum": row["Стоимость оказанных услуг(руб)"],
           "main_total_hours_plan": row["Рабочих часов по графику"],
           "main_total_hours_patient": row["Время с пациентом"],
           "main_total_patients": row["Кол-во пациентов"],
           "main_total_visits": row["Кол-во посещений"],
           "main_total_first": row["Кол-во первичных"]}
    if pot_hours:
        out["main_isp"] = row["Рабочих часов по графику"] / pot_hours * 100
    if pot_revenue:
        out["main_dost"] = row["Стоимость оказанных услуг(руб)"] / pot_revenue * 100
    return out


POT_COLUMNS = ["Специализация", "Норма ставки", "Количество ставок",
               "Рабочих часов всего", "Стоимость часа",
               "Целевая загрузка", "Потенциал"]


def parse_potential(df_raw):
    """Таблица потенциала (эталон, загружается файлом).
    Возвращает (часы_итого, потенциал_итого, сырой df, позиция_строки_итого).
    Итоги считаем ПО СТРОКАМ СПЕЦИАЛИЗАЦИЙ: в файлах, сгенерированных
    openpyxl, строка 'Итого' и столбцы 'Рабочих часов всего'/'Потенциал'
    — формулы без кэшированных значений, и pandas видит там NaN."""
    hr = find_header_row(df_raw, must_have=("потенциал", "рабочих часов"))
    t = build_table(df_raw, hr)
    c_h = find_col(t.columns, "pot_hours")
    c_r = find_col(t.columns, "pot_revenue")
    is_total = t.apply(lambda r: any(_is_total(v) for v in r.tolist()),
                       axis=1)
    if not is_total.any():
        raise ValueError("В таблице потенциала не найдена строка 'Итого'.")

    h_num = pd.to_numeric(t[c_h], errors="coerce")
    r_num = pd.to_numeric(t[c_r], errors="coerce")
    # 1) сумма по строкам специализаций (без строки Итого)
    hours = float(h_num[~is_total].fillna(0).sum())
    revenue = float(r_num[~is_total].fillna(0).sum())

    # 2) файл сгенерирован программно: 'Рабочих часов всего' (=C*D) и
    #    'Потенциал' (=E*F*G) — формулы без кэшированных значений.
    #    Пересчитываем из литеральных колонок: часы = ставки*часы_на_ставку,
    #    потенциал = часы * стоимость часа * целевая загрузка.
    if hours == 0 or revenue == 0:
        def _find(*subs, exclude=None):
            for col in t.columns:
                n = norm_text(col)
                if all(s in n for s in subs) and col != exclude:
                    return col
            return None
        c_n = _find("количество", "ставок")
        c_hd = _find("часов", exclude=c_h)          # часы на одну ставку
        c_pr = _find("стоимость", "часа")
        c_ld = _find("целевая", "загрузк")
        if c_n is not None and c_hd is not None:
            st = pd.to_numeric(t[c_n], errors="coerce").fillna(0)
            hd = pd.to_numeric(t[c_hd], errors="coerce").fillna(0)
            if hours == 0:
                hours = float((st * hd)[~is_total].sum())
            if revenue == 0 and c_pr is not None and c_ld is not None:
                pr = pd.to_numeric(t[c_pr], errors="coerce").fillna(0)
                ld = pd.to_numeric(t[c_ld], errors="coerce").fillna(0)
                revenue = float((st * hd * pr * ld)[~is_total].sum())

    # 3) крайний случай: берём саму строку Итого
    if hours == 0:
        hours = float(h_num[is_total].fillna(0).sum())
    if revenue == 0:
        revenue = float(r_num[is_total].fillna(0).sum())

    # позиция строки Итого в исходном df (1-based) для ссылок вида '1. Потенциал'!E14
    pot_total_row = int(t.index[is_total][0]) + 1
    return hours, revenue, df_raw, pot_total_row


def write_main_report(month_label, main_df, potential_raw, pot_hours,
                      pot_revenue, pot_total_row=None):
    """Excel с листом 'Потенциал' и листом 'Отчет по докторам'.
    Расчетные колонки — формулы Excel (как в референсе)."""
    th = CFG["overload_threshold"]
    wb = Workbook()
    # --- лист 1: Потенциал (копия загруженного эталона) ---
    ws0 = wb.active
    ws0.title = "1. Потенциал"
    if potential_raw is not None:
        for i, row in potential_raw.iterrows():
            for j, v in enumerate(row.tolist()):
                if pd.notna(v):
                    if isinstance(v, str) and v.strip() == "Фреболог":
                        v = "Флеболог"  # опечатка в исходном эталоне
                    ws0.cell(i + 1, j + 1, v)
        # Формулы пишем по РЕАЛЬНЫМ позициям колонок загруженного эталона
        # (порядок/набор колонок может отличаться): часы_всего = ставки *
        # часы_на_ставку, потенциал = часы_всего * цена * загрузка.
        # Ничего не хардкодим: позиции ищем по именам в строке заголовка.
        hr0 = find_header_row(potential_raw,
                              must_have=("потенциал", "рабочих часов"))
        header0 = [norm_text(c) for c in potential_raw.iloc[hr0].tolist()]

        def _col0(*subs, exclude=None):
            for idx, h in enumerate(header0):
                if all(s in h for s in subs) and idx != exclude:
                    return idx + 1  # 1-based номер колонки
            return None

        L = get_column_letter
        c_pot_h = _col0("часов", "всего") or _col0("часов")
        c_st = _col0("количество", "ставок")
        c_hr = _col0("часов", exclude=c_pot_h - 1 if c_pot_h else None)
        c_pr = _col0("стоимость", "часа")
        c_ld = _col0("загрузк")
        c_pot = _col0("потенциал")
        c_sp = _col0("специализация")
        first_spec = hr0 + 2
        last_spec = None
        for r in range(hr0 + 2, potential_raw.shape[0] + 1):
            spec = str(potential_raw.iloc[r - 1, (c_sp or 2) - 1]).strip()
            if spec and spec.lower() not in ("nan", "none", "специализация"):
                if not _is_total(spec):
                    if c_pot_h and c_st and c_hr:
                        ws0.cell(r, c_pot_h,
                                 f"={L(c_st)}{r}*{L(c_hr)}{r}")
                    if c_pot and c_pot_h and c_pr and c_ld:
                        ws0.cell(r, c_pot,
                                 f"={L(c_pot_h)}{r}*{L(c_pr)}{r}*"
                                 f"{L(c_ld)}{r}")
                    last_spec = r
                else:
                    # Итого: суммы по всем специализациям
                    for cc in (c_st, c_pot_h, c_pot):
                        if cc and last_spec:
                            ws0.cell(r, cc,
                                     f"=SUM({L(cc)}{first_spec}:"
                                     f"{L(cc)}{last_spec})")
                    break
    for j in range(1, 10):
        ws0.column_dimensions[get_column_letter(j)].width = 16
    bold = Font(bold=True)

    # --- лист 2: Отчет по докторам ---
    EXTRA = ["Главный вывод по врачу", "План действий по врачу"]
    ws = wb.create_sheet("Отчет по докторам")
    ws.cell(1, 1, f"Отчет по докторам, {month_label}").font = bold
    headers = ["№"] + MAIN_COLUMNS + CALC_COLUMNS + EXTRA
    for j, col in enumerate(headers, start=1):
        c = ws.cell(2, j, col)
        c.font = bold
        c.alignment = Alignment(wrap_text=True, vertical="center")
    first_r = 3
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
    # Итого: суммы по врачам — только D..G (часы, время, сумма, посещения).
    # H (пациенты) и I (первичные) — ЗНАЧЕНИЯ из отчетов 20 и 41, НЕ сумма
    # (один пациент мог быть у нескольких врачей) — выделяем цветом.
    for j in list(range(4, 8)) + [9]:
        L = get_column_letter(j)
        ws.cell(last_r, j, f"=SUM({L}{first_r}:{L}{last_r - 1})")
    yellow = PatternFill("solid", fgColor="FFF2CC")
    c = ws.cell(last_r, 8, int(main_df.iloc[-1]["Кол-во пациентов"]))
    c.fill = yellow
    c.font = bold

    for j in range(1, len(headers) + 1):
        ws.cell(last_r, j).font = bold

        # --- блок под таблицей (по спецификации, позиции от строки Итого) ---
    L_hours, L_sum = f"D{last_r}", f"F{last_r}"
    L_fh = f"J{last_r}"  # стоимость фактического часа в Итого
    r0 = last_r + 2      # первая строка блока (через 1 строку после Итого)
    ptr = pot_total_row  # строка Итого на листе Потенциал (1-based)
    # строка r0: Исп мощности (D), Дост потенциала (F)
    c = ws.cell(r0, 4, f"={L_hours}/'1. Потенциал'!E{ptr}*100")
    c.number_format = "0.00"; c.font = bold
    c = ws.cell(r0, 6, f"={L_sum}/'1. Потенциал'!H{ptr}*100")
    c.number_format = "0.00"; c.font = bold
    # строка r0+1: подписи
    ws.cell(r0 + 1, 4, "Исп мощности").font = bold
    ws.cell(r0 + 1, 6, "Дост потенциала").font = bold
    # строка r0+2: пусто
    # строка r0+3: Альтернативный потенциал (D — подпись, F — формула)
    ws.cell(r0 + 3, 4, "Альтернативный потенциал").font = bold
    c = ws.cell(r0 + 3, 6, f"={L_sum}/((D{r0}/100)*({L_fh}/{th}))")
    c.number_format = "0.00"; c.font = bold
    # строка r0+4: пусто
    # строка r0+5: Разница с классическим (D — подпись, F — формула)
    ws.cell(r0 + 5, 4, "Разница с классическим").font = bold
    c = ws.cell(r0 + 5, 6, f"=F{r0 + 3}-'1. Потенциал'!H{ptr}")
    c.number_format = "0.00"; c.font = bold
    # строка r0+6: пусто
    # строка r0+7: Вывод
    ws.cell(r0 + 7, 4, "Вывод").font = bold
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
          "коэффициент использования мощности": "main_isp",
          "% достижения потенциала": "main_dost",
          "стоимость помощи": "main_total_sum",
          "рабочие часы врачей": "main_total_hours_plan",
          "часы с пациентом": "main_total_hours_patient",
          "фл": "main_total_patients",
          "посещений": "main_total_visits",
          "первичных": "main_total_first"}


def _sheet1_extend(ws, ref, month_label, metrics):
    """Лист ДИНАМИКА: столбец месяца (новый или перезапись существующего).
    ВСЕ ячейки пишутся ВЫЧИСЛЕННЫМИ ЗНАЧЕНИЯМИ, без формул: openpyxl формулы
    не считает и кэшированных значений в файле нет — без пересчёта Excel такие
    ячейки выглядят пустыми. 'Изм в %' тоже считаем в Python.
    Заголовки ВСЕХ месяцев приводим к виду 'MM.YY' (08.25, 09.26): в исходнике
    они смешанные — Excel-сериалы, даты с часами, строки '12.25'/'09.2026'.
    Месяц ищется по дате (сериал/дата/'MM.YY'/'MM.YYYY') — если колонка уже
    есть (например, с некорректными данными), она ПЕРЕЗАПИСЫВАется на месте."""
    hdr = next(i for i in range(10) if ref.iloc[i].notna().sum() >= 3)
    hdr_row = hdr + 1                    # строка заголовка на листе (1-based)
    header = ref.iloc[hdr].tolist()
    param_col = next(i for i, v in enumerate(header)
                     if norm_text(v) == "параметр")
    izm_col = next(i for i, v in enumerate(header) if "изм" in norm_text(v))
    month_cols = [i for i in range(param_col + 1, izm_col)
                  if pd.notna(header[i])]
    # ВАЖНО: без dropna — пустые строки внутри тела (например, перед блоком
    # 'Альтернативный потенциал') сохраняют соответствие позиций листа
    body = ref.iloc[hdr + 1:]
    labels = body[param_col].tolist()
    first_data_row = hdr_row + 1

    def _month_year(v):
        """(месяц, год) из даты/Excel-сериала/'MM.YY'/'MM.YYYY' или None."""
        if isinstance(v, pd.Timestamp):
            return (int(v.month), int(v.year))
        if isinstance(v, (datetime, date)):
            return (int(v.strftime("%m")), int(v.strftime("%Y")))
        if isinstance(v, (int, float)) and not isinstance(v, bool) \
                and 40000 < float(v) < 60000:          # Excel-сериал
            d = date(1899, 12, 30) + timedelta(days=int(v))
            return (d.month, d.year)
        s = str(v).strip()
        if re.match(r"^\d{5}$", s):                    # сериал текстом
            return _month_year(int(s))
        m = re.match(r"^(\d{1,2})\.(\d{4})$", s)
        if m:
            return (int(m.group(1)), int(m.group(2)))
        m = re.match(r"^(\d{1,2})\.(\d{2})$", s)
        if m:
            return (int(m.group(1)), 2000 + int(m.group(2)))
        return None

    def _month_short(v):
        my = _month_year(v)
        return f"{my[0]:02d}.{my[1] % 100:02d}" if my else None

    # нормализуем заголовки ВСЕХ месяцев к 'MM.YY'
    for c in range(param_col + 2, ws.max_column + 1):
        s = _month_short(ws.cell(hdr_row, c).value)
        if s is not None:
            ws.cell(hdr_row, c, s)

    tgt = _month_year(month_label)
    tgt_short = f"{tgt[0]:02d}.{tgt[1] % 100:02d}"
    tgt_txts = {str(month_label).strip(), tgt_short,
                f"{tgt[0]:02d}.{tgt[1]}", f"{tgt[0]}.{tgt[1]}"}
    new_col = None
    dup_cols = []
    for c in range(3, ws.max_column + 1):
        v = ws.cell(hdr_row, c).value
        v_txt = str(v).strip().replace(" ", "")
        if (_month_year(v) == tgt) or (v_txt in tgt_txts):
            if new_col is None:
                new_col = c
            else:
                dup_cols.append(c)
    for c in sorted(dup_cols, reverse=True):
        ws.delete_cols(c)
    inserted = False
    if new_col is None:
        # вставляем новый столбец (после последнего месяца, перед Изм)
        insert_at = 3 + len(month_cols)
        ws.insert_cols(insert_at)
        new_col = insert_at
        inserted = True
    prev_col = new_col - 1
    izm_new = new_col + 1

    ws.cell(hdr_row, new_col, tgt_short).font = Font(bold=True)
    if inserted:
        ws.cell(hdr_row, izm_new, "Изм в %").font = Font(bold=True)

    # --- все производные метрики считаем здесь, в Python (не формулами!) ---
    def _g(k):
        v = metrics.get(k)
        return float(v) if v is not None else None

    def _div(a, b):
        if a is None or b is None or b == 0:
            return None
        return a / b

    def _pct(a, b):
        v = _div(a, b)
        return v * 100 if v is not None else None

    th = CFG["overload_threshold"]
    pot_h = _g("potential_hours_total")
    pot_r = _g("potential_revenue_total")
    s_sum = _g("main_total_sum")
    s_hours = _g("main_total_hours_plan")
    s_hp = _g("main_total_hours_patient")
    s_fl = _g("main_total_patients")
    s_vis = _g("main_total_visits")
    s_first = _g("main_total_first")
    isp_v = _pct(s_hours, pot_h)         # коэффициент использования мощности
    zagr_v = _pct(s_hp, s_hours)         # % загруженности
    alt_v = None
    if s_sum is not None and isp_v and zagr_v:
        denom = (isp_v / 100) * (zagr_v / 100) / (th / 100)
        if denom:
            alt_v = s_sum / denom
    derived_vals = {
        "коэффициент использования мощности": isp_v,
        "% достижения потенциала": _pct(s_sum, pot_r),
        "% загруженности клиники": zagr_v,
        "% загруженности врачей": zagr_v,
        "стоимость помощи на фл": _div(s_sum, s_fl),
        "посещений на фл": _div(s_vis, s_fl),
        "стоимость посещения": _div(s_sum, s_vis),
        "% первичных к фл": _pct(s_first, s_fl),
        "альтернативный потенциал": alt_v,
        "разница с классическим": (alt_v - pot_r)
        if alt_v is not None and pot_r is not None else None,
    }

    # значения предыдущего месяца (для 'Изм в %') берём из ref: pandas
    # читает кэшированные значения ячеек, формулы старых месяцев не мешают
    prev_vals = pd.to_numeric(body[prev_col - 1], errors="coerce")

    for r_off, label in enumerate(labels):
        r = first_data_row + r_off
        k = norm_text(label).replace("потеницал", "потенциал")
        val = None
        if k in MANUAL and metrics.get(MANUAL[k]) is not None:
            val = float(metrics[MANUAL[k]])
        elif k in derived_vals:
            val = derived_vals[k]
        if val is not None:
            # единый вид с историческими колонками (General: без разделителей
            # тысяч и принудительных нулей — как до сентября);
            # строки 18-19 (Альтернативный потенциал, Разница) — до целого
            if k in ("альтернативный потенциал", "разница с классическим"):
                rv = round(val)
            else:
                rv = round(val, 6)
            c = ws.cell(r, new_col, rv)
            c.number_format = "General"
        else:
            # строка не рассчитывается (пустая/служебная): стираем старое
            # значение перезаписываемого месяца, чтобы не осталось мусора
            ws.cell(r, new_col).value = None
        pv = prev_vals.iloc[r_off] if r_off < len(prev_vals) else None
        if (val is not None and pv is not None
                and pd.notna(pv) and float(pv) != 0):
            ws.cell(r, izm_new,
                    round(val / float(pv) - 1, 4)).number_format = "0.0%"
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




def _sheet1_style(ws):
    """Единый вид листа «Динамика»: тонкие границы по всей таблице
    и ширина колонок по длине ОТОБРАЖАЕМОГО текста (с учётом форматов:
    проценты считаются как '-3,2%', а не как '-0,0323')."""
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row,
                            min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = border

    def _disp(cell):
        v = cell.value
        if v is None:
            return ""
        if cell.number_format == "0.0%" and isinstance(v, (int, float)):
            return f"{v * 100:.1f}%"
        if isinstance(v, float):
            return f"{v:.6f}".rstrip("0").rstrip(".")
        if isinstance(v, (datetime, date)):
            return v.strftime("%d.%m.%Y")
        return str(v)

    for col in range(1, ws.max_column + 1):
        longest = max((len(_disp(ws.cell(r, col)))
                       for r in range(1, ws.max_row + 1)), default=0)
        ws.column_dimensions[get_column_letter(col)].width = min(
            max(longest + 2, 6), 45)


def update_dinamika_file(reference_bytes, month_label, metrics, main_df):
    """Загруженный файл «Динамика» -> только первый лист (ДИНАМИКА)
    с новым месяцем. Лист называется «динамика [месяц]».
    Возвращает bytes обновленного xlsx."""
    sheets = pd.read_excel(BytesIO(reference_bytes), sheet_name=None,
                           header=None)
    name0 = list(sheets.keys())[0]
    ref = sheets[name0]
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(title=f"динамика {month_label}")
    for i, row in ref.iterrows():
        for j, v in enumerate(row.tolist()):
            if pd.notna(v):
                ws.cell(i + 1, j + 1, v)
    _sheet1_extend(ws, ref, month_label, metrics)
    _sheet1_style(ws)                      # границы + ширина колонок по тексту
    wb.calculation.fullCalcOnLoad = True   # Excel пересчитает формулы при открытии
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
        pot_raw, pot_hours, pot_revenue, pot_total_row = None, None, None, None
        if f_pot is not None:
            pot_raw = read_any(f_pot)
            pot_hours, pot_revenue, _, pot_total_row = parse_potential(pot_raw)
        else:
            st.warning("⚠️ Таблица потенциала не загружена — лист «1. Потенциал» "
                       "не будет в отчете, а «Динамика» не сможет заполнить строки "
                       "«Мощность в часах» и «Потенциал по выручке». "
                       "Загрузите файл с эталоном потенциала.")
        zag_raw = read_any(f_zag)
        month_label = detect_month_label(zag_raw, default="")
        zag, norm_rate = parse_zagruzka(zag_raw)
        doctor_keys = set(zag["key"])
        obs_by_spec = parse_obschaya_summa(read_any(f_sum))
        unic_by_doc = parse_unic_patients(read_any(f_unic))
        perv_by_doc, perv_clinic_total = parse_pervoe_obr(read_any(f_perv))
        svod_by_doc, svod_clinic_total = parse_svodnyj_patients(
            read_any(f_svod), doctor_keys)
        main_df = build_main_table(zag, obs_by_spec, unic_by_doc,
                                   perv_by_doc, svod_by_doc, svod_clinic_total,
                                   perv_clinic_total)

    month_name = st.text_input("Месяц нового столбца «Динамики» "
                               "(подставлен из отчетов)",
                               value=month_label or "")

    st.subheader("Отчет по докторам")
    st.dataframe(main_df, use_container_width=True, hide_index=True)
    main_bytes = write_main_report(month_name, main_df, pot_raw,
                                   pot_hours or 0, pot_revenue or 0,
                                   pot_total_row)
    st.download_button("Скачать отчет (xlsx, 2 листа)", main_bytes,
                       "Отчет_по_докторам.xlsx")

    metrics = totals_for_dinamika(main_df, pot_hours, pot_revenue)
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
        import time as _time
        fname = f"Динамика_{month_name.replace('.', '_')}_{int(_time.time())}.xlsx"
        st.download_button("Скачать обновленную «Динамику» (xlsx)",
                           dyn_bytes, fname)
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
