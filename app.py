
# ================== УТИЛИТЫ И КОНФИГ ==================
import re
from io import BytesIO
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st
import yaml
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

CFG = yaml.safe_load("""
columns:
  doctor:       ["доктор", "врач", "фио врача"]
  spec:         ["специализация доктора", "специализация", "специальность"]
  hours_plan:   ["часов по табелю", "рабочих часов по графику"]
  hours_patient: ["часов по дошедшим", "время с пациентом"]
  visits:       ["количество услуг", "кол-во посещений", "посещений"]
  sum_rub:      ["сумма фактич", "общая сумма", "сумма"]
  patient:      ["пациентов", "кол-во пациентов", "фио пациента"]
  card:         ["комп. номер", "номер карты"]
  to_pay:       ["выставление в оплату"]
  service:      ["название услуги", "услуга"]
  first_visit:  ["первая услуга", "первое обращение", "первичн"]
service_keep: ["допплер", "узи", "прием", "приём", "эхокг", "эхо кг", "эхо-кг"]
to_pay_keep_value: "есть"
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
    return t.dropna(how="all")


def find_col(cols, key, required=True):
    for col in cols:
        if any(s in norm_text(col) for s in CFG["columns"][key]):
            return col
    if required:
        raise ValueError(f"Колонка '{key}' не найдена. Есть: {list(cols)}")
    return None

# ================== ПАРСЕРЫ ОТЧЕТОВ ==================

def read_any(uploaded):
    name = uploaded.name.lower()
    engine = "xlrd" if name.endswith(".xls") else "openpyxl"
    return pd.read_excel(BytesIO(uploaded.read()), sheet_name=0,
                         header=None, engine=engine)


def parse_zagruzka(df_raw):
    hr = find_header_row(df_raw, must_have=("доктор", "специализация"))
    t = build_table(df_raw, hr)
    c_doc = find_col(t.columns, "doctor")
    c_spec = find_col(t.columns, "spec")

    def is_doctor_row(r):
        doc = str(r[c_doc]).strip()
        spec = str(r[c_spec]).strip()
        if spec.lower() not in ("", "nan", "none"):
            return False
        if re.match(r"^\d{2}\.\d{2}\.\d{4}", doc):
            return False
        return not _is_total(doc)

    t = t[t.apply(is_doctor_row, axis=1)]

    def num(key):
        col = find_col(t.columns, key, required=False)
        if col is None:
            return 0.0
        return pd.to_numeric(t[col], errors="coerce").fillna(0)

    return pd.DataFrame({
        "key": t[c_doc].map(doctor_key),
        "ФИО врача": t[c_doc].astype(str).str.strip(),
        "Специализация": "",
        "Рабочих часов по графику": num("hours_plan"),
        "Время с пациентом": num("hours_patient"),
        "Кол-во посещений": num("visits"),
    })


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


def parse_pervoe_obr(df_raw):
    hr = find_header_row(df_raw, must_have=("доктор", "номер карты"))
    t = build_table(df_raw, hr)
    c_doc = find_col(t.columns, "doctor")
    t = t[t[c_doc].notna()]
    t = t[~t[c_doc].apply(_is_total)]
    return t[c_doc].map(doctor_key).value_counts()


def parse_svodnyj_patients(df_raw, doctor_keys):
    hr = find_header_row(df_raw, must_have=("врач", "пациент"))
    t = build_table(df_raw, hr)
    c_doc = find_col(t.columns, "doctor")
    c_pat = find_col(t.columns, "patient")
    c_pay = find_col(t.columns, "to_pay")
    c_srv = find_col(t.columns, "service", required=False)
    c_id = find_col(t.columns, "card", required=False) or c_pat
    t = t[t[c_doc].notna() & t[c_pat].notna()]
    t = t[t[c_doc].map(doctor_key).isin(doctor_keys)]
    t = t[t[c_pay].map(norm_text) == norm_text(CFG["to_pay_keep_value"])]
    if c_srv is not None:
        keep = [norm_text(s) for s in CFG["service_keep"]]
        t = t[t[c_srv].map(lambda v: any(k in norm_text(v) for k in keep))]
    t = t.drop_duplicates(subset=[c_doc, c_id])
    return t[c_doc].map(doctor_key).value_counts(), int(t[c_id].nunique())

# ================== ТАБЛИЦА 1 ==================

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
    df["Кол-во пациентов"] = df["key"].map(svod_by_doc).fillna(
        df["key"].map(unic_by_doc).fillna(0)).astype(int)
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

# ================== ТАБЛИЦА 2: ДИНАМИКА ==================

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


def build_dinamika(reference_bytes, new_col_title, metrics,
                   to_pay_label="Изм мес. к мес. в %"):
    ref = pd.read_excel(BytesIO(reference_bytes), sheet_name=0, header=None)
    hdr = next(i for i in range(10) if ref.iloc[i].notna().sum() >= 3)
    header = ref.iloc[hdr].tolist()
    param_col = next(i for i, v in enumerate(header)
                     if norm_text(v) == "параметр")
    izm_col = next(i for i, v in enumerate(header) if "изм" in norm_text(v))
    month_cols = [i for i in range(param_col + 1, izm_col)
                  if pd.notna(header[i])]
    body = ref.iloc[hdr + 1:].dropna(how="all")
    labels = body[param_col].tolist()

    wb = Workbook()
    ws = wb.active
    ws.title = "Динамика"
    ws.cell(1, 1, "№")
    ws.cell(1, 2, "Параметр")
    for j, c in enumerate(month_cols):
        ws.cell(1, 3 + j, _fmt_month(header[c]))
    new_col = 3 + len(month_cols)
    ws.cell(1, new_col, new_col_title)
    ws.cell(1, new_col + 1, to_pay_label)

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
        ws.cell(r, 1, r_off + 1)
        ws.cell(r, 2, label)
        for j, c in enumerate(month_cols):
            v = body.iloc[r_off, c]
            if pd.notna(v):
                ws.cell(r, 3 + j, v)
        k = norm_text(label).replace("потеницал", "потенциал")
        L = get_column_letter(new_col)
        if k in MANUAL and MANUAL[k] in metrics:
            ws.cell(r, new_col, metrics[MANUAL[k]])
        elif k in DERIVED:
            ws.cell(r, new_col, DERIVED[k].format(
                c=L, r_power=find_row("мощность в часах"),
                r_pot=find_row("потенциал по выручке"),
                r_sum=find_row("стоимость помощи"),
                r_hours=find_row("рабочие часы врачей"),
                r_hp=find_row("часы с пациентом"),
                r_fl=find_row("фл"), r_vis=find_row("посещений")))
        Lp = get_column_letter(new_col - 1)
        ws.cell(r, new_col + 1,
                f'=IFERROR({L}{r}/{Lp}{r}-1,"")').number_format = "0.0%"

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ================== ИНТЕРФЕЙС ==================

def to_excel_bytes(df, sheet_name="Отчет"):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=sheet_name)
    buf.seek(0)
    return buf.getvalue()


st.set_page_config(page_title="КВС: сборка отчетов", layout="wide")
st.title("КВС: сборка месячных отчетов")
st.caption("Загрузите 5 отчетов (названия файлов не важны — важна структура) "
           "+ опционально референс листа «Динамика».")

with st.sidebar:
    st.header("Параметры месяца")
    month_name = st.text_input("Название нового столбца (месяц)",
                               value="Сентябрь 2026")
    potential_hours = st.number_input("Мощность в часах (Итого по потенциалу)",
                                      min_value=0.0, value=0.0)
    potential_revenue = st.number_input("Потенциал по выручке (Итого)",
                                        min_value=0.0, value=0.0)
    st.divider()
    st.header("Файлы")
    f_zag = st.file_uploader("1. Загрузка врачей (без кабинетов)",
                             type=["xls", "xlsx"])
    f_sum = st.file_uploader("2. Общая сумма", type=["xls", "xlsx"])
    f_unic = st.file_uploader("3. Уникальных пациентов за период",
                              type=["xls", "xlsx"])
    f_perv = st.file_uploader("4. Первое обращение пациента",
                              type=["xls", "xlsx"])
    f_svod = st.file_uploader("5. Сводный по врачам с услугами и пациентами",
                              type=["xls", "xlsx"])
    f_dyn = st.file_uploader("Референс «Динамика» (опционально)",
                             type=["xls", "xlsx"])

files = [f_zag, f_sum, f_unic, f_perv, f_svod]
if not all(files):
    st.info("Загрузите все 5 отчетов — имена файлов могут быть любыми.")
    st.stop()

try:
    with st.spinner("Разбираю отчеты..."):
        zag = parse_zagruzka(read_any(f_zag))
        doctor_keys = set(zag["key"])
        sum_by_doc, spec_by_doc = parse_obschaya_summa(read_any(f_sum))
        unic_by_doc = parse_unic_patients(read_any(f_unic))
        perv_by_doc = parse_pervoe_obr(read_any(f_perv))
        svod_by_doc, svod_clinic_total = parse_svodnyj_patients(
            read_any(f_svod), doctor_keys)
        main_df = build_main_table(zag, sum_by_doc, spec_by_doc, unic_by_doc,
                                   perv_by_doc, svod_by_doc, svod_clinic_total)

    st.subheader("Таблица 1 — Отчет по докторам")
    st.dataframe(main_df, use_container_width=True, hide_index=True)
    st.download_button("Скачать таблицу 1 (xlsx)",
                       to_excel_bytes(main_df, "Отчет по докторам"),
                       "KVS_main.xlsx")

    metrics = totals_for_dinamika(main_df)
    metrics["potential_hours_total"] = potential_hours
    metrics["potential_revenue_total"] = potential_revenue

    st.subheader("Таблица 2 — лист «Динамика»")
    if f_dyn:
        dyn_bytes = build_dinamika(BytesIO(f_dyn.read()).getvalue(),
                                   month_name, metrics)
        st.download_button("Скачать таблицу 2 «Динамика» (xlsx)", dyn_bytes,
                           "KVS_dinamika.xlsx")
    else:
        st.warning("Референс «Динамики» не загружен.")
        simple = {k: metrics[v] for k, v in {
            "Мощность в часах": "potential_hours_total",
            "Потенциал по выручке": "potential_revenue_total",
            "Стоимость помощи": "main_total_sum",
            "Рабочие часы врачей": "main_total_hours_plan",
            "Часы с пациентом": "main_total_hours_patient",
            "ФЛ": "main_total_patients",
            "Посещений": "main_total_visits",
            "Первичных": "main_total_first",
        }.items() if v in metrics}
        dyn_df = pd.DataFrame({"Показатель": list(simple.keys()),
                               month_name: list(simple.values())})
        st.dataframe(dyn_df, hide_index=True)
        st.download_button("Скачать таблицу 2 (xlsx)",
                           to_excel_bytes(dyn_df, "Динамика"),
                           "KVS_dinamika.xlsx")

    with st.expander("Диагностика: как сматчились врачи"):
        c1, c2 = st.columns(2)
        c1.metric("Врачей в основном отчете", len(doctor_keys))
        c2.metric("Уникальных пациентов по клинике (из сводного)",
                  svod_clinic_total)
        missed = doctor_keys - set(svod_by_doc.index)
        if missed:
            st.write("Нет в сводном отчете:", sorted(missed))
except ValueError as e:
    st.error(str(e))
    st.stop()
