import io
from datetime import date

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Control de Importaciones y Entregas", page_icon="🌷", layout="wide")

REQUIRED_COLUMNS = ["CUSTOMER", "VARIETY", "PLANTS", "IMPORT WEEK", "DELIVERY WEEK"]

TIPO_DIRECTO = "IMPORTA Y ENTREGA (directo)"
TIPO_ENRAIZA = "IMPORTA Y ENRAIZA"
TIPO_REVISAR = "REVISAR"


def normalize_columns(df):
    df = df.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]
    return df


def parse_week(raw):
    """Devuelve (year, week) o (None, motivo_del_error)."""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return None, "vacío"
    s = str(raw).strip()
    if not s or s.lower() in ("nan", "none"):
        return None, "vacío"
    s = s.replace("\\", "/").replace(" ", "")
    parts = s.split("/")
    if len(parts) != 2:
        return None, f"formato no reconocido ('{raw}')"

    week_str, year_str = parts
    try:
        week = int(week_str)
    except ValueError:
        return None, f"semana no numérica ('{raw}')"

    if not year_str.isdigit():
        return None, f"año no numérico ('{raw}')"

    if len(year_str) == 4:
        year = int(year_str)
    elif len(year_str) == 3:
        year = int("2" + year_str)
    elif len(year_str) == 2:
        year = int("20" + year_str)
    else:
        year = int("200" + year_str)

    if not (1 <= week <= 53):
        return None, f"semana fuera de rango ('{raw}')"
    if not (2000 <= year <= 2100):
        return None, f"año fuera de rango ('{raw}')"

    return (year, week), None


def week_to_date(year, week):
    try:
        return date.fromisocalendar(year, week, 1)
    except ValueError:
        return date.fromisocalendar(year, 52, 1)


def format_week(parsed):
    if parsed is None:
        return None
    year, week = parsed
    return f"{week:02d}/{year}"


def classify_row(import_parsed, delivery_parsed, import_error, delivery_error):
    if import_parsed is None or delivery_parsed is None:
        motivos = [m for m in (import_error, delivery_error) if m]
        return TIPO_REVISAR, "; ".join(motivos), None

    import_date = week_to_date(*import_parsed)
    delivery_date = week_to_date(*delivery_parsed)

    if import_parsed == delivery_parsed:
        return TIPO_DIRECTO, "", 0
    if import_date < delivery_date:
        weeks = (delivery_date - import_date).days // 7
        return TIPO_ENRAIZA, "", weeks
    return TIPO_REVISAR, "la entrega aparece antes que la importación", None


@st.cache_data
def process_file(file_bytes):
    df = pd.read_excel(io.BytesIO(file_bytes))
    df = normalize_columns(df)

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        return None, missing

    import_parsed, import_errors = [], []
    delivery_parsed, delivery_errors = [], []
    for raw in df["IMPORT WEEK"]:
        p, e = parse_week(raw)
        import_parsed.append(p)
        import_errors.append(e)
    for raw in df["DELIVERY WEEK"]:
        p, e = parse_week(raw)
        delivery_parsed.append(p)
        delivery_errors.append(e)

    df["IMPORT WEEK (norm)"] = [format_week(p) for p in import_parsed]
    df["DELIVERY WEEK (norm)"] = [format_week(p) for p in delivery_parsed]

    tipos, motivos, semanas_enraizando = [], [], []
    for ip, dp, ie, de in zip(import_parsed, delivery_parsed, import_errors, delivery_errors):
        tipo, motivo, weeks = classify_row(ip, dp, ie, de)
        tipos.append(tipo)
        motivos.append(motivo)
        semanas_enraizando.append(weeks)

    df["TIPO"] = tipos
    df["MOTIVO REVISION"] = motivos
    df["SEMANAS ENRAIZANDO"] = semanas_enraizando
    df["_import_parsed"] = import_parsed
    df["_delivery_parsed"] = delivery_parsed

    if "CUTTINGS FROM" not in df.columns:
        df["CUTTINGS FROM"] = ""

    return df, []


def to_excel_bytes(df):
    output = io.BytesIO()
    export_df = df.drop(columns=["_import_parsed", "_delivery_parsed"], errors="ignore")
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Resultado")
    return output.getvalue()


st.title("🌷 Control de Importaciones y Entregas")

uploaded = st.file_uploader("Sube tu archivo Excel", type=["xlsx", "xls"])

if uploaded is None:
    st.info("Sube tu archivo Excel de ventas (mismo formato de VENTAS_2026.xlsx) para comenzar.")
    st.stop()

df, missing = process_file(uploaded.getvalue())

if df is None:
    st.error(
        "Al archivo le faltan columnas obligatorias: "
        + ", ".join(missing)
        + ". Revisa los encabezados de tu Excel."
    )
    st.stop()

st.sidebar.header("Filtros")

customers = sorted(df["CUSTOMER"].dropna().astype(str).unique())
selected_customers = st.sidebar.multiselect("Cliente", customers)

variety_search = st.sidebar.text_input("Buscar variedad")

origins = sorted(df["CUTTINGS FROM"].dropna().astype(str).unique())
origins = [o for o in origins if o.strip()]
selected_origins = st.sidebar.multiselect("Origen (CUTTINGS FROM)", origins)

tipos_disponibles = [TIPO_DIRECTO, TIPO_ENRAIZA, TIPO_REVISAR]
selected_tipos = st.sidebar.multiselect("Tipo de proceso", tipos_disponibles)

all_weeks = sorted(
    set(df["IMPORT WEEK (norm)"].dropna()) | set(df["DELIVERY WEEK (norm)"].dropna()),
    key=lambda w: (int(w.split("/")[1]), int(w.split("/")[0])),
)
week_options = ["Todas"] + all_weeks
selected_week = st.sidebar.selectbox("Semana", week_options)

filtered = df.copy()
if selected_customers:
    filtered = filtered[filtered["CUSTOMER"].astype(str).isin(selected_customers)]
if variety_search:
    filtered = filtered[filtered["VARIETY"].astype(str).str.contains(variety_search, case=False, na=False)]
if selected_origins:
    filtered = filtered[filtered["CUTTINGS FROM"].astype(str).isin(selected_origins)]
if selected_tipos:
    filtered = filtered[filtered["TIPO"].isin(selected_tipos)]

DISPLAY_ORDER = [
    "CUSTOMER",
    "VARIETY",
    "PLANTS",
    "CUTTINGS FROM",
    "IMPORT WEEK (norm)",
    "DELIVERY WEEK (norm)",
    "TIPO",
    "SEMANAS ENRAIZANDO",
    "MOTIVO REVISION",
]
DISPLAY_RENAME = {
    "CUSTOMER": "CLIENTE",
    "VARIETY": "PRODUCTO",
    "PLANTS": "CANTIDAD",
    "CUTTINGS FROM": "ORIGEN",
    "IMPORT WEEK (norm)": "SEMANA IMPORTACION",
    "DELIVERY WEEK (norm)": "SEMANA ENTREGA",
    "TIPO": "TIPO DE PROCESO",
    "SEMANAS ENRAIZANDO": "SEM. ENRAIZANDO",
}


def display_table(d):
    cols = [c for c in DISPLAY_ORDER if c in d.columns]
    return d[cols].rename(columns=DISPLAY_RENAME)


def week_breakdown(d, week_str):
    week_num, week_year = week_str.split("/")
    week_key = (int(week_year), int(week_num))
    week_date = week_to_date(*week_key)

    importa = d[d["_import_parsed"] == week_key]
    entrega = d[d["_delivery_parsed"] == week_key]
    enraizando = d[
        (d["TIPO"] == TIPO_ENRAIZA)
        & d["_import_parsed"].apply(lambda p: p is not None and week_to_date(*p) <= week_date)
        & d["_delivery_parsed"].apply(lambda p: p is not None and week_date < week_to_date(*p))
    ]
    return importa, entrega, enraizando


kpi1, kpi2, kpi3 = st.columns(3)
kpi1.metric("Cantidad total (filtrado)", int(filtered["PLANTS"].sum()))
kpi2.metric("Líneas", len(filtered))
kpi3.metric("Clientes", filtered["CUSTOMER"].nunique())

weeks_to_show = all_weeks if selected_week == "Todas" else [selected_week]

summary_rows = []
breakdown_by_week = {}
for w in weeks_to_show:
    importa, entrega, enraizando = week_breakdown(filtered, w)
    breakdown_by_week[w] = (importa, entrega, enraizando)
    summary_rows.append(
        {
            "SEMANA": w,
            "CANTIDAD IMPORTADA": int(importa["PLANTS"].sum()),
            "CANTIDAD ENTREGADA": int(entrega["PLANTS"].sum()),
            "LOTES ENRAIZANDO": len(enraizando),
        }
    )

st.subheader("📊 Resumen por semana")
st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

st.subheader("📅 Detalle por semana (origen, producto y cantidad)")
download_parts = []
for w in weeks_to_show:
    importa, entrega, enraizando = breakdown_by_week[w]
    download_parts.extend([importa, entrega, enraizando])
    with st.expander(f"Semana {w}", expanded=(len(weeks_to_show) == 1)):
        tab_import, tab_entrega, tab_enraiza = st.tabs(
            [
                f"📦 Importa ({len(importa)})",
                f"🚚 Entrega ({len(entrega)})",
                f"🌱 Enraizando ({len(enraizando)})",
            ]
        )
        with tab_import:
            st.dataframe(display_table(importa), use_container_width=True, hide_index=True)
        with tab_entrega:
            st.dataframe(display_table(entrega), use_container_width=True, hide_index=True)
        with tab_enraiza:
            st.dataframe(display_table(enraizando), use_container_width=True, hide_index=True)

table_for_download = pd.concat(download_parts).drop_duplicates() if download_parts else filtered.iloc[0:0]

revisar = filtered[filtered["TIPO"] == TIPO_REVISAR]
if not revisar.empty:
    st.subheader("⚠️ Filas para revisar en el Excel original")
    st.dataframe(display_table(revisar), use_container_width=True, hide_index=True)
    table_for_download = pd.concat([table_for_download, revisar]).drop_duplicates()

st.download_button(
    "⬇️ Descargar resultado filtrado (Excel)",
    data=to_excel_bytes(table_for_download),
    file_name="control_importaciones_entregas.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
