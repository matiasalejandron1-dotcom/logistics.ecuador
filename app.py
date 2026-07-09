import io
from datetime import date, timedelta

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


def week_key_from_date(d):
    year, week, _ = d.isocalendar()
    return (year, week)


def shift_week(week_key, offset):
    return week_key_from_date(week_to_date(*week_key) + timedelta(weeks=offset))


def format_week_key(week_key):
    year, week = week_key
    return f"{week:02d}/{year}"


def week_breakdown(d, week_key):
    week_date = week_to_date(*week_key)
    importa = d[d["_import_parsed"] == week_key]
    entrega_directa = d[(d["_delivery_parsed"] == week_key) & (d["TIPO"] == TIPO_DIRECTO)]
    entrega_enraizado = d[(d["_delivery_parsed"] == week_key) & (d["TIPO"] == TIPO_ENRAIZA)]
    enraizando = d[
        (d["TIPO"] == TIPO_ENRAIZA)
        & d["_import_parsed"].apply(lambda p: p is not None and week_to_date(*p) <= week_date)
        & d["_delivery_parsed"].apply(lambda p: p is not None and week_date < week_to_date(*p))
    ]
    return importa, entrega_directa, entrega_enraizado, enraizando


def render_by_origin(d):
    origins_here = sorted({str(o) for o in d["CUTTINGS FROM"].dropna() if str(o).strip()})
    if not origins_here:
        st.info("No hay filas para mostrar.")
        return
    origin_tabs = st.tabs([f"🌍 {o}" for o in origins_here])
    for tab, origin in zip(origin_tabs, origins_here):
        with tab:
            subset = d[d["CUTTINGS FROM"].astype(str) == origin]
            st.dataframe(display_table(subset), use_container_width=True, hide_index=True)


def render_week_view(d, week_key):
    st.caption(f"Semana {format_week_key(week_key)}")
    importa, entrega_directa, entrega_enraizado, enraizando = week_breakdown(d, week_key)

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("📦 Importa", int(importa["PLANTS"].sum()))
    k2.metric("🌱 Enraizando", len(enraizando))
    k3.metric("🚚 Entrega directa", int(entrega_directa["PLANTS"].sum()))
    k4.metric("🌷 Entrega de enraizado", int(entrega_enraizado["PLANTS"].sum()))

    combined = pd.concat([importa, entrega_directa, entrega_enraizado, enraizando])
    origins_here = sorted({str(o) for o in combined["CUTTINGS FROM"].dropna() if str(o).strip()})
    if not origins_here:
        st.info("No hay datos para esta semana.")
        return

    origin_tabs = st.tabs([f"🌍 {o}" for o in origins_here])
    for tab, origin in zip(origin_tabs, origins_here):
        with tab:
            o_importa = importa[importa["CUTTINGS FROM"].astype(str) == origin]
            o_enraizando = enraizando[enraizando["CUTTINGS FROM"].astype(str) == origin]
            o_entrega_directa = entrega_directa[entrega_directa["CUTTINGS FROM"].astype(str) == origin]
            o_entrega_enraizado = entrega_enraizado[entrega_enraizado["CUTTINGS FROM"].astype(str) == origin]

            sub_tabs = st.tabs(
                [
                    f"📦 Importa ({len(o_importa)})",
                    f"🌱 Enraizando ({len(o_enraizando)})",
                    f"🚚 Entrega directa ({len(o_entrega_directa)})",
                    f"🌷 Entrega de enraizado ({len(o_entrega_enraizado)})",
                ]
            )
            with sub_tabs[0]:
                st.dataframe(display_table(o_importa), use_container_width=True, hide_index=True)
            with sub_tabs[1]:
                st.dataframe(display_table(o_enraizando), use_container_width=True, hide_index=True)
            with sub_tabs[2]:
                st.dataframe(display_table(o_entrega_directa), use_container_width=True, hide_index=True)
            with sub_tabs[3]:
                st.dataframe(display_table(o_entrega_enraizado), use_container_width=True, hide_index=True)


kpi1, kpi2, kpi3 = st.columns(3)
kpi1.metric("Cantidad total (filtrado)", int(filtered["PLANTS"].sum()))
kpi2.metric("Líneas", len(filtered))
kpi3.metric("Clientes", filtered["CUSTOMER"].nunique())

today = date.today()
current_week_key = week_key_from_date(today)
past_week_key = shift_week(current_week_key, -1)
plus_weeks = [shift_week(current_week_key, i) for i in (1, 2, 3)]
horizon_end = plus_weeks[-1]

all_week_keys = sorted(
    {k for k in pd.concat([filtered["_import_parsed"], filtered["_delivery_parsed"]]) if k is not None}
)

tab_overview, tab_history, tab_past, tab_current, tab_p1, tab_p2, tab_p3, tab_future = st.tabs(
    [
        "📊 Overview",
        "🕰️ History",
        "⏮️ Past Week",
        "📍 Current Week",
        "➡️ Week +1",
        "➡️ Week +2",
        "➡️ Week +3",
        "🔮 Future Shipments",
    ]
)

with tab_overview:
    st.subheader("📊 Resumen por semana")
    summary_rows = []
    for w in all_week_keys:
        importa, entrega_directa, entrega_enraizado, enraizando = week_breakdown(filtered, w)
        summary_rows.append(
            {
                "SEMANA": format_week_key(w),
                "CANTIDAD IMPORTADA": int(importa["PLANTS"].sum()),
                "CANTIDAD ENTREGA DIRECTA": int(entrega_directa["PLANTS"].sum()),
                "CANTIDAD ENTREGA ENRAIZADO": int(entrega_enraizado["PLANTS"].sum()),
                "LOTES ENRAIZANDO": len(enraizando),
            }
        )
    st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

with tab_history:
    history_df = filtered[
        filtered["_delivery_parsed"].apply(lambda p: p is not None and p < past_week_key)
    ]
    st.caption(f"Entregas completadas antes de la semana {format_week_key(past_week_key)}")
    render_by_origin(history_df)

with tab_past:
    render_week_view(filtered, past_week_key)

with tab_current:
    render_week_view(filtered, current_week_key)

with tab_p1:
    render_week_view(filtered, plus_weeks[0])

with tab_p2:
    render_week_view(filtered, plus_weeks[1])

with tab_p3:
    render_week_view(filtered, plus_weeks[2])

with tab_future:
    future_df = filtered[
        filtered["_import_parsed"].apply(lambda p: p is not None and p > horizon_end)
    ]
    st.caption(f"Importaciones programadas después de la semana {format_week_key(horizon_end)}")
    render_by_origin(future_df)

revisar = filtered[filtered["TIPO"] == TIPO_REVISAR]
if not revisar.empty:
    st.subheader("⚠️ Filas para revisar en el Excel original")
    st.dataframe(display_table(revisar), use_container_width=True, hide_index=True)

st.download_button(
    "⬇️ Descargar resultado filtrado (Excel)",
    data=to_excel_bytes(filtered),
    file_name="control_importaciones_entregas.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
