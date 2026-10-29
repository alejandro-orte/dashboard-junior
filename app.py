"""
Dashboard Modernización 2026 - Bogotá
Lee la hoja 'Sitios' del archivo CONTROL excel y construye el embudo de
despliegue por zona (TSS Cargados -> TSS Aprobados -> Obra Civil -> Diseño TX
-> RFIC-OK -> Instalados -> Integrados -> InSrv), replicando y corrigiendo la
lógica de la pestaña 'Resumen' del Excel original.

Correr localmente:
    pip install -r requirements.txt
    streamlit run app.py

Desplegar en Streamlit Community Cloud: ver README.md
"""

import datetime as dt
from io import BytesIO

import openpyxl
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ----------------------------------------------------------------------------
# Configuración de página y estilos
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="Modernización 2026 · Bogotá",
    page_icon="📡",
    layout="wide",
)

# Paleta validada (misma que el dashboard de referencia)
COLOR = {
    "Oriente": "#2a78d6",    # blue
    "Norte": "#eb6834",      # orange
    "Sur": "#1baf7a",        # aqua
    "Occidente": "#eda100",  # yellow
}
GOOD = "#0ca30c"
WARNING = "#fab219"
CRITICAL = "#d03b3b"
MUTED = "#898781"
GRID = "#e1e0d9"
SEQ_RAMP = ["#86b6ef", "#6da7ec", "#5598e7", "#3987e5", "#2a78d6", "#256abf", "#1c5cab", "#184f95"]

ZONA_MAP = {
    "Bogota Oriente": "Oriente",
    "Bogota Norte": "Norte",
    "Bogota Sur": "Sur",
    "Bogota Occidente": "Occidente",
}
STAGES = [
    "Total Sitios", "TSS Cargados", "TSS Aprobados", "Obra Civil Ok",
    "Diseño TX", "RFIC-OK", "Instalados", "Integrados", "InSrv",
]
# columna (índice 0-based) en la hoja 'Sitios' para cada etapa, salvo Total
STAGE_COL = {
    "TSS Cargados": 5, "TSS Aprobados": 10, "Obra Civil Ok": 12,
    "Diseño TX": 14, "RFIC-OK": 13, "Instalados": 15,
    "Integrados": 16, "InSrv": 17,
}

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.5rem;}
    [data-testid="stMetricValue"] {font-size: 1.9rem; font-weight: 700;}
    .bug-banner {
        background:#fff8ec; border:1px solid #f3dca8; border-radius:10px;
        padding:12px 16px; font-size:0.85rem; color:#52514e; margin-bottom:1rem;
        line-height:1.5;
    }
    .bug-banner b {color:#0b0b0b;}
    table.detail {width:100%; border-collapse:collapse; font-size:0.92rem;}
    table.detail th {
        text-align:left; color:#898781; font-weight:600; font-size:0.72rem;
        text-transform:uppercase; letter-spacing:0.02em; padding:8px 10px;
        border-bottom:1px solid #e1e0d9;
    }
    table.detail td {padding:9px 10px; border-bottom:1px solid #e1e0d9; color:#0b0b0b;}
    table.detail tr:last-child td {border-bottom:none;}
    table.detail tr:hover td {background:#f9f9f7;}
    .tag {
        display:inline-block; padding:2px 10px; border-radius:999px;
        font-size:0.72rem; font-weight:700; letter-spacing:0.01em;
    }
    .tag-negado {background:#d03b3b; color:#fff;}
    .tag-cosite {background:#fab219; color:#0b0b0b;}
    .detail-card {
        background:#fcfcfb; border:1px solid rgba(11,11,11,0.08); border-radius:10px;
        padding:4px 14px; max-height:360px; overflow-y:auto;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ----------------------------------------------------------------------------
# Carga y procesamiento de datos
# ----------------------------------------------------------------------------
def has_date(v) -> bool:
    """True si la celda tiene una fecha real (no vacía/serial 0)."""
    return isinstance(v, (dt.datetime, dt.date)) and not isinstance(v, dt.time)


@st.cache_data(show_spinner="Leyendo Excel y recalculando el embudo...")
def load_sitios(file_bytes: bytes) -> dict:
    wb = openpyxl.load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
    if "Sitios" not in wb.sheetnames:
        raise ValueError("El archivo no tiene una hoja llamada 'Sitios'.")
    ws = wb["Sitios"]
    rows = list(ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True))

    zonas = list(ZONA_MAP.keys())
    matrix = {}
    negados_detalle, cosite_detalle = [], []
    sin_zona = 0

    # proyectos disponibles en el archivo (columna E), para selector
    proyectos = sorted({r[4] for r in rows if r[4]})

    return {"rows": rows, "proyectos": proyectos}


def compute_metrics(rows, proyecto: str) -> dict:
    zonas = list(ZONA_MAP.keys())
    matrix = {ZONA_MAP[z]: {} for z in zonas}
    negados_detalle, cosite_detalle = [], []
    sin_zona = 0

    filtered = [r for r in rows if r[4] == proyecto]
    for z in zonas:
        zshort = ZONA_MAP[z]
        rz = [r for r in filtered if r[1] == z]
        matrix[zshort]["Total Sitios"] = len(rz)
        for stage, col in STAGE_COL.items():
            matrix[zshort][stage] = sum(1 for r in rz if has_date(r[col]))
        for r in rz:
            carga = str(r[9]).strip().lower() if r[9] else ""
            if carga == "negado":
                negados_detalle.append({"sitio": r[0], "zona": zshort})
            elif carga == "cosite":
                cosite_detalle.append({"sitio": r[0], "zona": zshort})

    sin_zona = sum(1 for r in filtered if r[1] not in zonas)

    return {
        "matrix": matrix,
        "negados_detalle": negados_detalle,
        "cosite_detalle": cosite_detalle,
        "sin_zona": sin_zona,
        "total_filtrado": len(filtered),
    }


# ----------------------------------------------------------------------------
# Gráficas (Plotly)
# ----------------------------------------------------------------------------
def funnel_chart(matrix: dict) -> go.Figure:
    stages = [s for s in STAGES if s != "Total Sitios"]
    totals = [sum(matrix[z][s] for z in matrix) for s in stages]

    fig = go.Figure(
        go.Bar(
            x=totals,
            y=stages,
            orientation="h",
            marker=dict(color=SEQ_RAMP[: len(stages)]),
            text=[str(v) for v in totals],
            textposition="outside",
            hovertemplate="<b>%{y}</b><br>%{x} sitios<extra></extra>",
        )
    )
    fig.update_layout(
        height=380,
        margin=dict(l=10, r=40, t=10, b=10),
        plot_bgcolor="white",
        paper_bgcolor="white",
        yaxis=dict(autorange="reversed", showgrid=False),
        xaxis=dict(showgrid=True, gridcolor=GRID, zeroline=False, title="Sitios"),
        font=dict(family="system-ui, sans-serif", color="#0b0b0b"),
        showlegend=False,
    )
    return fig


def zone_stage_chart(matrix: dict, zonas_activas: list) -> go.Figure:
    fig = go.Figure()
    for z in zonas_activas:
        fig.add_trace(
            go.Bar(
                name=z,
                x=STAGES,
                y=[matrix[z][s] for s in STAGES],
                marker_color=COLOR[z],
                hovertemplate=f"<b>{z}</b> — %{{x}}<br>%{{y}} sitios<extra></extra>",
            )
        )
    fig.update_layout(
        barmode="group",
        height=420,
        margin=dict(l=10, r=10, t=30, b=10),
        plot_bgcolor="white",
        paper_bgcolor="white",
        xaxis=dict(showgrid=False),
        yaxis=dict(showgrid=True, gridcolor=GRID, zeroline=False, title="Sitios"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        font=dict(family="system-ui, sans-serif", color="#0b0b0b"),
    )
    return fig


def pct_insrv_chart(matrix: dict) -> go.Figure:
    zonas = list(matrix.keys())
    pcts = [round(100 * matrix[z]["InSrv"] / matrix[z]["Total Sitios"], 1) if matrix[z]["Total Sitios"] else 0 for z in zonas]
    labels = [f"{matrix[z]['InSrv']} / {matrix[z]['Total Sitios']}" for z in zonas]

    fig = go.Figure(
        go.Bar(
            x=pcts,
            y=zonas,
            orientation="h",
            marker=dict(color=GOOD),
            text=[f"{p}%  ({l})" for p, l in zip(pcts, labels)],
            textposition="outside",
            hovertemplate="<b>%{y}</b><br>%{x}%% en servicio<extra></extra>",
        )
    )
    fig.update_layout(
        height=260,
        margin=dict(l=10, r=10, t=10, b=10),
        plot_bgcolor="white",
        paper_bgcolor="white",
        xaxis=dict(showgrid=True, gridcolor=GRID, range=[0, 110], title="% InSrv"),
        yaxis=dict(showgrid=False),
        font=dict(family="system-ui, sans-serif", color="#0b0b0b"),
        showlegend=False,
    )
    return fig


# ----------------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------------
st.title("📡 Modernización 2026 · Bogotá")
st.caption("Embudo de despliegue por zona — se alimenta de la hoja **Sitios** del archivo de control.")

uploaded = st.sidebar.file_uploader(
    "Sube el Excel de control (.xlsx)", type=["xlsx"],
    help="El archivo nunca se guarda en el servidor: se procesa solo en memoria durante tu sesión.",
)

if uploaded is None:
    st.info(
        "⬅️ Sube el archivo **CONTROL...xlsx** en el panel izquierdo para generar el dashboard.\n\n"
        "El archivo no se almacena: se lee en memoria y se descarta al cerrar la sesión."
    )
    st.stop()

try:
    base = load_sitios(uploaded.getvalue())
except Exception as e:
    st.error(f"No se pudo leer el archivo: {e}")
    st.stop()

proyecto = st.sidebar.selectbox(
    "Proyecto", base["proyectos"],
    index=base["proyectos"].index("MODERNIZACION_2026") if "MODERNIZACION_2026" in base["proyectos"] else 0,
)

metrics = compute_metrics(base["rows"], proyecto)
matrix = metrics["matrix"]
zonas = list(matrix.keys())

zonas_sel = st.sidebar.multiselect("Zonas a mostrar", zonas, default=zonas)
if not zonas_sel:
    zonas_sel = zonas

total_sitios = sum(matrix[z]["Total Sitios"] for z in zonas)
total_insrv = sum(matrix[z]["InSrv"] for z in zonas)
total_cargado = sum(matrix[z]["TSS Cargados"] for z in zonas)
total_negados = len(metrics["negados_detalle"])
total_cosite = len(metrics["cosite_detalle"])

notas = []
if metrics["sin_zona"]:
    notas.append(f"{metrics['sin_zona']} sitio(s) del proyecto sin zona asignada (no se incluyen en los totales por zona)")
rfic_total = sum(matrix[z]["RFIC-OK"] for z in zonas)
if rfic_total == 0:
    notas.append("la etapa <b>RFIC-OK</b> está en 0 en todas las zonas — ningún sitio ha llegado aún a esa fase")
if notas:
    st.markdown(
        f"<div class='bug-banner'><b>Nota:</b> " + " · ".join(notas) + ".</div>",
        unsafe_allow_html=True,
    )

# KPIs
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Total sitios", total_sitios)
c2.metric("En servicio (InSrv)", total_insrv, f"{round(100*total_insrv/total_sitios) if total_sitios else 0}% del total")
c3.metric("En proceso", total_cargado - total_insrv)
c4.metric("Negados", total_negados)
c5.metric("Cosite", total_cosite)

st.divider()

st.subheader("Embudo total (todas las zonas)")
st.plotly_chart(funnel_chart({z: matrix[z] for z in zonas_sel}), use_container_width=True)

st.subheader("Avance por zona y etapa")
st.plotly_chart(zone_stage_chart(matrix, zonas_sel), use_container_width=True)

col_a, col_b = st.columns([1, 1])
with col_a:
    st.subheader("% InSrv por zona")
    st.plotly_chart(pct_insrv_chart({z: matrix[z] for z in zonas_sel}), use_container_width=True)
with col_b:
    st.subheader("Sitios Negados / Cosite")
    rows = (
        [{"Sitio": r["sitio"], "Zona": r["zona"], "Estado": "Negado"} for r in metrics["negados_detalle"]]
        + [{"Sitio": r["sitio"], "Zona": r["zona"], "Estado": "Cosite"} for r in metrics["cosite_detalle"]]
    )
    df = pd.DataFrame(rows)
    if zonas_sel and not df.empty:
        df = df[df["Zona"].isin(zonas_sel)]

    if df.empty:
        st.caption("No hay sitios negados ni en cosite para el filtro actual. ✅")
    else:
        tag_class = {"Negado": "tag-negado", "Cosite": "tag-cosite"}
        html_rows = "".join(
            f"<tr><td>{r.Sitio}</td><td>{r.Zona}</td>"
            f"<td><span class='tag {tag_class[r.Estado]}'>{r.Estado}</span></td></tr>"
            for r in df.itertuples()
        )
        st.markdown(
            f"""
            <div class="detail-card">
            <table class="detail">
                <thead><tr><th>Sitio</th><th>Zona</th><th>Estado</th></tr></thead>
                <tbody>{html_rows}</tbody>
            </table>
            </div>
            """,
            unsafe_allow_html=True,
        )

with st.expander("📋 Ver tabla completa (equivalente a la pestaña Resumen)"):
    df_resumen = pd.DataFrame({z: matrix[z] for z in zonas}).T
    df_resumen = df_resumen[STAGES]
    df_resumen.loc["TOTAL"] = df_resumen.sum()
    st.dataframe(df_resumen, use_container_width=True)

st.caption(
    f"Proyecto: **{proyecto}** · {total_sitios} sitios en {len(zonas)} zonas · "
    "Fuente: hoja 'Sitios' (consolidada desde 'Trabajos')."
)
