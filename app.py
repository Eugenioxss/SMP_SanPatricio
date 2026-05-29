from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="Dashboard San Patricio",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# --- SISTEMA DE LOGIN (Acceso Restringido) ---
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

if not st.session_state["autenticado"]:
    st.markdown("<br><br><br>", unsafe_allow_html=True)
    st.markdown("<h1 style='text-align: center; color: #FFFFFF;'>Colegio San Patricio</h1>", unsafe_allow_html=True)
    st.markdown("<h4 style='text-align: center; color: #A1A1AA;'>Portal Financiero Confidencial</h4>", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("login_form"):
            usuario = st.text_input("Usuario")
            password = st.text_input("Contraseña", type="password")
            submit_button = st.form_submit_button("Entrar al Dashboard", use_container_width=True)
            
            if submit_button:
                if usuario == "controller" and password == "SanPatricio2025":
                    st.session_state["autenticado"] = True
                    st.rerun()
                else:
                    st.error("Acceso denegado. Credenciales incorrectas.")
    st.stop()

# --- CSS MODERN FINTECH ---
st.markdown("""
<style>
    /* Forzar fondo claro en toda la app */
    .stApp {
        background-color: #F8FAFC !important;
    }
    
    /* Forzar textos oscuros para matar el Dark Mode automático */
    html, body, [class*="st-"] {
        color: #0F172A !important;
    }
    
    /* --- LO NUEVO: Ocultar TODO el menú, botones y barra de Streamlit --- */
    header[data-testid="stHeader"] {
        display: none !important;
    }
    [data-testid="stToolbar"] {
        display: none !important;
    }
    #MainMenu {
        visibility: hidden !important;
    }
    footer {
        visibility: hidden !important;
    }
    
    /* Reducir el espacio muerto superior e inferior */
    .block-container {
        padding-top: 2rem !important; 
        padding-bottom: 2rem !important;
    }
    
    /* Tarjetas de Métricas Fijas */
    [data-testid="stMetric"] {
        background-color: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-top: 3px solid #1E293B !important;
        padding: 1.2rem !important;
        border-radius: 8px !important;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1) !important;
    }
    
    [data-testid="stMetricLabel"] {
        color: #64748B !important;
    }
    [data-testid="stMetricValue"] {
        color: #0F172A !important;
    }
    
    /* --- ESTILOS PARA EL LOGIN --- */
    [data-testid="stTextInput"] > div > div > input {
        color: #0F172A !important; 
        background-color: #FFFFFF !important; 
        border: 1px solid #CBD5E1 !important; 
    }
    [data-testid="stTextInput"] label {
        color: #475569 !important; 
    }
    [data-testid="stTextInput"] p {
        color: #1E293B !important;
        font-weight: 600 !important;
    }
    
    /* Botón del Login estilo Apple (Oscuro y Minimalista) */
    [data-testid="stFormSubmitButton"] > button {
        background-color: #0F172A !important; 
        color: #FFFFFF !important; 
        border-radius: 8px !important;
        border: none !important;
        font-weight: 600 !important;
        transition: background-color 0.2s !important;
    }
    [data-testid="stFormSubmitButton"] > button:hover {
        background-color: #334155 !important; 
        color: #FFFFFF !important;
    }
    
    /* Tabla de Datos */
    [data-testid="stDataFrame"] {
        background-color: #FFFFFF !important;
    }
    
    /* Títulos */
    h1, h2, h3 {
        color: #0F172A !important;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif !important;
    }
    
    .technical-note {
        color: #64748B;
        font-size: 0.85rem;
        text-align: center;
        margin-top: 2rem;
    }
</style>
""", unsafe_allow_html=True)

BASE_DIR = Path(__file__).resolve().parent
EXTRACTED_DIR = BASE_DIR / "extracted_tables"

def _normalize_text(value: object) -> str:
    text = "" if value is None else str(value)
    return " ".join(text.lower().strip().split())

def _get_column(df: pd.DataFrame, column_name: str) -> str:
    normalized = {_normalize_text(col): col for col in df.columns}
    target = normalized.get(_normalize_text(column_name))
    if target is None:
        raise KeyError(f"Column not found: {column_name}. Available: {list(df.columns)}")
    return target

def _to_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")

def _read_csv(file_name: str) -> pd.DataFrame:
    file_path = EXTRACTED_DIR / file_name
    if not file_path.exists():
        raise FileNotFoundError(f"Missing required file: {file_path}")
    return pd.read_csv(file_path)

def _bucket_aging_label(label: str) -> str:
    text = _normalize_text(label)
    if "30" in text:
        return "30 dias"
    if "60" in text:
        return "60 dias"
    if "90" in text or "120" in text or "mas" in text:
        return "90+ dias"
    return "Otros"

def _currency(value: float) -> str:
    return f"${value:,.0f}"

def _percent(value: float) -> str:
    return f"{value * 100:.1f}%"

@st.cache_data
def load_data() -> dict[str, object]:
    balanza_febrero = _read_csv("Balanza_Febrero__tabla_1.csv")
    balanza_marzo = _read_csv("Balanza_Marzo__tabla_1.csv")
    niveles = _read_csv("Operativo_Marzo__tabla_1.csv")
    resumen = _read_csv("Operativo_Marzo__tabla_2.csv")
    antiguedad = _read_csv("Operativo_Marzo__tabla_3.csv")

    feb = balanza_febrero.loc[
        :,
        [
            _get_column(balanza_febrero, "Cuenta"),
            _get_column(balanza_febrero, "Descripción"),
            _get_column(balanza_febrero, "Saldo final"),
        ],
    ].rename(
        columns={
            _get_column(balanza_febrero, "Cuenta"): "Cuenta",
            _get_column(balanza_febrero, "Descripción"): "Descripcion Febrero",
            _get_column(balanza_febrero, "Saldo final"): "Saldo Febrero",
        }
    )
    mar = balanza_marzo.loc[
        :,
        [
            _get_column(balanza_marzo, "Cuenta"),
            _get_column(balanza_marzo, "Descripción"),
            _get_column(balanza_marzo, "Saldo final"),
        ],
    ].rename(
        columns={
            _get_column(balanza_marzo, "Cuenta"): "Cuenta",
            _get_column(balanza_marzo, "Descripción"): "Descripcion Marzo",
            _get_column(balanza_marzo, "Saldo final"): "Saldo Marzo",
        }
    )

    feb["Saldo Febrero"] = _to_numeric(feb["Saldo Febrero"])
    mar["Saldo Marzo"] = _to_numeric(mar["Saldo Marzo"])

    merged = feb.merge(mar, on="Cuenta", how="outer")
    merged["Saldo Febrero"] = merged["Saldo Febrero"].fillna(0)
    merged["Saldo Marzo"] = merged["Saldo Marzo"].fillna(0)
    merged["Mensual Marzo"] = merged["Saldo Marzo"] - merged["Saldo Febrero"]

    account_series = merged["Cuenta"].astype(str)
    revenue_mask = account_series.str.startswith("4")
    expense_mask = account_series.str.startswith("5")
    payroll_mask = account_series.str.startswith(("5110", "5120", "5130", "5210", "5220", "5230"))
    depreciation_mask = account_series.str.startswith("5900")

    revenue_monthly = -merged.loc[revenue_mask, "Mensual Marzo"].sum(min_count=1)
    expenses_monthly = merged.loc[expense_mask, "Mensual Marzo"].sum(min_count=1)
    depreciation_monthly = merged.loc[depreciation_mask, "Mensual Marzo"].sum(min_count=1)
    expenses_monthly_ex_depr = expenses_monthly - depreciation_monthly
    ebitda_monthly = revenue_monthly - expenses_monthly_ex_depr
    margin_monthly = ebitda_monthly / revenue_monthly if revenue_monthly else 0

    revenue_ytd = -merged.loc[revenue_mask, "Saldo Marzo"].sum(min_count=1)
    expenses_ytd = merged.loc[expense_mask, "Saldo Marzo"].sum(min_count=1)
    depreciation_ytd = merged.loc[depreciation_mask, "Saldo Marzo"].sum(min_count=1)
    ebitda_ytd = revenue_ytd - (expenses_ytd - depreciation_ytd)
    margin_ytd = ebitda_ytd / revenue_ytd if revenue_ytd else 0

    nivel_col = _get_column(resumen, "Nivel")
    alumnos_col = _get_column(resumen, "Alumnos")
    facturado_col = _get_column(resumen, "Facturado del mes")
    cobrado_col = _get_column(resumen, "Cobrado del mes")

    total_mask = resumen[nivel_col].astype(str).str.contains(r"^total$", case=False, na=False)
    detail_mask = ~total_mask

    total_alumnos = _to_numeric(resumen.loc[detail_mask, alumnos_col]).sum(min_count=1)
    total_facturado = _to_numeric(resumen.loc[detail_mask, facturado_col]).sum(min_count=1)
    total_cobrado = _to_numeric(resumen.loc[detail_mask, cobrado_col]).sum(min_count=1)

    if pd.isna(total_alumnos):
        total_alumnos = _to_numeric(resumen.loc[total_mask, alumnos_col]).sum(min_count=1)
    if pd.isna(total_facturado):
        total_facturado = _to_numeric(resumen.loc[total_mask, facturado_col]).sum(min_count=1)
    if pd.isna(total_cobrado):
        total_cobrado = _to_numeric(resumen.loc[total_mask, cobrado_col]).sum(min_count=1)

    total_alumnos = float(0 if pd.isna(total_alumnos) else total_alumnos)
    total_facturado = float(0 if pd.isna(total_facturado) else total_facturado)
    total_cobrado = float(0 if pd.isna(total_cobrado) else total_cobrado)

    cobranza_pct = total_cobrado / total_facturado if total_facturado else 0
    ticket_promedio = total_facturado / total_alumnos if total_alumnos else 0

    antig_label_col = _get_column(antiguedad, "Antigüedad")
    antig_saldo_col = _get_column(antiguedad, "Saldo marzo")

    total_antig_mask = antiguedad[antig_label_col].astype(str).str.contains(r"^total$", case=False, na=False)
    corriente_mask = antiguedad[antig_label_col].astype(str).str.contains(r"al corriente", case=False, na=False)
    overdue_mask = ~(total_antig_mask | corriente_mask)

    overdue_df = antiguedad.loc[overdue_mask, [antig_label_col, antig_saldo_col]].copy()
    overdue_df[antig_saldo_col] = _to_numeric(overdue_df[antig_saldo_col]).fillna(0)
    overdue_df["Bucket"] = overdue_df[antig_label_col].astype(str).map(_bucket_aging_label)

    donut_df = (
        overdue_df.groupby("Bucket", as_index=False)[antig_saldo_col]
        .sum()
        .rename(columns={antig_saldo_col: "Saldo"})
    )
    wanted_buckets = ["30 dias", "60 dias", "90+ dias"]
    donut_df = donut_df[donut_df["Bucket"].isin(wanted_buckets)]

    cartera_total = _to_numeric(antiguedad.loc[total_antig_mask, antig_saldo_col]).sum(min_count=1)
    if pd.isna(cartera_total) or cartera_total == 0:
        cartera_total = _to_numeric(antiguedad.loc[~total_antig_mask, antig_saldo_col]).sum(min_count=1)

    cartera_vencida = float(overdue_df[antig_saldo_col].sum(min_count=1))
    cartera_total = float(0 if pd.isna(cartera_total) else cartera_total)

    kpis_operativos = pd.DataFrame(
        {
            "KPI": ["Alumnos activos", "Ticket promedio", "Cartera vencida", "Cartera total"],
            "Valor": [
                f"{int(total_alumnos):,}",
                _currency(ticket_promedio),
                _currency(cartera_vencida),
                _currency(cartera_total),
            ],
        }
    )

    return {
        "revenue_monthly": float(revenue_monthly),
        "revenue_ytd": float(revenue_ytd),
        "ebitda_monthly": float(ebitda_monthly),
        "ebitda_ytd": float(ebitda_ytd),
        "margin_monthly": float(margin_monthly),
        "margin_ytd": float(margin_ytd),
        "cobranza_pct": float(cobranza_pct),
        "ticket_promedio": float(ticket_promedio),
        "cartera_vencida": float(cartera_vencida),
        "donut_df": donut_df,
        "kpis_operativos_df": kpis_operativos,
        "niveles": niveles,
    }


st.title("Dashboard - Colegio San Patricio")
st.markdown("---")

try:
    data = load_data()
except Exception as exc:
    st.error(f"No fue posible cargar los datos: {exc}")
    st.stop()

# --- BLOQUE 1: KPIs FINANCIEROS ---
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric(
        label="Revenue Mensual",
        value=_currency(data["revenue_monthly"]),
        delta=f"YTD: {_currency(data['revenue_ytd'])}",
    )

with col2:
    st.metric(
        label="EBITDA Mensual",
        value=_currency(data["ebitda_monthly"]),
        delta=f"YTD: {_currency(data['ebitda_ytd'])}",
    )

with col3:
    st.metric(
        label="Margen EBITDA",
        value=_percent(data["margin_monthly"]),
        delta=f"YTD: {_percent(data['margin_ytd'])}",
    )

with col4:
    target_cobranza = 0.95
    delta_cobranza = data["cobranza_pct"] - target_cobranza
    st.metric(
        label="Eficiencia de Cobranza",
        value=_percent(data["cobranza_pct"]),
        delta=f"vs meta 95%: {delta_cobranza * 100:+.1f} pp",
    )

st.markdown("<br><br>", unsafe_allow_html=True)

# --- BLOQUE 2: KPIs OPERATIVOS Y GRÁFICOS ---
col_izq, col_der = st.columns(2, gap="large")

with col_izq:
    st.subheader("Distribución de Cartera Vencida")
    if data["donut_df"].empty:
        st.info("No hay datos de cartera vencida para mostrar.")
    else:
        fig = px.pie(
            data["donut_df"],
            names="Bucket",
            values="Saldo",
            hole=0.65,
            color="Bucket",
            color_discrete_map={
                "30 dias": "#F59E0B",   # Ámbar
                "60 dias": "#F97316",   # Naranja
                "90+ dias": "#DC2626"   # Rojo
            }
        )
        
        # Fondo blanco fijo para matar la interferencia del dark mode
        fig.update_layout(
            paper_bgcolor='#FFFFFF',
            plot_bgcolor='#FFFFFF',
            margin=dict(t=10, b=20, l=0, r=0),
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=-0.15, xanchor="center", x=0.5, font=dict(color="#0F172A"))
        )
        
        fig.update_traces(
            textposition="inside",
            textinfo="percent",
            textfont_size=14,
            textfont_color="white",
            marker=dict(line=dict(color='#FFFFFF', width=2))
        )
        st.plotly_chart(fig, use_container_width=True)

with col_der:
    st.subheader("Indicadores Operativos")
    st.markdown("<br>", unsafe_allow_html=True)
    
    # 1. Definimos el diseño elegante usando Pandas Styling
    styled_df = data["kpis_operativos_df"].style.set_properties(**{
        'background-color': '#FFFFFF',
        'color': '#0F172A',
        'border-color': '#E2E8F0',
        'font-family': '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif',
        'padding': '12px',
        'border-bottom': '1px solid #E2E8F0'
    }).set_table_styles([{
        'selector': 'th',
        'props': [
            ('background-color', '#F8FAFC'),
            ('color', '#64748B'),
            ('font-weight', '600'),
            ('text-transform', 'uppercase'),
            ('font-size', '12px'),
            ('border-bottom', '2px solid #E2E8F0'),
            ('padding', '12px')
        ]
    }])
    
    # 2. Renderizamos la tabla estilizada
    st.dataframe(styled_df, use_container_width=True, hide_index=True)


# --- BLOQUE 3: EL FLEX MATEMÁTICO (CADENA DE MARKOV) ---
st.markdown("---")
st.subheader("Proyección de Riesgo (Abril 2025)")

# Extracción segura de los buckets para la heurística
donut_df = data["donut_df"]
def get_bucket_saldo(bucket_name):
    val = donut_df.loc[donut_df["Bucket"] == bucket_name, "Saldo"].sum()
    return float(val) if not pd.isna(val) else 0.0

s_30 = get_bucket_saldo("30 dias")
s_60 = get_bucket_saldo("60 dias")
s_90 = get_bucket_saldo("90+ dias")

# Lógica del Modelo de Markov Heurístico
trans_30_to_60 = s_30 * 0.40
trans_60_to_90 = s_60 * 0.70
stay_90 = s_90 * 0.90

riesgo_60_abril = trans_30_to_60
riesgo_90_abril = trans_60_to_90 + stay_90

total_vencido = s_30 + s_60 + s_90
recuperacion_estimada = total_vencido - (riesgo_60_abril + riesgo_90_abril)

col_r1, col_r2, col_r3 = st.columns(3)

with col_r1:
    st.metric(
        label="Recuperación Estimada",
        value=_currency(recuperacion_estimada),
        delta="Flujo proyectado a bancos",
        delta_color="normal"
    )
with col_r2:
    st.metric(
        label="Transición a 60 días",
        value=_currency(riesgo_60_abril),
        delta="Riesgo Medio",
        delta_color="off"
    )
with col_r3:
    st.metric(
        label="Deuda Crítica (90+ días)",
        value=_currency(riesgo_90_abril),
        delta="Riesgo Alto",
        delta_color="inverse"
    )

st.markdown("""
<div class="technical-note">
    <strong>Nota Técnica:</strong> Proyección calculada mediante un modelo de Markov heurístico. 
    Parámetros de transición (P): 40% (30→60 días), 70% (60→90 días) y 90% (retención estructural 90+).
</div>
""", unsafe_allow_html=True)