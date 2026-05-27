# dashboard_app.py — Dashboard Web PDF Amanda
# Streamlit + Plotly + SQLite
# Rodar: streamlit run dashboard_app.py

import streamlit as st
import pandas as pd
import sqlite3
import os
import io
from datetime import datetime, date
from pathlib import Path

# ── Configuração da página ──────────────────────────────────────
st.set_page_config(
    page_title="Dashboard PDF Amanda",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Localiza o banco de dados ───────────────────────────────────
PASTA_SCRIPT = Path(__file__).parent
# Procura o banco na pasta SAIDA ou na raiz (compatível com Render)
_banco_saida = PASTA_SCRIPT / "SAIDA" / "procedimentos.db"
_banco_raiz  = PASTA_SCRIPT / "procedimentos.db"
BANCO_PATH   = _banco_saida if _banco_saida.exists() else _banco_raiz

_pesq_saida  = PASTA_SCRIPT / "SAIDA" / "pesquisa_clinica.db"
_pesq_raiz   = PASTA_SCRIPT / "pesquisa_clinica.db"
PESQUISA_DB  = _pesq_saida if _pesq_saida.exists() else _pesq_raiz

# ── Carrega dados ───────────────────────────────────────────────
@st.cache_data(ttl=60)
def carregar_dados():
    if not BANCO_PATH.exists():
        return pd.DataFrame()
    conn = sqlite3.connect(str(BANCO_PATH))
    df = pd.read_sql_query("SELECT * FROM registros_estruturados", conn)
    conn.close()
    # Normaliza datas
    for col in ["data_inicio", "data_fim", "data_atendimento", "data_alta", "dt_nascimento"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
    return df

@st.cache_data(ttl=60)
def carregar_pesquisa():
    if not BANCO_PATH.exists():
        return pd.DataFrame()
    conn = sqlite3.connect(str(BANCO_PATH))
    tabelas = pd.read_sql_query(
        "SELECT name FROM sqlite_master WHERE type='table'", conn
    )["name"].tolist()
    if "pesquisa_clinica" in tabelas:
        df = pd.read_sql_query("SELECT * FROM pesquisa_clinica", conn)
    else:
        df = pd.DataFrame()
    conn.close()
    return df

def salvar_pesquisa(df_editado):
    conn = sqlite3.connect(str(BANCO_PATH))
    df_editado.to_sql("pesquisa_clinica", conn, if_exists="replace", index=False)
    conn.close()

def df_para_excel(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Dados")
    return buf.getvalue()

# ── CSS + Fonte ─────────────────────────────────────────────────
st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
/* ── Fonte global ── */
*, *::before, *::after, .stApp, .stMarkdown, .stDataFrame,
[data-testid], [class*="st-"] {
    font-family: 'Inter', sans-serif !important;
}

/* ── Variáveis de cor ── */
:root {
    --navy:   #0A2540;
    --blue:   #185FA5;
    --blue2:  #2278C9;
    --teal:   #1D9E75;
    --accent: #F0A500;
    --danger: #D85A30;
    --bg:     #F4F7FC;
    --card:   #FFFFFF;
    --border: #DDE4EE;
    --text:   #1A2B40;
    --muted:  #637383;
}

/* ── Esconde elementos Streamlit padrão ── */
#MainMenu          { visibility: hidden !important; }
footer             { visibility: hidden !important; }
[data-testid="stDecoration"]   { display: none !important; }
[data-testid="stToolbar"]      { display: none !important; }
[data-testid="stStatusWidget"] { display: none !important; }
header             { visibility: hidden !important; }

/* ── Scrollbar fina ── */
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: var(--bg); }
::-webkit-scrollbar-thumb { background: #C4CDD8; border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: var(--muted); }

/* ── Fundo geral ── */
.stApp { background-color: var(--bg) !important; }
.main .block-container { padding-top: 1.5rem !important; max-width: 1200px; }

/* ── Sidebar ── */
[data-testid="stSidebar"] {
    background: var(--navy) !important;
    border-right: 1px solid rgba(255,255,255,0.08) !important;
}
[data-testid="stSidebar"] * { color: #E8EFF8 !important; }
[data-testid="stSidebar"] .stMarkdown h2,
[data-testid="stSidebar"] .stMarkdown h3 { color: #FFFFFF !important; }
[data-testid="stSidebar"] .stMultiSelect [data-baseweb="select"] {
    background: rgba(255,255,255,0.08) !important;
    border-color: rgba(255,255,255,0.18) !important;
    border-radius: 6px !important;
}
[data-testid="stSidebar"] label { color: #A8BEDC !important; font-size: 0.8rem !important; font-weight: 500 !important; text-transform: uppercase !important; letter-spacing: 0.04em !important; }
[data-testid="stSidebar"] [data-testid="stDateInput"] input { background: rgba(255,255,255,0.08) !important; border-color: rgba(255,255,255,0.18) !important; color: #E8EFF8 !important; }
[data-testid="stSidebar"] button {
    background: rgba(24,95,165,0.5) !important;
    border: 1px solid rgba(255,255,255,0.15) !important;
    color: #FFFFFF !important;
    border-radius: 6px !important;
    width: 100% !important;
}
[data-testid="stSidebar"] button:hover { background: var(--blue) !important; }
[data-testid="stSidebar"] hr { border-color: rgba(255,255,255,0.1) !important; }

/* ── Sidebar — badge de status ── */
.sidebar-stat {
    background: rgba(255,255,255,0.07);
    border: 1px solid rgba(255,255,255,0.1);
    border-radius: 8px;
    padding: 0.6rem 0.9rem;
    margin-bottom: 0.5rem;
    display: flex;
    justify-content: space-between;
    align-items: center;
}
.sidebar-stat .s-label { font-size: 0.75rem; color: #A8BEDC; }
.sidebar-stat .s-value { font-size: 0.95rem; font-weight: 700; color: #FFFFFF; }

/* ── Header da página ── */
.page-header {
    background: linear-gradient(135deg, var(--navy) 0%, #1A3D6E 100%);
    border-radius: 12px;
    padding: 1.25rem 1.75rem;
    margin-bottom: 1.5rem;
    display: flex;
    align-items: center;
    justify-content: space-between;
    border: 1px solid rgba(255,255,255,0.08);
    box-shadow: 0 4px 20px rgba(10,37,64,0.15);
}
.page-header .h-left { display: flex; align-items: center; gap: 1rem; }
.page-header .h-icon { font-size: 2.2rem; }
.page-header .h-title { font-size: 1.4rem; font-weight: 700; color: #FFFFFF; margin: 0; line-height: 1.2; }
.page-header .h-sub   { font-size: 0.8rem; color: #A8BEDC; margin: 0; }
.page-header .h-badge {
    background: rgba(255,255,255,0.12);
    border: 1px solid rgba(255,255,255,0.18);
    border-radius: 20px;
    padding: 0.3rem 0.9rem;
    font-size: 0.78rem;
    color: #D0E4F7;
    font-weight: 500;
}

/* ── Cards de métrica ── */
.metric-card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 1.1rem 1.25rem;
    margin-bottom: 0.5rem;
    box-shadow: 0 2px 8px rgba(10,37,64,0.06);
    position: relative;
    overflow: hidden;
    transition: box-shadow 0.2s;
}
.metric-card:hover { box-shadow: 0 4px 16px rgba(10,37,64,0.12); }
.metric-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0;
    width: 4px; height: 100%;
    background: var(--accent-color, var(--blue));
    border-radius: 12px 0 0 12px;
}
.metric-card .mc-icon {
    font-size: 1.8rem;
    margin-bottom: 0.4rem;
    display: block;
}
.metric-card .mc-value {
    font-size: 1.9rem;
    font-weight: 700;
    color: var(--text);
    line-height: 1.1;
    display: block;
}
.metric-card .mc-label {
    font-size: 0.78rem;
    color: var(--muted);
    font-weight: 500;
    text-transform: uppercase;
    letter-spacing: 0.04em;
    display: block;
    margin-top: 0.15rem;
}

/* ── Abas ── */
[data-testid="stTabs"] [data-baseweb="tab"] {
    font-weight: 600 !important;
    font-size: 0.85rem !important;
    color: var(--muted) !important;
    border-radius: 8px 8px 0 0 !important;
    padding: 0.6rem 1rem !important;
}
[data-testid="stTabs"] [aria-selected="true"] {
    color: var(--blue) !important;
    border-bottom: 2px solid var(--blue) !important;
}

/* ── Títulos de seção ── */
.section-title {
    font-size: 0.95rem;
    font-weight: 700;
    color: var(--text);
    letter-spacing: 0.01em;
    margin-bottom: 0.75rem;
    padding-bottom: 0.4rem;
    border-bottom: 2px solid var(--blue);
    display: inline-block;
}

/* ── Ficha de paciente ── */
.ficha-box {
    background: #EBF3FB;
    border-left: 4px solid var(--blue);
    border-radius: 0 8px 8px 0;
    padding: 1rem 1.25rem;
    margin-bottom: 1rem;
    font-family: 'Courier New', monospace;
    font-size: 0.82rem;
    white-space: pre-wrap;
    color: var(--text);
    line-height: 1.6;
}

/* ── Badges ── */
.badge-d  { background:#E0EEF9; color:#0C447C; padding:3px 10px; border-radius:20px; font-size:11px; font-weight:700; letter-spacing:0.03em; }
.badge-e  { background:#E2F4EC; color:#1A6B40; padding:3px 10px; border-radius:20px; font-size:11px; font-weight:700; letter-spacing:0.03em; }
.badge-na { background:#F0EDE5; color:#5A4F3A; padding:3px 10px; border-radius:20px; font-size:11px; font-weight:700; letter-spacing:0.03em; }

/* ── Dataframe ── */
[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; border: 1px solid var(--border); }

/* ── Botão primário ── */
[data-testid="stButton"] button[kind="primary"] {
    background: var(--blue) !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
}
[data-testid="stButton"] button[kind="primary"]:hover {
    background: var(--blue2) !important;
}

/* ── Botão download ── */
[data-testid="stDownloadButton"] button {
    border: 1.5px solid var(--blue) !important;
    color: var(--blue) !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    background: transparent !important;
}
[data-testid="stDownloadButton"] button:hover {
    background: var(--blue) !important;
    color: #FFFFFF !important;
}

/* ── Métricas nativas Streamlit (fallback) ── */
[data-testid="stMetricValue"] { font-size: 1.9rem !important; font-weight: 700 !important; color: var(--text) !important; }
[data-testid="stMetricLabel"] { font-size: 0.78rem !important; color: var(--muted) !important; font-weight: 500 !important; text-transform: uppercase !important; }

/* ── Inputs de texto ── */
[data-testid="stTextInput"] input {
    border-radius: 8px !important;
    border: 1.5px solid var(--border) !important;
    background: white !important;
    color: var(--text) !important;
    font-size: 0.9rem !important;
    padding: 0.5rem 0.75rem !important;
    transition: border-color 0.2s, box-shadow 0.2s !important;
}
[data-testid="stTextInput"] input:focus {
    border-color: var(--blue) !important;
    box-shadow: 0 0 0 3px rgba(24,95,165,0.12) !important;
    outline: none !important;
}
[data-testid="stTextInput"] input::placeholder { color: #9EB0C4 !important; }

/* ── Expanders (fichas de paciente) ── */
[data-testid="stExpander"] {
    background: var(--card) !important;
    border: 1px solid var(--border) !important;
    border-radius: 10px !important;
    margin-bottom: 0.6rem !important;
    overflow: hidden !important;
    box-shadow: 0 1px 6px rgba(10,37,64,0.05) !important;
}
[data-testid="stExpander"]:hover {
    border-color: var(--blue) !important;
    box-shadow: 0 2px 12px rgba(24,95,165,0.1) !important;
}
[data-testid="stExpander"] summary {
    font-weight: 600 !important;
    color: var(--text) !important;
    font-size: 0.9rem !important;
    padding: 0.75rem 1rem !important;
}
[data-testid="stExpander"] summary:hover { background: #F0F5FB !important; }

/* ── Alertas / info boxes ── */
[data-testid="stInfo"] {
    background: #EBF3FB !important;
    border-left: 4px solid var(--blue) !important;
    border-radius: 0 8px 8px 0 !important;
    color: var(--text) !important;
}
[data-testid="stWarning"] {
    background: #FFF8EC !important;
    border-left: 4px solid var(--accent) !important;
    border-radius: 0 8px 8px 0 !important;
}
[data-testid="stSuccess"] {
    background: #EAFAF3 !important;
    border-left: 4px solid var(--teal) !important;
    border-radius: 0 8px 8px 0 !important;
}
[data-testid="stError"] {
    background: #FDF0EC !important;
    border-left: 4px solid var(--danger) !important;
    border-radius: 0 8px 8px 0 !important;
}

/* ── Chips de filtro ── */
.filter-chips { display: flex; flex-wrap: wrap; gap: 0.4rem; margin-bottom: 1rem; }
.chip {
    background: rgba(24,95,165,0.1);
    border: 1px solid rgba(24,95,165,0.25);
    color: var(--blue);
    border-radius: 20px;
    padding: 3px 10px;
    font-size: 0.75rem;
    font-weight: 600;
    display: inline-flex;
    align-items: center;
    gap: 4px;
}

/* ── Selectbox ── */
[data-baseweb="select"] {
    border-radius: 8px !important;
}
[data-baseweb="select"] > div {
    border-radius: 8px !important;
    border-color: var(--border) !important;
    background: white !important;
}
[data-baseweb="select"] > div:focus-within {
    border-color: var(--blue) !important;
    box-shadow: 0 0 0 3px rgba(24,95,165,0.12) !important;
}

/* ── Divider ── */
hr { border-color: var(--border) !important; margin: 1rem 0 !important; }

/* ── Caption / rodapé de tabela ── */
[data-testid="stCaptionContainer"] { color: var(--muted) !important; font-size: 0.78rem !important; }
</style>
""", unsafe_allow_html=True)

def render_metric_card(icon, value, label, accent_color="#185FA5"):
    """Renderiza card de métrica estilizado."""
    st.markdown(f"""
    <div class="metric-card" style="--accent-color:{accent_color}">
        <span class="mc-icon">{icon}</span>
        <span class="mc-value">{value}</span>
        <span class="mc-label">{label}</span>
    </div>
    """, unsafe_allow_html=True)

# ── Carrega dados ───────────────────────────────────────────────
df_total = carregar_dados()
banco_ok  = not df_total.empty

# ── Sidebar ─────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="padding:0.5rem 0 1.2rem 0; text-align:center;">
        <div style="font-size:2.4rem; margin-bottom:0.3rem;">🏥</div>
        <div style="font-size:1.1rem; font-weight:800; color:#FFFFFF; letter-spacing:0.02em;">PDF Amanda</div>
        <div style="font-size:0.75rem; color:#A8BEDC; margin-top:0.1rem; font-weight:400;">Dashboard de Produção Cirúrgica</div>
        <div style="font-size:0.72rem; color:#7A9EC0; margin-top:0.15rem;">HUPE · UERJ</div>
    </div>
    """, unsafe_allow_html=True)
    st.divider()

    if not banco_ok:
        st.error("Banco não encontrado.\nProcesse os PDFs primeiro\n(opção 1 do menu).")
        st.stop()

    # Stats rápidos
    ultima_str = "—"
    if "data_inicio" in df_total.columns:
        ultima = df_total["data_inicio"].max()
        if pd.notna(ultima):
            ultima_str = ultima.strftime("%d/%m/%Y")
    st.markdown(f"""
    <div class="sidebar-stat">
        <span class="s-label">📦 Registros</span>
        <span class="s-value">{len(df_total):,}</span>
    </div>
    <div class="sidebar-stat">
        <span class="s-label">📅 Último proc.</span>
        <span class="s-value">{ultima_str}</span>
    </div>
    """, unsafe_allow_html=True)
    st.divider()

    st.markdown("""
    <div style="font-size:0.72rem; font-weight:700; color:#A8BEDC; text-transform:uppercase; letter-spacing:0.08em; margin-bottom:0.75rem;">
        ⚙️ &nbsp;Filtros
    </div>
    """, unsafe_allow_html=True)

    # Tipo de procedimento
    tipos = sorted(df_total["tipo_procedimento"].dropna().unique().tolist()) \
            if "tipo_procedimento" in df_total.columns else []
    sel_tipo = st.multiselect("Tipo de procedimento", tipos)

    # Cirurgião
    cirurgioes = sorted(df_total["cirurgiao"].dropna().unique().tolist()) \
                 if "cirurgiao" in df_total.columns else []
    sel_cir = st.multiselect("Cirurgião", cirurgioes)

    # Período
    if "data_inicio" in df_total.columns:
        dmin = df_total["data_inicio"].min()
        dmax = df_total["data_inicio"].max()
        if pd.notna(dmin) and pd.notna(dmax):
            sel_periodo = st.date_input(
                "Período",
                value=(dmin.date(), dmax.date()),
                min_value=dmin.date(),
                max_value=dmax.date()
            )
        else:
            sel_periodo = None
    else:
        sel_periodo = None

    # Município
    municipios = sorted(df_total["municipio"].dropna().unique().tolist()) \
                 if "municipio" in df_total.columns else []
    sel_mun = st.multiselect("Município", municipios)

    # Lateralidade
    lats = sorted(df_total["lateralidade"].dropna().unique().tolist()) \
           if "lateralidade" in df_total.columns else []
    sel_lat = st.multiselect("Lateralidade", lats)

    # Tipo atendimento
    atends = sorted(df_total["tipo_atendimento"].dropna().unique().tolist()) \
             if "tipo_atendimento" in df_total.columns else []
    sel_atend = st.multiselect("Tipo de atendimento", atends)

    if st.button("🔄 Limpar filtros"):
        st.rerun()

# ── Aplica filtros ───────────────────────────────────────────────
df = df_total.copy()
if sel_tipo:
    df = df[df["tipo_procedimento"].isin(sel_tipo)]
if sel_cir:
    df = df[df["cirurgiao"].isin(sel_cir)]
if sel_periodo and len(sel_periodo) == 2 and "data_inicio" in df.columns:
    d0 = pd.Timestamp(sel_periodo[0])
    d1 = pd.Timestamp(sel_periodo[1])
    df = df[(df["data_inicio"] >= d0) & (df["data_inicio"] <= d1)]
if sel_mun:
    df = df[df["municipio"].isin(sel_mun)]
if sel_lat:
    df = df[df["lateralidade"].isin(sel_lat)]
if sel_atend:
    df = df[df["tipo_atendimento"].isin(sel_atend)]

# ── Header da página ────────────────────────────────────────────
filtros_ativos = sum([bool(sel_tipo), bool(sel_cir), bool(sel_mun), bool(sel_lat), bool(sel_atend)])
badge_filtro = f"<span class='h-badge'>🔍 {filtros_ativos} filtro(s) ativo(s)</span>" if filtros_ativos else "<span class='h-badge'>Todos os registros</span>"
st.markdown(f"""
<div class="page-header">
    <div class="h-left">
        <span class="h-icon">🏥</span>
        <div>
            <p class="h-title">Dashboard de Produção Cirúrgica</p>
            <p class="h-sub">HUPE · UERJ &nbsp;·&nbsp; Análise de Procedimentos Vasculares</p>
        </div>
    </div>
    {badge_filtro}
</div>
""", unsafe_allow_html=True)

# Chips de filtros ativos
chips_html = []
for label, valores in [
    ("Tipo", sel_tipo), ("Cirurgião", sel_cir),
    ("Município", sel_mun), ("Lado", sel_lat), ("Atendimento", sel_atend)
]:
    for v in (valores or []):
        chips_html.append(f'<span class="chip">🏷 {label}: {v}</span>')
if chips_html:
    st.markdown(f'<div class="filter-chips">{"".join(chips_html)}</div>', unsafe_allow_html=True)

# ── Abas principais ──────────────────────────────────────────────
aba1, aba2, aba3, aba4, aba5 = st.tabs([
    "📊 Visão Geral",
    "📋 Tabela de Registros",
    "👨‍⚕️ Cirurgiões",
    "👤 Pacientes",
    "🔬 Pesquisa Clínica"
])

# ════════════════════════════════════════
# ABA 1 — VISÃO GERAL
# ════════════════════════════════════════
with aba1:
    try:
        import plotly.express as px
        import plotly.graph_objects as go
        PLOTLY = True
    except ImportError:
        PLOTLY = False
        st.warning("Instale plotly para ver os gráficos: pip install plotly")

    # Métricas
    total   = len(df)
    unicos  = df["cpf"].nunique() if "cpf" in df.columns else 0
    cirurg  = df["cirurgiao"].nunique() if "cirurgiao" in df.columns else 0
    dur_med = int(df["duracao_minutos"].mean()) \
              if "duracao_minutos" in df.columns and df["duracao_minutos"].notna().any() else 0

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        render_metric_card("🩺", f"{total:,}", "Procedimentos", "#185FA5")
    with col2:
        render_metric_card("👤", f"{unicos:,}", "Pacientes únicos", "#1D9E75")
    with col3:
        render_metric_card("👨‍⚕️", str(cirurg), "Cirurgiões distintos", "#534AB7")
    with col4:
        render_metric_card("⏱️", f"{dur_med} min", "Duração média", "#F0A500")

    if PLOTLY and total > 0:
        st.divider()
        c1, c2 = st.columns(2)

        # Barras por tipo
        with c1:
            st.markdown('<span class="section-title">Procedimentos por tipo</span>', unsafe_allow_html=True)
            if "tipo_procedimento" in df.columns:
                ct = df["tipo_procedimento"].value_counts().reset_index()
                ct.columns = ["Tipo", "Total"]
                fig = px.bar(ct, x="Tipo", y="Total",
                             color="Tipo",
                             color_discrete_sequence=["#185FA5","#1D9E75","#D85A30","#534AB7"])
                fig.update_traces(hovertemplate="<b>%{x}</b><br>Total: %{y}<extra></extra>")
                fig.update_layout(showlegend=False, margin=dict(t=10,b=10),
                                  height=280, plot_bgcolor="white",
                                  paper_bgcolor="white",
                                  font=dict(family="Inter, sans-serif", size=12, color="#1A2B40"),
                                  hoverlabel=dict(bgcolor="white", bordercolor="#DDE4EE",
                                                  font_size=13, font_family="Inter, sans-serif"))
                fig.update_xaxes(showgrid=False)
                fig.update_yaxes(showgrid=True, gridcolor="#EEF0F4")
                st.plotly_chart(fig, use_container_width=True)

        # Rosca lateralidade
        with c2:
            st.markdown('<span class="section-title">Lateralidade</span>', unsafe_allow_html=True)
            if "lateralidade" in df.columns:
                cl = df["lateralidade"].value_counts().reset_index()
                cl.columns = ["Lado", "Total"]
                fig2 = px.pie(cl, names="Lado", values="Total", hole=0.5,
                              color_discrete_sequence=["#185FA5","#1D9E75","#534AB7","#D85A30"])
                fig2.update_layout(margin=dict(t=10,b=10), height=280,
                                   paper_bgcolor="white",
                                   font=dict(family="sans-serif", size=12, color="#1A2B40"))
                fig2.update_traces(textposition="outside", textinfo="percent+label")
                st.plotly_chart(fig2, use_container_width=True)

        # Linha mensal
        st.markdown('<span class="section-title">Evolução mensal por tipo de procedimento</span>', unsafe_allow_html=True)
        if "data_inicio" in df.columns and "tipo_procedimento" in df.columns:
            df_m = df.dropna(subset=["data_inicio"]).copy()
            df_m["mes"] = df_m["data_inicio"].dt.to_period("M").astype(str)
            mensal = df_m.groupby(["mes","tipo_procedimento"]).size().reset_index(name="Total")
            fig3 = px.line(mensal, x="mes", y="Total", color="tipo_procedimento",
                           color_discrete_sequence=["#185FA5","#1D9E75","#D85A30","#534AB7"],
                           markers=True)
            fig3.update_traces(hovertemplate="<b>%{x}</b><br>Procedimentos: %{y}<extra></extra>")
            fig3.update_layout(margin=dict(t=10,b=40), height=300,
                               xaxis_title="", yaxis_title="Procedimentos",
                               legend_title="", plot_bgcolor="white",
                               paper_bgcolor="white",
                               font=dict(family="Inter, sans-serif", size=12, color="#1A2B40"),
                               hoverlabel=dict(bgcolor="white", bordercolor="#DDE4EE",
                                               font_size=13, font_family="Inter, sans-serif"),
                               legend=dict(orientation="h", yanchor="bottom", y=1.02,
                                           xanchor="right", x=1))
            fig3.update_xaxes(tickangle=45, showgrid=False)
            fig3.update_yaxes(showgrid=True, gridcolor="#EEF0F4")
            st.plotly_chart(fig3, use_container_width=True)

        c3, c4 = st.columns(2)

        # Top cirurgiões
        with c3:
            st.markdown('<span class="section-title">Top 10 cirurgiões</span>', unsafe_allow_html=True)
            if "cirurgiao" in df.columns:
                top_c = df["cirurgiao"].value_counts().head(10).reset_index()
                top_c.columns = ["Cirurgião", "Total"]
                top_c["Nome"] = top_c["Cirurgião"].str.split().str[:2].str.join(" ")
                fig4 = px.bar(top_c, x="Total", y="Nome", orientation="h",
                              color_discrete_sequence=["#185FA5"])
                fig4.update_traces(hovertemplate="<b>%{y}</b><br>Procedimentos: %{x}<extra></extra>")
                fig4.update_layout(showlegend=False, margin=dict(t=10,b=10),
                                   height=320, yaxis_title="",
                                   plot_bgcolor="white", paper_bgcolor="white",
                                   font=dict(family="Inter, sans-serif", size=12, color="#1A2B40"),
                                   hoverlabel=dict(bgcolor="white", bordercolor="#DDE4EE",
                                                   font_size=13, font_family="Inter, sans-serif"))
                fig4.update_xaxes(showgrid=True, gridcolor="#EEF0F4")
                fig4.update_yaxes(showgrid=False)
                st.plotly_chart(fig4, use_container_width=True)

        # Internação vs ambulatorial
        with c4:
            st.markdown('<span class="section-title">Tipo de atendimento</span>', unsafe_allow_html=True)
            if "tipo_atendimento" in df.columns:
                ta = df["tipo_atendimento"].value_counts().reset_index()
                ta.columns = ["Tipo", "Total"]
                fig5 = px.pie(ta, names="Tipo", values="Total", hole=0.5,
                              color_discrete_sequence=["#534AB7","#F0A500"])
                fig5.update_layout(margin=dict(t=10,b=10), height=320,
                                   paper_bgcolor="white",
                                   font=dict(family="sans-serif", size=12, color="#1A2B40"))
                fig5.update_traces(textposition="outside", textinfo="percent+label")
                st.plotly_chart(fig5, use_container_width=True)

        # Top municípios
        st.markdown('<span class="section-title">Top 10 municípios de origem</span>', unsafe_allow_html=True)
        if "municipio" in df.columns:
            top_m = df["municipio"].value_counts().head(10).reset_index()
            top_m.columns = ["Município", "Total"]
            fig6 = px.bar(top_m, x="Município", y="Total",
                          color_discrete_sequence=["#0A2540"])
            fig6.update_traces(hovertemplate="<b>%{x}</b><br>Procedimentos: %{y}<extra></extra>")
            fig6.update_layout(showlegend=False, margin=dict(t=10,b=60),
                               height=300, plot_bgcolor="white",
                               paper_bgcolor="white",
                               font=dict(family="Inter, sans-serif", size=12, color="#1A2B40"),
                               hoverlabel=dict(bgcolor="white", bordercolor="#DDE4EE",
                                               font_size=13, font_family="Inter, sans-serif"))
            fig6.update_xaxes(tickangle=35, showgrid=False)
            fig6.update_yaxes(showgrid=True, gridcolor="#EEF0F4")
            st.plotly_chart(fig6, use_container_width=True)

# ════════════════════════════════════════
# ABA 2 — TABELA DE REGISTROS
# ════════════════════════════════════════
with aba2:
    st.markdown("#### Registros filtrados")
    busca = st.text_input("🔍 Buscar (nome, prontuário, CPF, cirurgião, município)",
                          placeholder="Ex: CRISTIANE ou 2578507")

    df_view = df.copy()
    if busca:
        mask = pd.Series([False] * len(df_view))
        for col in ["nome","prontuario","cpf","cirurgiao","municipio"]:
            if col in df_view.columns:
                mask |= df_view[col].astype(str).str.upper().str.contains(
                    busca.upper(), na=False)
        df_view = df_view[mask]

    st.markdown(f"**{len(df_view):,} registro(s) encontrado(s)**")

    # Colunas de exibição
    cols_exib = [c for c in [
        "tipo_procedimento","prontuario","nome","cpf","dt_nascimento",
        "idade","sexo","municipio","tipo_atendimento","data_inicio",
        "data_fim","duracao_minutos","lateralidade","cirurgiao",
        "anestesista","centro_sala"
    ] if c in df_view.columns]

    st.dataframe(
        df_view[cols_exib].rename(columns={
            "tipo_procedimento":"Tipo","prontuario":"Prontuário",
            "nome":"Nome","cpf":"CPF","dt_nascimento":"Dt Nasc.",
            "idade":"Idade","sexo":"Sexo","municipio":"Município",
            "tipo_atendimento":"Atendimento","data_inicio":"Início",
            "data_fim":"Fim","duracao_minutos":"Dur.(min)",
            "lateralidade":"Lado","cirurgiao":"Cirurgião",
            "anestesista":"Anestesista","centro_sala":"Sala"
        }),
        use_container_width=True,
        height=420
    )

    st.download_button(
        label="⬇️ Exportar para Excel",
        data=df_para_excel(df_view[cols_exib]),
        file_name=f"registros_filtrados_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ════════════════════════════════════════
# ABA 3 — CIRURGIÕES
# ════════════════════════════════════════
with aba3:
    st.markdown("#### Ranking de cirurgiões")

    if "cirurgiao" not in df.columns:
        st.info("Sem dados de cirurgiões.")
    else:
        try:
            import plotly.express as px
            PLOTLY = True
        except ImportError:
            PLOTLY = False

        ranking = df.groupby("cirurgiao").agg(
            Total=("prontuario","count"),
            Pacientes=("cpf","nunique"),
            Dur_media=("duracao_minutos","mean"),
            Primeiro=("data_inicio","min"),
            Ultimo=("data_inicio","max")
        ).reset_index().sort_values("Total", ascending=False)

        # Por tipo
        if "tipo_procedimento" in df.columns:
            for tp in df["tipo_procedimento"].unique():
                col_nome = tp[:12] if tp else "Outro"
                ranking[col_nome] = df[df["tipo_procedimento"]==tp]\
                    .groupby("cirurgiao")["prontuario"].count().reindex(
                    ranking["cirurgiao"]).values

        ranking["Dur_media"] = ranking["Dur_media"].fillna(0).astype(int)
        for dc in ["Primeiro","Ultimo"]:
            if dc in ranking.columns:
                ranking[dc] = pd.to_datetime(ranking[dc], errors="coerce")\
                    .dt.strftime("%d/%m/%Y")

        st.dataframe(ranking.rename(columns={
            "cirurgiao":"Cirurgião","Total":"Total","Pacientes":"Pacientes únicos",
            "Dur_media":"Duração média (min)","Primeiro":"Primeiro proc.",
            "Ultimo":"Último proc."
        }), use_container_width=True, height=350)

        st.divider()
        st.markdown("#### Evolução de um cirurgião")
        cir_sel = st.selectbox("Selecione o cirurgião",
                               sorted(df["cirurgiao"].dropna().unique()))
        df_cir = df[df["cirurgiao"]==cir_sel].copy()
        if PLOTLY and not df_cir.empty and "data_inicio" in df_cir.columns:
            df_cir["mes"] = df_cir["data_inicio"].dt.to_period("M").astype(str)
            mensal_c = df_cir.groupby("mes").size().reset_index(name="Total")
            fig_c = px.bar(mensal_c, x="mes", y="Total",
                           title=f"Procedimentos mensais — {cir_sel.split()[0]}",
                           color_discrete_sequence=["#185FA5"])
            fig_c.update_xaxes(tickangle=45)
            fig_c.update_layout(height=280, margin=dict(t=40,b=40))
            st.plotly_chart(fig_c, use_container_width=True)

# ════════════════════════════════════════
# ABA 4 — PACIENTES
# ════════════════════════════════════════
with aba4:
    st.markdown("#### Busca de paciente")
    st.caption("Digite nome, prontuário ou CPF para ver todos os procedimentos realizados.")

    termo = st.text_input("🔍 Nome, prontuário ou CPF",
                          placeholder="Ex: SEVERINO ou 2578507 ou 63867451753",
                          key="busca_pac")

    if termo:
        mask = pd.Series([False] * len(df_total))
        for col in ["nome","prontuario","cpf"]:
            if col in df_total.columns:
                mask |= df_total[col].astype(str).str.upper().str.contains(
                    termo.upper(), na=False)
        resultado = df_total[mask].copy()

        if resultado.empty:
            st.warning(f"Nenhum paciente encontrado para '{termo}'.")
        else:
            # Agrupa por CPF
            cpfs = resultado["cpf"].dropna().unique() if "cpf" in resultado.columns \
                   else resultado["prontuario"].unique()

            for cpf in cpfs:
                if "cpf" in resultado.columns:
                    pac = resultado[resultado["cpf"]==cpf]
                else:
                    pac = resultado[resultado["prontuario"]==cpf]

                if pac.empty:
                    continue

                linha = pac.iloc[0]
                nome  = linha.get("nome","—")
                pront = linha.get("prontuario","—")
                cpf_f = str(cpf) if cpf else "—"
                nasc  = pd.to_datetime(linha.get("dt_nascimento"), errors="coerce")
                nasc_s = nasc.strftime("%d/%m/%Y") if pd.notna(nasc) else "—"
                idade = linha.get("idade","—")
                sexo  = linha.get("sexo","—")
                mun   = linha.get("municipio","—")

                ficha = (
                    f"Nome      : {nome}\n"
                    f"Prontuário: {pront}   CPF: {cpf_f}\n"
                    f"Nasc.     : {nasc_s}   Idade: {idade}   Sexo: {sexo}\n"
                    f"Município : {mun}\n"
                    f"{'─'*52}\n"
                    f"PROCEDIMENTOS REALIZADOS ({len(pac)})\n"
                    f"{'─'*52}\n"
                )
                for i, (_, row) in enumerate(pac.iterrows(), 1):
                    di = pd.to_datetime(row.get("data_inicio"), errors="coerce")
                    di_s = di.strftime("%d/%m/%Y") if pd.notna(di) else "—"
                    dur  = row.get("duracao_minutos","—")
                    lado = row.get("lateralidade","—")
                    cir  = row.get("cirurgiao","—")
                    tp   = row.get("tipo_procedimento","—")
                    int_ = pd.to_datetime(row.get("data_atendimento"), errors="coerce")
                    int_s = int_.strftime("%d/%m/%Y") if pd.notna(int_) else "—"
                    alt_ = pd.to_datetime(row.get("data_alta"), errors="coerce")
                    alt_s = alt_.strftime("%d/%m/%Y") if pd.notna(alt_) else "—"
                    ficha += (
                        f"\n{i}. {tp}\n"
                        f"   Data: {di_s}  Lado: {lado}  Duração: {dur} min\n"
                        f"   Cirurgião: {cir}\n"
                        f"   Internação: {int_s}   Alta: {alt_s}\n"
                    )

                with st.expander(f"📋 {nome} — Prontuário {pront} — {len(pac)} procedimento(s)", expanded=True):
                    st.markdown(f'<div class="ficha-box">{ficha}</div>',
                                unsafe_allow_html=True)
                    st.download_button(
                        "⬇️ Exportar ficha (Excel)",
                        data=df_para_excel(pac),
                        file_name=f"ficha_{pront}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"exp_{cpf}"
                    )
    else:
        # Pacientes com múltiplos procedimentos
        st.markdown("#### Pacientes com múltiplos procedimentos")
        if "cpf" in df_total.columns:
            multi = df_total.groupby(["cpf","nome"]).agg(
                Total=("prontuario","count"),
                Tipos=("tipo_procedimento", lambda x: " | ".join(sorted(x.dropna().unique()))),
                Prontuario=("prontuario","first")
            ).reset_index()
            multi = multi[multi["Total"] > 1].sort_values("Total", ascending=False)
            st.dataframe(multi.rename(columns={
                "cpf":"CPF","nome":"Nome","Total":"Procedimentos",
                "Tipos":"Tipos realizados","Prontuario":"Prontuário"
            }), use_container_width=True, height=380)
            st.caption(f"{len(multi)} pacientes com mais de 1 procedimento registrado.")

# ════════════════════════════════════════
# ABA 5 — PESQUISA CLÍNICA
# ════════════════════════════════════════
with aba5:
    st.markdown("#### Planilha de Pesquisa Clínica")
    st.caption("Campos **IMC** e **Tab** são editáveis. Clique em 'Salvar' após editar.")

    df_pesq = carregar_pesquisa()

    if df_pesq.empty:
        st.info("Planilha de pesquisa não encontrada no banco.\n"
                "Use a opção [7] do menu para gerar a planilha.")
    else:
        colunas_edit = {}
        for col in df_pesq.columns:
            if col in ["IMC","Tab","imc","tab"]:
                colunas_edit[col] = st.column_config.TextColumn(col, width="small")
            elif col in ["ID","id"]:
                colunas_edit[col] = st.column_config.TextColumn(col, width="small")
            else:
                colunas_edit[col] = st.column_config.TextColumn(col, disabled=True)

        df_editado = st.data_editor(
            df_pesq,
            column_config=colunas_edit,
            use_container_width=True,
            height=420,
            key="editor_pesquisa"
        )

        st.info("⚠️ **Atenção:** alterações salvas aqui são temporárias na nuvem (Streamlit Cloud). Para persistência permanente, exporte em Excel e reimporte localmente.", icon="💡")

        c1, c2 = st.columns(2)
        with c1:
            if st.button("💾 Salvar alterações", type="primary"):
                try:
                    salvar_pesquisa(df_editado)
                    st.success("Alterações salvas com sucesso!")
                    st.cache_data.clear()
                except Exception as e:
                    st.error(f"Erro ao salvar: {e}")
        with c2:
            st.download_button(
                "⬇️ Exportar planilha (Excel)",
                data=df_para_excel(df_editado),
                file_name=f"pesquisa_clinica_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
