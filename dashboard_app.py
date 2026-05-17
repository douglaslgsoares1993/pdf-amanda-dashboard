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
BANCO_PATH   = PASTA_SCRIPT / "SAIDA" / "procedimentos.db"
PESQUISA_DB  = PASTA_SCRIPT / "SAIDA" / "pesquisa_clinica.db"

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

# ── CSS customizado ─────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stMetricValue"] { font-size: 2rem !important; font-weight: 600; }
[data-testid="stMetricLabel"] { font-size: 0.85rem !important; color: #555; }
.ficha-box {
    background: #f0f7ff;
    border-left: 4px solid #185FA5;
    border-radius: 8px;
    padding: 1rem 1.25rem;
    margin-bottom: 1rem;
    font-family: monospace;
    font-size: 0.85rem;
    white-space: pre-wrap;
}
.badge-d  { background:#E6F1FB; color:#0C447C; padding:2px 8px;
            border-radius:4px; font-size:12px; font-weight:600; }
.badge-e  { background:#EAF3DE; color:#27500A; padding:2px 8px;
            border-radius:4px; font-size:12px; font-weight:600; }
.badge-na { background:#F1EFE8; color:#444; padding:2px 8px;
            border-radius:4px; font-size:12px; font-weight:600; }
</style>
""", unsafe_allow_html=True)

# ── Carrega dados ───────────────────────────────────────────────
df_total = carregar_dados()
banco_ok  = not df_total.empty

# ── Sidebar ─────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🏥 PDF Amanda")
    st.markdown("**Dashboard de Produção Cirúrgica**")
    st.markdown("HUPE / UERJ")
    st.divider()

    if not banco_ok:
        st.error("Banco de dados não encontrado.\nProcesse os PDFs primeiro\n(opção 1 do menu).")
        st.stop()

    st.markdown(f"**Base:** {len(df_total):,} registros")
    if "data_inicio" in df_total.columns:
        ultima = df_total["data_inicio"].max()
        if pd.notna(ultima):
            st.markdown(f"**Último proc.:** {ultima.strftime('%d/%m/%Y')}")
    st.divider()

    st.markdown("### Filtros")

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
    col1, col2, col3, col4 = st.columns(4)
    total = len(df)
    unicos = df["cpf"].nunique() if "cpf" in df.columns else 0
    cirurg = df["cirurgiao"].nunique() if "cirurgiao" in df.columns else 0
    dur_med = int(df["duracao_minutos"].mean()) \
              if "duracao_minutos" in df.columns and df["duracao_minutos"].notna().any() else 0

    col1.metric("Total de procedimentos", f"{total:,}")
    col2.metric("Pacientes únicos (CPF)", f"{unicos:,}")
    col3.metric("Cirurgiões distintos", str(cirurg))
    col4.metric("Duração média", f"{dur_med} min")

    if PLOTLY and total > 0:
        st.divider()
        c1, c2 = st.columns(2)

        # Barras por tipo
        with c1:
            st.markdown("##### Procedimentos por tipo")
            if "tipo_procedimento" in df.columns:
                ct = df["tipo_procedimento"].value_counts().reset_index()
                ct.columns = ["Tipo", "Total"]
                fig = px.bar(ct, x="Tipo", y="Total",
                             color="Tipo",
                             color_discrete_sequence=["#378ADD","#1D9E75","#D85A30"])
                fig.update_layout(showlegend=False, margin=dict(t=10,b=10),
                                  height=280)
                st.plotly_chart(fig, use_container_width=True)

        # Rosca lateralidade
        with c2:
            st.markdown("##### Lateralidade")
            if "lateralidade" in df.columns:
                cl = df["lateralidade"].value_counts().reset_index()
                cl.columns = ["Lado", "Total"]
                fig2 = px.pie(cl, names="Lado", values="Total", hole=0.45,
                              color_discrete_sequence=["#378ADD","#1D9E75","#888780","#D85A30"])
                fig2.update_layout(margin=dict(t=10,b=10), height=280)
                st.plotly_chart(fig2, use_container_width=True)

        # Linha mensal
        st.markdown("##### Evolução mensal por tipo de procedimento")
        if "data_inicio" in df.columns and "tipo_procedimento" in df.columns:
            df_m = df.dropna(subset=["data_inicio"]).copy()
            df_m["mes"] = df_m["data_inicio"].dt.to_period("M").astype(str)
            mensal = df_m.groupby(["mes","tipo_procedimento"]).size().reset_index(name="Total")
            fig3 = px.line(mensal, x="mes", y="Total", color="tipo_procedimento",
                           color_discrete_sequence=["#378ADD","#1D9E75","#D85A30"],
                           markers=True)
            fig3.update_layout(margin=dict(t=10,b=40), height=300,
                               xaxis_title="", yaxis_title="Procedimentos",
                               legend_title="")
            fig3.update_xaxes(tickangle=45)
            st.plotly_chart(fig3, use_container_width=True)

        c3, c4 = st.columns(2)

        # Top cirurgiões
        with c3:
            st.markdown("##### Top 10 cirurgiões")
            if "cirurgiao" in df.columns:
                top_c = df["cirurgiao"].value_counts().head(10).reset_index()
                top_c.columns = ["Cirurgião", "Total"]
                top_c["Nome"] = top_c["Cirurgião"].str.split().str[:2].str.join(" ")
                fig4 = px.bar(top_c, x="Total", y="Nome", orientation="h",
                              color_discrete_sequence=["#534AB7"])
                fig4.update_layout(showlegend=False, margin=dict(t=10,b=10),
                                   height=320, yaxis_title="")
                st.plotly_chart(fig4, use_container_width=True)

        # Internação vs ambulatorial
        with c4:
            st.markdown("##### Tipo de atendimento")
            if "tipo_atendimento" in df.columns:
                ta = df["tipo_atendimento"].value_counts().reset_index()
                ta.columns = ["Tipo", "Total"]
                fig5 = px.pie(ta, names="Tipo", values="Total", hole=0.45,
                              color_discrete_sequence=["#534AB7","#BA7517"])
                fig5.update_layout(margin=dict(t=10,b=10), height=320)
                st.plotly_chart(fig5, use_container_width=True)

        # Top municípios
        st.markdown("##### Top 10 municípios de origem")
        if "municipio" in df.columns:
            top_m = df["municipio"].value_counts().head(10).reset_index()
            top_m.columns = ["Município", "Total"]
            fig6 = px.bar(top_m, x="Município", y="Total",
                          color_discrete_sequence=["#185FA5"])
            fig6.update_layout(showlegend=False, margin=dict(t=10,b=60),
                               height=300)
            fig6.update_xaxes(tickangle=35)
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
