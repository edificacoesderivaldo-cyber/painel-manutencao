import streamlit as st
import pandas as pd
from datetime import datetime

# ==============================================================================
# CONFIGURAÇÃO DA PÁGINA
# ==============================================================================
st.set_page_config(
    page_title="Gestão de Chamados | Equipe Interna & PROHAB",
    page_icon="🛠️",
    layout="wide"
)

st.title("🛠️ Gestão Operacional de Chamados Prediais")
st.markdown("Triagem e direcionamento de demandas: **Equipe Interna** versus **Contrato PROHAB** por unidade.")
st.divider()

# ==============================================================================
# ÁREA DE ARRASTAR E SOLTAR O ARQUIVO (UPLOAD DIRETO)
# ==============================================================================
st.sidebar.title("📥 Entrada de Dados")
arquivo_carregado = st.sidebar.file_uploader(
    "Arraste a exportação do SysAid (.csv)",
    type=["csv"],
    help="Exporte os chamados no SysAid em formato CSV e solte o arquivo aqui."
)

# ==============================================================================
# PROCESSAMENTO DOS DADOS
# ==============================================================================
@st.cache_data
def processar_base(arquivo):
    try:
        df = pd.read_csv(arquivo, encoding="latin1", sep=",")
    except Exception:
        try:
            df = pd.read_csv(arquivo, encoding="utf-8", sep=",")
        except Exception:
            return pd.DataFrame()

    # Tratamento de localidade
    df["Local_Tratado"] = df["Local"].fillna(df["Empresa"]).fillna("Não Informado")

    # Tratamento de datas
    fmt_data = "%d-%m-%Y %H:%M:%S"
    df["dt_solicitacao"] = pd.to_datetime(df["Hora da solicitação"], format=fmt_data, errors="coerce")
    df["dt_conclusao_prev"] = pd.to_datetime(df["Data de conclusão"], format=fmt_data, errors="coerce")

    hoje = datetime.now()
    df["Dias_Aberto"] = (hoje - df["dt_solicitacao"]).dt.days.fillna(0).astype(int)

    # Classificação de Frente: Equipe Interna vs. OS PROHAB
    def definir_tipo(linha):
        sub = str(linha.get("Subcategoria", "")).lower()
        terceiro = str(linha.get("Categoria de terceiro nível", "")).lower()
        tit = str(linha.get("Título", "")).lower()

        termos_pesados = [
            "reforma", "alvenaria", "piso", "calçada", "construção",
            "adequação", "projeto", "orçamento", "rufo", "passarela", "bomba"
        ]
        if any(t in sub or t in terceiro or t in tit for t in termos_pesados):
            return "OS Contratual (PROHAB)"
        
        if "telhado" in sub and any(t in tit for t in ["reforma", "troca", "vazamento"]):
            return "OS Contratual (PROHAB)"

        return "Equipe Interna (Própria)"

    df["Tipo_Atendimento"] = df.apply(definir_tipo, axis=1)

    # Farol de SLA
    def avaliar_sla(linha):
        if linha["Status"] == "Solicitação Atendida":
            return "Concluído"
        if pd.isna(linha["dt_conclusao_prev"]):
            return "Sem Prazo"
        if linha["dt_conclusao_prev"] < hoje:
            return "Vencido"
        return "No Prazo"

    df["Farol_SLA"] = df.apply(avaliar_sla, axis=1)
    return df

# Carga do ficheiro carregado ou do ficheiro padrão do repositório
df_chamados = pd.DataFrame()

if arquivo_carregado is not None:
    df_chamados = processar_base(arquivo_carregado)
else:
    try:
        df_chamados = processar_base("export (1).csv")
    except Exception:
        df_chamados = pd.DataFrame()

# Mensagem orientativa caso não haja ficheiro carregado
if df_chamados.empty:
    st.info("👈 Por favor, **arraste o arquivo CSV exportado do SysAid** no campo à esquerda para visualizar o painel.")
    st.stop()

# ==============================================================================
# FILTROS OPERACIONAIS NA LATERAL
# ==============================================================================
st.sidebar.divider()
st.sidebar.title("🔍 Filtros de Visualização")

filtro_atendimento = st.sidebar.radio(
    "Frente de Trabalho",
    ["Todas", "Equipe Interna (Própria)", "OS Contratual (PROHAB)"]
)

unidades = ["Todas"] + sorted(list(df_chamados["Local_Tratado"].dropna().unique()))
filtro_unidade = st.sidebar.selectbox("Unidade / Prédio", unidades)

disciplinas = ["Todas"] + sorted(list(df_chamados["Subcategoria"].dropna().unique()))
filtro_disciplina = st.sidebar.selectbox("Especialidade", disciplinas)

status_opcoes = ["Todos"] + sorted(list(df_chamados["Status"].dropna().unique()))
filtro_status = st.sidebar.selectbox("Status do Chamado", status_opcoes)

# Aplicação dos filtros
df_view = df_chamados.copy()

if filtro_atendimento != "Todas":
    df_view = df_view[df_view["Tipo_Atendimento"] == filtro_atendimento]
if filtro_unidade != "Todas":
    df_view = df_view[df_view["Local_Tratado"] == filtro_unidade]
if filtro_disciplina != "Todas":
    df_view = df_view[df_view["Subcategoria"] == filtro_disciplina]
if filtro_status != "Todos":
    df_view = df_view[df_view["Status"] == filtro_status]

# ==============================================================================
# CARTÕES DE INDICADORES (KPIS)
# ==============================================================================
tot = len(df_view)
internos = len(df_view[df_view["Tipo_Atendimento"] == "Equipe Interna (Própria)"])
prohab = len(df_view[df_view["Tipo_Atendimento"] == "OS Contratual (PROHAB)"])
vencidos = len(df_view[df_view["Farol_SLA"] == "Vencido"])

c1, c2, c3, c4 = st.columns(4)
c1.metric("Total de Chamados", f"{tot}")
c2.metric("Equipe Interna", f"{internos}", delta="Reparos Rápidos")
c3.metric("Frente PROHAB (OS)", f"{prohab}", delta="Manutenção Pesada", delta_color="normal")
c4.metric("SLA Vencido", f"{vencidos}", delta="Atenção", delta_color="inverse")

st.divider()

# ==============================================================================
# GRÁFICOS ANALÍTICOS (NATIVOS DO STREAMLIT)
# ==============================================================================
col_g1, col_g2 = st.columns(2)

with col_g1:
    st.subheader("📍 Fila por Unidade")
    dados_unidade = df_view["Local_Tratado"].value_counts().head(8)
    st.bar_chart(dados_unidade, horizontal=True)

with col_g2:
    st.subheader("🧱 Demandas por Especialidade")
    dados_disciplina = df_view["Subcategoria"].value_counts().head(8)
    st.bar_chart(dados_disciplina)

# ==============================================================================
# TABELA DE GESTÃO DIÁRIA
# ==============================================================================
st.subheader("📋 Relação de Chamados Filtrados")

colunas = [
    "#",
    "Tipo_Atendimento",
    "Local_Tratado",
    "Subcategoria",
    "Título",
    "Usuário Solicitante",
    "Dias_Aberto",
    "Farol_SLA",
    "Status"
]

tabela = df_view[colunas].sort_values("Dias_Aberto", ascending=False)
st.dataframe(tabela, use_container_width=True, height=450)

# Exportação dos dados filtrados
csv_baixar = tabela.to_csv(index=False).encode("utf-8")
st.download_button(
    label="📥 Baixar Dados Filtrados (CSV)",
    data=csv_baixar,
    file_name=f"chamados_filtrados_{datetime.now().strftime('%Y%m%d')}.csv",
    mime="text/csv"
)
