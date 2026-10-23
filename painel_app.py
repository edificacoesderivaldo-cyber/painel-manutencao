import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA STREAMLIT
# ==============================================================================
st.set_page_config(
    page_title="Engenharia Predial | Painel Integrado OS & Help Desk",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==============================================================================
# 2. FUNÇÕES DE CARGA E TRATAMENTO DE DADOS (ISOLADAS)
# ==============================================================================

# --- Módulo A: Base do Google Sheets (Suas OS Existentes) ---
@st.cache_data(ttl=600)
def carregar_dados_os():
    """
    Substitua esta função pela sua lógica atual de conexão com o Google Sheets
    (ex.: gspread ou url de exportação pública/privada da planilha).
    """
    # Exemplo ilustrativo de estrutura do seu DataFrame de OS existente:
    df_exemplo_os = pd.DataFrame({
        "Nº OS": ["OS-2026-001", "OS-2026-002", "OS-2026-003"],
        "Chamado_SysAid": ["198086", "197288", "197540"],
        "Unidade": ["SESI/SENAI SIA", "SESI/SENAI SIA", "SESI/SENAI SIA"],
        "Objeto": ["Recuperação de Forro 3º Andar", "Remanejamento Trocador de Calor", "Recuperação de Revestimento"],
        "Empresa_Contratada": ["Engenharia Alfa Ltda", "Construções Beta", "Engenharia Alfa Ltda"],
        "Valor_Total_SINAPI": [14500.00, 8900.50, 4200.00],
        "Status_OS": ["Em Execução", "Medição Aprovada", "Concluída"]
    })
    return df_exemplo_os


# --- Módulo B: Base do Help Desk SysAid (CSV) ---
@st.cache_data
def carregar_dados_sysaid(origem_arquivo):
    """
    Lê a exportação do SysAid em latin1/utf-8 e formata prazos, localidades e SLA.
    """
    try:
        df = pd.read_csv(origem_arquivo, encoding="latin1", sep=",")
    except Exception:
        df = pd.read_csv(origem_arquivo, encoding="utf-8", sep=",")

    # Unificação de Unidade/Local (preenche nulos de Local com Empresa)
    df["Local_Tratado"] = df["Local"].fillna(df["Empresa"]).fillna("Não Informado")

    # Conversão de datas padrão SysAid (DD-MM-YYYY HH:MM:SS)
    formato_data = "%d-%m-%Y %H:%M:%S"
    df["dt_solicitacao"] = pd.to_datetime(df["Hora da solicitação"], format=formato_data, errors="coerce")
    df["dt_conclusao_prev"] = pd.to_datetime(df["Data de conclusão"], format=formato_data, errors="coerce")
    df["dt_modificacao"] = pd.to_datetime(df["Hora de modificação"], format=formato_data, errors="coerce")

    agora = datetime.now()

    # Cálculo da idade do chamado (dias em aberto)
    df["Dias_Aberto"] = (agora - df["dt_solicitacao"]).dt.days.fillna(0).astype(int)

    # Classificação do Farol de SLA
    def definir_sla(linha):
        if linha["Status"] == "Solicitação Atendida":
            return "Concluído"
        if pd.isna(linha["dt_conclusao_prev"]):
            return "Sem Prazo"
        horas_restantes = (linha["dt_conclusao_prev"] - agora).total_seconds() / 3600
        if horas_restantes < 0:
            return "Vencido"
        elif horas_restantes <= 48:
            return "Crítico (Até 48h)"
        else:
            return "No Prazo"

    df["Farol_SLA"] = df.apply(definir_sla, axis=1)
    df["ID_Chamado"] = df["#"].astype(str)
    return df


# ==============================================================================
# 3. BARRA LATERAL (FILTROS E UPLOAD)
# ==============================================================================
st.sidebar.title("🏢 Gestão Predial")
st.sidebar.markdown("**SESI • SENAI • FIBRA**")
st.sidebar.divider()

# Upload opcional do Help Desk (se não houver arquivo local fixo)
arquivo_upload = st.sidebar.file_uploader(
    "Carregar exportação SysAid (.csv)",
    type=["csv"],
    help="Exporte o relatório do SysAid e arraste aqui para atualizar os chamados."
)

# Carregamento seguro da base do SysAid
try:
    if arquivo_upload is not None:
        df_helpdesk_raw = carregar_dados_sysaid(arquivo_upload)
    else:
        df_helpdesk_raw = carregar_dados_sysaid("export (1).csv")
except Exception:
    df_helpdesk_raw = pd.DataFrame()

# Carregamento das OS do Google Sheets
df_os_raw = carregar_dados_os()

# Filtros para o Help Desk na Barra Lateral
st.sidebar.subheader("Filtros do Help Desk")
if not df_helpdesk_raw.empty:
    filtro_unidade = st.sidebar.selectbox(
        "Unidade / Prédio",
        ["Todas"] + sorted(list(df_helpdesk_raw["Local_Tratado"].unique()))
    )
    filtro_status = st.sidebar.selectbox(
        "Status do Chamado",
        ["Todos"] + sorted(list(df_helpdesk_raw["Status"].unique()))
    )
    filtro_sla = st.sidebar.selectbox(
        "Situação do SLA",
        ["Todos", "Vencido", "Crítico (Até 48h)", "No Prazo", "Concluído"]
    )
else:
    filtro_unidade, filtro_status, filtro_sla = "Todas", "Todos", "Todos"

st.sidebar.divider()
st.sidebar.caption("Engenharia de Manutenção Predial | v2.0")


# ==============================================================================
# 4. NAVEGAÇÃO EM ABAS PRINCIPAIS
# ==============================================================================
tab_os, tab_helpdesk, tab_cruzamento = st.tabs([
    "📋 Controle Financeiro de OS (Google Sheets)",
    "🛠️ Gestão de Chamados Help Desk (SysAid)",
    "🔗 Vínculo Chamado ↔ Ordem de Serviço"
])


# ------------------------------------------------------------------------------
# ABA 1: CONTROLE DE ORDENS DE SERVIÇO
# ------------------------------------------------------------------------------
with tab_os:
    st.header("Controle de Execução Contratual e Ordens de Serviço")
    st.markdown("Acompanhamento das medições de contratos corporativos, saldos e empenhos.")
    
    # KPIs da base de OS
    col_os1, col_os2, col_os3 = st.columns(3)
    total_os = len(df_os_raw)
    valor_total_os = df_os_raw["Valor_Total_SINAPI"].sum()
    col_os1.metric("Ordens de Serviço Emitidas", f"{total_os}")
    col_os2.metric("Montante Alocado (R$)", f"R$ {valor_total_os:,.2f}")
    col_os3.metric("OS em Execução", len(df_os_raw[df_os_raw["Status_OS"] == "Em Execução"]))
    
    st.divider()
    st.subheader("Relação de Ordens de Serviço Ativas")
    st.dataframe(df_os_raw, use_container_width=True)


# ------------------------------------------------------------------------------
# ABA 2: HELP DESK SYSAID
# ------------------------------------------------------------------------------
with tab_helpdesk:
    st.header("Central de Chamados Operacionais | SysAid")
    st.markdown("Triagem, classificação por especialidade e monitoramento de tempo de fila.")

    if df_helpdesk_raw.empty:
        st.warning("Nenhum dado do SysAid carregado. Por favor, adicione o arquivo 'export (1).csv' ou faça o upload na barra lateral.")
    else:
        # Aplicando filtros
        df_hd = df_helpdesk_raw.copy()
        if filtro_unidade != "Todas":
            df_hd = df_hd[df_hd["Local_Tratado"] == filtro_unidade]
        if filtro_status != "Todos":
            df_hd = df_hd[df_hd["Status"] == filtro_status]
        if filtro_sla != "Todos":
            df_hd = df_hd[df_hd["Farol_SLA"] == filtro_sla]

        # KPIs Executivos do Help Desk
        total_hd = len(df_hd)
        ativos_hd = len(df_hd[df_hd["Status"] != "Solicitação Atendida"])
        vencidos_hd = len(df_hd[df_hd["Farol_SLA"] == "Vencido"])
        atendidos_hd = len(df_hd[df_hd["Status"] == "Solicitação Atendida"])

        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric("Total no Filtro", f"{total_hd}")
        kpi2.metric("Chamados Ativos (Fila)", f"{ativos_hd}", delta=f"{(ativos_hd/total_hd*100):.1f}% do total" if total_hd else "0%")
        kpi3.metric("SLA Vencido", f"{vencidos_hd}", delta="- Repactuar Prazo" if vencidos_hd > 0 else "Regular", delta_color="inverse")
        kpi4.metric("Solicitações Atendidas", f"{atendidos_hd}")

        st.divider()

        # Gráficos Analíticos (Linha dupla)
        g_col1, g_col2 = st.columns([1.1, 0.9])

        with g_col1:
            st.subheader("Fila Ativa por Unidade Operacional")
            df_unid_plot = df_hd[df_hd["Status"] != "Solicitação Atendida"]["Local_Tratado"].value_counts().reset_index()
            df_unid_plot.columns = ["Unidade", "Chamados"]
            
            fig_bar = px.bar(
                df_unid_plot.head(8).sort_values("Chamados", ascending=True),
                x="Chamados",
                y="Unidade",
                orientation="h",
                text="Chamados",
                color="Chamados",
                color_continuous_scale="Blues"
            )
            fig_bar.update_layout(yaxis={"categoryorder": "total ascending"}, showlegend=False, height=360, margin=dict(l=20, r=20, t=30, b=20))
            st.plotly_chart(fig_bar, use_container_width=True)

        with g_col2:
            st.subheader("Demandas por Especialidade Técnica")
            df_sub_plot = df_hd[df_hd["Status"] != "Solicitação Atendida"]["Subcategoria"].value_counts().reset_index()
            df_sub_plot.columns = ["Especialidade", "Total"]
            
            fig_pie = px.pie(
                df_sub_plot.head(6),
                names="Especialidade",
                values="Total",
                hole=0.45,
                color_discrete_sequence=px.colors.qualitative.Safe
            )
            fig_pie.update_layout(height=360, margin=dict(l=20, r=20, t=30, b=20))
            st.plotly_chart(fig_pie, use_container_width=True)

        # Fila de Triagem Operacional
        st.subheader("Fila de Chamados e Ações Recomendadas")
        
        # Mapeamento de ação de engenharia
        def sugerir_acao(subcategoria):
            sub = str(subcategoria).lower()
            if any(termo in sub for termo in ["reforma", "alvenaria", "piso", "calçada", "forro"]):
                return "Emitir OS SINAPI"
            elif any(termo in sub for termo in ["elétrica", "lâmpada", "disjuntor", "curto"]):
                return "Equipe Volante Elétrica"
            elif any(termo in sub for termo in ["esquadrias", "porta", "janela", "vidro"]):
                return "Serralheria / Vidraçaria"
            elif any(termo in sub for termo in ["telhado", "vazamento", "calha"]):
                return "Reparo Cobertura"
            return "Vistoria Técnica"

        df_hd["Acao_Recomendada"] = df_hd["Subcategoria"].apply(sugerir_acao)

        colunas_exibir = [
            "ID_Chamado", "Local_Tratado", "Subcategoria", "Título",
            "Usuário Solicitante", "Hora da solicitação", "Farol_SLA",
            "Dias_Aberto", "Acao_Recomendada"
        ]

        df_tabela = df_hd[colunas_exibir].sort_values("Dias_Aberto", ascending=False)
        st.dataframe(df_tabela, use_container_width=True, height=400)


# ------------------------------------------------------------------------------
# ABA 3: CRUZAMENTO DIRETO (CHAMADO SYSAID ↔ ORDEM DE SERVIÇO)
# ------------------------------------------------------------------------------
with tab_cruzamento:
    st.header("Rastreabilidade: Do Chamado ao Pagamento Contratual")
    st.markdown("Identifique quais chamados abertos no SysAid já viraram Ordens de Serviço com medição orçamentária.")

    if not df_helpdesk_raw.empty and not df_os_raw.empty:
        # Cruzamento (Merge) entre as duas bases pela chave do Chamado
        df_integrado = pd.merge(
            df_helpdesk_raw[["ID_Chamado", "Local_Tratado", "Subcategoria", "Título", "Hora da solicitação", "Status"]],
            df_os_raw,
            left_on="ID_Chamado",
            right_on="Chamado_SysAid",
            how="left"
        )

        # Filtro de chamados com ou sem OS
        status_vinculo = st.radio(
            "Visualizar:",
            ["Todos os Chamados", "Chamados com OS Gerada", "Chamados Aguardando Emissão de OS"],
            horizontal=True
        )

        if status_vinculo == "Chamados com OS Gerada":
            df_integrado = df_integrado[df_integrado["Nº OS"].notna()]
        elif status_vinculo == "Chamados Aguardando Emissão de OS":
            df_integrado = df_integrado[df_integrado["Nº OS"].isna()]

        st.dataframe(
            df_integrado[[
                "ID_Chamado", "Local_Tratado", "Subcategoria", "Título",
                "Status", "Nº OS", "Empresa_Contratada", "Valor_Total_SINAPI", "Status_OS"
            ]],
            use_container_width=True
        )
    else:
        st.info("Carregue ambas as bases para visualizar o cruzamento de dados.")
