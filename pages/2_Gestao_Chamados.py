import streamlit as st
import pandas as pd
from datetime import datetime

# ==============================================================================
# CONFIGURAÇÃO DA PÁGINA E ESTILOS DE IMPRESSÃO
# ==============================================================================
st.set_page_config(
    page_title="Gestão de Chamados | Equipe Interna & PROHAB",
    page_icon="🛠️",
    layout="wide"
)

# CSS para otimização da impressão direta na página (Ctrl + P)
st.markdown("""
<style>
@media print {
    section[data-testid="stSidebar"],
    header[data-testid="stHeader"],
    footer,
    .stDownloadButton,
    button {
        display: none !important;
    }
    .main .block-container {
        max-width: 100% !important;
        padding: 0 !important;
        margin: 0 !important;
    }
    body {
        background-color: white !important;
        color: black !important;
    }
}
</style>
""", unsafe_allow_html=True)

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

# Carga dos dados
df_chamados = pd.DataFrame()

if arquivo_carregado is not None:
    df_chamados = processar_base(arquivo_carregado)
else:
    try:
        df_chamados = processar_base("export (1).csv")
    except Exception:
        df_chamados = pd.DataFrame()

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

# Aplicando os filtros
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
# GRÁFICOS ANALÍTICOS
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

# ==============================================================================
# GERADOR DO RELATÓRIO EXECUTIVO (IMPRESSÃO / SALVAR EM PDF)
# ==============================================================================
def gerar_relatorio_impressao(df_dados, filtros_info):
    linhas_tabela = ""
    for _, row in df_dados.iterrows():
        sla_val = str(row.get("Farol_SLA", ""))
        cor_sla = "#b02a37" if sla_val == "Vencido" else ("#198754" if sla_val == "Concluído" else "#0d6efd")
        fundo_sla = "#f8d7da" if sla_val == "Vencido" else ("#d1e7dd" if sla_val == "Concluído" else "#cfe2ff")

        linhas_tabela += f"""
        <tr>
            <td style="text-align: center; font-weight: bold; color: #0d6efd;">{row.get('#', '')}</td>
            <td>{row.get('Tipo_Atendimento', '')}</td>
            <td>{row.get('Local_Tratado', '')}</td>
            <td>{row.get('Subcategoria', '')}</td>
            <td>{row.get('Título', '')}</td>
            <td>{row.get('Usuário Solicitante', '')}</td>
            <td style="text-align: center;">{row.get('Dias_Aberto', '')} d</td>
            <td style="text-align: center;"><span style="background-color: {fundo_sla}; color: {cor_sla}; padding: 3px 8px; border-radius: 4px; font-weight: bold; font-size: 11px;">{sla_val}</span></td>
            <td style="text-align: center;">{row.get('Status', '')}</td>
        </tr>
        """

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <title>Relatório de Chamados Prediais</title>
    <style>
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif;
            margin: 20px;
            color: #212529;
            background-color: #ffffff;
        }}
        .no-print {{
            margin-bottom: 20px;
            padding: 12px 18px;
            background: #e7f1ff;
            border: 1px solid #b6d4fe;
            border-radius: 6px;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }}
        .btn-imprimir {{
            background-color: #0d6efd;
            color: white;
            border: none;
            padding: 10px 22px;
            font-size: 14px;
            font-weight: bold;
            border-radius: 5px;
            cursor: pointer;
        }}
        .btn-imprimir:hover {{
            background-color: #0b5ed7;
        }}
        .cabecalho {{
            border-bottom: 3px solid #0d6efd;
            padding-bottom: 8px;
            margin-bottom: 12px;
            display: flex;
            justify-content: space-between;
            align-items: flex-end;
        }}
        .cabecalho h1 {{
            font-size: 19px;
            margin: 0;
            color: #0d6efd;
            text-transform: uppercase;
        }}
        .cabecalho p {{
            font-size: 12px;
            color: #6c757d;
            margin: 4px 0 0 0;
        }}
        .meta-box {{
            background-color: #f8f9fa;
            border: 1px solid #dee2e6;
            border-radius: 6px;
            padding: 10px 14px;
            margin-bottom: 14px;
            font-size: 12px;
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 10px;
        }}
        .meta-item strong {{
            color: #0d6efd;
            display: block;
            margin-bottom: 2px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 11px;
        }}
        th {{
            background-color: #0d6efd;
            color: white;
            text-align: left;
            padding: 7px 6px;
            font-weight: 600;
            border: 1px solid #0d6efd;
        }}
        td {{
            padding: 6px;
            border: 1px solid #dee2e6;
            vertical-align: middle;
        }}
        tr:nth-child(even) {{
            background-color: #f8f9fa;
        }}
        .rodape {{
            margin-top: 20px;
            font-size: 11px;
            color: #6c757d;
            text-align: right;
            border-top: 1px solid #dee2e6;
            padding-top: 6px;
        }}
        @media print {{
            .no-print {{
                display: none !important;
            }}
            body {{
                margin: 0;
                padding: 10px;
            }}
            @page {{
                size: landscape;
                margin: 8mm;
            }}
            th {{
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }}
        }}
    </style>
</head>
<body>
    <div class="no-print">
        <div>
            <strong style="color: #084298;">Relatório Gerencial de Chamados</strong>
            <p style="margin: 2px 0 0 0; font-size: 12px; color: #055160;">Clique no botão para imprimir ou selecione <b>"Salvar como PDF"</b> na tela de impressão.</p>
        </div>
        <button class="btn-imprimir" onclick="window.print()">🖨️ Imprimir / Salvar em PDF</button>
    </div>

    <div class="cabecalho">
        <div>
            <h1>Sistema Integrado de Engenharia Predial</h1>
            <p>SESI • SENAI • FIBRA | Relatório Operacional de Chamados</p>
        </div>
        <div style="text-align: right; font-size: 12px;">
            <b>Emissão:</b> {filtros_info.get('data_hora', '')}<br>
            <b>Total Filtrado:</b> {filtros_info.get('total', 0)} chamados
        </div>
    </div>

    <div class="meta-box">
        <div class="meta-item"><strong>Frente Operacional:</strong> {filtros_info.get('frente', 'Todas')}</div>
        <div class="meta-item"><strong>Unidade / Prédio:</strong> {filtros_info.get('unidade', 'Todas')}</div>
        <div class="meta-item"><strong>Especialidade:</strong> {filtros_info.get('especialidade', 'Todas')}</div>
        <div class="meta-item"><strong>Status:</strong> {filtros_info.get('status', 'Todos')}</div>
    </div>

    <table>
        <thead>
            <tr>
                <th style="width: 60px; text-align: center;">Nº Chamado</th>
                <th style="width: 120px;">Frente</th>
                <th style="width: 100px;">Unidade</th>
                <th style="width: 110px;">Especialidade</th>
                <th>Título / Descrição da Demanda</th>
                <th style="width: 130px;">Solicitante</th>
                <th style="width: 60px; text-align: center;">Idade</th>
                <th style="width: 75px; text-align: center;">Farol SLA</th>
                <th style="width: 85px; text-align: center;">Status</th>
            </tr>
        </thead>
        <tbody>
            {linhas_tabela}
        </tbody>
    </table>

    <div class="rodape">
        Relatório emitido pela Central de Gestão Predial SESI/SENAI/FIBRA.
    </div>
</body>
</html>
"""
    return html

# Metadados dos filtros ativos na tela
info_filtros = {
    "frente": filtro_atendimento,
    "unidade": filtro_unidade,
    "especialidade": filtro_disciplina,
    "status": filtro_status,
    "total": len(tabela),
    "data_hora": datetime.now().strftime("%d/%m/%Y %H:%M")
}

# ==============================================================================
# BOTÕES DE EXPORTAÇÃO
# ==============================================================================
col_btn1, col_btn2 = st.columns(2)

with col_btn1:
    csv_baixar = tabela.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Baixar Dados Filtrados (CSV)",
        data=csv_baixar,
        file_name=f"chamados_filtrados_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
        mime="text/csv",
        use_container_width=True
    )

with col_btn2:
    html_relatorio = gerar_relatorio_impressao(tabela, info_filtros)
    st.download_button(
        label="📄 Baixar Relatório (Imprimir / Salvar em PDF)",
        data=html_relatorio.encode("utf-8"),
        file_name=f"relatorio_chamados_{datetime.now().strftime('%Y%m%d_%H%M')}.html",
        mime="text/html",
        use_container_width=True
    )
