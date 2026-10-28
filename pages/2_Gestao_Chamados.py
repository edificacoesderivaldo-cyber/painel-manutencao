import html as html_lib
import os

import streamlit as st
import pandas as pd
from datetime import datetime

# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA E ESTILOS VISUAIS
# ==============================================================================
st.set_page_config(
    page_title="Monitoramento de Chamados | Engenharia Predial",
    page_icon="🏢",
    layout="wide"
)

# Estilização CSS para modernização dos cartões, botões e impressão direta
st.markdown("""
<style>
    /* Estilo do título e cabeçalho */
    .header-box {
        padding: 12px 18px;
        background: linear-gradient(90deg, #0d47a1 0%, #1976d2 100%);
        color: white;
        border-radius: 8px;
        margin-bottom: 18px;
    }
    .header-box h2 {
        margin: 0;
        font-size: 22px;
        font-weight: 700;
        color: #ffffff;
    }
    .header-box p {
        margin: 4px 0 0 0;
        font-size: 13px;
        color: #e3f2fd;
    }

    /* Otimização para impressão via navegador */
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
    }
</style>
""", unsafe_allow_html=True)

# Cabeçalho institucional moderno
st.markdown("""
<div class="header-box">
    <h2>🏢 Central de Monitoramento de Chamados Prediais</h2>
    <p>SESI • SENAI • FIBRA | Direcionamento Operacional: Equipe Própria & Contrato de Manutenção PROHAB</p>
</div>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. CARREGAMENTO E PROCESSAMENTO DA BASE SYSAID
# ==============================================================================
st.sidebar.title("📥 Base de Dados")
arquivo_carregado = st.sidebar.file_uploader(
    "Arraste a exportação do SysAid (.csv)",
    type=["csv"],
    help="Exporte o relatório no SysAid em formato CSV e solte aqui para atualizar."
)

ARQUIVO_PADRAO = "export (1).csv"
STATUS_CONCLUIDO = "Solicitação Atendida"
FORMATO_DATA = "%d-%m-%Y %H:%M:%S"
COLUNAS_OBRIGATORIAS = [
    "#", "Título", "Usuário Solicitante", "Subcategoria", "Status",
    "Hora da solicitação", "Data de conclusão",
]


@st.cache_data
def processar_base(arquivo):
    """Lê o CSV e faz a classificação que NÃO depende da data atual (pode ser cacheada)."""
    df = None
    for encoding in ("utf-8-sig", "latin1"):
        try:
            if hasattr(arquivo, "seek"):
                arquivo.seek(0)
            df = pd.read_csv(arquivo, encoding=encoding, sep=",")
            break
        except Exception:
            continue
    if df is None:
        raise ValueError("Não foi possível ler o CSV. Verifique se o arquivo é a exportação do SysAid (separador vírgula).")

    faltando = [c for c in COLUNAS_OBRIGATORIAS if c not in df.columns]
    if faltando:
        raise ValueError("Colunas obrigatórias não encontradas no CSV: " + ", ".join(faltando))

    # Localidade: usa "Local" e cai para "Empresa" quando vazio
    local = df["Local"] if "Local" in df.columns else pd.Series(index=df.index, dtype=object)
    empresa = df["Empresa"] if "Empresa" in df.columns else pd.Series(index=df.index, dtype=object)
    df["Local_Tratado"] = local.fillna(empresa).fillna("Não Informado")

    # Conversão de datas
    df["dt_solicitacao"] = pd.to_datetime(df["Hora da solicitação"], format=FORMATO_DATA, errors="coerce")
    df["dt_conclusao_prev"] = pd.to_datetime(df["Data de conclusão"], format=FORMATO_DATA, errors="coerce")

    # Classificação operacional da demanda
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
    return df


def calcular_indicadores(df):
    """Dias em aberto e farol de SLA dependem da data atual: calculados a cada execução, sem cache."""
    hoje = datetime.now()
    df = df.copy()
    df["Dias_Aberto"] = (hoje - df["dt_solicitacao"]).dt.days.fillna(0).astype(int)

    def avaliar_sla(linha):
        if linha.get("Status") == STATUS_CONCLUIDO:
            return "Concluído"
        if pd.isna(linha["dt_conclusao_prev"]):
            return "Sem Prazo"
        if linha["dt_conclusao_prev"] < hoje:
            return "Vencido"
        return "No Prazo"

    df["Farol_SLA"] = df.apply(avaliar_sla, axis=1) if not df.empty else pd.Series(dtype=object)
    return df


# Carga de dados
df_chamados = pd.DataFrame()
erro_carga = None
try:
    if arquivo_carregado is not None:
        df_chamados = processar_base(arquivo_carregado)
    elif os.path.exists(ARQUIVO_PADRAO):
        df_chamados = processar_base(ARQUIVO_PADRAO)
except ValueError as e:
    erro_carga = str(e)

if erro_carga:
    st.error(f"❌ {erro_carga}")
    st.stop()

if df_chamados.empty:
    st.info("👈 Por favor, **arraste o arquivo CSV exportado do SysAid** no campo à esquerda para visualizar o painel.")
    st.stop()

df_chamados = calcular_indicadores(df_chamados)

# ==============================================================================
# 3. HELPER PARA BOTÕES DE FILTRO DINÂMICO (PILLS / SEGMENTED CONTROL)
# ==============================================================================
def criar_botoes_filtro(label, opcoes, default, key):
    """Cria botões de seleção tipo pills nativos do Streamlit com fallback seguro."""
    if hasattr(st, "pills"):
        valor = st.pills(label, opcoes, default=default, key=key)
        return valor if valor is not None else default
    elif hasattr(st, "segmented_control"):
        valor = st.segmented_control(label, opcoes, default=default, key=key)
        return valor if valor is not None else default
    else:
        idx = opcoes.index(default) if default in opcoes else 0
        return st.radio(label, opcoes, index=idx, horizontal=True, key=key)

# ==============================================================================
# 4. PAINEL DE FILTRAGEM RÁPIDA EM BOTÕES (CORPO DA PÁGINA)
# ==============================================================================
with st.container(border=True):
    st.markdown("#### 🎯 Filtros Rápidos de Operação")
    
    col_f1, col_f2 = st.columns([1.2, 1])
    with col_f1:
        frentes = ["Todas", "Equipe Interna (Própria)", "OS Contratual (PROHAB)"]
        filtro_atendimento = criar_botoes_filtro("Frente de Trabalho:", frentes, "Todas", "f_atend")
    with col_f2:
        slas = ["Todos", "Vencido", "No Prazo", "Concluído"]
        filtro_sla = criar_botoes_filtro("Farol de Prazo (SLA):", slas, "Todos", "f_sla")

    # Linha de botões de Status
    status_principais = ["Todos", "Em Atendimento", "Aberto", "Aguardando Aprovação", "Solicitação Atendida"]
    filtro_status = criar_botoes_filtro("Status do Chamado:", status_principais, "Todos", "f_status")

    # Linha de botões das Unidades Operacionais
    unidades_base = ["Todas"] + sorted([u for u in df_chamados["Local_Tratado"].dropna().unique() if u != "Não Informado"])
    filtro_unidade = criar_botoes_filtro("Unidade / Prédio Operacional:", unidades_base, "Todas", "f_unid")

    # Barra de busca por texto livre
    busca_texto = st.text_input(
        "🔍 Busca Rápida (pesquise por ID do chamado, solicitante ou termo no título):",
        placeholder="Ex.: esgoto, lâmpada, disjuntor, telhado, Fabiola, 198086..."
    ).strip().lower()

# ==============================================================================
# 5. APLICAÇÃO DOS FILTROS NA BASE
# ==============================================================================
df_view = df_chamados.copy()

if filtro_atendimento != "Todas":
    df_view = df_view[df_view["Tipo_Atendimento"] == filtro_atendimento]
if filtro_sla != "Todos":
    df_view = df_view[df_view["Farol_SLA"] == filtro_sla]
if filtro_status != "Todos":
    df_view = df_view[df_view["Status"] == filtro_status]
if filtro_unidade != "Todas":
    df_view = df_view[df_view["Local_Tratado"] == filtro_unidade]

if busca_texto:
    df_view = df_view[
        df_view["Título"].str.lower().str.contains(busca_texto, na=False) |
        df_view["Usuário Solicitante"].str.lower().str.contains(busca_texto, na=False) |
        df_view["#"].astype(str).str.contains(busca_texto, na=False) |
        df_view["Subcategoria"].str.lower().str.contains(busca_texto, na=False)
    ]

# ==============================================================================
# 6. CARTÕES DE INDICADORES EXECUTIVOS (KPIS)
# ==============================================================================
tot = len(df_view)
internos = len(df_view[df_view["Tipo_Atendimento"] == "Equipe Interna (Própria)"])
prohab = len(df_view[df_view["Tipo_Atendimento"] == "OS Contratual (PROHAB)"])
vencidos = len(df_view[df_view["Farol_SLA"] == "Vencido"])
concluidos = len(df_view[df_view["Status"] == "Solicitação Atendida"])

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Total no Filtro", f"{tot}")
col2.metric("Equipe Interna", f"{internos}", delta="Reparos Rápidos", delta_color="normal")
col3.metric("Frente PROHAB (OS)", f"{prohab}", delta="Manutenção Civil", delta_color="normal")
col4.metric("SLA Vencido", f"{vencidos}", delta="Atenção" if vencidos > 0 else "OK", delta_color="inverse")
col5.metric("Atendidos", f"{concluidos}", delta="Concluídos")

st.divider()

# ==============================================================================
# 7. GRÁFICOS ANALÍTICOS (VISÃO TERRITORIAL E TÉCNICA)
# ==============================================================================
g1, g2 = st.columns(2)

with g1:
    st.markdown("##### 📍 Volume de Chamados por Unidade")
    dados_unidade = df_view["Local_Tratado"].value_counts().head(8)
    if not dados_unidade.empty:
        try:
            st.bar_chart(dados_unidade, horizontal=True)
        except TypeError:  # versões antigas do Streamlit não aceitam 'horizontal'
            st.bar_chart(dados_unidade)
    else:
        st.info("Nenhum dado para o filtro selecionado.")

with g2:
    st.markdown("##### 🧱 Demandas por Especialidade Técnica")
    dados_disciplina = df_view["Subcategoria"].value_counts().head(8)
    if not dados_disciplina.empty:
        st.bar_chart(dados_disciplina)
    else:
        st.info("Nenhum dado para o filtro selecionado.")

# ==============================================================================
# 8. TABELA OPERACIONAL DE ATENDIMENTO
# ==============================================================================
st.markdown(f"#### 📋 Fila de Chamados Filtrados ({tot} registros encontrados)")

colunas_tabela = [
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

tabela = df_view[colunas_tabela].sort_values("Dias_Aberto", ascending=False)
st.dataframe(tabela, use_container_width=True, height=430)

# ==============================================================================
# 9. GERADOR DE RELATÓRIO FORMATADO (IMPRESSÃO / PDF)
# ==============================================================================
def gerar_relatorio_impressao(df_dados, filtros_info):
    e = lambda v: "" if pd.isna(v) else html_lib.escape(str(v))
    linhas_tabela = ""
    for _, row in df_dados.iterrows():
        sla_val = str(row.get("Farol_SLA", ""))
        cor_sla = "#b02a37" if sla_val == "Vencido" else ("#198754" if sla_val == "Concluído" else "#0d6efd")
        fundo_sla = "#f8d7da" if sla_val == "Vencido" else ("#d1e7dd" if sla_val == "Concluído" else "#cfe2ff")

        linhas_tabela += f"""
        <tr>
            <td style="text-align: center; font-weight: bold; color: #0d6efd;">{e(row.get('#', ''))}</td>
            <td>{e(row.get('Tipo_Atendimento', ''))}</td>
            <td>{e(row.get('Local_Tratado', ''))}</td>
            <td>{e(row.get('Subcategoria', ''))}</td>
            <td>{e(row.get('Título', ''))}</td>
            <td>{e(row.get('Usuário Solicitante', ''))}</td>
            <td style="text-align: center;">{e(row.get('Dias_Aberto', ''))} d</td>
            <td style="text-align: center;"><span style="background-color: {fundo_sla}; color: {cor_sla}; padding: 3px 8px; border-radius: 4px; font-weight: bold; font-size: 11px;">{e(sla_val)}</span></td>
            <td style="text-align: center;">{e(row.get('Status', ''))}</td>
        </tr>
        """

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <title>Relatório Gerencial de Chamados</title>
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
            <b>Emissão:</b> {e(filtros_info.get('data_hora', ''))}<br>
            <b>Total Filtrado:</b> {e(filtros_info.get('total', 0))} chamados
        </div>
    </div>

    <div class="meta-box">
        <div class="meta-item"><strong>Frente Operacional:</strong> {e(filtros_info.get('frente', 'Todas'))}</div>
        <div class="meta-item"><strong>Unidade / Prédio:</strong> {e(filtros_info.get('unidade', 'Todas'))}</div>
        <div class="meta-item"><strong>Farol SLA:</strong> {e(filtros_info.get('sla', 'Todos'))}</div>
        <div class="meta-item"><strong>Status:</strong> {e(filtros_info.get('status', 'Todos'))}</div>
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

# Metadados para o relatório impresso
info_filtros = {
    "frente": filtro_atendimento,
    "unidade": filtro_unidade,
    "sla": filtro_sla,
    "status": filtro_status,
    "total": len(tabela),
    "data_hora": datetime.now().strftime("%d/%m/%Y %H:%M")
}

# ==============================================================================
# 10. BOTÕES DE EXPORTAÇÃO
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
