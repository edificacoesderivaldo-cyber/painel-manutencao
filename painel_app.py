import streamlit as st
import pandas as pd
import json
from datetime import datetime
from io import BytesIO
import base64

st.set_page_config(page_title="Painel de Manutenção Predial", layout="wide", initial_sidebar_state="collapsed")

st.title("📊 Painel de Manutenção Predial")
st.markdown("**Análise de chamados e controle financeiro — SESI e SENAI**  |  Envie sua planilha Excel para gerar o relatório completo")

uploaded_file = st.file_uploader("📎 Envie o arquivo Excel (CONTROLE_DE_O_S.xlsx)", type=["xlsx", "xls"])

if not uploaded_file:
    st.info("👉 Clique no botão acima para fazer upload da sua planilha Excel com os chamados de manutenção.")
    st.stop()

try:
    df_os = pd.read_excel(uploaded_file, sheet_name='O.S', header=2)
except Exception as e:
    st.error(f"❌ Erro ao carregar aba 'O.S': {e}")
    st.stop()

try:
    df_saldo = pd.read_excel(uploaded_file, sheet_name='SALDO', header=0)
except Exception:
    df_saldo = None

df_os = df_os.dropna(subset=['O.S']).copy()

def norm_status(s):
    if pd.isna(s) or s == '': return "SEM STATUS"
    s = str(s).strip().upper()
    mapping = {"PARALIZADO": "PARALISADO", "CONCLUÍDO": "CONCLUIDO"}
    return mapping.get(s, s)

def norm_unidade(u):
    if pd.isna(u): return "NÃO INFORMADA"
    return str(u).strip().upper().replace("GUARA", "GUARÁ")

status_col = next((c for c in df_os.columns if 'STATUS' in c.upper()), 'STATUS')
unidade_col = next((c for c in df_os.columns if 'UNIDADE' in c.upper()), 'UNIDADE')

df_os['STATUS'] = df_os[status_col].apply(norm_status)
df_os['UNIDADE'] = df_os[unidade_col].apply(norm_unidade)

valor_col = next((c for c in df_os.columns if 'VALOR' in c.upper() and 'INICIAL' in c.upper()), None)
if valor_col:
    df_os[valor_col] = pd.to_numeric(df_os[valor_col], errors='coerce').fillna(0.0)

desc_col = next((c for c in df_os.columns if 'DESCRI' in c.upper()), None)
nr_col = next((c for c in df_os.columns if c.strip().upper() == 'NR'), None)
os_col = next((c for c in df_os.columns if 'O.S' in c.upper() or 'OS' in c.upper()), 'O.S')

df_os['NR'] = pd.to_numeric(df_os[nr_col], errors='coerce').fillna(0).astype(int) if nr_col else 0
df_os['O.S'] = pd.to_numeric(df_os[os_col], errors='coerce').fillna(0).astype(int)

# Detecção de coluna de data de envio
data_col = next((c for c in df_os.columns if 'DATA' in c.upper() and 'ENVIO' in c.upper()), None)
if data_col:
    df_os[data_col] = pd.to_datetime(df_os[data_col], errors='coerce')

# Detecção de coluna de mês de emissão / nota fiscal / conclusão
mes_col = next((c for c in df_os.columns if any(k in c.upper() for k in ['MÊS', 'MES', 'EMISSÃO', 'EMISSAO', 'COMPETÊNCIA', 'COMPETENCIA'])), None)
nf_col = next((c for c in df_os.columns if any(k in c.upper() for k in ['NF', 'NFE', 'NOTA FISCAL'])), None)

hoje = datetime.now()
total_chamados = len(df_os)
concluido = len(df_os[df_os['STATUS'] == 'CONCLUIDO'])
em_execucao = len(df_os[df_os['STATUS'] == 'EM EXECUÇÃO'])
paralisado = len(df_os[df_os['STATUS'] == 'PARALISADO'])
aguard_orcamento = len(df_os[df_os['STATUS'] == 'AGUARDANDO ORÇAMENTO'])
valor_total = float(df_os[valor_col].sum() if valor_col else 0.0)

abertos_30 = 0
if data_col:
    for _, row in df_os.iterrows():
        if row['STATUS'] != 'CONCLUIDO' and pd.notna(row[data_col]):
            dias = (hoje - pd.to_datetime(row[data_col])).days
            if dias > 30:
                abertos_30 += 1

MESES_PT = {
    1: 'Jan', 2: 'Fev', 3: 'Mar', 4: 'Abr', 5: 'Mai', 6: 'Jun',
    7: 'Jul', 8: 'Ago', 9: 'Set', 10: 'Out', 11: 'Nov', 12: 'Dez'
}

data = {
    'chamados': [],
    'status_dist': {},
    'unidades': {},
    'totais': {},
    'contracts': {
        'sesi': {'contrato': 1440000.0, 'utilizado': 0.0, 'saldo': 1440000.0},
        'senai': {'contrato': 1440000.0, 'utilizado': 0.0, 'saldo': 1440000.0}
    }
}

for _, row in df_os.iterrows():
    valor = float(row[valor_col]) if valor_col and pd.notna(row[valor_col]) else 0.0
    status = str(row['STATUS'])
    dias_abertos = None
    data_envio_str = ''
    mes_emissao = 'Não Definido'
    
    if data_col and pd.notna(row[data_col]):
        data_envio_dt = pd.to_datetime(row[data_col])
        data_envio_str = data_envio_dt.strftime('%Y-%m-%d')
        mes_emissao = f"{MESES_PT.get(data_envio_dt.month, data_envio_dt.month)}/{data_envio_dt.year}"
        if status != 'CONCLUIDO':
            dias_abertos = (hoje - data_envio_dt).days
            
    if mes_col and pd.notna(row[mes_col]) and str(row[mes_col]).strip() != '':
        mes_emissao = str(row[mes_col]).strip()

    num_nf = str(row[nf_col]).strip() if nf_col and pd.notna(row[nf_col]) else '-'
    if num_nf.lower() == 'nan': num_nf = '-'

    # Regra de liberação de pagamento para NFE
    if status == 'CONCLUIDO':
        status_pagamento = 'LIBERADO P/ NFE'
    elif status in ['EM EXECUÇÃO', 'LIBERADO']:
        status_pagamento = 'EM MEDIÇÃO'
    elif status == 'PARALISADO':
        status_pagamento = 'BLOQUEADO'
    else:
        status_pagamento = 'PENDENTE ORÇAMENTO'

    desc_val = str(row[desc_col]).strip() if desc_col and pd.notna(row[desc_col]) else ''
    if desc_val.lower() == 'nan':
        desc_val = '-'

    unidade_str = str(row['UNIDADE'])
    casa = 'SESI' if 'SESI' in unidade_str.upper() else ('SENAI' if 'SENAI' in unidade_str.upper() else 'OUTROS')

    data['chamados'].append({
        'nr': int(row['NR']),
        'os': int(row['O.S']),
        'unidade': unidade_str,
        'casa': casa,
        'status': status,
        'descricao': desc_val,
        'valor': round(valor, 2),
        'data_envio': data_envio_str,
        'dias_abertos': dias_abertos,
        'mes_emissao': mes_emissao,
        'status_pagamento': status_pagamento,
        'nota_fiscal': num_nf
    })

for status, count in df_os['STATUS'].value_counts().items():
    data['status_dist'][str(status)] = int(count)

for unit in sorted(df_os['UNIDADE'].unique()):
    df_u = df_os[df_os['UNIDADE'] == unit]
    data['unidades'][str(unit)] = {
        'total': len(df_u),
        'valor': round(float(df_u[valor_col].sum() if valor_col else 0.0), 2),
        'com_orcamento': len(df_u[df_u[valor_col] > 0]) if valor_col else 0,
        'status': {str(k): int(v) for k, v in df_u['STATUS'].value_counts().items()}
    }

df_sesi = df_os[df_os['UNIDADE'].str.contains('SESI', case=False, na=False)]
df_senai = df_os[df_os['UNIDADE'].str.contains('SENAI', case=False, na=False)]

sesi_utilizado = round(float(df_sesi[valor_col].sum() if valor_col else 0.0), 2)
senai_utilizado = round(float(df_senai[valor_col].sum() if valor_col else 0.0), 2)

data['contracts']['sesi']['utilizado'] = sesi_utilizado
data['contracts']['senai']['utilizado'] = senai_utilizado
data['contracts']['sesi']['saldo'] = round(data['contracts']['sesi']['contrato'] - sesi_utilizado, 2)
data['contracts']['senai']['saldo'] = round(data['contracts']['senai']['contrato'] - senai_utilizado, 2)

data['totais'] = {
    'total_chamados': total_chamados,
    'concluido': concluido,
    'em_execucao': em_execucao,
    'paralisado': paralisado,
    'aguard_orcamento': aguard_orcamento,
    'abertos_30dias': abertos_30,
    'valor_total': round(valor_total, 2)
}

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Painel de Manutenção Predial e Pagamentos - SESI/SENAI</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        :root {
            --primary: #3987e5;
            --success: #0ca30c;
            --warning: #fab219;
            --danger: #e34948;
            --info: #1baf7a;
            --light: #f5f5f5;
            --dark: #0b0b0b;
            --text: #333;
            --border: #ddd;
            --pay-bg: #f0f7ff;
            --pay-border: #b8daff;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
            background: #f8f9fa;
            color: var(--text);
            line-height: 1.6;
        }

        .container {
            max-width: 1400px;
            margin: 0 auto;
            padding: 20px;
        }

        .header {
            display: flex;
            align-items: center;
            gap: 16px;
            margin-bottom: 30px;
            padding: 20px;
            background: white;
            border-radius: 8px;
            border-left: 4px solid var(--primary);
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }

        .header h1 {
            font-size: 28px;
            margin: 0;
        }

        .header p {
            color: #666;
            margin: 0;
            font-size: 14px;
        }

        .kpi-section {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }

        .kpi-card {
            background: white;
            padding: 20px;
            border-radius: 8px;
            border: 1px solid var(--border);
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            transition: all 0.3s;
        }

        .kpi-card:hover {
            border-color: var(--primary);
            box-shadow: 0 4px 8px rgba(57, 135, 229, 0.1);
        }

        .kpi-label {
            font-size: 12px;
            color: #666;
            margin-bottom: 8px;
            font-weight: 500;
            text-transform: uppercase;
        }

        .kpi-value {
            font-size: 32px;
            font-weight: bold;
            color: var(--dark);
            margin-bottom: 4px;
        }

        .kpi-percent {
            font-size: 12px;
            color: #999;
        }

        .grid2 {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(350px, 1fr));
            gap: 24px;
            margin-bottom: 24px;
        }

        .card {
            background: white;
            padding: 24px;
            border-radius: 8px;
            border: 1px solid var(--border);
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            margin-bottom: 24px;
        }

        .payment-card {
            border: 2px solid var(--primary);
            box-shadow: 0 4px 12px rgba(57, 135, 229, 0.12);
            background: #ffffff;
        }

        .payment-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 12px;
            padding-bottom: 14px;
            border-bottom: 2px solid var(--light);
            margin-bottom: 20px;
        }

        .payment-badge-status {
            background: #e7f3ff;
            color: #0d6efd;
            padding: 6px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 700;
            letter-spacing: 0.5px;
        }

        .card-title {
            font-size: 18px;
            font-weight: 600;
            margin-bottom: 20px;
            padding-bottom: 12px;
            border-bottom: 2px solid var(--light);
        }

        .contract-stat {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 12px;
            font-size: 13px;
        }

        .contract-label {
            color: #666;
        }

        .contract-value {
            font-weight: 600;
            color: var(--dark);
        }

        .progress-container {
            margin-bottom: 16px;
        }

        .progress-label {
            display: flex;
            justify-content: space-between;
            margin-bottom: 6px;
            font-size: 12px;
            color: #666;
        }

        .progress-bar {
            width: 100%;
            height: 8px;
            background: #e9ecef;
            border-radius: 4px;
            overflow: hidden;
        }

        .progress-fill {
            height: 100%;
            background: linear-gradient(90deg, var(--info), var(--success));
            border-radius: 4px;
            transition: width 0.3s;
        }

        .chart-wrapper {
            position: relative;
            height: 320px;
            margin: 20px 0;
        }

        .filter-section {
            display: flex;
            flex-direction: column;
            gap: 16px;
        }

        .filter-group {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            align-items: center;
        }

        .filter-label {
            font-weight: 600;
            font-size: 13px;
            color: #666;
            display: block;
            width: 100%;
            margin-bottom: 4px;
        }

        .filter-chip {
            display: inline-block;
            padding: 6px 12px;
            background: white;
            border: 1px solid var(--border);
            border-radius: 20px;
            font-size: 13px;
            cursor: pointer;
            transition: all 0.2s;
            user-select: none;
        }

        .filter-chip:hover {
            border-color: var(--primary);
            background: var(--light);
        }

        .filter-chip.active {
            background: var(--primary);
            color: white;
            border-color: var(--primary);
        }

        .filter-chip-pay {
            border-color: #bee5eb;
            background: #f8fbff;
        }

        .filter-chip-pay.active {
            background: #0d6efd;
            border-color: #0d6efd;
            color: white;
        }

        .filter-input {
            padding: 8px 12px;
            border: 1px solid var(--border);
            border-radius: 4px;
            font-size: 13px;
            width: 100%;
            max-width: 300px;
        }

        .filter-input:focus {
            outline: none;
            border-color: var(--primary);
            box-shadow: 0 0 0 3px rgba(57, 135, 229, 0.1);
        }

        .table-wrapper {
            overflow-x: auto;
            margin-top: 16px;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }

        thead {
            background: var(--light);
        }

        th {
            padding: 12px;
            text-align: left;
            color: #666;
            font-weight: 600;
            border-bottom: 2px solid var(--border);
        }

        td {
            padding: 12px;
            border-bottom: 1px solid var(--border);
        }

        tbody tr:hover {
            background: var(--light);
        }

        .status-badge {
            display: inline-block;
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 11px;
            font-weight: 600;
        }

        .status-concluido { background: #d4edda; color: #155724; }
        .status-em-execução, .status-em-execucao { background: #d1ecf1; color: #0c5460; }
        .status-paralisado { background: #f8d7da; color: #721c24; }
        .status-aguardando-orçamento, .status-aguardando-orcamento { background: #fff3cd; color: #856404; }
        .status-liberado { background: #d4edda; color: #155724; }
        .status-projeto { background: #e2e3e5; color: #383d41; }
        .status-planejamento { background: #d6d8db; color: #383d41; }
        .status-aguardando-aprovação, .status-aguardando-aprovacao { background: #ffe5d0; color: #a04000; }
        .status-sem-status { background: #e9ecef; color: #495057; }

        /* Badges de Pagamento / NFE */
        .pay-badge-liberado {
            background: #d1e7dd;
            color: #0f5132;
            border: 1px solid #badbcc;
            font-weight: 700;
        }
        .pay-badge-medicao {
            background: #cff4fc;
            color: #055160;
            border: 1px solid #b6effb;
            font-weight: 600;
        }
        .pay-badge-bloqueado {
            background: #f8d7da;
            color: #842029;
            border: 1px solid #f5c2c7;
            font-weight: 600;
        }
        .pay-badge-pendente {
            background: #fff3cd;
            color: #664d03;
            border: 1px solid #ffecb5;
        }

        .unit-card {
            padding: 16px;
            background: white;
            border: 1px solid var(--border);
            border-radius: 8px;
            box-shadow: 0 1px 2px rgba(0,0,0,0.05);
        }

        .unit-card-name {
            font-weight: 600;
            color: #333;
            margin-bottom: 12px;
            font-size: 14px;
        }

        .unit-card-info {
            display: flex;
            justify-content: space-between;
            align-items: flex-end;
        }

        .unit-card-label {
            font-size: 12px;
            color: #666;
            margin-bottom: 4px;
        }

        .unit-card-value {
            font-weight: 600;
            color: #333;
            font-size: 16px;
        }

        .unit-card-value-right {
            font-weight: 600;
            color: var(--primary);
            font-size: 14px;
        }

        .footer {
            margin-top: 40px;
            padding: 20px;
            text-align: center;
            color: #999;
            font-size: 12px;
        }

        @media print {
            body { background: white; }
            .filter-section { display: none; }
            .kpi-card:hover { box-shadow: none; }
        }

        @media (max-width: 768px) {
            .grid2 { grid-template-columns: 1fr; }
            .kpi-section { grid-template-columns: repeat(2, 1fr); }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div>
                <h1>📊 Painel de Manutenção Predial & Liberação de Pagamentos</h1>
                <p>SESI e SENAI — Acompanhamento de Chamados, Emissão de NFE e Gestão Orçamentária</p>
            </div>
        </div>

        <!-- KPIs Gerais -->
        <div class="kpi-section" id="kpis"></div>

        <!-- Saldo em Contrato + Tempo em Aberto -->
        <div class="grid2">
            <div class="card" style="margin-bottom: 0;">
                <div class="card-title">💰 Saldo em Contrato</div>
                <div id="contracts"></div>
            </div>

            <div class="card" style="margin-bottom: 0;">
                <div class="card-title">⏱️ Tempo Médio em Aberto</div>
                <div id="tempoAberto"></div>
            </div>
        </div>

        <!-- Distribuição de Status -->
        <div class="card">
            <div class="card-title">📈 Distribuição de Status</div>
            <div class="chart-wrapper">
                <canvas id="statusChart" role="img" aria-label="Distribuição de chamados por status"></canvas>
            </div>
        </div>

        <!-- Filtros Globais Interativos -->
        <div class="card">
            <div class="card-title">🔍 Filtros Interativos Globais</div>
            <div class="filter-section" id="filters"></div>
        </div>

        <!-- Chamados por Unidade -->
        <div class="card">
            <div class="card-title">🏢 Chamados por Unidade (com investimento)</div>
            <div class="chart-wrapper" style="height: 500px; margin-bottom: 20px;">
                <canvas id="unitChart" role="img" aria-label="Distribuição de chamados por unidade"></canvas>
            </div>
            <div id="unitLegendBottom"></div>
        </div>

        <!-- ======================================================= -->
        <!-- NOVO PAINEL: CONTROLE DE PAGAMENTO / LIBERAÇÃO DE NFE  -->
        <!-- ======================================================= -->
        <div class="card payment-card" id="painelPagamentos">
            <div class="payment-header">
                <div>
                    <h2 style="font-size: 20px; font-weight: 700; color: #0d6efd; display: flex; align-items: center; gap: 8px;">
                        💳 Controle de Pagamentos & Liberação para NFE
                    </h2>
                    <p style="font-size: 13px; color: #666; margin-top: 4px;">
                        Valide os serviços concluídos por mês e CNPJ (SESI/SENAI) para autorizar a empresa a emitir a Nota Fiscal
                    </p>
                </div>
                <div class="payment-badge-status" id="paySummaryBadge">
                    Carregando resumo financeiro...
                </div>
            </div>

            <!-- Mini KPIs de Pagamento -->
            <div class="kpi-section" style="margin-bottom: 20px;">
                <div class="kpi-card" style="border-left: 4px solid var(--success);">
                    <div class="kpi-label">Liberado para Emitir NFE</div>
                    <div class="kpi-value" id="payValLiberado" style="color: var(--success); font-size: 26px;">R$ 0,00</div>
                    <div class="kpi-percent" id="payCountLiberado">0 O.S. aptas para faturamento</div>
                </div>
                <div class="kpi-card" style="border-left: 4px solid var(--primary);">
                    <div class="kpi-label">Em Medição / Andamento</div>
                    <div class="kpi-value" id="payValMedicao" style="color: var(--primary); font-size: 26px;">R$ 0,00</div>
                    <div class="kpi-percent" id="payCountMedicao">0 O.S. em execução</div>
                </div>
                <div class="kpi-card" style="border-left: 4px solid var(--danger);">
                    <div class="kpi-label">Bloqueado / Paralisado</div>
                    <div class="kpi-value" id="payValBloqueado" style="color: var(--danger); font-size: 26px;">R$ 0,00</div>
                    <div class="kpi-percent" id="payCountBloqueado">0 O.S. paralisadas</div>
                </div>
                <div class="kpi-card" style="border-left: 4px solid #6c757d;">
                    <div class="kpi-label">Total do Filtro de Pagamento</div>
                    <div class="kpi-value" id="payValTotal" style="font-size: 26px;">R$ 0,00</div>
                    <div class="kpi-percent" id="payCountTotal">0 chamados selecionados</div>
                </div>
            </div>

            <!-- Filtro de Entidade / CNPJ, Mês de Emissão e Status -->
            <div style="background: #f8f9fa; padding: 16px; border-radius: 8px; margin-bottom: 20px; border: 1px solid var(--border);">
                <div style="font-weight: 600; font-size: 13px; color: #333; margin-bottom: 8px;">
                    🏛️ Filtrar por Entidade / CNPJ de Faturamento (Casa):
                </div>
                <div class="filter-group" id="payCasaFilters" style="margin-bottom: 14px;"></div>

                <div style="font-weight: 600; font-size: 13px; color: #333; margin-bottom: 8px;">
                    📅 Filtrar por Mês de Emissão / Competência da Nota:
                </div>
                <div class="filter-group" id="payMonthFilters"></div>

                <div style="font-weight: 600; font-size: 13px; color: #333; margin-top: 14px; margin-bottom: 8px;">
                    📌 Filtrar por Status de Faturamento:
                </div>
                <div class="filter-group" id="payStatusFilters"></div>
            </div>

            <!-- Tabela de Liberação de Pagamentos -->
            <div class="table-wrapper">
                <table id="payTable">
                    <thead>
                        <tr style="background: #e9ecef;">
                            <th>O.S</th>
                            <th>NR</th>
                            <th>Entidade / CNPJ</th>
                            <th>Unidade</th>
                            <th>Descrição</th>
                            <th>Mês Competência</th>
                            <th>Status O.S</th>
                            <th>Liberação p/ NFE</th>
                            <th style="text-align: right;">Valor a Faturar</th>
                        </tr>
                    </thead>
                    <tbody id="payTableBody"></tbody>
                </table>
            </div>
        </div>

        <!-- Tabela Completa de Chamados Geral -->
        <div class="card">
            <div class="card-title">📋 Lista Completa de Chamados</div>
            
            <div style="margin-bottom: 16px;">
                <div style="font-weight: 600; font-size: 13px; color: #666; margin-bottom: 10px;">Filtrar por Status na Tabela</div>
                <div class="filter-group" id="listStatusFilter"></div>
            </div>

            <div style="margin-bottom: 16px; display: flex; gap: 12px; align-items: flex-end; flex-wrap: wrap;">
                <div style="flex: 1; min-width: 220px;">
                    <label style="display: block; margin-bottom: 6px; font-size: 12px; font-weight: 600; color: #666;">🔍 Buscar por O.S.</label>
                    <input type="text" class="filter-input" id="osSearch" placeholder="Ex: 194876" oninput="filterByOS(this.value)" />
                </div>
                <div style="font-size: 13px; color: #666;">
                    Exibindo <strong id="ticketCount">0</strong> de <strong id="totalTickets">0</strong> chamados
                </div>
            </div>
            <div class="table-wrapper">
                <table id="ticketTable">
                    <thead>
                        <tr>
                            <th>O.S</th>
                            <th>NR</th>
                            <th>Unidade</th>
                            <th>Descrição</th>
                            <th>Status</th>
                            <th>Data Envio</th>
                            <th>Dias Abertos</th>
                            <th style="text-align: right;">Valor</th>
                        </tr>
                    </thead>
                    <tbody id="tableBody"></tbody>
                </table>
            </div>
        </div>

        <div class="footer">
            <p>Atualizado em <span id="updateTime"></span> | Painel Integrado de Gestão Predial & Faturamento SESI/SENAI</p>
        </div>
    </div>

    <script>
        const COLOR_MAP = {
            'CONCLUIDO': '#0ca30c',
            'EM EXECUÇÃO': '#3987e5',
            'PARALISADO': '#e34948',
            'AGUARDANDO ORÇAMENTO': '#fab219',
            'LIBERADO': '#1baf7a',
            'PROJETO': '#6c757d',
            'PLANEJAMENTO': '#9085e9',
            'AGUARDANDO APROVAÇÃO': '#fd7e14',
            'SEM STATUS': '#999'
        };

        const DATA = __DATA_PLACEHOLDER__;

        let state = {
            tickets: [],
            filters: { status: [], unidade: [], os: '', listStatus: '' },
            payFilters: { mes: '', statusPagamento: '', casa: '' }
        };

        function init() {
            state.tickets = DATA.chamados;
            document.getElementById('updateTime').textContent = new Date().toLocaleString('pt-BR');
            document.getElementById('totalTickets').textContent = state.tickets.length;
            renderFilters();
            renderListStatusFilter();
            renderPaymentCasaFilters();
            renderPaymentMonthFilters();
            render();
        }

        function getFiltered() {
            if (state.filters.os !== '') {
                return state.tickets.filter(t => t.os.toString().includes(state.filters.os));
            }
            return state.tickets.filter(t => {
                const matchStatus = state.filters.status.length === 0 || state.filters.status.includes(t.status);
                const matchUnidade = state.filters.unidade.length === 0 || state.filters.unidade.includes(t.unidade);
                const matchListStatus = state.filters.listStatus === '' || t.status === state.filters.listStatus;
                return matchStatus && matchUnidade && matchListStatus;
            });
        }

        function getPayFiltered() {
            return state.tickets.filter(t => {
                const matchUnidade = state.filters.unidade.length === 0 || state.filters.unidade.includes(t.unidade);
                const matchCasa = state.payFilters.casa === '' || t.casa === state.payFilters.casa;
                const matchMes = state.payFilters.mes === '' || t.mes_emissao === state.payFilters.mes;
                const matchPayStatus = state.payFilters.statusPagamento === '' || t.status_pagamento === state.payFilters.statusPagamento;
                return matchUnidade && matchCasa && matchMes && matchPayStatus;
            });
        }

        function getPayFiltered() {
            return state.tickets.filter(t => {
                const matchUnidade = state.filters.unidade.length === 0 || state.filters.unidade.includes(t.unidade);
                const matchMes = state.payFilters.mes === '' || t.mes_emissao === state.payFilters.mes;
                const matchPayStatus = state.payFilters.statusPagamento === '' || t.status_pagamento === state.payFilters.statusPagamento;
                return matchUnidade && matchMes && matchPayStatus;
            });
        }

        function renderKPIs() {
            const filtered = getFiltered();
            const total = filtered.length;
            const concluido = filtered.filter(t => t.status === 'CONCLUIDO').length;
            const emExec = filtered.filter(t => t.status === 'EM EXECUÇÃO').length;
            const paralisado = filtered.filter(t => t.status === 'PARALISADO').length;
            const muitoAntigos = filtered.filter(t => t.dias_abertos !== null && t.dias_abertos > 30).length;
            const valorTotal = filtered.reduce((s, t) => s + (t.valor || 0), 0);
            const comOrcamento = filtered.filter(t => t.valor > 0).length;

            const html = `
                <div class="kpi-card">
                    <div class="kpi-label">Total de Chamados</div>
                    <div class="kpi-value">${total}</div>
                    <div class="kpi-percent">Chamados filtrados</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Concluídos</div>
                    <div class="kpi-value" style="color: var(--success);">${concluido}</div>
                    <div class="kpi-percent">${total ? Math.round(concluido * 100 / total) : 0}% do total</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Em Execução</div>
                    <div class="kpi-value" style="color: var(--primary);">${emExec}</div>
                    <div class="kpi-percent">${total ? Math.round(emExec * 100 / total) : 0}% em andamento</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">⚠️ Paralisados</div>
                    <div class="kpi-value" style="color: var(--danger);">${paralisado}</div>
                    <div class="kpi-percent">Necessita atenção</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">🔴 Abertos > 30 dias</div>
                    <div class="kpi-value" style="color: #e34948;">${muitoAntigos}</div>
                    <div class="kpi-percent">Cobrar celeridade</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Valor Investido</div>
                    <div class="kpi-value" style="font-size: 22px; color: var(--primary);">
                        R$ ${(valorTotal / 1000).toFixed(1)}K
                    </div>
                    <div class="kpi-percent">${comOrcamento} chamados orçados</div>
                </div>
            `;
            document.getElementById('kpis').innerHTML = html;
        }

        function renderContracts() {
            const contracts = DATA.contracts;
            let html = '';
            for (const [key, cData] of Object.entries(contracts)) {
                const name = key === 'sesi' ? 'SESI' : 'SENAI';
                const utilizado = cData.utilizado;
                const saldo = cData.saldo;
                const contrato = cData.contrato;
                const pctUtilizado = contrato > 0 ? Math.min(100, Math.round((utilizado * 100) / contrato)) : 0;

                html += `
                    <div style="margin-bottom: 24px;">
                        <h3 style="font-size: 14px; margin-bottom: 12px; color: var(--primary); font-weight: 600; text-transform: uppercase;">${name}</h3>
                        <div class="contract-stat">
                            <span class="contract-label">Valor do Contrato</span>
                            <span class="contract-value">R$ ${new Intl.NumberFormat('pt-BR', {minimumFractionDigits: 2}).format(contrato)}</span>
                        </div>
                        <div class="contract-stat">
                            <span class="contract-label">Utilizado</span>
                            <span class="contract-value">R$ ${new Intl.NumberFormat('pt-BR', {minimumFractionDigits: 2}).format(utilizado)}</span>
                        </div>
                        <div class="progress-container">
                            <div class="progress-label">
                                <span>Progresso</span>
                                <span>${pctUtilizado}% utilizado</span>
                            </div>
                            <div class="progress-bar">
                                <div class="progress-fill" style="width: ${pctUtilizado}%"></div>
                            </div>
                        </div>
                        <div class="contract-stat" style="margin-bottom: 0;">
                            <span class="contract-label">Saldo Disponível</span>
                            <span class="contract-value" style="color: var(--info);">R$ ${new Intl.NumberFormat('pt-BR', {minimumFractionDigits: 2}).format(saldo)}</span>
                        </div>
                    </div>
                `;
            }
            document.getElementById('contracts').innerHTML = html;
        }

        function renderTempoAberto() {
            const filtered = getFiltered();
            const abertos = filtered.filter(t => t.status !== 'CONCLUIDO' && t.dias_abertos !== null);

            if (abertos.length === 0) {
                document.getElementById('tempoAberto').innerHTML = '<p style="color: #999; text-align: center; padding: 40px 0;">Nenhum chamado aberto nos filtros atuais</p>';
                return;
            }

            const mediaDias = Math.round(abertos.reduce((s, t) => s + t.dias_abertos, 0) / abertos.length);
            const range030 = abertos.filter(t => t.dias_abertos <= 30).length;
            const range3060 = abertos.filter(t => t.dias_abertos > 30 && t.dias_abertos <= 60).length;
            const range60plus = abertos.filter(t => t.dias_abertos > 60).length;

            const html = `
                <div style="text-align: center; margin-bottom: 24px; padding: 20px; background: #f8f9fa; border-radius: 8px;">
                    <div style="font-size: 12px; color: #666; margin-bottom: 8px; text-transform: uppercase;">Média de Dias em Aberto</div>
                    <div style="font-size: 48px; font-weight: bold; color: var(--primary); margin-bottom: 4px;">${mediaDias}</div>
                    <div style="font-size: 13px; color: #999;">dias (em ${abertos.length} chamados abertos)</div>
                </div>
                <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px;">
                    <div style="background: #d4edda; border-radius: 8px; padding: 14px; text-align: center; border-left: 4px solid #0ca30c;">
                        <div style="font-size: 26px; font-weight: bold; color: #155724;">${range030}</div>
                        <div style="font-size: 12px; color: #155724; font-weight: 600; margin-top: 4px;">0-30 dias</div>
                        <div style="font-size: 11px; color: #666; margin-top: 2px;">${Math.round(range030 * 100 / abertos.length)}%</div>
                    </div>
                    <div style="background: #fff3cd; border-radius: 8px; padding: 14px; text-align: center; border-left: 4px solid #fab219;">
                        <div style="font-size: 26px; font-weight: bold; color: #856404;">${range3060}</div>
                        <div style="font-size: 12px; color: #856404; font-weight: 600; margin-top: 4px;">30-60 dias</div>
                        <div style="font-size: 11px; color: #666; margin-top: 2px;">${Math.round(range3060 * 100 / abertos.length)}%</div>
                    </div>
                    <div style="background: #f8d7da; border-radius: 8px; padding: 14px; text-align: center; border-left: 4px solid #e34948;">
                        <div style="font-size: 26px; font-weight: bold; color: #721c24;">${range60plus}</div>
                        <div style="font-size: 12px; color: #721c24; font-weight: 600; margin-top: 4px;">>60 dias</div>
                        <div style="font-size: 11px; color: #666; margin-top: 2px;">${Math.round(range60plus * 100 / abertos.length)}%</div>
                    </div>
                </div>
            `;
            document.getElementById('tempoAberto').innerHTML = html;
        }

        function renderFilters() {
            const statuses = [...new Set(state.tickets.map(t => t.status))].sort();
            const unidades = [...new Set(state.tickets.map(t => t.unidade))].sort();

            let html = `
                <div class="filter-group">
                    <span class="filter-label">Status (clique para alternar)</span>
                    ${statuses.map(s => `
                        <span class="filter-chip ${state.filters.status.includes(s) ? 'active' : ''}" onclick="toggleFilter('status', '${s}')">
                            ${s}
                        </span>
                    `).join('')}
                </div>
                <div class="filter-group" style="margin-top: 10px;">
                    <span class="filter-label">Unidade (clique para alternar)</span>
                    ${unidades.map(u => `
                        <span class="filter-chip ${state.filters.unidade.includes(u) ? 'active' : ''}" onclick="toggleFilter('unidade', '${u}')">
                            ${u}
                        </span>
                    `).join('')}
                </div>
            `;
            document.getElementById('filters').innerHTML = html;
        }

        function renderPaymentCasaFilters() {
            const opcoes = [
                { id: '', label: '🏢 Todas as Entidades (SESI + SENAI)' },
                { id: 'SESI', label: '🔵 SESI (CNPJ SESI)' },
                { id: 'SENAI', label: '🟠 SENAI (CNPJ SENAI)' }
            ];

            let html = '';
            opcoes.forEach(op => {
                const isActive = state.payFilters.casa === op.id;
                html += `
                    <span class="filter-chip filter-chip-pay ${isActive ? 'active' : ''}" onclick="setPayCasa('${op.id}')">
                        ${op.label}
                    </span>
                `;
            });
            document.getElementById('payCasaFilters').innerHTML = html;
        }

        function setPayCasa(casa) {
            state.payFilters.casa = casa;
            renderPaymentCasaFilters();
            renderPaymentPanel();
        }

        function renderPaymentMonthFilters() {
            const mesesValidos = [...new Set(state.tickets.map(t => t.mes_emissao))].filter(m => m && m !== 'Não Definido').sort();
            
            let htmlMes = `
                <span class="filter-chip filter-chip-pay ${state.payFilters.mes === '' ? 'active' : ''}" onclick="setPayMonth('')">
                    Todos os Meses
                </span>
            `;
            mesesValidos.forEach(m => {
                const isActive = state.payFilters.mes === m;
                htmlMes += `
                    <span class="filter-chip filter-chip-pay ${isActive ? 'active' : ''}" onclick="setPayMonth('${m}')">
                        📅 ${m}
                    </span>
                `;
            });
            document.getElementById('payMonthFilters').innerHTML = htmlMes;

            const payStatuses = ['LIBERADO P/ NFE', 'EM MEDIÇÃO', 'BLOQUEADO', 'PENDENTE ORÇAMENTO'];
            let htmlStatus = `
                <span class="filter-chip filter-chip-pay ${state.payFilters.statusPagamento === '' ? 'active' : ''}" onclick="setPayStatus('')">
                    Todos os Status
                </span>
            `;
            payStatuses.forEach(s => {
                const isActive = state.payFilters.statusPagamento === s;
                htmlStatus += `
                    <span class="filter-chip filter-chip-pay ${isActive ? 'active' : ''}" onclick="setPayStatus('${s}')">
                        ${s}
                    </span>
                `;
            });
            document.getElementById('payStatusFilters').innerHTML = htmlStatus;
        }

        function setPayMonth(m) {
            state.payFilters.mes = m;
            renderPaymentMonthFilters();
            renderPaymentPanel();
        }

        function setPayStatus(s) {
            state.payFilters.statusPagamento = s;
            renderPaymentMonthFilters();
            renderPaymentPanel();
        }

        function renderPaymentPanel() {
            const payList = getPayFiltered();

            const liberados = payList.filter(t => t.status_pagamento === 'LIBERADO P/ NFE');
            const emMedicao = payList.filter(t => t.status_pagamento === 'EM MEDIÇÃO');
            const bloqueados = payList.filter(t => t.status_pagamento === 'BLOQUEADO');

            const valLiberado = liberados.reduce((acc, t) => acc + (t.valor || 0), 0);
            const valMedicao = emMedicao.reduce((acc, t) => acc + (t.valor || 0), 0);
            const valBloqueado = bloqueados.reduce((acc, t) => acc + (t.valor || 0), 0);
            const valTotal = payList.reduce((acc, t) => acc + (t.valor || 0), 0);

            const fmt = v => new Intl.NumberFormat('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2}).format(v);

            document.getElementById('payValLiberado').textContent = `R$ ${fmt(valLiberado)}`;
            document.getElementById('payCountLiberado').textContent = `${liberados.length} O.S. prontas para faturar`;

            document.getElementById('payValMedicao').textContent = `R$ ${fmt(valMedicao)}`;
            document.getElementById('payCountMedicao').textContent = `${emMedicao.length} O.S. em andamento`;

            document.getElementById('payValBloqueado').textContent = `R$ ${fmt(valBloqueado)}`;
            document.getElementById('payCountBloqueado').textContent = `${bloqueados.length} O.S. paralisadas`;

            document.getElementById('payValTotal').textContent = `R$ ${fmt(valTotal)}`;
            document.getElementById('payCountTotal').textContent = `${payList.length} chamados filtrados`;

            const casaTxt = state.payFilters.casa ? ` [${state.payFilters.casa}]` : ' [SESI + SENAI]';
            const mesTxt = state.payFilters.mes ? `Competência: ${state.payFilters.mes}` : 'Todos os Meses';
            document.getElementById('paySummaryBadge').textContent = `Faturamento: R$ ${fmt(valLiberado)}${casaTxt} (${mesTxt})`;

            let rowsHTML = payList
                .sort((a, b) => {
                    // Ordenar primeiro os liberados para faturamento
                    if (a.status_pagamento === 'LIBERADO P/ NFE' && b.status_pagamento !== 'LIBERADO P/ NFE') return -1;
                    if (a.status_pagamento !== 'LIBERADO P/ NFE' && b.status_pagamento === 'LIBERADO P/ NFE') return 1;
                    return b.valor - a.valor;
                })
                .map(t => {
                    let badgeClass = 'pay-badge-pendente';
                    let icone = '⏳';
                    if (t.status_pagamento === 'LIBERADO P/ NFE') {
                        badgeClass = 'pay-badge-liberado';
                        icone = '✅';
                    } else if (t.status_pagamento === 'EM MEDIÇÃO') {
                        badgeClass = 'pay-badge-medicao';
                        icone = '🔄';
                    } else if (t.status_pagamento === 'BLOQUEADO') {
                        badgeClass = 'pay-badge-bloqueado';
                        icone = '⛔';
                    }

                    const statusClass = 'status-' + t.status.toLowerCase().replace(/\s+/g, '-').replace(/[^\w-]/g, '');
                    const casaBadgeColor = t.casa === 'SESI' ? '#0d6efd' : (t.casa === 'SENAI' ? '#e65100' : '#6c757d');
                    const casaBadgeBg = t.casa === 'SESI' ? '#e7f3ff' : (t.casa === 'SENAI' ? '#fff3e0' : '#f8f9fa');

                    return `
                        <tr>
                            <td style="font-weight: 700; color: #0d6efd;">#${t.os}</td>
                            <td>${t.nr}</td>
                            <td>
                                <span style="display: inline-block; font-size: 11px; font-weight: 700; padding: 2px 8px; border-radius: 4px; background: ${casaBadgeBg}; color: ${casaBadgeColor}; border: 1px solid ${casaBadgeColor}40;">
                                    ${t.casa}
                                </span>
                            </td>
                            <td><strong>${t.unidade}</strong></td>
                            <td style="max-width: 260px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${t.descricao}">${t.descricao}</td>
                            <td><span style="font-size: 12px; font-weight: 600; color: #495057;">${t.mes_emissao}</span></td>
                            <td><span class="status-badge ${statusClass}">${t.status}</span></td>
                            <td><span class="status-badge ${badgeClass}">${icone} ${t.status_pagamento}</span></td>
                            <td style="text-align: right; font-weight: 700; color: ${t.status_pagamento === 'LIBERADO P/ NFE' ? '#0ca30c' : '#333'};">
                                R$ ${fmt(t.valor)}
                            </td>
                        </tr>
                    `;
                }).join('');

            if (payList.length === 0) {
                rowsHTML = '<tr><td colspan="9" style="text-align:center; padding: 24px; color: #999;">Nenhum chamado de pagamento localizado para este filtro.</td></tr>';
            }

            document.getElementById('payTableBody').innerHTML = rowsHTML;
        }

        function toggleFilter(type, value) {
            if (state.filters[type].includes(value)) {
                state.filters[type] = state.filters[type].filter(v => v !== value);
            } else {
                state.filters[type].push(value);
            }
            renderFilters();
            render();
        }

        function filterByOS(value) {
            state.filters.os = value.trim();
            render();
        }

        function toggleListStatus(status) {
            state.filters.listStatus = (state.filters.listStatus === status) ? '' : status;
            renderListStatusFilter();
            render();
        }

        function renderListStatusFilter() {
            const statuses = [...new Set(state.tickets.map(t => t.status))].sort();
            let html = `
                <span class="filter-chip ${state.filters.listStatus === '' ? 'active' : ''}" onclick="toggleListStatus('')">
                    Todos
                </span>
            `;
            statuses.forEach(s => {
                const isActive = state.filters.listStatus === s;
                html += `
                    <span class="filter-chip ${isActive ? 'active' : ''}" onclick="toggleListStatus('${s}')">
                        ${s}
                    </span>
                `;
            });
            document.getElementById('listStatusFilter').innerHTML = html;
        }

        function renderCharts() {
            const filtered = getFiltered();

            // Status Doughnut Chart
            const statusCounts = {};
            filtered.forEach(t => {
                statusCounts[t.status] = (statusCounts[t.status] || 0) + 1;
            });
            const statusLabels = Object.keys(statusCounts).sort();
            const statusValues = statusLabels.map(s => statusCounts[s]);
            const statusColors = statusLabels.map(s => COLOR_MAP[s] || '#999');

            if (window.statusChartInstance) window.statusChartInstance.destroy();
            window.statusChartInstance = new Chart(document.getElementById('statusChart'), {
                type: 'doughnut',
                data: {
                    labels: statusLabels,
                    datasets: [{
                        data: statusValues,
                        backgroundColor: statusColors,
                        borderColor: '#fff',
                        borderWidth: 2
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { position: 'bottom' }
                    }
                }
            });

            // Unit Bar Chart
            const unitData = {};
            filtered.forEach(t => {
                if (!unitData[t.unidade]) {
                    unitData[t.unidade] = { status: {}, valor: 0 };
                }
                unitData[t.unidade].status[t.status] = (unitData[t.unidade].status[t.status] || 0) + 1;
                unitData[t.unidade].valor += t.valor;
            });

            const unitLabels = Object.entries(unitData)
                .sort((a, b) => {
                    const totalA = Object.values(a[1].status).reduce((s, v) => s + v, 0);
                    const totalB = Object.values(b[1].status).reduce((s, v) => s + v, 0);
                    return totalB - totalA;
                })
                .map(e => e[0]);

            const allStatuses = [...new Set(filtered.map(t => t.status))].sort();
            const datasets = allStatuses.map(status => ({
                label: status,
                data: unitLabels.map(unit => unitData[unit].status[status] || 0),
                backgroundColor: COLOR_MAP[status] || '#999',
                borderColor: '#fff',
                borderWidth: 1
            }));

            if (window.unitChartInstance) window.unitChartInstance.destroy();
            window.unitChartInstance = new Chart(document.getElementById('unitChart'), {
                type: 'bar',
                data: { labels: unitLabels, datasets: datasets },
                options: {
                    indexAxis: 'y',
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        x: { beginAtZero: true, stacked: true, ticks: { stepSize: 1 } },
                        y: { stacked: true }
                    },
                    plugins: { legend: { display: false } }
                }
            });

            renderUnitCards(unitData, unitLabels, allStatuses);
        }

        function renderUnitCards(unitData, unitLabels, allStatuses) {
            let legendaHTML = '<div style="display: flex; gap: 16px; flex-wrap: wrap; margin-bottom: 20px; justify-content: center;">';
            allStatuses.forEach(status => {
                const color = COLOR_MAP[status] || '#999';
                legendaHTML += `
                    <div style="display: flex; align-items: center; gap: 8px; font-size: 12px;">
                        <div style="width: 16px; height: 16px; background: ${color}; border-radius: 2px;"></div>
                        <span>${status}</span>
                    </div>
                `;
            });
            legendaHTML += '</div>';

            let cardsHTML = legendaHTML + '<div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 16px;">';
            unitLabels.forEach(unit => {
                const total = Object.values(unitData[unit].status).reduce((s, v) => s + v, 0);
                const valor = unitData[unit].valor;
                const valorFmt = new Intl.NumberFormat('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2}).format(valor);

                cardsHTML += `
                    <div class="unit-card">
                        <div class="unit-card-name">${unit}</div>
                        <div class="unit-card-info">
                            <div>
                                <div class="unit-card-label">Chamados</div>
                                <div class="unit-card-value">${total}</div>
                            </div>
                            <div style="text-align: right;">
                                <div class="unit-card-label">Investido</div>
                                <div class="unit-card-value-right">R$ ${valorFmt}</div>
                            </div>
                        </div>
                    </div>
                `;
            });
            cardsHTML += '</div>';
            document.getElementById('unitLegendBottom').innerHTML = cardsHTML;
        }

        function renderTable() {
            const filtered = getFiltered();
            document.getElementById('ticketCount').textContent = filtered.length;

            let html = filtered
                .sort((a, b) => {
                    if (a.status === 'CONCLUIDO' && b.status !== 'CONCLUIDO') return 1;
                    if (a.status !== 'CONCLUIDO' && b.status === 'CONCLUIDO') return -1;
                    if (a.dias_abertos !== null && b.dias_abertos !== null) {
                        return b.dias_abertos - a.dias_abertos;
                    }
                    return a.unidade.localeCompare(b.unidade) || a.os - b.os;
                })
                .map(t => {
                    const statusClass = 'status-' + t.status.toLowerCase().replace(/\s+/g, '-').replace(/[^\w-]/g, '');
                    let diasHTML = '';
                    if (t.dias_abertos !== null) {
                        let cor = '#0ca30c';
                        let fundo = '#d4edda';
                        if (t.dias_abertos > 60) {
                            cor = '#e34948';
                            fundo = '#f8d7da';
                        } else if (t.dias_abertos > 30) {
                            cor = '#fab219';
                            fundo = '#fff3cd';
                        }
                        diasHTML = `<td style="font-weight: 600; color: ${cor}; background: ${fundo}; border-radius: 4px; padding: 6px 10px; text-align: center;">${t.dias_abertos}d</td>`;
                    } else {
                        diasHTML = `<td style="text-align: center; color: #999;">-</td>`;
                    }

                    const valorFmt = new Intl.NumberFormat('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2}).format(t.valor);

                    return `
                        <tr>
                            <td style="font-weight: 600;">#${t.os}</td>
                            <td>${t.nr}</td>
                            <td>${t.unidade}</td>
                            <td style="max-width: 280px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${t.descricao}">${t.descricao}</td>
                            <td><span class="status-badge ${statusClass}">${t.status}</span></td>
                            <td>${t.data_envio || '-'}</td>
                            ${diasHTML}
                            <td style="text-align: right; font-weight: 600;">R$ ${valorFmt}</td>
                        </tr>
                    `;
                }).join('');

            document.getElementById('tableBody').innerHTML = html;
        }

        function render() {
            renderKPIs();
            renderContracts();
            renderTempoAberto();
            renderCharts();
            renderPaymentPanel();
            renderTable();
        }

        init();
    </script>
</body>
</html>"""

st.subheader("📊 Dashboard Interativo")

html_content = HTML_TEMPLATE.replace('__DATA_PLACEHOLDER__', json.dumps(data, ensure_ascii=False))

st.components.v1.html(html_content, height=4400, scrolling=True)

st.subheader("📥 Download dos Arquivos")
col1, col2 = st.columns(2)

with col1:
    st.download_button(
        label="📄 Download HTML Completo",
        data=html_content,
        file_name=f"painel_manutencao_pagamentos_{datetime.now().strftime('%Y%m%d')}.html",
        mime="text/html"
    )

with col2:
    st.info("✅ Painel de Pagamentos e Chamados gerado com sucesso!")
