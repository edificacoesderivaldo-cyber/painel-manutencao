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

# Detecção de coluna de mês de emissão / competência
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
            max-width: 100%;
            background: white;
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

        tfoot td {
            padding: 14px 12px;
            font-weight: 700;
            border-top: 2px solid #333;
            background: #f1f5f9;
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

        .units-split-container {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 24px;
            margin-top: 10px;
        }

        @media (max-width: 900px) {
            .units-split-container {
                grid-template-columns: 1fr;
            }
        }

        .units-column {
            background: #ffffff;
            border-radius: 10px;
            border: 1px solid var(--border);
            padding: 16px;
        }

        .units-column-sesi {
            border-top: 4px solid #0d6efd;
            background: #fbfdff;
        }

        .units-column-senai {
            border-top: 4px solid #e65100;
            background: #fffbf9;
        }

        .units-column-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding-bottom: 12px;
            margin-bottom: 14px;
            border-bottom: 2px solid var(--border);
        }

        .units-column-title {
            font-size: 16px;
            font-weight: 700;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .units-column-title.sesi { color: #0d6efd; }
        .units-column-title.senai { color: #e65100; }

        .units-column-subtotal {
            font-size: 12px;
            font-weight: 600;
            padding: 4px 10px;
            border-radius: 20px;
        }

        .units-column-subtotal.sesi {
            background: #e7f3ff;
            color: #0a58ca;
        }

        .units-column-subtotal.senai {
            background: #fff0e6;
            color: #c44000;
        }

        .unit-card {
            padding: 14px 16px;
            background: white;
            border: 1px solid var(--border);
            border-radius: 8px;
            box-shadow: 0 1px 2px rgba(0,0,0,0.04);
            transition: transform 0.2s, box-shadow 0.2s;
        }

        .unit-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 10px rgba(0,0,0,0.08);
        }

        .unit-card-sesi {
            border-left: 4px solid #0d6efd;
        }

        .unit-card-senai {
            border-left: 4px solid #e65100;
        }

        .unit-card-name {
            font-weight: 700;
            color: #212529;
            margin-bottom: 10px;
            font-size: 13px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .unit-card-info {
            display: flex;
            justify-content: space-between;
            align-items: flex-end;
        }

        .unit-card-label {
            font-size: 11px;
            color: #6c757d;
            margin-bottom: 2px;
            text-transform: uppercase;
            font-weight: 600;
        }

        .unit-card-value {
            font-weight: 700;
            color: #212529;
            font-size: 18px;
            line-height: 1;
        }

        .unit-card-value-right {
            font-weight: 700;
            font-size: 14px;
        }

        .unit-card-value-right.sesi { color: #0d6efd; }
        .unit-card-value-right.senai { color: #e65100; }

        .footer {
            margin-top: 40px;
            padding: 20px;
            text-align: center;
            color: #999;
            font-size: 12px;
        }

        .btn-action-copy {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 9px 16px;
            background: #0d6efd;
            color: #ffffff;
            border: none;
            border-radius: 6px;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
            box-shadow: 0 2px 4px rgba(13, 110, 253, 0.2);
        }

        .btn-action-copy:hover {
            background: #0b5ed7;
            transform: translateY(-1px);
            box-shadow: 0 4px 8px rgba(13, 110, 253, 0.3);
        }

        .btn-action-copy:active {
            transform: translateY(0);
        }

        .copy-toast {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            background: #d1e7dd;
            color: #0f5132;
            padding: 8px 14px;
            border-radius: 6px;
            font-size: 13px;
            font-weight: 600;
            border: 1px solid #badbcc;
            opacity: 0;
            transition: opacity 0.3s ease;
        }

        .copy-toast.show {
            opacity: 1;
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
        <!-- PAINEL: CONTROLE DE PAGAMENTO / LIBERAÇÃO DE NFE        -->
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
                    <tfoot id="payTableFoot"></tfoot>
                </table>
            </div>

            <!-- Faixa de Resumo do Fechamento de Faturamento -->
            <div id="paySummaryFooter" style="margin-top: 18px; padding: 14px 20px; background: #e8f4fd; border-radius: 8px; border-left: 5px solid #0d6efd; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px;">
                <div id="paySummaryFooterText" style="font-size: 14px; font-weight: 600; color: #0a58ca;">
                    Fechamento de Faturamento
                </div>
                <div id="paySummaryFooterVal" style="font-size: 18px; font-weight: 800; color: #0ca30c;">
                    R$ 0,00
                </div>
            </div>

            <!-- Botão de Cópia Única para E-mail -->
            <div style="margin-top: 16px; padding: 14px 18px; background: #f8fafc; border: 1px dashed #cbd5e1; border-radius: 8px; display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 12px;">
                <div style="display: flex; gap: 10px; flex-wrap: wrap; align-items: center;">
                    <button class="btn-action-copy" onclick="copiarTabelaEmail()" title="Copia apenas as O.S. liberadas para emissão de NF-e e formata para colar no e-mail ou Excel">
                        📋 Copiar Tabela p/ E-mail
                    </button>
                </div>
                <div id="copyToast" class="copy-toast">
                    ✅ Tabela copiada! Pressione Ctrl + V no seu e-mail.
                </div>
            </div>
        </div>

        <!-- ======================================================= -->
        <!-- TABELA COMPLETA DE CHAMADOS GERAL                       -->
        <!-- ======================================================= -->
        <div class="card">
            <div class="card-title">📋 Lista Completa de Chamados</div>
            
            <div style="margin-bottom: 16px;">
                <div style="font-weight: 600; font-size: 13px; color: #666; margin-bottom: 10px;">Filtrar por Status na Tabela</div>
                <div class="filter-group" id="listStatusFilter"></div>
            </div>

            <!-- Controles de Busca: O.S. e Unidade -->
            <div style="margin-bottom: 16px; display: flex; gap: 14px; align-items: flex-end; flex-wrap: wrap;">
                <div style="flex: 1; min-width: 200px;">
                    <label style="display: block; margin-bottom: 6px; font-size: 12px; font-weight: 600; color: #666;">🔍 Buscar por O.S.</label>
                    <input type="text" class="filter-input" id="osSearch" placeholder="Ex: 194876" oninput="filterByOS(this.value)" />
                </div>

                <div style="flex: 1.5; min-width: 260px;">
                    <label style="display: block; margin-bottom: 6px; font-size: 12px; font-weight: 600; color: #666;">🏢 Filtrar por Unidade</label>
                    <select class="filter-input" id="unitSelect" onchange="filterByListUnit(this.value)" style="cursor: pointer;">
                        <option value="">Carregando unidades...</option>
                    </select>
                </div>

                <div style="font-size: 13px; color: #666; padding-bottom: 8px;">
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
            filters: { status: [], unidade: [], os: '', listStatus: '', listUnit: '' },
            payFilters: { mes: '', statusPagamento: '', casa: '' }
        };

        function init() {
            state.tickets = DATA.chamados;
            document.getElementById('updateTime').textContent = new Date().toLocaleString('pt-BR');
            document.getElementById('totalTickets').textContent = state.tickets.length;
            renderFilters();
            renderListStatusFilter();
            renderListUnitFilter();
            renderPaymentCasaFilters();
            renderPaymentMonthFilters();
            render();
        }

        function getFiltered() {
            return state.tickets.filter(t => {
                const matchOS = state.filters.os === '' || t.os.toString().includes(state.filters.os);
                const matchStatus = state.filters.status.length === 0 || state.filters.status.includes(t.status);
                const matchUnidadeGlobal = state.filters.unidade.length === 0 || state.filters.unidade.includes(t.unidade);
                const matchListStatus = state.filters.listStatus === '' || t.status === state.filters.listStatus;
                const matchListUnit = state.filters.listUnit === '' || t.unidade === state.filters.listUnit;
                return matchOS && matchStatus && matchUnidadeGlobal && matchListStatus && matchListUnit;
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
                    if (a.status_pagamento === 'LIBERADO P/ NFE' && b.status_pagamento !== 'LIBERADO P/ NFE') return -1;
                    if (a.status_pagamento !== 'LIBERADO P/ NFE' && b.status_pagamento === 'LIBERADO P/ NFE') return 1;
                    return b.valor - a.valor;
                })
                .map(t => {
                    let badgeClass = 'status-aguardando-orcamento';
                    let icone = '⏳';
                    if (t.status_pagamento === 'LIBERADO P/ NFE') {
                        badgeClass = 'status-concluido';
                        icone = '✅';
                    } else if (t.status_pagamento === 'EM MEDIÇÃO') {
                        badgeClass = 'status-em-execucao';
                        icone = '🔄';
                    } else if (t.status_pagamento === 'BLOQUEADO') {
                        badgeClass = 'status-paralisado';
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

            // Linha Final de Totalização (tfoot)
            const footHTML = `
                <tr>
                    <td colspan="4" style="font-size: 13px; text-transform: uppercase; color: #1e293b;">
                        📌 TOTAL GERAL DO FILTRO: <span style="color: #0d6efd;">${state.payFilters.casa || 'SESI + SENAI'}</span> | <span style="color: #055160;">${state.payFilters.mes || 'TODOS OS MESES'}</span>
                    </td>
                    <td colspan="4" style="text-align: right; color: #475569; font-size: 12px;">
                        Liberado para NFE: <strong style="color: #0ca30c;">R$ ${fmt(valLiberado)}</strong> (${liberados.length} O.S.) &nbsp;|&nbsp; Total da Seleção (${payList.length} O.S.):
                    </td>
                    <td style="text-align: right; font-size: 15px; color: #0f172a; background: #e2e8f0;">
                        R$ ${fmt(valTotal)}
                    </td>
                </tr>
            `;
            document.getElementById('payTableFoot').innerHTML = footHTML;

            const casaNome = state.payFilters.casa ? state.payFilters.casa : 'SESI + SENAI';
            const mesNome = state.payFilters.mes ? state.payFilters.mes : 'Todos os Meses';
            document.getElementById('paySummaryFooterText').innerHTML = `
                🏷️ <strong>Resumo do Faturamento:</strong> Entidade: <u>${casaNome}</u> | Mês de Competência: <u>${mesNome}</u> | Aptos para NFE: <strong>${liberados.length} chamados</strong>
            `;
            document.getElementById('paySummaryFooterVal').innerHTML = `
                Total Faturamento Liberado: R$ ${fmt(valLiberado)}
            `;
        }

        function showCopyToast(msg) {
            const toast = document.getElementById('copyToast');
            toast.textContent = msg;
            toast.classList.add('show');
            setTimeout(() => {
                toast.classList.remove('show');
            }, 4000);
        }

        function copiarTabelaEmail() {
            const payList = getPayFiltered();
            const liberados = payList.filter(t => t.status_pagamento === 'LIBERADO P/ NFE');

            if (liberados.length === 0) {
                showCopyToast('⚠️ Nenhuma O.S. liberada para emissão de NF-e neste filtro.');
                return;
            }

            const valLiberado = liberados.reduce((acc, t) => acc + (t.valor || 0), 0);
            const fmt = v => new Intl.NumberFormat('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2}).format(v);

            const casaTxt = state.payFilters.casa ? state.payFilters.casa : 'SESI / SENAI';
            const mesTxt = state.payFilters.mes ? state.payFilters.mes : 'Todos os Meses';

            let html = `
                <div style="font-family: Arial, sans-serif; color: #333333; line-height: 1.5;">
                    <p style="font-size: 14px; margin-bottom: 8px;">
                        Prezados,<br><br>
                        Segue a relação de serviços conferidos e <strong>LIBERADOS PARA EMISSÃO DE NOTA FISCAL ELETRÔNICA (NF-e)</strong> referente à medição:
                    </p>
                    <p style="font-size: 13px; color: #555555; margin-bottom: 14px;">
                        • <strong>Entidade / CNPJ:</strong> ${casaTxt}<br>
                        • <strong>Mês de Competência:</strong> ${mesTxt}<br>
                        • <strong>Quantidade de O.S. Liberadas:</strong> ${liberados.length}<br>
                        • <strong>Valor Total Liberado para NF-e:</strong> <span style="color: #0ca30c; font-weight: bold;">R$ ${fmt(valLiberado)}</span>
                    </p>
                    <table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; font-size: 12px; border: 1px solid #cccccc; font-family: Arial, sans-serif;">
                        <thead>
                            <tr style="background-color: #0d6efd; color: #ffffff; text-align: left;">
                                <th style="padding: 8px; border: 1px solid #b8daff;">O.S</th>
                                <th style="padding: 8px; border: 1px solid #b8daff;">NR</th>
                                <th style="padding: 8px; border: 1px solid #b8daff;">Entidade / CNPJ</th>
                                <th style="padding: 8px; border: 1px solid #b8daff;">Unidade</th>
                                <th style="padding: 8px; border: 1px solid #b8daff;">Descrição do Serviço</th>
                                <th style="padding: 8px; border: 1px solid #b8daff;">Mês Competência</th>
                                <th style="padding: 8px; border: 1px solid #b8daff;">Liberação p/ NF-e</th>
                                <th style="padding: 8px; border: 1px solid #b8daff; text-align: right;">Valor Autorizado (R$)</th>
                            </tr>
                        </thead>
                        <tbody>
            `;

            let plain = `RELAÇÃO DE SERVIÇOS LIBERADOS PARA EMISSÃO DE NF-E\nEntidade: ${casaTxt} | Competência: ${mesTxt}\n\n`;
            plain += `O.S\tNR\tCasa\tUnidade\tDescrição\tCompetência\tStatus Liberação\tValor Autorizado\n`;

            liberados.forEach((t, i) => {
                const bg = i % 2 === 0 ? '#ffffff' : '#f8f9fa';

                html += `
                    <tr style="background-color: ${bg};">
                        <td style="padding: 8px; border: 1px solid #dddddd; font-weight: bold; color: #0d6efd;">#${t.os}</td>
                        <td style="padding: 8px; border: 1px solid #dddddd;">${t.nr}</td>
                        <td style="padding: 8px; border: 1px solid #dddddd; font-weight: bold;">${t.casa}</td>
                        <td style="padding: 8px; border: 1px solid #dddddd;"><strong>${t.unidade}</strong></td>
                        <td style="padding: 8px; border: 1px solid #dddddd;">${t.descricao}</td>
                        <td style="padding: 8px; border: 1px solid #dddddd;">${t.mes_emissao}</td>
                        <td style="padding: 8px; border: 1px solid #dddddd; text-align: center;">
                            <span style="background: #d1e7dd; color: #0f5132; padding: 4px 8px; border-radius: 4px; font-weight: bold; font-size: 11px; border: 1px solid #badbcc;">
                                ✅ LIBERADO P/ NFE
                            </span>
                        </td>
                        <td style="padding: 8px; border: 1px solid #dddddd; text-align: right; font-weight: bold; color: #0ca30c;">R$ ${fmt(t.valor)}</td>
                    </tr>
                `;

                plain += `${t.os}\t${t.nr}\t${t.casa}\t${t.unidade}\t${t.descricao}\t${t.mes_emissao}\tLIBERADO P/ NFE\tR$ ${fmt(t.valor)}\n`;
            });

            html += `
                        </tbody>
                        <tfoot>
                            <tr style="background-color: #e8f4fd; font-weight: bold;">
                                <td colspan="7" style="padding: 10px; border: 1px solid #b8daff; text-align: right; font-size: 13px;">
                                    TOTAL AUTORIZADO PARA EMISSÃO DE NF-E (${liberados.length} O.S.):
                                </td>
                                <td style="padding: 10px; border: 1px solid #b8daff; text-align: right; font-size: 14px; color: #0ca30c;">
                                    R$ ${fmt(valLiberado)}
                                </td>
                            </tr>
                        </tfoot>
                    </table>
                    <p style="font-size: 12px; color: #777777; margin-top: 12px;">
                        * Emitir a Nota Fiscal Eletrônica (NF-e) no CNPJ correspondente contendo estritamente os serviços e valores discriminados acima.
                    </p>
                </div>
            `;

            plain += `\nTOTAL LIBERADO PARA NF-E: R$ ${fmt(valLiberado)} (${liberados.length} O.S.)`;

            try {
                const blobHtml = new Blob([html], { type: 'text/html' });
                const blobText = new Blob([plain], { type: 'text/plain' });
                const data = [new ClipboardItem({ 'text/html': blobHtml, 'text/plain': blobText })];

                navigator.clipboard.write(data).then(() => {
                    showCopyToast(`📋 Tabela copiada (${liberados.length} O.S. liberadas)! Cole no e-mail (Ctrl+V).`);
                }).catch(() => {
                    navigator.clipboard.writeText(plain).then(() => {
                        showCopyToast(`📋 Dados copiados (${liberados.length} O.S.)! Cole no e-mail (Ctrl+V).`);
                    });
                });
            } catch (err) {
                navigator.clipboard.writeText(plain).then(() => {
                    showCopyToast(`📋 Dados copiados! Cole no e-mail (Ctrl+V).`);
                });
            }
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

        function filterByListUnit(unit) {
            state.filters.listUnit = unit;
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

        function renderListUnitFilter() {
            const select = document.getElementById('unitSelect');
            if (!select) return;

            const todasUnidades = [...new Set(state.tickets.map(t => t.unidade))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
            const sesi = todasUnidades.filter(u => u.toUpperCase().includes('SESI'));
            const senai = todasUnidades.filter(u => u.toUpperCase().includes('SENAI'));
            const outras = todasUnidades.filter(u => !u.toUpperCase().includes('SESI') && !u.toUpperCase().includes('SENAI'));

            let html = `<option value="">🏢 Todas as Unidades (${todasUnidades.length})</option>`;

            if (sesi.length > 0) {
                html += `<optgroup label="🔵 UNIDADES SESI">`;
                sesi.forEach(u => {
                    const sel = state.filters.listUnit === u ? 'selected' : '';
                    html += `<option value="${u}" ${sel}>${u}</option>`;
                });
                html += `</optgroup>`;
            }

            if (senai.length > 0) {
                html += `<optgroup label="🟠 UNIDADES SENAI">`;
                senai.forEach(u => {
                    const sel = state.filters.listUnit === u ? 'selected' : '';
                    html += `<option value="${u}" ${sel}>${u}</option>`;
                });
                html += `</optgroup>`;
            }

            if (outras.length > 0) {
                html += `<optgroup label="⚪ OUTRAS">`;
                outras.forEach(u => {
                    const sel = state.filters.listUnit === u ? 'selected' : '';
                    html += `<option value="${u}" ${sel}>${u}</option>`;
                });
                html += `</optgroup>`;
            }

            select.innerHTML = html;
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

            renderUnitCards(unitData, allStatuses);
        }

        function renderUnitCards(unitData, allStatuses) {
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

            const todasUnidades = Object.keys(unitData);
            const unidadesSesi = todasUnidades
                .filter(u => u.toUpperCase().includes('SESI'))
                .sort((a, b) => a.localeCompare(b, 'pt-BR'));

            const unidadesSenai = todasUnidades
                .filter(u => u.toUpperCase().includes('SENAI'))
                .sort((a, b) => a.localeCompare(b, 'pt-BR'));

            const unidadesOutras = todasUnidades
                .filter(u => !u.toUpperCase().includes('SESI') && !u.toUpperCase().includes('SENAI'))
                .sort((a, b) => a.localeCompare(b, 'pt-BR'));

            if (unidadesOutras.length > 0) {
                unidadesSesi.push(...unidadesOutras);
            }

            const fmt = v => new Intl.NumberFormat('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2}).format(v);

            let totalChamadosSesi = 0;
            let totalInvestidoSesi = 0;
            unidadesSesi.forEach(u => {
                totalChamadosSesi += Object.values(unitData[u].status).reduce((s, v) => s + v, 0);
                totalInvestidoSesi += unitData[u].valor;
            });

            let totalChamadosSenai = 0;
            let totalInvestidoSenai = 0;
            unidadesSenai.forEach(u => {
                totalChamadosSenai += Object.values(unitData[u].status).reduce((s, v) => s + v, 0);
                totalInvestidoSenai += unitData[u].valor;
            });

            const buildCard = (unit, casaClass) => {
                const total = Object.values(unitData[unit].status).reduce((s, v) => s + v, 0);
                const valor = unitData[unit].valor;
                return `
                    <div class="unit-card unit-card-${casaClass}">
                        <div class="unit-card-name" title="${unit}">${unit}</div>
                        <div class="unit-card-info">
                            <div>
                                <div class="unit-card-label">Chamados</div>
                                <div class="unit-card-value">${total}</div>
                            </div>
                            <div style="text-align: right;">
                                <div class="unit-card-label">Investido</div>
                                <div class="unit-card-value-right ${casaClass}">R$ ${fmt(valor)}</div>
                            </div>
                        </div>
                    </div>
                `;
            };

            const cardsSesiHTML = unidadesSesi.length > 0 
                ? unidadesSesi.map(u => buildCard(u, 'sesi')).join('')
                : '<p style="color: #999; font-size: 13px; text-align: center; padding: 20px;">Nenhuma unidade SESI localizada.</p>';

            const cardsSenaiHTML = unidadesSenai.length > 0 
                ? unidadesSenai.map(u => buildCard(u, 'senai')).join('')
                : '<p style="color: #999; font-size: 13px; text-align: center; padding: 20px;">Nenhuma unidade SENAI localizada.</p>';

            let layoutHTML = legendaHTML + `
                <div class="units-split-container">
                    <!-- Coluna SESI (Esquerda / Azul) -->
                    <div class="units-column units-column-sesi">
                        <div class="units-column-header">
                            <div class="units-column-title sesi">
                                🔵 UNIDADES SESI
                            </div>
                            <div class="units-column-subtotal sesi">
                                ${totalChamadosSesi} chamados &nbsp;|&nbsp; R$ ${fmt(totalInvestidoSesi)}
                            </div>
                        </div>
                        <div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(210px, 1fr)); gap: 12px;">
                            ${cardsSesiHTML}
                        </div>
                    </div>

                    <!-- Coluna SENAI (Direita / Laranja) -->
                    <div class="units-column units-column-senai">
                        <div class="units-column-header">
                            <div class="units-column-title senai">
                                🟠 UNIDADES SENAI
                            </div>
                            <div class="units-column-subtotal senai">
                                ${totalChamadosSenai} chamados &nbsp;|&nbsp; R$ ${fmt(totalInvestidoSenai)}
                            </div>
                        </div>
                        <div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(210px, 1fr)); gap: 12px;">
                            ${cardsSenaiHTML}
                        </div>
                    </div>
                </div>
            `;

            document.getElementById('unitLegendBottom').innerHTML = layoutHTML;
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
                            <td style="font-weight: 600; color: #0d6efd;">#${t.os}</td>
                            <td>${t.nr}</td>
                            <td><strong>${t.unidade}</strong></td>
                            <td style="max-width: 280px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${t.descricao}">${t.descricao}</td>
                            <td><span class="status-badge ${statusClass}">${t.status}</span></td>
                            <td>${t.data_envio || '-'}</td>
                            ${diasHTML}
                            <td style="text-align: right; font-weight: 600;">R$ ${valorFmt}</td>
                        </tr>
                    `;
                }).join('');

            if (filtered.length === 0) {
                html = '<tr><td colspan="8" style="text-align:center; padding: 24px; color: #999;">Nenhum chamado localizado para os filtros selecionados.</td></tr>';
            }

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
