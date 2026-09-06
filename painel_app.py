import streamlit as st
import pandas as pd
import json
from datetime import datetime
from io import BytesIO
import base64

st.set_page_config(page_title="Painel de Manutenção Predial", layout="wide", initial_sidebar_state="collapsed")

st.title("📊 Painel de Manutenção Predial")
st.markdown("**Análise de chamados — SESI e SENAI**  |  Envie sua planilha Excel para gerar o relatório")

# ============== UPLOAD ==============
uploaded_file = st.file_uploader("📎 Envie o arquivo Excel (CONTROLE_DE_O_S.xlsx)", type=["xlsx", "xls"])

if not uploaded_file:
    st.info("👉 Clique no botão acima para fazer upload da sua planilha Excel com os chamados de manutenção.")
    st.stop()

# ============== CARREGAR E PROCESSAR DADOS ==============
try:
    df_os = pd.read_excel(uploaded_file, sheet_name='O.S', header=2)
    df_saldo = pd.read_excel(uploaded_file, sheet_name='SALDO', header=0)
except Exception as e:
    st.error(f"❌ Erro ao carregar arquivo: {e}")
    st.stop()

df_os = df_os.dropna(subset=['O.S']).copy()

# ============== NORMALIZAR DADOS ==============
def norm_status(s):
    if pd.isna(s) or s == '': return "SEM STATUS"
    s = str(s).strip().upper()
    mapping = {"PARALIZADO": "PARALISADO", "CONCLUÍDO": "CONCLUIDO"}
    return mapping.get(s, s)

def norm_unidade(u):
    if pd.isna(u): return "NÃO INFORMADA"
    return str(u).strip().upper().replace("GUARA", "GUARÁ")

df_os['STATUS'] = df_os['STATUS '].apply(norm_status)
df_os['UNIDADE'] = df_os['UNIDADE'].apply(norm_unidade)

# Coluna de valor
valor_col = next((c for c in df_os.columns if 'VALOR' in c.upper() and 'INICIAL' in c.upper()), None)
if valor_col:
    df_os[valor_col] = pd.to_numeric(df_os[valor_col], errors='coerce')

df_os['NR'] = pd.to_numeric(df_os['NR'], errors='coerce').fillna(0).astype(int)
df_os['O.S'] = pd.to_numeric(df_os['O.S'], errors='coerce').fillna(0).astype(int)

# Data de envio
data_col = next((c for c in df_os.columns if 'DATA' in c.upper() and 'ENVIO' in c.upper()), None)
if data_col:
    df_os[data_col] = pd.to_datetime(df_os[data_col], errors='coerce')

# ============== CALCULAR MÉTRICAS ==============
hoje = datetime.now()
total_chamados = len(df_os)
concluido = len(df_os[df_os['STATUS'] == 'CONCLUIDO'])
em_execucao = len(df_os[df_os['STATUS'] == 'EM EXECUÇÃO'])
paralisado = len(df_os[df_os['STATUS'] == 'PARALISADO'])

valor_total = float(df_os[valor_col].sum() if valor_col else 0)

# Abertos com mais de 30 dias
abertos_30 = 0
if data_col:
    for _, row in df_os.iterrows():
        if row['STATUS'] != 'CONCLUIDO' and pd.notna(row[data_col]):
            dias = (hoje - pd.to_datetime(row[data_col])).days
            if dias > 30:
                abertos_30 += 1

# ============== PREPARAR DADOS JSON ==============
data = {
    'chamados': [],
    'status_dist': {},
    'unidades': {},
    'totais': {},
    'contracts': {
        'sesi': {'contrato': 1440000, 'utilizado': 0, 'saldo': 1188238.58},
        'senai': {'contrato': 1440000, 'utilizado': 0, 'saldo': 1066830.71}
    }
}

for _, row in df_os.iterrows():
    valor = float(row[valor_col]) if valor_col and pd.notna(row[valor_col]) else 0
    status = str(row['STATUS'])
    dias_abertos = None
    if status != 'CONCLUIDO' and data_col and pd.notna(row[data_col]):
        dias_abertos = (hoje - pd.to_datetime(row[data_col])).days
    
    data['chamados'].append({
        'nr': int(row['NR']),
        'os': int(row['O.S']),
        'unidade': str(row['UNIDADE']),
        'status': status,
        'descricao': str(row.get('DESCRIÇÃO DO SERVIÇO', '')).strip(),
        'valor': round(valor, 2),
        'data_envio': str(row[data_col]).split()[0] if data_col and pd.notna(row[data_col]) else '',
        'dias_abertos': dias_abertos
    })

for status, count in df_os['STATUS'].value_counts().items():
    data['status_dist'][str(status)] = int(count)

for unit in sorted(df_os['UNIDADE'].unique()):
    df_u = df_os[df_os['UNIDADE'] == unit]
    data['unidades'][str(unit)] = {
        'total': len(df_u),
        'valor': round(float(df_u[valor_col].sum() if valor_col else 0), 2),
        'com_orcamento': len(df_u[df_u[valor_col] > 0]) if valor_col else 0,
        'status': {str(k): int(v) for k, v in df_u['STATUS'].value_counts().items()}
    }

df_sesi = df_os[df_os['UNIDADE'].str.contains('SESI', case=False, na=False)]
df_senai = df_os[df_os['UNIDADE'].str.contains('SENAI', case=False, na=False)]
data['contracts']['sesi']['utilizado'] = round(float(df_sesi[valor_col].sum() if valor_col else 0), 2)
data['contracts']['senai']['utilizado'] = round(float(df_senai[valor_col].sum() if valor_col else 0), 2)
data['contracts']['sesi']['saldo'] = data['contracts']['sesi']['contrato'] - data['contracts']['sesi']['utilizado']
data['contracts']['senai']['saldo'] = data['contracts']['senai']['contrato'] - data['contracts']['senai']['utilizado']

data['totais'] = {
    'total_chamados': total_chamados,
    'concluido': concluido,
    'em_execucao': em_execucao,
    'paralisado': paralisado,
    'abertos_30dias': abertos_30,
    'valor_total': round(valor_total, 2)
}

# ============== TEMPLATE HTML ==============
HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Painel de Manutenção Predial - SESI/SENAI</title>
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
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
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
            font-size: 24px;
            margin: 0;
        }
        .header .subtitle {
            font-size: 12px;
            color: #999;
            margin-top: 4px;
        }
        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 16px;
            margin-bottom: 30px;
        }
        .kpi-card {
            background: white;
            padding: 20px;
            border-radius: 8px;
            border-left: 4px solid var(--primary);
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .kpi-label {
            font-size: 12px;
            color: #999;
            margin-bottom: 8px;
            text-transform: uppercase;
        }
        .kpi-value {
            font-size: 28px;
            font-weight: bold;
            color: var(--primary);
        }
        .grid-2col {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(500px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }
        .card {
            background: white;
            padding: 24px;
            border-radius: 8px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .card-title {
            font-size: 16px;
            font-weight: 600;
            margin-bottom: 16px;
            color: var(--text);
        }
        .chart-wrapper {
            height: 300px;
            position: relative;
        }
        .contract-stat {
            display: flex;
            justify-content: space-between;
            margin-bottom: 12px;
            font-size: 14px;
        }
        .contract-label {
            color: #666;
        }
        .contract-value {
            font-weight: 600;
            color: var(--primary);
        }
        .progress-container {
            margin: 16px 0;
        }
        .progress-label {
            display: flex;
            justify-content: space-between;
            font-size: 12px;
            margin-bottom: 8px;
            color: #666;
        }
        .progress-bar {
            width: 100%;
            height: 24px;
            background: #e0e0e0;
            border-radius: 12px;
            overflow: hidden;
        }
        .progress-fill {
            height: 100%;
            background: linear-gradient(90deg, var(--success), var(--info));
            transition: width 0.3s;
        }
        .unit-legend {
            display: flex;
            gap: 20px;
            flex-wrap: wrap;
            margin-bottom: 20px;
            justify-content: center;
        }
        .legend-item {
            display: flex;
            align-items: center;
            gap: 8px;
            font-size: 12px;
        }
        .legend-color {
            width: 20px;
            height: 20px;
            border-radius: 2px;
        }
        .unit-cards {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(220px, 1fr));
            gap: 16px;
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
        .unit-card-left {
            flex: 1;
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
        .unit-card-right {
            text-align: right;
        }
        .unit-card-value-right {
            font-weight: 600;
            color: var(--primary);
            font-size: 14px;
        }
        @media print {
            body { background: white; }
            .container { padding: 10px; }
            .card { box-shadow: none; border: 1px solid #ddd; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div>
                <h1>📊 Painel de Manutenção Predial</h1>
                <div class="subtitle">SESI e SENAI — Gestão de Chamados de Manutenção</div>
            </div>
        </div>

        <!-- KPIs -->
        <div class="kpi-grid">
            <div class="kpi-card">
                <div class="kpi-label">Total de Chamados</div>
                <div class="kpi-value" id="kpi-total">0</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Concluídos</div>
                <div class="kpi-value" id="kpi-concluido">0</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Em Execução</div>
                <div class="kpi-value" id="kpi-execucao">0</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Paralisados</div>
                <div class="kpi-value" id="kpi-paralisado">0</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Abertos >30 dias</div>
                <div class="kpi-value" id="kpi-30dias">0</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Valor Investido</div>
                <div class="kpi-value" id="kpi-valor">R$ 0</div>
            </div>
        </div>

        <!-- Saldo em Contrato -->
        <div class="grid-2col">
            <div class="card">
                <div class="card-title">💰 Saldo em Contrato</div>
                <div id="contracts"></div>
            </div>
            <div class="card">
                <div class="card-title">⏱️ Tempo Médio em Aberto</div>
                <div id="tempoAberto"></div>
            </div>
        </div>

        <!-- Status Geral -->
        <div class="grid-2col">
            <div class="card">
                <div class="card-title">📈 Distribuição de Status</div>
                <div class="chart-wrapper">
                    <canvas id="statusChart"></canvas>
                </div>
            </div>
        </div>

        <!-- Chamados por Unidade -->
        <div class="card">
            <div class="card-title">🏢 Chamados por Unidade (com investimento)</div>
            <div class="chart-wrapper" style="height: 500px; margin-bottom: 20px;">
                <canvas id="unitChart"></canvas>
            </div>
            <div id="unitLegendBottom"></div>
        </div>
    </div>

    <script>
        const DATA = __DATA_PLACEHOLDER__;
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

        function renderKPIs() {
            const totais = DATA.totais;
            document.getElementById('kpi-total').textContent = totais.total_chamados;
            document.getElementById('kpi-concluido').textContent = totais.concluido;
            document.getElementById('kpi-execucao').textContent = totais.em_execucao;
            document.getElementById('kpi-paralisado').textContent = totais.paralisado;
            document.getElementById('kpi-30dias').textContent = totais.abertos_30dias;
            const valorFormatado = new Intl.NumberFormat('pt-BR', {style: 'currency', currency: 'BRL'}).format(totais.valor_total);
            document.getElementById('kpi-valor').textContent = valorFormatado;
        }

        function renderContracts() {
            const contracts = DATA.contracts;
            let html = '';
            for (const [key, data] of Object.entries(contracts)) {
                const pct = Math.round(data.saldo * 100 / data.contrato);
                const name = key === 'sesi' ? 'SESI' : 'SENAI';
                const utilizado = data.contrato - data.saldo;
                const utilizadoPct = 100 - pct;
                html += `
                    <div style="margin-bottom: 24px;">
                        <h3 style="font-size: 14px; margin-bottom: 12px; color: #3987e5; font-weight: 600; text-transform: uppercase;">${name}</h3>
                        <div style="display: flex; justify-content: space-between; margin-bottom: 8px; font-size: 14px;">
                            <span style="color: #666;">Valor do Contrato</span>
                            <span style="font-weight: 600; color: #3987e5;">R$ ${new Intl.NumberFormat('pt-BR').format(data.contrato)}</span>
                        </div>
                        <div style="display: flex; justify-content: space-between; margin-bottom: 8px; font-size: 14px;">
                            <span style="color: #666;">Utilizado</span>
                            <span style="font-weight: 600; color: #3987e5;">R$ ${new Intl.NumberFormat('pt-BR').format(utilizado)}</span>
                        </div>
                        <div style="margin: 16px 0;">
                            <div style="display: flex; justify-content: space-between; font-size: 12px; margin-bottom: 8px; color: #666;">
                                <span>Progresso</span>
                                <span>${utilizadoPct}% utilizado</span>
                            </div>
                            <div style="width: 100%; height: 24px; background: #e0e0e0; border-radius: 12px; overflow: hidden;">
                                <div style="height: 100%; width: ${utilizadoPct}%; background: linear-gradient(90deg, #0ca30c, #1baf7a); transition: width 0.3s;"></div>
                            </div>
                        </div>
                        <div style="display: flex; justify-content: space-between; margin-bottom: 0; font-size: 14px;">
                            <span style="color: #666;">Saldo Disponível</span>
                            <span style="font-weight: 600; color: #1baf7a;">R$ ${new Intl.NumberFormat('pt-BR').format(data.saldo)}</span>
                        </div>
                    </div>
                `;
            }
            document.getElementById('contracts').innerHTML = html;
        }

        function renderTempoAberto() {
            const abertos = DATA.chamados.filter(t => t.status !== 'CONCLUIDO' && t.dias_abertos !== null);
            if (abertos.length === 0) {
                document.getElementById('tempoAberto').innerHTML = '<p style="color: #999;">Nenhum chamado aberto</p>';
                return;
            }
            const mediaDias = Math.round(abertos.reduce((s, t) => s + (t.dias_abertos || 0), 0) / abertos.length);
            const range030 = abertos.filter(t => t.dias_abertos <= 30).length;
            const range3060 = abertos.filter(t => t.dias_abertos > 30 && t.dias_abertos <= 60).length;
            const range60plus = abertos.filter(t => t.dias_abertos > 60).length;
            const html = `
                <div style="text-align: center; margin-bottom: 24px; padding: 20px; background: #f8f9fa; border-radius: 8px;">
                    <div style="font-size: 12px; color: #666; margin-bottom: 8px; text-transform: uppercase;">Média de Dias em Aberto</div>
                    <div style="font-size: 48px; font-weight: bold; color: #3987e5; margin-bottom: 4px;">${mediaDias}</div>
                    <div style="font-size: 13px; color: #999;">dias (em ${abertos.length} chamados abertos)</div>
                </div>
                <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-top: 16px;">
                    <div style="background: #d4edda; border-radius: 8px; padding: 16px; text-align: center; border-left: 4px solid #0ca30c;">
                        <div style="font-size: 28px; font-weight: bold; color: #155724;">${range030}</div>
                        <div style="font-size: 12px; color: #155724; margin-top: 8px; font-weight: 600;">0-30 dias</div>
                    </div>
                    <div style="background: #fff3cd; border-radius: 8px; padding: 16px; text-align: center; border-left: 4px solid #fab219;">
                        <div style="font-size: 28px; font-weight: bold; color: #856404;">${range3060}</div>
                        <div style="font-size: 12px; color: #856404; margin-top: 8px; font-weight: 600;">30-60 dias</div>
                    </div>
                    <div style="background: #f8d7da; border-radius: 8px; padding: 16px; text-align: center; border-left: 4px solid #e34948;">
                        <div style="font-size: 28px; font-weight: bold; color: #721c24;">${range60plus}</div>
                        <div style="font-size: 12px; color: #721c24; margin-top: 8px; font-weight: 600;">>60 dias</div>
                    </div>
                </div>
            `;
            document.getElementById('tempoAberto').innerHTML = html;
        }

        function renderCharts() {
            // Status Chart
            const statusCounts = {};
            DATA.chamados.forEach(t => {
                statusCounts[t.status] = (statusCounts[t.status] || 0) + 1;
            });
            const statusLabels = Object.keys(statusCounts).sort();
            const statusValues = statusLabels.map(s => statusCounts[s]);
            const statusColors = statusLabels.map(s => COLOR_MAP[s]);

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
                    plugins: { legend: { position: 'bottom' } }
                }
            });

            // Unit Chart
            const unitData = {};
            DATA.chamados.forEach(t => {
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

            const allStatuses = [...new Set(DATA.chamados.map(t => t.status))].sort();
            const datasets = allStatuses.map(status => ({
                label: status,
                data: unitLabels.map(unit => unitData[unit].status[status] || 0),
                backgroundColor: COLOR_MAP[status],
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

            renderUnitTable(unitData, unitLabels, allStatuses);
        }

        function renderUnitTable(unitData, unitLabels, allStatuses) {
            let legendaHTML = '<div style="display: flex; gap: 20px; flex-wrap: wrap; margin-bottom: 20px; justify-content: center;">';
            allStatuses.forEach(status => {
                const color = COLOR_MAP[status];
                legendaHTML += `
                    <div style="display: flex; align-items: center; gap: 8px; font-size: 12px;">
                        <div style="width: 20px; height: 20px; background: ${color}; border-radius: 2px;"></div>
                        <span>${status}</span>
                    </div>
                `;
            });
            legendaHTML += '</div>';

            let cardsHTML = legendaHTML + '<div class="unit-cards">';
            unitLabels.forEach(unit => {
                const total = Object.values(unitData[unit].status).reduce((s, v) => s + v, 0);
                const valor = unitData[unit].valor;
                const valorFormatado = new Intl.NumberFormat('pt-BR', {style: 'currency', currency: 'BRL'}).format(valor).replace('R$', '').trim();

                cardsHTML += `
                    <div class="unit-card">
                        <div class="unit-card-name">${unit}</div>
                        <div class="unit-card-info">
                            <div class="unit-card-left">
                                <div class="unit-card-label">Chamados</div>
                                <div class="unit-card-value">${total}</div>
                            </div>
                            <div class="unit-card-right">
                                <div class="unit-card-label">Investido</div>
                                <div class="unit-card-value-right">R$ ${valorFormatado}</div>
                            </div>
                        </div>
                    </div>
                `;
            });
            cardsHTML += '</div>';
            document.getElementById('unitLegendBottom').innerHTML = cardsHTML;
        }

        function render() {
            renderKPIs();
            renderContracts();
            renderTempoAberto();
            renderCharts();
        }

        render();
    </script>
</body>
</html>"""

        # ============== RENDERIZAR DASHBOARD ==============
        st.subheader("📊Dashboard Interativo")
        
        # Gerar HTML
        html_content = HTML_TEMPLATE.replace('__DATA_PLACEHOLDER__', json.dumps(data, ensure_ascii=False))
        
        # Mostrar em iframe
        st.components.v1.html(html_content, height=3000, scrolling=True)
        
        # ============== DOWNLOADS ==============
        st.subheader("📥 Download dos Arquivos")
        col1, col2 = st.columns(2)
        
        with col1:
            # Download HTML
            st.download_button(
                label="📄 Download HTML",
                data=html_content,
                file_name=f"painel_manutencao_{datetime.now().strftime('%Y%m%d')}.html",
                mime="text/html"
            )
        
        with col2:
            st.info("✅ Dashboard gerado com sucesso! Use os botões acima para fazer download.")
            "Atualizar app com nova estrutura de dashboard"
