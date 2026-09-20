import streamlit as st
import pandas as pd
import json
import re
from datetime import datetime

st.set_page_config(
    page_title="Painel de Manutenção Predial",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.title("📊 Painel de Manutenção Predial")
st.markdown("**Análise de chamados e controle financeiro — SESI e SENAI** | Envie sua planilha Excel para gerar o relatório completo")

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

def parse_data_br(val):
    """
    Interpreta datas estritamente no padrão brasileiro DD/MM/AAAA.
    Corrige automaticamente células em que o Excel com regionalização US
    inverteu o Dia pelo Mês quando o dia digitado é <= 12 (ex: 08/09 virando 09 de Agosto).
    """
    if pd.isna(val) or val == '' or str(val).strip().lower() in ['nan', 'nat', '-', 'none']:
        return None

    hoje_ref = datetime.now()

    # 1. Se recebido como texto / string
    if isinstance(val, str):
        val_clean = val.strip()
        match_br = re.match(r'^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})', val_clean)
        if match_br:
            d, m, y = int(match_br.group(1)), int(match_br.group(2)), int(match_br.group(3))
            if y < 100: y += 2000
            try:
                return pd.Timestamp(year=y, month=m, day=d)
            except Exception:
                pass
        try:
            dt = pd.to_datetime(val_clean, dayfirst=True, errors='coerce')
            if pd.notna(dt):
                return dt
        except Exception:
            pass

    # 2. Se recebido como objeto datetime/Timestamp do Excel
    if isinstance(val, (datetime, pd.Timestamp)):
        y = val.year
        m = val.month
        d = val.day

        # Se d > 12, com certeza 'd' é o dia
        if d > 12:
            return pd.Timestamp(year=y, month=m, day=d)

        # Se d <= 12 e m <= 12, reverte inversão de digitação regional do Excel
        dt_orig = pd.Timestamp(year=y, month=m, day=d)
        dt_swapped = None
        try:
            dt_swapped = pd.Timestamp(year=y, month=d, day=m)
        except Exception:
            dt_swapped = None

        if dt_swapped is not None:
            if dt_orig > hoje_ref and dt_swapped <= hoje_ref:
                return dt_swapped
            return dt_swapped

        return dt_orig

    return None

data_col = next((c for c in df_os.columns if 'DATA' in c.upper() and any(k in c.upper() for k in ['ENVIO', 'CHAMADO', 'ABERTURA'])), None)
if data_col:
    df_os[data_col] = df_os[data_col].apply(parse_data_br)

data_orc_col = next((c for c in df_os.columns if any(k in c.upper() for k in [
    'DATA ORÇAMENTO', 'DATA ORCAMENTO', 'DATA ENTREGA ORÇAMENTO', 'DATA ENTREGA ORCAMENTO',
    'DATA ENVIO ORÇAMENTO', 'DATA ENVIO ORCAMENTO', 'ENTREGA DO ORÇAMENTO', 'ENTREGA DO ORCAMENTO',
    'DATA DE ENTREGA DO ORÇAMENTO', 'ENTREGA ORÇAMENTO', 'ENTREGA ORCAMENTO', 'DATA ORÇ.', 'DATA ORC.'
])), None)
if data_orc_col:
    df_os[data_orc_col] = df_os[data_orc_col].apply(parse_data_br)

mes_col = next((c for c in df_os.columns if any(k in c.upper() for k in [
    'MÊS', 'MES', 'EMISSÃO', 'EMISSAO', 'COMPETÊNCIA', 'COMPETENCIA', 
    'FATURAMENTO', 'FATURA', 'REFERÊNCIA', 'REFERENCIA', 'REF'
])), None)
nf_col = next((c for c in df_os.columns if any(k in c.upper() for k in ['NF', 'NFE', 'NOTA FISCAL'])), None)

ORDEM_MESES = {
    'JANEIRO': 1, 'JAN': 1, 'FEVEREIRO': 2, 'FEV': 2, 'MARÇO': 3, 'MARCO': 3, 'MAR': 3,
    'ABRIL': 4, 'ABR': 4, 'MAIO': 5, 'MAI': 5, 'JUNHO': 6, 'JUN': 6, 'JULHO': 7, 'JUL': 7,
    'AGOSTO': 8, 'AGO': 8, 'SETEMBRO': 9, 'SET': 9, 'OUTUBRO': 10, 'OUT': 10,
    'NOVEMBRO': 11, 'NOV': 11, 'DEZEMBRO': 12, 'DEZ': 12
}

def obter_peso_mes(m_str):
    m_clean = str(m_str).strip().upper()
    ano = datetime.now().year
    match_ano = re.search(r'(20\d\d|\b\d{2}\b)', m_clean)
    if match_ano:
        ano_val = int(match_ano.group(1))
        ano = 2000 + ano_val if ano_val < 100 else ano_val

    mes_num = 99
    for nome, peso in ORDEM_MESES.items():
        if nome in m_clean:
            mes_num = peso
            break
            
    if mes_num == 99:
        match_num = re.search(r'\b(0?[1-9]|1[0-2])\b', m_clean)
        if match_num:
            mes_num = int(match_num.group(1))

    return ano * 100 + mes_num

meses_existentes = []
if mes_col:
    valores_mes = df_os[mes_col].dropna().astype(str).str.strip().unique()
    for m in valores_mes:
        m_upper = m.upper()
        if m_upper not in ['', 'NAN', 'NONE', '-', 'NÃO DEFINIDO', 'NAO DEFINIDO']:
            if m_upper not in meses_existentes:
                meses_existentes.append(m_upper)
    meses_existentes = sorted(meses_existentes, key=obter_peso_mes)

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
            dias = (hoje - row[data_col]).days
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
    'meses_existentes': meses_existentes,
    'tem_coluna_orcamento': data_orc_col is not None,
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
    data_orc_str = ''
    dias_orcamento = None
    
    mes_emissao = '-'
    if mes_col and pd.notna(row[mes_col]):
        m_txt = str(row[mes_col]).strip().upper()
        if m_txt not in ['', 'NAN', 'NONE', '-']:
            mes_emissao = m_txt
    elif not mes_col and data_col and pd.notna(row[data_col]):
        data_envio_dt = row[data_col]
        mes_emissao = f"{MESES_PT.get(data_envio_dt.month, data_envio_dt.month)}/{data_envio_dt.year}".upper()
    
    if data_col and pd.notna(row[data_col]):
        data_envio_dt = row[data_col]
        data_envio_str = data_envio_dt.strftime('%d/%m/%Y')
        if status != 'CONCLUIDO':
            dias_abertos = max(0, (hoje - data_envio_dt).days)

    if data_orc_col and pd.notna(row[data_orc_col]):
        data_orc_dt = row[data_orc_col]
        data_orc_str = data_orc_dt.strftime('%d/%m/%Y')
        if data_col and pd.notna(row[data_col]):
            dias_orcamento = max(0, (data_orc_dt - row[data_col]).days)
    elif status == 'AGUARDANDO ORÇAMENTO' and data_col and pd.notna(row[data_col]):
        dias_orcamento = max(0, (hoje - row[data_col]).days)

    num_nf = str(row[nf_col]).strip() if nf_col and pd.notna(row[nf_col]) else '-'
    if num_nf.lower() == 'nan': num_nf = '-'

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
        'data_orcamento': data_orc_str,
        'dias_orcamento': dias_orcamento,
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
        * { margin: 0; padding: 0; box-sizing: border-box; }
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
            padding-bottom: 20px;
        }
        .container { max-width: 1400px; margin: 0 auto; padding: 20px; }
        .header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 16px;
            margin-bottom: 16px;
            padding: 20px;
            background: white;
            border-radius: 8px;
            border-left: 4px solid var(--primary);
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .header h1 { font-size: 24px; margin: 0; }
        .header p { color: #666; margin: 0; font-size: 14px; }
        
        /* Barra de Navegação por Abas */
        .nav-tabs-bar {
            display: flex;
            gap: 10px;
            margin-bottom: 20px;
            flex-wrap: wrap;
        }
        .nav-tab-btn {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 10px 18px;
            background: #ffffff;
            color: #475569;
            border: 1px solid #cbd5e1;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
            box-shadow: 0 1px 2px rgba(0,0,0,0.04);
        }
        .nav-tab-btn:hover { background: #f1f5f9; color: #0f172a; border-color: #94a3b8; }
        .nav-tab-btn.active {
            background: #0d6efd;
            color: #ffffff;
            border-color: #0d6efd;
            box-shadow: 0 2px 6px rgba(13, 110, 253, 0.25);
        }

        /* Cards de Atalho do Painel Inicial */
        .portal-cards-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 18px;
            margin-bottom: 24px;
        }
        .portal-card {
            background: #ffffff;
            border: 1px solid var(--border);
            border-radius: 10px;
            padding: 18px 22px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            box-shadow: 0 2px 4px rgba(0,0,0,0.04);
            cursor: pointer;
            transition: all 0.2s ease;
        }
        .portal-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 14px rgba(0,0,0,0.08);
        }
        .portal-card-pay { border-left: 5px solid #0d6efd; }
        .portal-card-med { border-left: 5px solid #0284c7; }
        .portal-card-list { border-left: 5px solid #10b981; }

        .btn-back-nav {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            padding: 8px 16px;
            background: #f8fafc;
            color: #334155;
            border: 1px solid #cbd5e1;
            border-radius: 6px;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
        }
        .btn-back-nav:hover { background: #e2e8f0; color: #0f172a; }

        .card-title-bar {
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 12px;
            margin-bottom: 20px;
            padding-bottom: 12px;
            border-bottom: 2px solid var(--light);
        }
        .btn-print {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            padding: 7px 14px;
            background: #f8fafc;
            color: #334155;
            border: 1px solid #cbd5e1;
            border-radius: 6px;
            font-size: 12px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
            box-shadow: 0 1px 2px rgba(0,0,0,0.05);
        }
        .btn-print:hover { background: #e2e8f0; color: #0f172a; border-color: #94a3b8; transform: translateY(-1px); }
        .btn-print-primary { background: #0d6efd; color: #ffffff; border-color: #0d6efd; }
        .btn-print-primary:hover { background: #0b5ed7; color: #ffffff; border-color: #0a58ca; }
        
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
        .kpi-card:hover { border-color: var(--primary); box-shadow: 0 4px 8px rgba(57, 135, 229, 0.1); }
        .kpi-label { font-size: 12px; color: #666; margin-bottom: 8px; font-weight: 500; text-transform: uppercase; }
        .kpi-value { font-size: 30px; font-weight: bold; color: var(--dark); margin-bottom: 4px; }
        .kpi-percent { font-size: 12px; color: #999; }
        
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

        .medicoes-card {
            background: #ffffff;
            border-radius: 10px;
            border: 2px solid #0284c7;
            box-shadow: 0 4px 12px rgba(2, 132, 199, 0.1);
            padding: 24px;
            margin-bottom: 24px;
        }
        .mini-kpi-med {
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            padding: 14px 16px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.03);
            border-left: 4px solid #cbd5e1;
        }
        .mini-kpi-med.success { border-left-color: #10b981; }
        .mini-kpi-med.primary { border-left-color: #0284c7; }
        .mini-kpi-med.warning { border-left-color: #f59e0b; }
        .mini-kpi-med.dark { border-left-color: #475569; }
        .mini-kpi-med.indigo { border-left-color: #6366f1; }
        .mini-kpi-med.amber { border-left-color: #d97706; }
        .mini-kpi-med.emerald { border-left-color: #059669; }

        .previsao-box {
            background: #f8fafc;
            border: 1.5px solid #cbd5e1;
            border-radius: 10px;
            padding: 18px 20px;
            margin-bottom: 22px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.03);
        }
        .previsao-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 10px;
            margin-bottom: 14px;
            padding-bottom: 10px;
            border-bottom: 1px solid #e2e8f0;
        }
        .btn-inspect-toggle {
            background: #ffffff;
            border: 1px solid #cbd5e1;
            padding: 5px 12px;
            border-radius: 6px;
            font-size: 11px;
            font-weight: 700;
            color: #475569;
            cursor: pointer;
            transition: all 0.2s;
            display: inline-flex;
            align-items: center;
            gap: 5px;
        }
        .btn-inspect-toggle:hover {
            background: #e2e8f0;
            color: #0f172a;
        }

        .filter-chip-med {
            padding: 7px 14px;
            background: white;
            border: 1px solid #cbd5e1;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 700;
            color: #475569;
            cursor: pointer;
            transition: all 0.2s;
            user-select: none;
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }
        .filter-chip-med:hover { background: #e2e8f0; color: #0f172a; }
        .filter-chip-med.active-sesi {
            background: #0284c7;
            color: white;
            border-color: #0284c7;
            box-shadow: 0 2px 6px rgba(2, 132, 199, 0.25);
        }
        .filter-chip-med.active-senai {
            background: #ea580c;
            color: white;
            border-color: #ea580c;
            box-shadow: 0 2px 6px rgba(234, 88, 12, 0.25);
        }
        .filter-chip-med.active-anual {
            background: #0f172a;
            color: white;
            border-color: #0f172a;
            box-shadow: 0 2px 6px rgba(15, 23, 42, 0.25);
        }

        .summary-ribbon {
            margin-top: 16px;
            padding: 14px 18px;
            background: #f0fdf4;
            border: 1px solid #bbf7d0;
            border-radius: 8px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 10px;
        }
        .summary-ribbon .left { font-size: 13px; color: #166534; font-weight: 600; }
        .summary-ribbon .right { font-size: 18px; font-weight: 800; color: #15803d; }

        .card-title {
            font-size: 18px;
            font-weight: 600;
            margin-bottom: 20px;
            padding-bottom: 12px;
            border-bottom: 2px solid var(--light);
        }
        .contract-stat { display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; font-size: 13px; }
        .contract-label { color: #666; }
        .contract-value { font-weight: 600; color: var(--dark); }
        .progress-container { margin-bottom: 16px; }
        .progress-label { display: flex; justify-content: space-between; margin-bottom: 6px; font-size: 12px; color: #666; }
        .progress-bar { width: 100%; height: 8px; background: #e9ecef; border-radius: 4px; overflow: hidden; }
        .progress-fill { height: 100%; background: linear-gradient(90deg, var(--info), var(--success)); border-radius: 4px; transition: width 0.3s; }
        .chart-wrapper { position: relative; height: 320px; margin: 20px 0; }
        
        .filter-section { display: flex; flex-direction: column; gap: 16px; }
        .filter-group { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; }
        .filter-label { font-weight: 600; font-size: 13px; color: #666; display: block; width: 100%; margin-bottom: 4px; }
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
        .filter-chip:hover { border-color: var(--primary); background: var(--light); }
        .filter-chip.active { background: var(--primary); color: white; border-color: var(--primary); }
        .filter-chip-pay { border-color: #bee5eb; background: #f8fbff; }
        .filter-chip-pay.active { background: #0d6efd; border-color: #0d6efd; color: white; }
        
        .filter-input {
            padding: 8px 12px;
            border: 1px solid var(--border);
            border-radius: 4px;
            font-size: 13px;
            width: 100%;
            max-width: 100%;
            background: white;
        }
        .filter-input:focus { outline: none; border-color: var(--primary); box-shadow: 0 0 0 3px rgba(57, 135, 229, 0.1); }
        
        .table-wrapper { 
            max-height: 520px;
            overflow-y: auto;
            overflow-x: auto; 
            margin-top: 16px; 
            border: 1px solid var(--border);
            border-radius: 8px;
            background: #ffffff;
            position: relative;
        }
        table { width: 100%; border-collapse: collapse; font-size: 13px; }
        thead th {
            position: sticky;
            top: 0;
            z-index: 10;
            background: #f1f5f9;
            box-shadow: 0 1px 2px rgba(0,0,0,0.05);
            padding: 12px;
            text-align: left;
            color: #475569;
            font-weight: 600;
            border-bottom: 2px solid var(--border);
        }
        td { padding: 12px; border-bottom: 1px solid var(--border); }
        tfoot td {
            position: sticky;
            bottom: 0;
            z-index: 10;
            padding: 14px 12px;
            font-weight: 700;
            border-top: 2px solid #333;
            background: #f8fafc;
            box-shadow: 0 -1px 3px rgba(0,0,0,0.05);
        }
        tbody tr:hover { background: var(--light); }
        
        .status-badge { display: inline-block; padding: 4px 8px; border-radius: 4px; font-size: 11px; font-weight: 600; }
        .status-concluido { background: #d4edda; color: #155724; }
        .status-em-execução, .status-em-execucao { background: #d1ecf1; color: #0c5460; }
        .status-paralisado { background: #f8d7da; color: #721c24; }
        .status-aguardando-orçamento, .status-aguardando-orcamento { background: #fff3cd; color: #856404; }
        .status-liberado { background: #d4edda; color: #155724; }
        .status-projeto { background: #e2e3e5; color: #383d41; }
        .status-planejamento { background: #d6d8db; color: #383d41; }
        .status-aguardando-aprovação, .status-aguardando-aprovacao { background: #ffe5d0; color: #a04000; }
        .status-sem-status { background: #e9ecef; color: #495057; }
        
        .units-split-container { display: grid; grid-template-columns: 1fr 1fr; gap: 24px; margin-top: 10px; }
        @media (max-width: 900px) { .units-split-container { grid-template-columns: 1fr; } }
        .units-column { background: #ffffff; border-radius: 10px; border: 1px solid var(--border); padding: 16px; }
        .units-column-sesi { border-top: 4px solid #0d6efd; background: #fbfdff; }
        .units-column-senai { border-top: 4px solid #e65100; background: #fffbf9; }
        .units-column-header { display: flex; justify-content: space-between; align-items: center; padding-bottom: 12px; margin-bottom: 14px; border-bottom: 2px solid var(--border); }
        .units-column-title { font-size: 16px; font-weight: 700; display: flex; align-items: center; gap: 8px; }
        .units-column-title.sesi { color: #0d6efd; }
        .units-column-title.senai { color: #e65100; }
        .units-column-subtotal { font-size: 12px; font-weight: 600; padding: 4px 10px; border-radius: 20px; }
        .units-column-subtotal.sesi { background: #e7f3ff; color: #0a58ca; }
        .units-column-subtotal.senai { background: #fff0e6; color: #c44000; }
        
        .unit-card { padding: 14px 16px; background: white; border: 1px solid var(--border); border-radius: 8px; box-shadow: 0 1px 2px rgba(0,0,0,0.04); transition: transform 0.2s, box-shadow 0.2s; }
        .unit-card:hover { transform: translateY(-2px); box-shadow: 0 4px 10px rgba(0,0,0,0.08); }
        .unit-card-sesi { border-left: 4px solid #0d6efd; }
        .unit-card-senai { border-left: 4px solid #e65100; }
        .unit-card-name { font-weight: 700; color: #212529; margin-bottom: 10px; font-size: 13px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
        .unit-card-info { display: flex; justify-content: space-between; align-items: flex-end; }
        .unit-card-label { font-size: 11px; color: #6c757d; margin-bottom: 2px; text-transform: uppercase; font-weight: 600; }
        .unit-card-value { font-weight: 700; color: #212529; font-size: 18px; line-height: 1; }
        .unit-card-value-right { font-weight: 700; font-size: 14px; }
        .unit-card-value-right.sesi { color: #0d6efd; }
        .unit-card-value-right.senai { color: #e65100; }
        
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
        .btn-action-copy:hover { background: #0b5ed7; transform: translateY(-1px); box-shadow: 0 4px 8px rgba(13, 110, 253, 0.3); }
        .btn-action-copy:active { transform: translateY(0); }
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
        .copy-toast.show { opacity: 1; }
        
        .footer { margin-top: 30px; padding: 15px; text-align: center; color: #999; font-size: 12px; }
        .print-header-stamp { display: none; }
        
        @media print {
            body { background: white !important; color: #000 !important; }
            .no-print, .btn-print, .btn-action-copy, .copy-toast, .filter-section, #filters, #payCasaFilters, #payMonthFilters, #payStatusFilters, #listStatusFilter, #medCasaFilters, #medPeriodoFilters, .filter-input, .nav-tabs-bar, .portal-cards-grid, .btn-back-nav { display: none !important; }
            .tab-view { display: block !important; }
            .table-wrapper { max-height: none !important; overflow: visible !important; border: none !important; }
            .print-header-stamp { display: block !important; margin-bottom: 18px; padding-bottom: 12px; border-bottom: 2px solid #333; }
            .print-header-stamp h2 { font-size: 18px; color: #111; margin: 0 0 4px 0; }
            .print-header-stamp p { font-size: 12px; color: #555; margin: 0; }
            body.is-printing-panel .container > *:not(.target-print-active),
            body.is-printing-panel .tab-view > *:not(.target-print-active) { display: none !important; }
            body.is-printing-panel .target-print-active { display: block !important; width: 100% !important; margin: 0 !important; padding: 0 !important; border: none !important; box-shadow: none !important; }
            .card, .kpi-card, .medicoes-card { box-shadow: none !important; border: 1px solid #ccc !important; page-break-inside: avoid; }
            table { page-break-inside: auto; }
            tr { page-break-inside: avoid; page-break-after: auto; }
        }
    </style>
</head>
<body>
    <div class="container">
        <!-- Header Geral -->
        <div class="header">
            <div>
                <h1>📊 Painel de Manutenção Predial & Liberação de Pagamentos</h1>
                <p>SESI e SENAI — Acompanhamento de Chamados, Emissão de NFE e Gestão Orçamentária</p>
            </div>
            <div class="no-print">
                <button class="btn-print btn-print-primary" onclick="imprimirRelatorioGeral()" title="Imprime todo o painel consolidado">
                    🖨️ Imprimir Relatório Completo
                </button>
            </div>
        </div>

        <!-- Barra de Navegação por Abas -->
        <div class="nav-tabs-bar no-print">
            <button class="nav-tab-btn active" id="nav-inicial" onclick="switchTab('inicial')">
                📊 Painel Inicial (Visão Geral)
            </button>
            <button class="nav-tab-btn" id="nav-pagamentos" onclick="switchTab('pagamentos')">
                💳 Controle de Pagamentos & NF-e
            </button>
            <button class="nav-tab-btn" id="nav-medicoes" onclick="switchTab('medicoes')">
                📐 MEDIÇÕES (Controle Contratual)
            </button>
            <button class="nav-tab-btn" id="nav-chamados" onclick="switchTab('chamados')">
                📋 Lista Completa de Chamados
            </button>
        </div>

        <!-- ======================================================= -->
        <!-- ABA 1: PAINEL INICIAL (VISÃO GERAL)                     -->
        <!-- ======================================================= -->
        <div class="tab-view" id="view-inicial" style="display: block;">
            
            <!-- Cards de Acesso Rápido aos Painéis -->
            <div class="portal-cards-grid no-print">
                <div class="portal-card portal-card-pay" onclick="switchTab('pagamentos')">
                    <div style="display: flex; align-items: center; gap: 14px;">
                        <div style="font-size: 32px;">💳</div>
                        <div>
                            <div style="font-size: 16px; font-weight: 700; color: #0d6efd;">Controle de Pagamentos & NF-e</div>
                            <div style="font-size: 12px; color: #64748b;">Conferência por Mês e CNPJ (SESI/SENAI) para autorizar emissão de Nota Fiscal</div>
                        </div>
                    </div>
                    <button class="btn-print btn-print-primary" style="pointer-events: none; margin-top: 14px;">Acessar Faturamento ➔</button>
                </div>

                <div class="portal-card portal-card-med" onclick="switchTab('medicoes')">
                    <div style="display: flex; align-items: center; gap: 14px;">
                        <div style="font-size: 32px;">📐</div>
                        <div>
                            <div style="font-size: 16px; font-weight: 700; color: #0284c7;">Módulo MEDIÇÕES (Contrato)</div>
                            <div style="font-size: 12px; color: #64748b;">Livro Oficial de Medições mensais e acumulado anual com saldo contratual</div>
                        </div>
                    </div>
                    <button class="btn-print" style="pointer-events: none; margin-top: 14px; border-color: #0284c7; color: #0284c7;">Acessar Medições ➔</button>
                </div>

                <div class="portal-card portal-card-list" onclick="switchTab('chamados')">
                    <div style="display: flex; align-items: center; gap: 14px;">
                        <div style="font-size: 32px;">📋</div>
                        <div>
                            <div style="font-size: 16px; font-weight: 700; color: #10b981;">Lista Completa de Chamados</div>
                            <div style="font-size: 12px; color: #64748b;">Tabela com busca por O.S., filtro por Unidade SESI/SENAI e acompanhamento de prazos</div>
                        </div>
                    </div>
                    <button class="btn-print" style="pointer-events: none; margin-top: 14px; border-color: #10b981; color: #10b981;">Ver Chamados ➔</button>
                </div>
            </div>

            <!-- KPIs Gerais -->
            <div class="kpi-section" id="kpis"></div>

            <!-- Saldo em Contrato + Tempo Médio em Aberto -->
            <div class="grid2" id="painelContratosPrazos">
                <div class="card" style="margin-bottom: 0;">
                    <div class="card-title-bar">
                        <div class="card-title" style="margin-bottom: 0; padding-bottom: 0; border: none;">💰 Saldo em Contrato</div>
                        <button class="btn-print no-print" onclick="imprimirPainel('painelContratosPrazos', 'Saldos em Contrato e Tempo Médio')">🖨️ Imprimir</button>
                    </div>
                    <div id="contracts"></div>
                </div>

                <div class="card" style="margin-bottom: 0;">
                    <div class="card-title-bar">
                        <div class="card-title" style="margin-bottom: 0; padding-bottom: 0; border: none;">⏱️ Tempo Médio em Aberto</div>
                        <button class="btn-print no-print" onclick="imprimirPainel('painelContratosPrazos', 'Saldos em Contrato e Tempo Médio')">🖨️ Imprimir</button>
                    </div>
                    <div id="tempoAberto"></div>
                </div>
            </div>

            <!-- Distribuição de Status -->
            <div class="card" id="painelStatus">
                <div class="card-title-bar">
                    <div class="card-title" style="margin-bottom: 0; padding-bottom: 0; border: none;">📈 Distribuição de Status</div>
                    <button class="btn-print no-print" onclick="imprimirPainel('painelStatus', 'Distribuição de Status dos Chamados')">🖨️ Imprimir Gráfico</button>
                </div>
                <div class="chart-wrapper">
                    <canvas id="statusChart" role="img" aria-label="Distribuição de chamados por status"></canvas>
                </div>
            </div>

            <!-- Filtros Globais Interativos -->
            <div class="card no-print">
                <div class="card-title">🔍 Filtros Interativos Globais</div>
                <div class="filter-section" id="filters"></div>
            </div>

            <!-- Chamados por Unidade (SESI à esquerda e SENAI à direita de A a Z) -->
            <div class="card" id="painelUnidades">
                <div class="card-title-bar">
                    <div class="card-title" style="margin-bottom: 0; padding-bottom: 0; border: none;">🏢 Chamados por Unidade (com investimento)</div>
                    <button class="btn-print no-print" onclick="imprimirPainel('painelUnidades', 'Investimento e Chamados por Unidade (SESI e SENAI)')">🖨️ Imprimir Unidades</button>
                </div>
                <div class="chart-wrapper" style="height: 500px; margin-bottom: 20px;">
                    <canvas id="unitChart" role="img" aria-label="Distribuição de chamados por unidade"></canvas>
                </div>
                <div id="unitLegendBottom"></div>
            </div>
        </div>

        <!-- ======================================================= -->
        <!-- ABA 2: PAINEL DE CONTROLE DE PAGAMENTOS                -->
        <!-- ======================================================= -->
        <div class="tab-view" id="view-pagamentos" style="display: none;">
            <div class="no-print" style="margin-bottom: 14px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px;">
                <button class="btn-back-nav" onclick="switchTab('inicial')">⬅️ Retornar ao Painel Inicial</button>
                <button class="btn-print" style="border-color: #0284c7; color: #0284c7;" onclick="switchTab('medicoes')">
                    📐 Ir para Módulo MEDIÇÕES (Boletim Oficial) ➔
                </button>
            </div>

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
                    <div style="display: flex; gap: 10px; align-items: center; flex-wrap: wrap;">
                        <button class="btn-print no-print" onclick="imprimirPainel('painelPagamentos', 'Controle de Pagamentos e Liberação para NF-e')" title="Imprimir este painel de faturamento">
                            🖨️ Imprimir Faturamento
                        </button>
                        <div class="payment-badge-status" id="paySummaryBadge">
                            Carregando resumo financeiro...
                        </div>
                    </div>
                </div>

                <!-- Mini KPIs de Pagamento -->
                <div class="kpi-section" style="margin-bottom: 20px;">
                    <div class="kpi-card" style="border-left: 4px solid var(--success);">
                        <div class="kpi-label">Liberado para Emitir NFE</div>
                        <div class="kpi-value" id="payValLiberado" style="color: var(--success); font-size: 26px;">R$ 0,00</div>
                        <div class="kpi-percent" id="payCountLiberado">0 O.S. prontas para faturar</div>
                    </div>
                    <div class="kpi-card" style="border-left: 4px solid var(--primary);">
                        <div class="kpi-label">Em Medição / Andamento</div>
                        <div class="kpi-value" id="payValMedicao" style="color: var(--primary); font-size: 26px;">R$ 0,00</div>
                        <div class="kpi-percent" id="payCountMedicao">0 O.S. em andamento</div>
                    </div>
                    <div class="kpi-card" style="border-left: 4px solid var(--danger);">
                        <div class="kpi-label">Bloqueado / Paralisado</div>
                        <div class="kpi-value" id="payValBloqueado" style="color: var(--danger); font-size: 26px;">R$ 0,00</div>
                        <div class="kpi-percent" id="payCountBloqueado">0 O.S. paralisadas</div>
                    </div>
                    <div class="kpi-card" style="border-left: 4px solid #6c757d;">
                        <div class="kpi-label">Total do Filtro de Pagamento</div>
                        <div class="kpi-value" id="payValTotal" style="font-size: 26px;">R$ 0,00</div>
                        <div class="kpi-percent" id="payCountTotal">0 chamados filtrados</div>
                    </div>
                </div>

                <!-- Filtros de Entidade e Mês -->
                <div class="no-print" style="background: #f8f9fa; padding: 16px; border-radius: 8px; margin-bottom: 20px; border: 1px solid var(--border);">
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

                <!-- Tabela de Liberação de Pagamentos com Scroll Interno -->
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

                <!-- Faixa de Resumo do Fechamento -->
                <div id="paySummaryFooter" style="margin-top: 18px; padding: 14px 20px; background: #e8f4fd; border-radius: 8px; border-left: 5px solid #0d6efd; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px;">
                    <div id="paySummaryFooterText" style="font-size: 14px; font-weight: 600; color: #0a58ca;">Fechamento de Faturamento</div>
                    <div id="paySummaryFooterVal" style="font-size: 18px; font-weight: 800; color: #0ca30c;">R$ 0,00</div>
                </div>

                <!-- Botão de Cópia Única para E-mail -->
                <div class="no-print" style="margin-top: 16px; padding: 14px 18px; background: #f8fafc; border: 1px dashed #cbd5e1; border-radius: 8px; display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 12px;">
                    <div style="display: flex; gap: 10px; flex-wrap: wrap; align-items: center;">
                        <button class="btn-action-copy" onclick="copiarTabelaEmail()" title="Copia apenas as O.S. liberadas para emissão de NF-e formatadas para colar no e-mail">
                            📋 Copiar Tabela p/ E-mail
                        </button>
                    </div>
                    <div id="copyToast" class="copy-toast">✅ Tabela copiada! Pressione Ctrl + V no seu e-mail.</div>
                </div>
            </div>

            <div class="no-print" style="margin-top: 14px; text-align: center;">
                <button class="btn-back-nav" onclick="switchTab('inicial')">⬅️ Retornar ao Painel Inicial</button>
            </div>
        </div>

        <!-- ======================================================= -->
        <!-- ABA 3: MÓDULO MEDIÇÕES (LIVRO OFICIAL DO CONTRATO)      -->
        <!-- ======================================================= -->
        <div class="tab-view" id="view-medicoes" style="display: none;">
            <div class="no-print" style="margin-bottom: 14px; display: flex; justify-content: space-between; align-items: center;">
                <button class="btn-back-nav" onclick="switchTab('inicial')">⬅️ Retornar ao Painel Inicial</button>
            </div>

            <div class="medicoes-card" id="painelMedicoes">
                <div class="payment-header" style="border-bottom: 2px solid #e0f2fe;">
                    <div>
                        <h2 style="font-size: 22px; font-weight: 800; color: #0284c7; display: flex; align-items: center; gap: 10px;">
                            📐 MEDIÇÕES — Controle de Medição Contratual
                        </h2>
                        <p style="font-size: 13px; color: #64748b; margin-top: 4px;">
                            Livro Oficial de Medições: itens faturados (LIBERADO P/ NFE) com teto e saldo contratual de R$ 1.440.000,00
                        </p>
                    </div>
                    <div style="display: flex; gap: 10px; align-items: center; flex-wrap: wrap;">
                        <button class="btn-print no-print" onclick="imprimirPainel('painelMedicoes', 'Boletim de Medição Contratual')" title="Imprimir Relatório Oficial de Medição">
                            🖨️ Imprimir Medição
                        </button>
                        <div style="font-size: 12px; background: #e0f2fe; color: #0369a1; padding: 6px 12px; border-radius: 20px; font-weight: 700;">
                            Contrato: R$ 1.440.000,00 por Entidade
                        </div>
                    </div>
                </div>

                <!-- KPIs Financeiros da Medição e Contrato -->
                <div style="font-size: 11px; font-weight: 800; color: #64748b; text-transform: uppercase; margin-bottom: 8px; letter-spacing: 0.5px;">
                    📋 Medições Efetivadas & Faturamento Homologado (Status Atual):
                </div>
                <div class="kpi-section" style="margin-bottom: 18px;">
                    <div class="mini-kpi-med success">
                        <div class="kpi-label" style="color: #059669; font-weight: 700;">Valor da Medição Selecionada</div>
                        <div class="kpi-value" id="medKpiValor" style="color: #10b981; font-size: 26px;">R$ 0,00</div>
                        <div class="kpi-percent" id="medKpiSub" style="color: #059669; font-weight: 600;">0 O.S. aprovadas</div>
                    </div>

                    <div class="mini-kpi-med primary">
                        <div class="kpi-label" style="color: #0284c7; font-weight: 700;">Acumulado Anual Liberado</div>
                        <div class="kpi-value" id="medKpiAnual" style="color: #0284c7; font-size: 26px;">R$ 0,00</div>
                        <div class="kpi-percent" id="medKpiAnualSub" style="color: #0369a1; font-weight: 600;">0 O.S. no exercício</div>
                    </div>

                    <div class="mini-kpi-med dark">
                        <div class="kpi-label" style="color: #475569; font-weight: 700;">Teto Contratual Total</div>
                        <div class="kpi-value" id="medKpiTeto" style="color: #1e293b; font-size: 26px;">R$ 1.440.000,00</div>
                        <div class="kpi-percent" style="color: #64748b;">Valor limite homologado</div>
                    </div>

                    <div class="mini-kpi-med warning">
                        <div class="kpi-label" style="color: #b45309; font-weight: 700;">Saldo Disponível em Contrato</div>
                        <div class="kpi-value" id="medKpiSaldo" style="color: #d97706; font-size: 26px;">R$ 1.440.000,00</div>
                        <div class="kpi-percent" id="medKpiSaldoPct" style="color: #b45309; font-weight: 600;">100% ainda disponível</div>
                    </div>
                </div>

                <!-- Painel de Previsão de Medições Futuras & Saldo Estimado -->
                <div class="previsao-box">
                    <div class="previsao-header">
                        <div style="font-size: 13px; font-weight: 800; color: #0f172a; display: flex; align-items: center; gap: 8px;">
                            🛡️ GESTÃO DE RISCO: PREVISÃO DE MEDIÇÕES FUTURAS & SALDO ESTIMADO (<span id="medPrevCasaLabel" style="color: #0284c7;">SENAI</span>)
                        </div>
                        <div style="display: flex; align-items: center; gap: 10px;">
                            <button type="button" class="btn-inspect-toggle no-print" onclick="toggleInspecaoPrevisao()" id="btnToggleInspect">
                                👁️ Inspecionar O.S. Futuras (0)
                            </button>
                            <div id="medRiscoBadge" style="font-size: 11px; font-weight: 800; padding: 4px 10px; border-radius: 14px; background: #dcfce7; color: #15803d; border: 1px solid #bbf7d0;">
                                🟢 SALDO ESTIMADO SEGURO
                            </div>
                        </div>
                    </div>

                    <!-- 3 Cards de Projeção Financeira -->
                    <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 14px; margin-bottom: 16px;">
                        <div class="mini-kpi-med indigo">
                            <div class="kpi-label" style="color: #4f46e5; font-weight: 700;">Previsão em Investimentos (O.S. Futuras)</div>
                            <div class="kpi-value" id="medKpiPrevisao" style="color: #4338ca; font-size: 25px;">R$ 0,00</div>
                            <div class="kpi-percent" id="medKpiPrevisaoSub" style="color: #4f46e5; font-weight: 600;">0 O.S. em andamento/orçadas</div>
                        </div>

                        <div class="mini-kpi-med amber">
                            <div class="kpi-label" style="color: #b45309; font-weight: 700;">Total Geral Comprometido (Medido + Previsto)</div>
                            <div class="kpi-value" id="medKpiComprometido" style="color: #d97706; font-size: 25px;">R$ 0,00</div>
                            <div class="kpi-percent" id="medKpiComprometidoSub" style="color: #b45309; font-weight: 600;">0% do contrato consumido</div>
                        </div>

                        <div class="mini-kpi-med emerald" id="medCardSaldoEstimado">
                            <div class="kpi-label" style="color: #059669; font-weight: 700;">Saldo Estimado em Contrato (Margem Real Livre)</div>
                            <div class="kpi-value" id="medKpiSaldoEstimado" style="color: #047857; font-size: 25px;">R$ 0,00</div>
                            <div class="kpi-percent" id="medKpiSaldoEstimadoSub" style="color: #059669; font-weight: 600;">Margem segura p/ novas programações</div>
                        </div>
                    </div>

                    <!-- Barra de Consumo do Teto Contratual -->
                    <div>
                        <div style="display: flex; justify-content: space-between; font-size: 11px; font-weight: 700; color: #475569; margin-bottom: 6px;">
                            <span>Consumo do Teto Contratual (R$ 1.440.000,00):</span>
                            <span id="medBarraComprometidoLabel">0% comprometido (Medições + Previsão)</span>
                        </div>
                        <div style="width: 100%; height: 10px; background: #e2e8f0; border-radius: 5px; overflow: hidden; display: flex;">
                            <div id="medBarraMedido" style="height: 100%; background: #0284c7; width: 0%; transition: width 0.4s;" title="Faturado / Liberado"></div>
                            <div id="medBarraPrevisto" style="height: 100%; background: #f59e0b; width: 0%; transition: width 0.4s;" title="Previsão de O.S. Futuras"></div>
                        </div>
                        <div style="display: flex; justify-content: space-between; flex-wrap: wrap; gap: 8px; font-size: 11px; color: #64748b; margin-top: 6px;">
                            <span style="display: inline-flex; align-items: center; gap: 5px;">
                                <span style="width: 8px; height: 8px; background: #0284c7; border-radius: 2px; display: inline-block;"></span> 
                                Liberado/Medido: <strong id="medLegendaMedido" style="color: #0284c7;">R$ 0,00</strong>
                            </span>
                            <span style="display: inline-flex; align-items: center; gap: 5px;">
                                <span style="width: 8px; height: 8px; background: #f59e0b; border-radius: 2px; display: inline-block;"></span> 
                                Previsão Futura: <strong id="medLegendaPrevisto" style="color: #d97706;">R$ 0,00</strong>
                            </span>
                            <span style="display: inline-flex; align-items: center; gap: 5px;">
                                <span style="width: 8px; height: 8px; background: #10b981; border-radius: 2px; display: inline-block;"></span> 
                                Margem Livre Estimada: <strong id="medLegendaSaldoEstimado" style="color: #059669;">R$ 0,00</strong>
                            </span>
                        </div>
                    </div>

                    <!-- Gaveta de Inspeção das O.S. Futuras -->
                    <div id="medInspecaoBox" style="display: none; margin-top: 14px; padding-top: 14px; border-top: 1px dashed #cbd5e1;">
                        <div style="font-size: 12px; font-weight: 700; color: #334155; margin-bottom: 8px; display: flex; justify-content: space-between; align-items: center;">
                            <span>🔍 O.S. com Investimentos em Andamento / Previsão Futura:</span>
                            <span style="font-size: 11px; color: #64748b;">(Itens orçados ainda não liberados p/ NFE)</span>
                        </div>
                        <div style="max-height: 220px; overflow-y: auto; border: 1px solid #e2e8f0; border-radius: 6px;">
                            <table style="font-size: 11px; width: 100%; border-collapse: collapse;">
                                <thead style="background: #f1f5f9; position: sticky; top: 0;">
                                    <tr>
                                        <th style="padding: 6px 10px;">O.S</th>
                                        <th style="padding: 6px 10px;">Unidade</th>
                                        <th style="padding: 6px 10px;">Descrição</th>
                                        <th style="padding: 6px 10px;">Status O.S</th>
                                        <th style="padding: 6px 10px; text-align: right;">Previsão (R$)</th>
                                    </tr>
                                </thead>
                                <tbody id="medInspecaoBody"></tbody>
                            </table>
                        </div>
                    </div>
                </div>

                <!-- Painel de Filtros de Casa e Período -->
                <div class="no-print" style="background: #f8fafc; padding: 16px; border-radius: 8px; margin-bottom: 20px; border: 1px solid #e2e8f0;">
                    <div style="font-weight: 700; font-size: 12px; color: #475569; margin-bottom: 8px; text-transform: uppercase;">
                        🏛️ Entidade Contratual (Casa):
                    </div>
                    <div class="filter-group" id="medCasaFilters" style="margin-bottom: 14px;"></div>

                    <div style="font-weight: 700; font-size: 12px; color: #475569; margin-bottom: 8px; text-transform: uppercase;">
                        📅 Medição Mensal de Competência ou Exercício Anual:
                    </div>
                    <div class="filter-group" id="medPeriodoFilters"></div>
                </div>

                <!-- Tabela de Medições Exata da Planilha de Gerenciamento -->
                <div class="table-wrapper">
                    <table id="medTable">
                        <thead>
                            <tr style="background: #f1f5f9;">
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
                        <tbody id="medTableBody"></tbody>
                        <tfoot id="medTableFoot"></tfoot>
                    </table>
                </div>

                <!-- Faixa de Resumo da Medição -->
                <div class="summary-ribbon">
                    <div class="left" id="medRibbonText">
                        🏷️ <strong>Resumo da Medição:</strong> Entidade: <u>SESI</u> | Competência: <u>SETEMBRO</u> | Aptos para NFE: <strong>0 chamados</strong>
                    </div>
                    <div class="right" id="medRibbonVal">
                        Total Medição Liberada: R$ 0,00
                    </div>
                </div>
            </div>

            <div class="no-print" style="margin-top: 14px; text-align: center;">
                <button class="btn-back-nav" onclick="switchTab('inicial')">⬅️ Retornar ao Painel Inicial</button>
            </div>
        </div>

        <!-- ======================================================= -->
        <!-- ABA 4: LISTA COMPLETA DE CHAMADOS                       -->
        <!-- ======================================================= -->
        <div class="tab-view" id="view-chamados" style="display: none;">
            <div class="no-print" style="margin-bottom: 14px; display: flex; justify-content: space-between; align-items: center;">
                <button class="btn-back-nav" onclick="switchTab('inicial')">⬅️ Retornar ao Painel Inicial</button>
            </div>

            <div class="card" id="painelChamados">
                <div class="card-title-bar">
                    <div class="card-title" style="margin-bottom: 0; padding-bottom: 0; border: none;">📋 Lista Completa de Chamados</div>
                    <button class="btn-print no-print" onclick="imprimirPainel('painelChamados', 'Lista Completa de Chamados')">🖨️ Imprimir Lista</button>
                </div>
                
                <div class="no-print" style="margin-bottom: 16px;">
                    <div style="font-weight: 600; font-size: 13px; color: #666; margin-bottom: 10px;">Filtrar por Status na Tabela</div>
                    <div class="filter-group" id="listStatusFilter"></div>
                </div>

                <div class="no-print" style="margin-bottom: 16px; display: flex; gap: 14px; align-items: flex-end; flex-wrap: wrap;">
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
                                <th>Data Envio (DD/MM/AAAA)</th>
                                <th>Entrega Orçamento</th>
                                <th>Dias Abertos</th>
                                <th style="text-align: right;">Valor</th>
                            </tr>
                        </thead>
                        <tbody id="tableBody"></tbody>
                    </table>
                </div>
            </div>

            <div class="no-print" style="margin-top: 14px; text-align: center;">
                <button class="btn-back-nav" onclick="switchTab('inicial')">⬅️ Retornar ao Painel Inicial</button>
            </div>
        </div>

        <div class="footer">
            <p>Atualizado em <span id="updateTime"></span> | Painel Integrado de Gestão Predial, Medições & Faturamento SESI/SENAI</p>
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
            payFilters: { mes: '', statusPagamento: '', casa: '' },
            medFilters: { casa: 'SESI', periodo: 'ANUAL' }
        };

        function switchTab(tabId) {
            document.querySelectorAll('.tab-view').forEach(el => el.style.display = 'none');
            document.querySelectorAll('.nav-tab-btn').forEach(btn => btn.classList.remove('active'));

            const targetView = document.getElementById('view-' + tabId);
            if (targetView) targetView.style.display = 'block';

            const targetNav = document.getElementById('nav-' + tabId);
            if (targetNav) targetNav.classList.add('active');

            window.scrollTo({ top: 0, behavior: 'smooth' });

            if (tabId === 'inicial' && window.statusChartInstance && window.unitChartInstance) {
                window.statusChartInstance.resize();
                window.unitChartInstance.resize();
            }
        }

        function init() {
            state.tickets = DATA.chamados;
            document.getElementById('updateTime').textContent = new Date().toLocaleString('pt-BR');
            document.getElementById('totalTickets').textContent = state.tickets.length;

            // Define período inicial padrão para Medições (último mês existente ou ANUAL)
            if (DATA.meses_existentes && DATA.meses_existentes.length > 0) {
                state.medFilters.periodo = DATA.meses_existentes[DATA.meses_existentes.length - 1];
            } else {
                state.medFilters.periodo = 'ANUAL';
            }

            renderFilters();
            renderListStatusFilter();
            renderListUnitFilter();
            renderPaymentCasaFilters();
            renderPaymentMonthFilters();
            renderMedicoesCasaFilters();
            renderMedicoesPeriodoFilters();
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
            let mesesValidos = DATA.meses_existentes || [];
            if (mesesValidos.length === 0) {
                mesesValidos = [...new Set(state.tickets.map(t => t.mes_emissao))]
                    .filter(m => m && !['-', 'NÃO DEFINIDO', 'NAN', ''].includes(m.toUpperCase()));
            }
            
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

        function renderMedicoesCasaFilters() {
            const opcoes = [
                { id: 'SESI', label: '🔵 SESI (CNPJ SESI)' },
                { id: 'SENAI', label: '🟠 SENAI (CNPJ SENAI)' }
            ];

            let html = '';
            opcoes.forEach(op => {
                const isActive = state.medFilters.casa === op.id;
                const activeClass = isActive ? (op.id === 'SESI' ? ' active-sesi' : ' active-senai') : '';
                html += `
                    <span class="filter-chip-med${activeClass}" onclick="setMedCasa('${op.id}')">
                        ${op.label}
                    </span>
                `;
            });
            document.getElementById('medCasaFilters').innerHTML = html;
        }

        function setMedCasa(casa) {
            state.medFilters.casa = casa;
            renderMedicoesCasaFilters();
            renderMedicoesPeriodoFilters();
            renderMedicoesPanel();
        }

        function renderMedicoesPeriodoFilters() {
            let mesesValidos = DATA.meses_existentes || [];
            if (mesesValidos.length === 0) {
                mesesValidos = [...new Set(state.tickets.map(t => t.mes_emissao))]
                    .filter(m => m && !['-', 'NÃO DEFINIDO', 'NAN', ''].includes(m.toUpperCase()));
            }

            const isSesi = state.medFilters.casa === 'SESI';
            const activeMonthClass = isSesi ? ' active-sesi' : ' active-senai';

            let html = '';
            mesesValidos.forEach(m => {
                const isActive = state.medFilters.periodo === m;
                html += `
                    <span class="filter-chip-med ${isActive ? activeMonthClass : ''}" onclick="setMedPeriodo('${m}')">
                        📅 ${m}
                    </span>
                `;
            });

            const isAnual = state.medFilters.periodo === 'ANUAL';
            html += `
                <span class="filter-chip-med ${isAnual ? 'active-anual' : ''}" onclick="setMedPeriodo('ANUAL')">
                    📊 ACUMULADO ANUAL (Todas as Medições)
                </span>
            `;

            document.getElementById('medPeriodoFilters').innerHTML = html;
        }

        function setMedPeriodo(p) {
            state.medFilters.periodo = p;
            renderMedicoesPeriodoFilters();
            renderMedicoesPanel();
        }

        function toggleInspecaoPrevisao() {
            const box = document.getElementById('medInspecaoBox');
            const btn = document.getElementById('btnToggleInspect');
            if (!box) return;
            if (box.style.display === 'none' || box.style.display === '') {
                box.style.display = 'block';
                btn.innerHTML = '✖️ Ocultar O.S. Futuras';
            } else {
                box.style.display = 'none';
                btn.innerHTML = `👁️ Inspecionar O.S. Futuras (${window.lastCountPrevisao || 0})`;
            }
        }

        function renderMedicoesPanel() {
            const fmt = v => new Intl.NumberFormat('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2}).format(v);
            const casaAtual = state.medFilters.casa;

            // 1. Filtrar estritamente chamados LIBERADOS P/ NFE da Casa
            const liberadosBase = state.tickets.filter(t => t.status_pagamento === 'LIBERADO P/ NFE');
            const casaList = liberadosBase.filter(t => t.casa === casaAtual);

            // Total anual efetivamente medido/faturado da casa selecionada
            const totalAnualCasa = casaList.reduce((acc, t) => acc + (t.valor || 0), 0);
            const countAnualCasa = casaList.length;

            // Filtrar pelo período selecionado (mês ou anual)
            let medList = casaList;
            if (state.medFilters.periodo !== 'ANUAL') {
                medList = casaList.filter(t => t.mes_emissao === state.medFilters.periodo);
            }

            const totalMedicao = medList.reduce((acc, t) => acc + (t.valor || 0), 0);
            const countMedicao = medList.length;

            // 2. Cálculos da Previsão de Investimentos Futuros (O.S. em andamento / orçadas com valor > 0 não concluídas)
            const ticketsCasa = state.tickets.filter(t => t.casa === casaAtual);
            const chamadosFuturos = ticketsCasa.filter(t => t.status_pagamento !== 'LIBERADO P/ NFE' && t.status !== 'CONCLUIDO' && (t.valor || 0) > 0);
            const valorPrevisaoFutura = chamadosFuturos.reduce((acc, t) => acc + (t.valor || 0), 0);
            const countPrevisaoFutura = chamadosFuturos.length;
            window.lastCountPrevisao = countPrevisaoFutura;

            // Chamados abertos sem valor registrado
            const countSemValor = ticketsCasa.filter(t => t.status !== 'CONCLUIDO' && ((t.valor || 0) === 0)).length;

            // 3. Teto, Total Comprometido e Saldo Estimado
            const tetoContrato = 1440000.00;
            const saldoContratoReal = Math.max(0, tetoContrato - totalAnualCasa);
            const pctSaldoReal = ((saldoContratoReal / tetoContrato) * 100).toFixed(1);

            const totalComprometido = totalAnualCasa + valorPrevisaoFutura;
            const saldoEstimado = tetoContrato - totalComprometido;
            const pctComprometido = Math.min(100, Math.max(0, (totalComprometido / tetoContrato) * 100)).toFixed(1);
            const pctMedido = Math.min(100, Math.max(0, (totalAnualCasa / tetoContrato) * 100));
            const pctPrevisto = Math.min(100 - pctMedido, Math.max(0, (valorPrevisaoFutura / tetoContrato) * 100));
            const pctSaldoEstimado = ((saldoEstimado / tetoContrato) * 100).toFixed(1);

            // 4. Atualizar os 4 KPIs Superiores Tradicionais
            document.getElementById('medKpiValor').textContent = `R$ ${fmt(totalMedicao)}`;
            const periodoLabelTxt = state.medFilters.periodo === 'ANUAL' ? 'no acumulado anual' : 'nesta medição mensal';
            document.getElementById('medKpiSub').textContent = `${countMedicao} O.S. liberadas ${periodoLabelTxt}`;

            document.getElementById('medKpiAnual').textContent = `R$ ${fmt(totalAnualCasa)}`;
            document.getElementById('medKpiAnualSub').textContent = `${countAnualCasa} O.S. faturadas no exercício (${casaAtual})`;

            document.getElementById('medKpiTeto').textContent = `R$ ${fmt(tetoContrato)}`;
            document.getElementById('medKpiSaldo').textContent = `R$ ${fmt(saldoContratoReal)}`;
            document.getElementById('medKpiSaldoPct').textContent = `${pctSaldoReal}% disponível no teto homologado`;

            // 5. Atualizar o Bloco de Gestão de Risco e Previsão Futura
            const labelCasaEl = document.getElementById('medPrevCasaLabel');
            if (labelCasaEl) labelCasaEl.textContent = casaAtual;

            document.getElementById('medKpiPrevisao').textContent = `R$ ${fmt(valorPrevisaoFutura)}`;
            const subPrevisaoTxt = countSemValor > 0 
                ? `${countPrevisaoFutura} O.S. orçadas (+${countSemValor} aguard. orçamento)`
                : `${countPrevisaoFutura} O.S. orçadas em andamento`;
            document.getElementById('medKpiPrevisaoSub').textContent = subPrevisaoTxt;

            document.getElementById('medKpiComprometido').textContent = `R$ ${fmt(totalComprometido)}`;
            document.getElementById('medKpiComprometidoSub').textContent = `${pctComprometido}% do teto consumido (Medido + Previsto)`;

            const cardSaldoEstimado = document.getElementById('medCardSaldoEstimado');
            const kpiSaldoEstimadoVal = document.getElementById('medKpiSaldoEstimado');
            const kpiSaldoEstimadoSub = document.getElementById('medKpiSaldoEstimadoSub');
            const badgeRisco = document.getElementById('medRiscoBadge');

            kpiSaldoEstimadoVal.textContent = `R$ ${fmt(saldoEstimado)}`;

            // Termômetro de Alerta Orçamentário
            if (saldoEstimado < 0) {
                kpiSaldoEstimadoVal.style.color = '#dc2626';
                kpiSaldoEstimadoSub.innerHTML = `⚠️ <strong style="color: #dc2626;">ESTOURO PREVISTO:</strong> Excede o teto em R$ ${fmt(Math.abs(saldoEstimado))}`;
                if (cardSaldoEstimado) cardSaldoEstimado.style.borderLeftColor = '#dc2626';
                if (badgeRisco) {
                    badgeRisco.style.background = '#fee2e2';
                    badgeRisco.style.color = '#b91c1c';
                    badgeRisco.style.borderColor = '#fecaca';
                    badgeRisco.textContent = '⛔ ESTOURO PROJETADO (CONTRATO ESGOTADO)';
                }
            } else if (saldoEstimado < 100000) {
                kpiSaldoEstimadoVal.style.color = '#ea580c';
                kpiSaldoEstimadoSub.innerHTML = `⚠️ <strong style="color: #ea580c;">MARGEM CRÍTICA:</strong> Apenas ${pctSaldoEstimado}% livre no teto`;
                if (cardSaldoEstimado) cardSaldoEstimado.style.borderLeftColor = '#ea580c';
                if (badgeRisco) {
                    badgeRisco.style.background = '#ffedd5';
                    badgeRisco.style.color = '#c2410c';
                    badgeRisco.style.borderColor = '#fed7aa';
                    badgeRisco.textContent = '🔴 ALERTA CRÍTICO (LIMITE QUASE ESGOTADO)';
                }
            } else if (saldoEstimado < 300000) {
                kpiSaldoEstimadoVal.style.color = '#d97706';
                kpiSaldoEstimadoSub.innerHTML = `Margem de atenção: ${pctSaldoEstimado}% livre (${fmt(saldoEstimado)})`;
                if (cardSaldoEstimado) cardSaldoEstimado.style.borderLeftColor = '#f59e0b';
                if (badgeRisco) {
                    badgeRisco.style.background = '#fef3c7';
                    badgeRisco.style.color = '#b45309';
                    badgeRisco.style.borderColor = '#fde68a';
                    badgeRisco.textContent = '🟡 ATENÇÃO AO SALDO ESTIMADO';
                }
            } else {
                kpiSaldoEstimadoVal.style.color = '#047857';
                kpiSaldoEstimadoSub.innerHTML = `Margem segura: ${pctSaldoEstimado}% livre (${fmt(saldoEstimado)})`;
                if (cardSaldoEstimado) cardSaldoEstimado.style.borderLeftColor = '#10b981';
                if (badgeRisco) {
                    badgeRisco.style.background = '#dcfce7';
                    badgeRisco.style.color = '#15803d';
                    badgeRisco.style.borderColor = '#bbf7d0';
                    badgeRisco.textContent = '🟢 SALDO ESTIMADO SEGURO';
                }
            }

            // Atualizar Barra Visual de Consumo Contratual
            document.getElementById('medBarraComprometidoLabel').textContent = `${pctComprometido}% do teto consumido (R$ ${fmt(totalComprometido)} de R$ 1,44M)`;
            document.getElementById('medBarraMedido').style.width = `${pctMedido}%`;
            document.getElementById('medBarraPrevisto').style.width = `${pctPrevisto}%`;

            document.getElementById('medLegendaMedido').textContent = `R$ ${fmt(totalAnualCasa)} (${pctMedido.toFixed(1)}%)`;
            document.getElementById('medLegendaPrevisto').textContent = `R$ ${fmt(valorPrevisaoFutura)} (${pctPrevisto.toFixed(1)}%)`;
            document.getElementById('medLegendaSaldoEstimado').textContent = `R$ ${fmt(saldoEstimado)} (${pctSaldoEstimado}%)`;

            // Atualizar Botão e Tabela de Inspeção de O.S. Futuras
            const btnInspect = document.getElementById('btnToggleInspect');
            if (btnInspect) {
                const boxInspect = document.getElementById('medInspecaoBox');
                if (boxInspect && boxInspect.style.display !== 'none') {
                    btnInspect.innerHTML = '✖️ Ocultar O.S. Futuras';
                } else {
                    btnInspect.innerHTML = `👁️ Inspecionar O.S. Futuras (${countPrevisaoFutura})`;
                }
            }

            const inspecaoBody = document.getElementById('medInspecaoBody');
            if (inspecaoBody) {
                if (chamadosFuturos.length === 0) {
                    inspecaoBody.innerHTML = '<tr><td colspan="5" style="text-align:center; padding: 12px; color: #94a3b8;">Nenhuma O.S. orçada pendente para esta entidade.</td></tr>';
                } else {
                    inspecaoBody.innerHTML = chamadosFuturos
                        .sort((a, b) => b.valor - a.valor)
                        .map(t => `
                            <tr>
                                <td style="padding: 6px 10px; font-weight: 700; color: #4338ca;">#${t.os}</td>
                                <td style="padding: 6px 10px; font-weight: 600;">${t.unidade}</td>
                                <td style="padding: 6px 10px; max-width: 250px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${t.descricao}">${t.descricao}</td>
                                <td style="padding: 6px 10px;"><span style="font-size: 10px; font-weight: 700; background: #e0e7ff; color: #3730a3; padding: 2px 6px; border-radius: 4px;">${t.status}</span></td>
                                <td style="padding: 6px 10px; text-align: right; font-weight: 700; color: #d97706;">R$ ${fmt(t.valor)}</td>
                            </tr>
                        `).join('');
                }
            }

            // 6. Renderizar Linhas da Tabela Oficial de Medições (Itens Liberados p/ NFE)
            let rowsHTML = medList
                .sort((a, b) => b.valor - a.valor)
                .map(t => {
                    const casaBadgeColor = t.casa === 'SESI' ? '#0284c7' : '#ea580c';
                    const casaBadgeBg = t.casa === 'SESI' ? '#e0f2fe' : '#ffedd5';

                    return `
                        <tr>
                            <td style="font-weight: 800; color: #0284c7;">#${t.os}</td>
                            <td style="font-weight: 600;">${t.nr}</td>
                            <td>
                                <span style="display: inline-block; font-size: 11px; font-weight: 800; padding: 2px 8px; border-radius: 4px; background: ${casaBadgeBg}; color: ${casaBadgeColor}; border: 1px solid ${casaBadgeColor}40;">
                                    ${t.casa}
                                </span>
                            </td>
                            <td><strong>${t.unidade}</strong></td>
                            <td style="max-width: 320px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${t.descricao}">${t.descricao}</td>
                            <td><span style="font-size: 12px; font-weight: 700; color: #475569;">${t.mes_emissao}</span></td>
                            <td><span class="status-badge status-concluido">CONCLUIDO</span></td>
                            <td>
                                <span style="font-size: 11px; font-weight: 700; padding: 3px 8px; border-radius: 4px; display: inline-flex; align-items: center; gap: 4px; background: #dcfce7; color: #15803d; border: 1px solid #bbf7d0;">
                                    ✅ LIBERADO P/ NFE
                                </span>
                            </td>
                            <td style="text-align: right; font-weight: 800; color: #10b981; font-size: 13px;">
                                R$ ${fmt(t.valor)}
                            </td>
                        </tr>
                    `;
                }).join('');

            if (medList.length === 0) {
                rowsHTML = '<tr><td colspan="9" style="text-align:center; padding: 24px; color: #999;">Nenhuma Ordem de Serviço com status LIBERADO P/ NFE para a entidade e período selecionados.</td></tr>';
            }

            document.getElementById('medTableBody').innerHTML = rowsHTML;

            // Rodapé Fixo da Tabela
            const periodoLabel = state.medFilters.periodo === 'ANUAL' ? 'ACUMULADO ANUAL DO EXERCÍCIO' : state.medFilters.periodo;
            const footHTML = `
                <tr>
                    <td colspan="4" style="color: #0f172a; font-size: 13px;">
                        📌 TOTAL GERAL DA MEDIÇÃO: <span style="color: #0284c7; font-weight: 800;">${casaAtual} | ${periodoLabel}</span>
                    </td>
                    <td colspan="4" style="text-align: right; color: #475569; font-size: 12px;">
                        Liberado para NFE: <strong style="color: #10b981;">R$ ${fmt(totalMedicao)} (${countMedicao} O.S.)</strong> &nbsp;|&nbsp; Total da Seleção (${countMedicao} O.S.):
                    </td>
                    <td style="text-align: right; font-size: 15px; color: #0f172a; background: #e2e8f0; font-weight: 800;">
                        R$ ${fmt(totalMedicao)}
                    </td>
                </tr>
            `;
            document.getElementById('medTableFoot').innerHTML = footHTML;

            // Faixa Inferior de Resumo com Alerta de Margem Livre
            document.getElementById('medRibbonText').innerHTML = `
                🏷️ <strong>Resumo da Medição:</strong> Entidade: <u>${casaAtual}</u> | Competência: <u>${periodoLabel}</u> | O.S. Faturadas: <strong>${countMedicao} itens</strong> &nbsp;|&nbsp; 🛡️ Margem Real Livre Restante: <strong style="color: #15803d;">R$ ${fmt(saldoEstimado)}</strong>
            `;
            document.getElementById('medRibbonVal').textContent = `Total Medição: R$ ${fmt(totalMedicao)}`;
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

            let plain = `RELAÇÃO DE SERVIÇOS LIBERADOS PARA EMISSÃO DE NF-E\\nEntidade: ${casaTxt} | Competência: ${mesTxt}\\n\\n`;
            plain += `O.S\\tNR\\tCasa\\tUnidade\\tDescrição\\tCompetência\\tStatus Liberação\\tValor Autorizado\\n`;

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

                plain += `${t.os}\\t${t.nr}\\t${t.casa}\\t${t.unidade}\\t${t.descricao}\\t${t.mes_emissao}\\tLIBERADO P/ NFE\\tR$ ${fmt(t.valor)}\\n`;
            });

            html += `
                        </tbody>
                        <tfoot>
                            <tr style="background-color: #f1f5f9; font-weight: bold;">
                                <td colspan="7" style="padding: 10px; border: 1px solid #cccccc; text-align: right;">VALOR TOTAL AUTORIZADO:</td>
                                <td style="padding: 10px; border: 1px solid #cccccc; text-align: right; color: #0ca30c; font-size: 14px;">R$ ${fmt(valLiberado)}</td>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            `;
            plain += `\\nVALOR TOTAL AUTORIZADO: R$ ${fmt(valLiberado)}\\n`;

            try {
                if (navigator.clipboard && window.ClipboardItem) {
                    const blobHtml = new Blob([html], { type: 'text/html' });
                    const blobText = new Blob([plain], { type: 'text/plain' });
                    const item = new ClipboardItem({ 'text/html': blobHtml, 'text/plain': blobText });
                    navigator.clipboard.write([item]).then(() => {
                        showCopyToast('✅ Tabela copiada! Cole com Ctrl + V no seu e-mail.');
                    }).catch(() => {
                        copiarFallback(plain);
                    });
                } else {
                    copiarFallback(plain);
                }
            } catch (err) {
                copiarFallback(plain);
            }
        }

        function copiarFallback(text) {
            const ta = document.createElement('textarea');
            ta.value = text;
            ta.style.position = 'fixed';
            ta.style.opacity = '0';
            document.body.appendChild(ta);
            ta.select();
            document.execCommand('copy');
            document.body.removeChild(ta);
            showCopyToast('✅ Dados copiados! Pressione Ctrl + V no seu e-mail.');
        }

        function imprimirPainel(elementId, titulo) {
            const el = document.getElementById(elementId);
            if (!el) return;

            const parentView = el.closest('.tab-view');
            if (parentView) parentView.style.display = 'block';

            document.querySelectorAll('.target-print-active').forEach(n => n.classList.remove('target-print-active'));
            document.querySelectorAll('.print-header-stamp').forEach(n => n.remove());

            const stamp = document.createElement('div');
            stamp.className = 'print-header-stamp';
            stamp.innerHTML = `
                <h2>🏢 SESI / SENAI — Gestão de Manutenção Predial</h2>
                <p><strong>Relatório:</strong> ${titulo || 'Painel'} | <strong>Emissão:</strong> ${new Date().toLocaleString('pt-BR')}</p>
            `;
            el.insertBefore(stamp, el.firstChild);

            el.classList.add('target-print-active');
            document.body.classList.add('is-printing-panel');

            const tituloOriginal = document.title;
            if (titulo) {
                document.title = `${titulo} - SESI SENAI`;
            }

            window.print();

            setTimeout(() => {
                document.body.classList.remove('is-printing-panel');
                el.classList.remove('target-print-active');
                if (stamp.parentNode) stamp.remove();
                document.title = tituloOriginal;
            }, 1000);
        }

        function imprimirRelatorioGeral() {
            document.body.classList.remove('is-printing-panel');
            document.querySelectorAll('.target-print-active').forEach(n => n.classList.remove('target-print-active'));
            window.print();
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
                    const statusClass = 'status-' + t.status.toLowerCase().replace(/\\s+/g, '-').replace(/[^\\w-]/g, '');
                    
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
                        diasHTML = `<td style="font-weight: 700; color: ${cor}; background: ${fundo}; border-radius: 4px; padding: 6px 10px; text-align: center;">${t.dias_abertos}d</td>`;
                    } else {
                        diasHTML = `<td style="text-align: center; color: #999;">-</td>`;
                    }

                    let orcHTML = '';
                    if (t.data_orcamento) {
                        const diasTxt = t.dias_orcamento !== null ? ` <span style="font-size: 11px; font-weight: 700; color: #0d6efd;">(${t.dias_orcamento}d)</span>` : '';
                        orcHTML = `<td style="text-align: center; font-size: 12px; font-weight: 600; color: #1e293b;">📅 ${t.data_orcamento}${diasTxt}</td>`;
                    } else if (t.status === 'AGUARDANDO ORÇAMENTO' && t.dias_orcamento !== null) {
                        orcHTML = `<td style="text-align: center; font-size: 11px; font-weight: 700; color: #b45309; background: #fef3c7; border-radius: 4px;">⏳ Aguardando (${t.dias_orcamento}d)</td>`;
                    } else {
                        orcHTML = `<td style="text-align: center; color: #999;">-</td>`;
                    }

                    const valorFmt = new Intl.NumberFormat('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2}).format(t.valor);

                    return `
                        <tr>
                            <td style="font-weight: 600; color: #0d6efd;">#${t.os}</td>
                            <td>${t.nr}</td>
                            <td><strong>${t.unidade}</strong></td>
                            <td style="max-width: 280px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${t.descricao}">${t.descricao}</td>
                            <td><span class="status-badge ${statusClass}">${t.status}</span></td>
                            <td style="text-align: center; font-size: 12px; font-weight: 600; color: #334155;">${t.data_envio || '-'}</td>
                            ${orcHTML}
                            ${diasHTML}
                            <td style="text-align: right; font-weight: 600;">R$ ${valorFmt}</td>
                        </tr>
                    `;
                }).join('');

            if (filtered.length === 0) {
                html = '<tr><td colspan="9" style="text-align:center; padding: 24px; color: #999;">Nenhum chamado localizado para os filtros selecionados.</td></tr>';
            }

            document.getElementById('tableBody').innerHTML = html;
        }

        function render() {
            renderKPIs();
            renderContracts();
            renderTempoAberto();
            renderCharts();
            renderPaymentPanel();
            renderMedicoesPanel();
            renderTable();
        }

        init();
    </script>
</body>
</html>"""

st.subheader("📊 Dashboard Interativo")

html_content = HTML_TEMPLATE.replace('__DATA_PLACEHOLDER__', json.dumps(data, ensure_ascii=False))

# Altura calibrada para a extensão real de cada aba com rolagem ativada, eliminando o vácuo de tela
st.components.v1.html(html_content, height=1400, scrolling=True)

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
