import io
import json
import re
from datetime import date, datetime
import numpy as np
import openpyxl
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# ==============================================================================
# CONFIGURAÇÃO DA PÁGINA STREAMLIT
# ==============================================================================
st.set_page_config(
    page_title="Painel de Manutenção Predial & Liberação de Pagamentos",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilização da interface Streamlit
st.markdown("""
<style>
    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 2rem;
        padding-left: 2rem;
        padding-right: 2rem;
    }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)


# ==============================================================================
# FUNÇÕES DE PROCESSAMENTO E NORMALIZAÇÃO DE DADOS
# ==============================================================================
def norm_str(s):
    if pd.isna(s) or s is None:
        return ""
    return str(s).strip()

def norm_unidade(u):
    if pd.isna(u) or u is None:
        return "NÃO INFORMADA"
    s = str(u).strip().upper()
    s = re.sub(r'\s+', ' ', s)
    if "CENTRAL" in s and "PARK" in s:
        return "SESI CENTRAL PARK"
    if "SIA" in s:
        return "SENAI SIA"
    if "SIG" in s:
        return "SENAI SIG"
    if "FIBRA" in s:
        if "SEDE" in s or "EDIF" in s:
            return "EDIFÍCIO FIBRA (SEDE)"
        return "FIBRA"
    return s

def detectar_casa(unidade_str, entidade_cnpj=""):
    u = (unidade_str or "").upper()
    e = (entidade_cnpj or "").upper()
    if "FIBRA" in e or "FIBRA" in u:
        return "FIBRA"
    if "SENAI" in e or "SENAI" in u:
        return "SENAI"
    if "SESI" in e or "SESI" in u:
        return "SESI"
    return "SESI"

def parse_data_br(val):
    """
    Interpretação rigorosa de datas brasileiras no padrão DIA/MÊS/ANO (DD/MM/AAAA).
    Não realiza inversão automática entre dia e mês.
    """
    if pd.isna(val) or val is None:
        return None
    if isinstance(val, (datetime, date, pd.Timestamp)):
        return val.strftime('%d/%m/%Y')
    
    s = str(val).strip().split(' ')[0]
    if not s or s.lower() in ['nan', 'none', '-', 'nat']:
        return None
    
    parts = re.split(r'[/.-]', s)
    if len(parts) == 3:
        p0, p1, p2 = parts[0], parts[1], parts[2]
        try:
            # Caso esteja no formato AAAA-MM-DD
            if len(p0) == 4:
                ano = int(p0)
                mes = int(p1)
                dia = int(p2)
                return f"{dia:02d}/{mes:02d}/{ano:04d}"
            # Caso padrão brasileiro DD/MM/AAAA
            elif len(p2) == 4 or len(p2) == 2:
                dia = int(p0)
                mes = int(p1)
                ano = int(p2)
                if ano < 100:
                    ano += 2000
                return f"{dia:02d}/{mes:02d}/{ano:04d}"
        except Exception:
            pass
    return s

def calcular_dias_abertos(data_envio_str, status_str):
    if not data_envio_str or status_str == 'CONCLUIDO':
        return None
    try:
        parts = data_envio_str.split('/')
        if len(parts) == 3:
            dt = datetime(int(parts[2]), int(parts[1]), int(parts[0])).date()
            hoje = date.today()
            diff = (hoje - dt).days
            return max(0, diff)
    except Exception:
        pass
    return None

def calcular_prazo_orcamento(data_envio_str, data_orc_str):
    if not data_envio_str:
        return "-"
    try:
        parts_env = data_envio_str.split('/')
        dt_env = datetime(int(parts_env[2]), int(parts_env[1]), int(parts_env[0])).date()
        
        if data_orc_str and data_orc_str != "-":
            parts_orc = data_orc_str.split('/')
            dt_orc = datetime(int(parts_orc[2]), int(parts_orc[1]), int(parts_orc[0])).date()
            dias = (dt_orc - dt_env).days
            return f"{data_orc_str} ({max(0, dias)}d)"
        else:
            hoje = date.today()
            dias = (hoje - dt_env).days
            return f"⏳ Aguardando ({max(0, dias)}d)"
    except Exception:
        return data_orc_str or "-"


# ==============================================================================
# CARREGAMENTO DA PLANILHA E TRATAMENTO SEGURO
# ==============================================================================
uploaded_file = st.file_uploader("📂 Envie o arquivo Excel (CONTROLE_DE_O_S.xlsx)", type=["xlsx", "xls"])

if uploaded_file is None:
    st.info("👋 Por favor, faça o upload da sua planilha Excel de Manutenção Predial acima para gerar o painel interativo.")
    st.stop()

try:
    xl = pd.ExcelFile(uploaded_file)
    target_sheet = 'O.S' if 'O.S' in xl.sheet_names else xl.sheet_names[0]
    
    # Detecção dinâmica da linha do cabeçalho nas primeiras 15 linhas
    df_raw = xl.parse(target_sheet, header=None, nrows=15)
    header_idx = 0
    for idx, row in df_raw.iterrows():
        row_str = " ".join([str(v).upper() for v in row.values if pd.notna(v)])
        if ("O.S" in row_str or "OS" in row_str) and ("STATUS" in row_str or "UNIDADE" in row_str or "DESCRI" in row_str or "VALOR" in row_str):
            header_idx = idx
            break
            
    df_os = xl.parse(target_sheet, header=header_idx)
    df_os = df_os.dropna(how='all', axis=1)
except Exception as e:
    st.error(f"Erro ao processar o arquivo Excel: {str(e)}")
    st.stop()

# Localização flexível e protegida das colunas necessárias
cols_upper = {str(c).strip().upper(): c for c in df_os.columns}

def find_col(possible_names):
    for name in possible_names:
        for c_up, orig in cols_upper.items():
            if name in c_up:
                return orig
    return None

os_col = find_col(['O.S', 'Nº O.S', 'NUMERO O.S', 'NUMERO DA O.S', 'OS'])
nr_col = find_col(['NR', 'Nº', 'ITEM', 'NUM'])
unidade_col = find_col(['UNIDADE', 'UNID', 'ESCOLA', 'LOCAL', 'PREDIO'])
status_col = find_col(['STATUS DA O.S', 'STATUS O.S', 'STATUS', 'SITUAÇÃO'])
valor_col = find_col(['VALOR A FATURAR', 'VALOR FATURAR', 'VALOR ORÇADO', 'VALOR', 'INVESTIMENTO'])
mes_comp_col = find_col(['MÊS COMPETÊNCIA', 'MES COMPETENCIA', 'COMPETÊNCIA', 'COMPETENCIA', 'MÊS', 'MES'])
nfe_status_col = find_col(['LIBERAÇÃO P/ NFE', 'LIBERACAO P/ NFE', 'NFE', 'NOTA FISCAL', 'LIBERAÇÃO', 'LIBERACAO'])
data_envio_col = find_col(['DATA ENVIO', 'DATA DE ENVIO', 'DT ENVIO', 'DATA ABERTURA'])
data_orc_col = find_col(['ENTREGA ORÇAMENTO', 'ENTREGA ORCAMENTO', 'DATA ORÇAMENTO', 'DT ORCAMENTO'])

# Busca prioritária para a Descrição Resumida (evitando colunas de data/hora/valor)
desc_col = None
for c_up, orig in cols_upper.items():
    if "DESCRI" in c_up and "RESUM" in c_up:
        desc_col = orig
        break
if not desc_col:
    for c_up, orig in cols_upper.items():
        if ("DESCRI" in c_up or "SERVIÇO" in c_up or "SERVICO" in c_up) and not any(k in c_up for k in ["DATA", "DT", "HORA", "VALOR", "STATUS", "PREVISÃO"]):
            desc_col = orig
            break

# Tetos Contratuais Homologados
saldo_sesi = 1440000.0
saldo_senai = 1440000.0
saldo_fibra = 80000.0 # Teto contratual oficial da FIBRA: R$ 80.000,00

if 'SALDO' in xl.sheet_names:
    try:
        df_saldo = xl.parse('SALDO')
        for idx, row in df_saldo.iterrows():
            r_str = " ".join([str(v).upper() for v in row.values if pd.notna(v)])
            if "FIBRA" in r_str:
                for v in row.values:
                    if isinstance(v, (int, float)) and v > 1000:
                        saldo_fibra = float(v)
            elif "SENAI" in r_str:
                for v in row.values:
                    if isinstance(v, (int, float)) and v > 50000:
                        saldo_senai = float(v)
            elif "SESI" in r_str:
                for v in row.values:
                    if isinstance(v, (int, float)) and v > 50000:
                        saldo_sesi = float(v)
    except Exception:
        pass


# ==============================================================================
# CONSTRUÇÃO DO CONJUNTO DE DADOS
# ==============================================================================
tickets = []
valid_rows = df_os[df_os[os_col].notna()] if os_col else df_os

for _, row in valid_rows.iterrows():
    raw_os = str(row[os_col]).split('.')[0].strip() if os_col and pd.notna(row[os_col]) else ""
    if not raw_os or raw_os.upper() in ["TOTAL", "SUBTOTAL"]:
        continue
        
    nr_val = str(row[nr_col]).split('.')[0].strip() if nr_col and pd.notna(row[nr_col]) else "-"
    unidade_val = norm_unidade(row[unidade_col]) if unidade_col and pd.notna(row[unidade_col]) else "NÃO INFORMADA"
    status_val = str(row[status_col]).strip().upper() if status_col and pd.notna(row[status_col]) else "SEM STATUS"
    
    # Tratamento da Descrição Resumida (eliminando datas acidentais)
    desc_val = ""
    if desc_col and pd.notna(row[desc_col]):
        raw_desc = str(row[desc_col]).strip()
        if not re.match(r'^\d{4}-\d{2}-\d{2}', raw_desc) and not re.match(r'^\d{2}/\d{2}/\d{4}', raw_desc):
            desc_val = raw_desc
    if not desc_val:
        for c_up, orig in cols_upper.items():
            if orig != desc_col and any(k in c_up for k in ["DESC", "SERVIC", "OBJETO", "ITEM"]):
                if not any(k in c_up for k in ["DATA", "DT", "HORA", "VALOR", "STATUS", "PREVISÃO"]):
                    cand = str(row[orig]).strip() if pd.notna(row[orig]) else ""
                    if cand and not re.match(r'^\d{4}-\d{2}-\d{2}', cand) and not re.match(r'^\d{2}/\d{2}/\d{4}', cand):
                        desc_val = cand
                        break
            
    # Tratamento de valores financeiros
    val_num = 0.0
    if valor_col and pd.notna(row[valor_col]):
        try:
            val_clean = str(row[valor_col]).replace('R$', '').replace('.', '').replace(',', '.').strip()
            val_num = float(val_clean)
        except Exception:
            try:
                val_num = float(row[valor_col])
            except Exception:
                val_num = 0.0
                
    # Competência e NFE
    mes_comp = str(row[mes_comp_col]).strip().upper() if mes_comp_col and pd.notna(row[mes_comp_col]) else "NÃO DEFINIDO"
    nfe_val = str(row[nfe_status_col]).strip().upper() if nfe_status_col and pd.notna(row[nfe_status_col]) else ""
    
    # Datas no padrão brasileiro
    dt_envio_br = parse_data_br(row[data_envio_col]) if data_envio_col and pd.notna(row[data_envio_col]) else None
    dt_orc_br = parse_data_br(row[data_orc_col]) if data_orc_col and pd.notna(row[data_orc_col]) else None
    
    prazo_orc_txt = calcular_prazo_orcamento(dt_envio_br, dt_orc_br)
    dias_abertos = calcular_dias_abertos(dt_envio_br, status_val)
    casa_val = detectar_casa(unidade_val, str(row.get('ENTIDADE', '')))

    tickets.append({
        "os": raw_os,
        "nr": nr_val,
        "unidade": unidade_val,
        "casa": casa_val,
        "descricao": desc_val,
        "status": status_val,
        "mes_competencia": mes_comp,
        "liberado_nfe": "LIBERADO P/ NFE" if ("LIBERADO" in nfe_val and "NFE" in nfe_val) else nfe_val,
        "valor": val_num,
        "data_envio": dt_envio_br or "-",
        "entrega_orcamento": prazo_orc_txt,
        "dias_abertos": dias_abertos
    })

payload_data = {
    "chamados": tickets,
    "contratos": {
        "SESI": {"teto": saldo_sesi},
        "SENAI": {"teto": saldo_senai},
        "FIBRA": {"teto": saldo_fibra}
    }
}
DATA_JSON = json.dumps(payload_data, ensure_ascii=False)


# ==============================================================================
# TEMPLATE HTML / CSS / JAVASCRIPT DO PAINEL
# ==============================================================================
HTML_TEMPLATE = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Painel de Manutenção Predial & Liberação de Pagamentos</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
            background: #f8fafc;
            color: #1e293b;
            padding: 16px;
            line-height: 1.45;
        }}
        .container {{ max-width: 1440px; margin: 0 auto; }}

        /* Top Banner */
        .top-banner {{
            background: #ffffff;
            border-radius: 12px;
            padding: 18px 24px;
            margin-bottom: 16px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            border-left: 5px solid #0d6efd;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 14px;
        }}
        .top-banner h1 {{ font-size: 20px; font-weight: 800; color: #0f172a; display: flex; align-items: center; gap: 10px; }}
        .top-banner p {{ font-size: 13px; color: #64748b; margin-top: 2px; }}

        /* Barra de Navegação por Abas */
        .nav-tabs {{
            display: flex;
            gap: 10px;
            margin-bottom: 20px;
            flex-wrap: wrap;
        }}
        .tab-btn {{
            padding: 10px 18px;
            background: #ffffff;
            border: 1px solid #cbd5e1;
            border-radius: 8px;
            font-size: 13px;
            font-weight: 700;
            color: #475569;
            cursor: pointer;
            transition: all 0.2s;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            user-select: none;
        }}
        .tab-btn:hover {{ background: #f1f5f9; color: #0f172a; }}
        .tab-btn.active {{
            background: #0d6efd;
            color: #ffffff;
            border-color: #0d6efd;
            box-shadow: 0 2px 6px rgba(13, 110, 253, 0.25);
        }}

        /* Cards de Atalho do Painel Inicial */
        .shortcut-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 14px;
            margin-bottom: 20px;
        }}
        .shortcut-card {{
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 10px;
            padding: 16px 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 1px 3px rgba(0,0,0,0.03);
            border-left: 4px solid #0d6efd;
            cursor: pointer;
            transition: all 0.2s;
        }}
        .shortcut-card:hover {{ transform: translateY(-2px); box-shadow: 0 4px 10px rgba(0,0,0,0.06); }}
        .shortcut-card.med {{ border-left-color: #0284c7; }}
        .shortcut-card.list {{ border-left-color: #10b981; }}
        .shortcut-title {{ font-size: 14px; font-weight: 800; color: #0f172a; }}
        .shortcut-desc {{ font-size: 12px; color: #64748b; margin-top: 2px; }}
        .shortcut-btn {{
            padding: 6px 12px;
            background: #eff6ff;
            color: #0d6efd;
            font-size: 12px;
            font-weight: 700;
            border-radius: 6px;
            white-space: nowrap;
        }}

        /* Estrutura de Cards */
        .card {{
            background: #ffffff;
            border-radius: 10px;
            border: 1px solid #e2e8f0;
            padding: 20px;
            margin-bottom: 20px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        }}
        .card-title-bar {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 16px;
            padding-bottom: 12px;
            border-bottom: 1px solid #f1f5f9;
            flex-wrap: wrap;
            gap: 10px;
        }}
        .card-title {{ font-size: 16px; font-weight: 800; color: #0f172a; display: flex; align-items: center; gap: 8px; }}

        /* KPIs */
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 14px;
            margin-bottom: 20px;
        }}
        .kpi-card {{
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            padding: 14px 16px;
            border-left: 4px solid #cbd5e1;
        }}
        .kpi-card.blue {{ border-left-color: #0d6efd; }}
        .kpi-card.green {{ border-left-color: #10b981; }}
        .kpi-card.amber {{ border-left-color: #f59e0b; }}
        .kpi-card.red {{ border-left-color: #ef4444; }}
        .kpi-card.purple {{ border-left-color: #7c3aed; }}
        .kpi-card.dark {{ border-left-color: #475569; }}
        .kpi-label {{ font-size: 11px; font-weight: 700; text-transform: uppercase; color: #64748b; margin-bottom: 4px; }}
        .kpi-val {{ font-size: 22px; font-weight: 800; color: #0f172a; line-height: 1.2; }}
        .kpi-val.green {{ color: #10b981; }}
        .kpi-val.blue {{ color: #0d6efd; }}
        .kpi-val.red {{ color: #ef4444; }}
        .kpi-val.purple {{ color: #7c3aed; }}
        .kpi-sub {{ font-size: 12px; color: #64748b; margin-top: 4px; }}

        /* Botões de Ação de Entidade na Lista */
        .list-casa-btn {{
            padding: 6px 13px;
            background: #ffffff;
            border: 1px solid #cbd5e1;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 700;
            color: #475569;
            cursor: pointer;
            transition: all 0.2s;
            display: inline-flex;
            align-items: center;
            gap: 5px;
            user-select: none;
        }}
        .list-casa-btn:hover {{ background: #f1f5f9; color: #0f172a; }}
        .list-casa-btn.active {{
            background: #0f172a;
            color: #ffffff;
            border-color: #0f172a;
        }}
        .list-casa-btn.active-sesi {{
            background: #0d6efd;
            color: #ffffff;
            border-color: #0d6efd;
        }}
        .list-casa-btn.active-senai {{
            background: #ea580c;
            color: #ffffff;
            border-color: #ea580c;
        }}
        .list-casa-btn.active-fibra {{
            background: #7c3aed;
            color: #ffffff;
            border-color: #7c3aed;
        }}

        /* Filtros Gerais */
        .filter-group {{ display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 14px; }}
        .filter-chip {{
            padding: 6px 14px;
            background: #ffffff;
            border: 1px solid #cbd5e1;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 700;
            color: #475569;
            cursor: pointer;
            transition: all 0.2s;
            user-select: none;
        }}
        .filter-chip:hover {{ background: #e2e8f0; color: #0f172a; }}
        .filter-chip.active {{ background: #0d6efd; color: #fff; border-color: #0d6efd; }}
        .filter-chip.active-senai {{ background: #ea580c; color: #fff; border-color: #ea580c; }}
        .filter-chip.active-fibra {{ background: #7c3aed; color: #fff; border-color: #7c3aed; }}
        .filter-chip.active-dark {{ background: #0f172a; color: #fff; border-color: #0f172a; }}

        .filter-input {{
            padding: 8px 12px;
            border: 1px solid #cbd5e1;
            border-radius: 6px;
            font-size: 13px;
            color: #1e293b;
            width: 100%;
        }}

        /* Tabelas */
        .table-wrapper {{
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            overflow-x: auto;
            max-height: 520px;
            background: #ffffff;
        }}
        table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
        thead th {{
            position: sticky;
            top: 0;
            background: #f8fafc;
            color: #475569;
            font-weight: 700;
            padding: 10px 12px;
            text-align: left;
            border-bottom: 2px solid #cbd5e1;
            z-index: 5;
        }}
        td {{ padding: 9px 12px; border-bottom: 1px solid #f1f5f9; }}
        tbody tr:hover {{ background: #f8fafc; }}

        /* Status Pills */
        .pill {{
            padding: 3px 8px;
            border-radius: 4px;
            font-size: 11px;
            font-weight: 700;
            display: inline-block;
        }}
        .pill.concluido {{ background: #dcfce7; color: #15803d; }}
        .pill.execucao {{ background: #e0f2fe; color: #0369a1; }}
        .pill.paralisado {{ background: #fee2e2; color: #b91c1c; }}
        .pill.orcamento {{ background: #fef3c7; color: #b45309; }}
        .pill.liberado-nfe {{ background: #dcfce7; color: #15803d; border: 1px solid #bbf7d0; }}

        /* Botões de Ação */
        .btn-action {{
            padding: 7px 12px;
            background: #ffffff;
            border: 1px solid #cbd5e1;
            border-radius: 6px;
            font-size: 12px;
            font-weight: 700;
            color: #475569;
            cursor: pointer;
            transition: all 0.2s;
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }}
        .btn-action:hover {{ background: #f1f5f9; color: #0f172a; }}

        .btn-back-nav {{
            padding: 8px 14px;
            background: #ffffff;
            border: 1px solid #cbd5e1;
            border-radius: 6px;
            font-size: 13px;
            font-weight: 700;
            color: #475569;
            cursor: pointer;
        }}
        .btn-back-nav:hover {{ background: #f1f5f9; }}

        /* Toast de Feedback */
        .toast-msg {{
            position: fixed;
            bottom: 24px;
            right: 24px;
            background: #0f172a;
            color: #ffffff;
            padding: 12px 20px;
            border-radius: 8px;
            font-size: 13px;
            font-weight: 600;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            display: none;
            z-index: 9999;
        }}

        /* Gestão de Risco e Previsão */
        .previsao-box {{
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            padding: 16px;
            margin-bottom: 18px;
        }}
        .progress-bar-wrap {{
            height: 12px;
            background: #e2e8f0;
            border-radius: 6px;
            overflow: hidden;
            display: flex;
            margin: 10px 0;
        }}
        .progress-bar-seg {{ height: 100%; transition: width 0.3s; }}

        /* Gaveta Retrátil de O.S. Futuras */
        .drawer-futuras {{
            margin-top: 14px;
            padding-top: 14px;
            border-top: 1px dashed #cbd5e1;
            display: none;
        }}
        .drawer-futuras.open {{ display: block; }}

        /* Isolamento para Impressão sem páginas em branco */
        #printSection {{ display: none; }}
        @media print {{
            body * {{ visibility: hidden; }}
            #printSection, #printSection * {{ visibility: visible; }}
            #printSection {{
                position: absolute;
                left: 0;
                top: 0;
                width: 100%;
                display: block !important;
                background: #ffffff !important;
                padding: 10px !important;
            }}
            .no-print {{ display: none !important; }}
            table {{ font-size: 10px !important; }}
            th, td {{ padding: 6px 8px !important; }}
            @page {{ size: landscape; margin: 10mm; }}
        }}
    </style>
</head>
<body>
<div class="container" id="mainContainer">
    
    <!-- Topo Institucional -->
    <div class="top-banner">
        <div>
            <h1>🏢 Painel de Manutenção Predial & Liberação de Pagamentos</h1>
            <p>SESI, SENAI e FIBRA — Acompanhamento de Chamados, Emissão de NFE e Gestão Orçamentária</p>
        </div>
        <button class="btn-action no-print" onclick="imprimirRelatorioCompleto()">🖨️ Imprimir Relatório Completo</button>
    </div>

    <!-- Navegação em Abas -->
    <div class="nav-tabs no-print">
        <button class="tab-btn active" id="tabBtn-inicial" onclick="switchTab('inicial')">📊 Painel Inicial (Visão Geral)</button>
        <button class="tab-btn" id="tabBtn-pagamentos" onclick="switchTab('pagamentos')">💳 Controle de Pagamentos & NF-e</button>
        <button class="tab-btn" id="tabBtn-medicoes" onclick="switchTab('medicoes')">📐 MEDIÇÕES (Controle Contratual)</button>
        <button class="tab-btn" id="tabBtn-chamados" onclick="switchTab('chamados')">📋 Lista Completa de Chamados</button>
    </div>

    <!-- ============================================================== -->
    <!-- ABA 1: PAINEL INICIAL (VISÃO GERAL)                             -->
    <!-- ============================================================== -->
    <div class="tab-view" id="view-inicial">
        <div class="shortcut-grid no-print">
            <div class="shortcut-card" onclick="switchTab('pagamentos')">
                <div>
                    <div class="shortcut-title">💳 Controle de Pagamentos & NF-e</div>
                    <div class="shortcut-desc">Conferência por Mês e CNPJ para autorizar emissão de NF</div>
                </div>
                <div class="shortcut-btn">Acessar Faturamento →</div>
            </div>
            <div class="shortcut-card med" onclick="switchTab('medicoes')">
                <div>
                    <div class="shortcut-title">📐 Módulo MEDIÇÕES (Contrato)</div>
                    <div class="shortcut-desc">Livro Oficial de Medições mensais e saldo contratual</div>
                </div>
                <div class="shortcut-btn">Acessar Medições →</div>
            </div>
            <div class="shortcut-card list" onclick="switchTab('chamados')">
                <div>
                    <div class="shortcut-title">📋 Lista Completa de Chamados</div>
                    <div class="shortcut-desc">Tabela detalhada com filtros de status, prazos e busca</div>
                </div>
                <div class="shortcut-btn">Ver Chamados →</div>
            </div>
        </div>

        <div class="kpi-grid" id="mainKPIs"></div>

        <div class="card" id="painelSaldos">
            <div class="card-title-bar">
                <div class="card-title">💰 Saldo em Contrato</div>
                <button class="btn-action no-print" onclick="imprimirPainel('painelSaldos', 'Saldo em Contrato')">🖨️ Imprimir</button>
            </div>
            <div id="contractCards"></div>
        </div>

        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px;">
            <div class="card" id="painelTempoAberto">
                <div class="card-title-bar">
                    <div class="card-title">⏱️ Tempo Médio em Aberto</div>
                    <button class="btn-action no-print" onclick="imprimirPainel('painelTempoAberto', 'Tempo Médio em Aberto')">🖨️ Imprimir</button>
                </div>
                <div id="tempoAbertoContent"></div>
            </div>

            <div class="card" id="painelGraficoStatus">
                <div class="card-title-bar">
                    <div class="card-title">📊 Chamados por Status</div>
                    <button class="btn-action no-print" onclick="imprimirPainel('painelGraficoStatus', 'Status dos Chamados')">🖨️ Imprimir</button>
                </div>
                <div style="height: 240px; position: relative;">
                    <canvas id="statusChart"></canvas>
                </div>
            </div>
        </div>

        <div class="card" id="painelUnidades">
            <div class="card-title-bar">
                <div class="card-title">🏢 Chamados por Unidade (com investimento)</div>
                <button class="btn-action no-print" onclick="imprimirPainel('painelUnidades', 'Chamados por Unidade')">🖨️ Imprimir Unidades</button>
            </div>
            <div id="unitsGrid"></div>
        </div>
    </div>

    <!-- ============================================================== -->
    <!-- ABA 2: CONTROLE DE PAGAMENTOS & NF-E                            -->
    <!-- ============================================================== -->
    <div class="tab-view" id="view-pagamentos" style="display: none;">
        <div class="no-print" style="margin-bottom: 14px;">
            <button class="btn-back-nav" onclick="switchTab('inicial')">⬅️ Retornar ao Painel Inicial</button>
        </div>

        <div class="card" id="painelPagamentos">
            <div class="card-title-bar">
                <div class="card-title">💳 Controle de Pagamentos & Liberação para NFE</div>
                <div style="display: flex; gap: 8px;" class="no-print">
                    <button class="btn-action" onclick="copiarTabelaEmail()">📋 Copiar p/ E-mail</button>
                    <button class="btn-action" onclick="copiarTabelaExcel()">📊 Copiar p/ Planilha</button>
                    <button class="btn-action" onclick="imprimirPainel('painelPagamentos', 'Controle de Pagamentos')">🖨️ Imprimir Faturamento</button>
                </div>
            </div>

            <div class="kpi-grid" id="payKPIs"></div>

            <div class="no-print" style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 14px; margin-bottom: 16px;">
                <div style="font-size: 12px; font-weight: 700; color: #475569; margin-bottom: 6px;">🏛️ ENTIDADE / CNPJ (CASA):</div>
                <div class="filter-group" id="payCasaFilter"></div>

                <div style="font-size: 12px; font-weight: 700; color: #475569; margin-bottom: 6px;">📅 MÊS DE COMPETÊNCIA:</div>
                <div class="filter-group" id="payMesFilter"></div>
            </div>

            <div class="table-wrapper">
                <table>
                    <thead>
                        <tr>
                            <th>O.S</th>
                            <th>NR</th>
                            <th>Entidade / CNPJ</th>
                            <th>Unidade</th>
                            <th>Descrição Resumida</th>
                            <th>Mês Competência</th>
                            <th>Status O.S</th>
                            <th>Liberação p/ NFE</th>
                            <th style="text-align: right;">Valor a Faturar</th>
                        </tr>
                    </thead>
                    <tbody id="payTableBody"></tbody>
                    <tfoot>
                        <tr style="background: #f1f5f9; font-weight: 800;">
                            <td colspan="7" id="payFootDesc">TOTAL</td>
                            <td id="payFootLib" style="color: #10b981;">-</td>
                            <td id="payFootTotal" style="text-align: right; color: #0f172a; font-size: 14px;">-</td>
                        </tr>
                    </tfoot>
                </table>
            </div>

            <div style="margin-top: 14px; padding: 12px 16px; background: #eff6ff; border-radius: 6px; display: flex; justify-content: space-between; align-items: center;" id="paySummaryRibbon">
                <span style="font-size: 13px; color: #1e40af; font-weight: 600;" id="payRibbonTxt">Resumo</span>
                <span style="font-size: 16px; font-weight: 800; color: #1d4ed8;" id="payRibbonVal">R$ 0,00</span>
            </div>
        </div>
    </div>

    <!-- ============================================================== -->
    <!-- ABA 3: MEDIÇÕES (CONTROLE CONTRATUAL)                           -->
    <!-- ============================================================== -->
    <div class="tab-view" id="view-medicoes" style="display: none;">
        <div class="no-print" style="margin-bottom: 14px;">
            <button class="btn-back-nav" onclick="switchTab('inicial')">⬅️ Retornar ao Painel Inicial</button>
        </div>

        <div class="card" id="painelMedicoes">
            <div class="card-title-bar">
                <div>
                    <div class="card-title">📐 MEDIÇÕES — Controle de Medição Contratual</div>
                    <div style="font-size: 12px; color: #64748b; margin-top: 2px;" id="medTetoBadge">Livro Oficial de Medições com Teto Contratual</div>
                </div>
                <div style="display: flex; gap: 8px; flex-wrap: wrap;" class="no-print">
                    <button class="btn-action" onclick="copiarMedicaoEmail()">📋 Copiar p/ E-mail</button>
                    <button class="btn-action" onclick="copiarMedicaoExcel()">📊 Copiar p/ Planilha</button>
                    <button class="btn-action" onclick="imprimirPainel('painelMedicoes', 'Boletim de Medição Contratual')">🖨️ Imprimir Medição</button>
                </div>
            </div>

            <!-- 4 KPIs Originais Contratuais -->
            <div class="kpi-grid" id="medKPIs"></div>

            <!-- Bloco de Gestão de Risco e Previsão de Saldo Estimado -->
            <div class="previsao-box" id="previsaoBox">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; flex-wrap: wrap; gap: 10px;">
                    <div style="font-size: 13px; font-weight: 800; color: #0f172a;" id="prevTitle">🛡️ GESTÃO DE RISCO: PREVISÃO DE MEDIÇÕES FUTURAS & SALDO ESTIMADO</div>
                    <div style="display: flex; align-items: center; gap: 8px;" class="no-print">
                        <button class="btn-action" id="btnToggleFuturas" onclick="toggleFuturasDrawer()">👁️ Inspecionar O.S. Futuras (<span id="prevCountBadge">0</span>)</button>
                        <span class="pill" id="prevRiskBadge" style="background: #dcfce7; color: #15803d;">🟢 MARGEM SEGURA</span>
                        <button class="btn-action" onclick="imprimirFuturas()">🖨️ Imprimir Previsão</button>
                    </div>
                </div>
                <div class="kpi-grid" id="prevKPIs" style="margin-bottom: 10px;"></div>
                <div style="font-size: 12px; color: #64748b; margin-bottom: 4px;">Consumo Proporcional do Teto Homologado:</div>
                <div class="progress-bar-wrap" id="prevProgressBar"></div>
                <div style="display: flex; justify-content: space-between; font-size: 11px; color: #64748b; margin-top: 4px;">
                    <span>🔵 Liberado/Medido</span>
                    <span>🟠 Previsão Futura</span>
                    <span>🟢 Margem Livre Restante</span>
                </div>

                <!-- Gaveta de Inspeção das O.S. Futuras -->
                <div class="drawer-futuras" id="drawerFuturas">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                        <strong style="font-size: 12px; color: #0f172a;">🔍 Detalhamento das O.S. com Investimentos em Andamento / Previsão Futura:</strong>
                        <div style="display: flex; gap: 6px;" class="no-print">
                            <button class="btn-action" onclick="copiarFuturasEmail()">📋 Copiar Futuras p/ E-mail</button>
                            <button class="btn-action" onclick="copiarFuturasExcel()">📊 Copiar Futuras p/ Planilha</button>
                        </div>
                    </div>
                    <div class="table-wrapper" style="max-height: 380px;">
                        <table>
                            <thead>
                                <tr>
                                    <th>O.S</th>
                                    <th>NR</th>
                                    <th>Unidade</th>
                                    <th>Descrição Resumida</th>
                                    <th>Status O.S</th>
                                    <th style="text-align: right;">Previsão (R$)</th>
                                </tr>
                            </thead>
                            <tbody id="futurasTableBody"></tbody>
                            <tfoot>
                                <tr style="background: #f8fafc; font-weight: 800;">
                                    <td colspan="5" id="futurasFootDesc">TOTAL DE PREVISÕES FUTURAS</td>
                                    <td id="futurasFootVal" style="text-align: right; color: #ea580c; font-size: 13px;">R$ 0,00</td>
                                </tr>
                            </tfoot>
                        </table>
                    </div>
                </div>
            </div>

            <!-- Filtros do Módulo de Medição -->
            <div class="no-print" style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 14px; margin-bottom: 16px;">
                <div style="font-size: 12px; font-weight: 700; color: #475569; margin-bottom: 6px;">🏛️ ENTIDADE CONTRATUAL (CASA):</div>
                <div class="filter-group" id="medCasaFilter"></div>

                <div style="font-size: 12px; font-weight: 700; color: #475569; margin-bottom: 6px;">📅 MEDIÇÃO MENSAL DE COMPETÊNCIA OU EXERCÍCIO ANUAL:</div>
                <div class="filter-group" id="medPeriodoFilter"></div>
            </div>

            <!-- Tabela Oficial de Medições -->
            <div class="table-wrapper">
                <table>
                    <thead>
                        <tr>
                            <th>O.S</th>
                            <th>NR</th>
                            <th>Entidade / CNPJ</th>
                            <th>Unidade</th>
                            <th>Descrição Resumida</th>
                            <th>Mês Competência</th>
                            <th>Status O.S</th>
                            <th>Liberação p/ NFE</th>
                            <th style="text-align: right;">Valor a Faturar</th>
                        </tr>
                    </thead>
                    <tbody id="medTableBody"></tbody>
                    <tfoot>
                        <tr style="background: #f1f5f9; font-weight: 800;">
                            <td colspan="7" id="medFootDesc">TOTAL DA MEDIÇÃO</td>
                            <td id="medFootLib" style="color: #10b981;">-</td>
                            <td id="medFootTotal" style="text-align: right; color: #0f172a; font-size: 14px;">-</td>
                        </tr>
                    </tfoot>
                </table>
            </div>

            <div class="no-print" style="margin-top: 14px; display: flex; justify-content: flex-end; gap: 8px;">
                <button class="btn-action" onclick="copiarMedicaoEmail()">📋 Copiar Medição p/ E-mail</button>
                <button class="btn-action" onclick="copiarMedicaoExcel()">📊 Copiar Medição p/ Planilha</button>
            </div>
        </div>
    </div>

    <!-- ============================================================== -->
    <!-- ABA 4: LISTA COMPLETA DE CHAMADOS                               -->
    <!-- ============================================================== -->
    <div class="tab-view" id="view-chamados" style="display: none;">
        <div class="no-print" style="margin-bottom: 14px;">
            <button class="btn-back-nav" onclick="switchTab('inicial')">⬅️ Retornar ao Painel Inicial</button>
        </div>

        <div class="card" id="painelChamados">
            <div class="card-title-bar">
                <div style="display: flex; align-items: center; gap: 14px; flex-wrap: wrap;">
                    <div class="card-title" style="margin-bottom: 0;">📋 Lista Completa de Chamados</div>
                    
                    <div class="no-print" id="listCasaFilters" style="display: inline-flex; align-items: center; gap: 6px; flex-wrap: wrap;">
                        <button type="button" class="list-casa-btn active" id="btnCasaAll" onclick="setListCasa('')">Todas as Entidades</button>
                        <button type="button" class="list-casa-btn" id="btnCasaSesi" onclick="setListCasa('SESI')">🔵 SESI</button>
                        <button type="button" class="list-casa-btn" id="btnCasaSenai" onclick="setListCasa('SENAI')">🟠 SENAI</button>
                        <button type="button" class="list-casa-btn" id="btnCasaFibra" onclick="setListCasa('FIBRA')">🟣 FIBRA</button>
                    </div>
                </div>

                <div style="display: flex; gap: 8px; align-items: center; flex-wrap: wrap;" class="no-print">
                    <button class="btn-action" onclick="copiarChamadosEmail()">📋 Copiar p/ E-mail</button>
                    <button class="btn-action" onclick="copiarChamadosExcel()">📊 Copiar p/ Planilha</button>
                    <button class="btn-action" onclick="imprimirPainel('painelChamados', 'Lista Completa de Chamados')">🖨️ Imprimir Lista</button>
                </div>
            </div>

            <!-- Botões Rápidos de Status -->
            <div class="no-print" style="margin-bottom: 14px;">
                <div style="font-size: 12px; font-weight: 700; color: #64748b; margin-bottom: 6px;">Filtrar por Status na Tabela:</div>
                <div class="filter-group" id="listStatusFilter"></div>
            </div>

            <div class="no-print" style="margin-bottom: 16px; display: flex; gap: 14px; align-items: flex-end; flex-wrap: wrap;">
                <div style="flex: 1; min-width: 200px;">
                    <label style="display: block; margin-bottom: 4px; font-size: 12px; font-weight: 700; color: #64748b;">🔍 Buscar por O.S.</label>
                    <input type="text" class="filter-input" placeholder="Ex: 194876" oninput="filterByOS(this.value)" />
                </div>
                <div style="flex: 1.5; min-width: 260px;">
                    <label style="display: block; margin-bottom: 4px; font-size: 12px; font-weight: 700; color: #64748b;">🏢 Filtrar por Unidade</label>
                    <select class="filter-input" id="unitSelect" onchange="filterByListUnit(this.value)">
                        <option value="">Carregando unidades...</option>
                    </select>
                </div>
                <div style="font-size: 13px; color: #64748b; padding-bottom: 8px;">
                    Exibindo <strong id="ticketCount">0</strong> de <strong id="totalTickets">0</strong> chamados
                </div>
            </div>

            <div class="table-wrapper">
                <table>
                    <thead>
                        <tr>
                            <th>O.S</th>
                            <th>NR</th>
                            <th>Unidade</th>
                            <th>Descrição Resumida</th>
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

            <div class="no-print" style="margin-top: 14px; display: flex; justify-content: flex-end; gap: 8px;">
                <button class="btn-action" onclick="copiarChamadosEmail()">📋 Copiar Lista p/ E-mail</button>
                <button class="btn-action" onclick="copiarChamadosExcel()">📊 Copiar Lista p/ Planilha</button>
            </div>
        </div>
    </div>

</div>

<!-- Contêiner Isolado de Impressão -->
<div id="printSection"></div>

<!-- Toast -->
<div class="toast-msg" id="toastMessage"></div>

<script>
    const DATA = {DATA_JSON};

    let state = {{
        tickets: DATA.chamados || [],
        contratos: DATA.contratos || {{
            SESI: {{ teto: 1440000.0 }},
            SENAI: {{ teto: 1440000.0 }},
            FIBRA: {{ teto: 80000.0 }}
        }},
        filters: {{
            os: '',
            listUnit: '',
            listStatus: '',
            listCasa: ''
        }},
        payFilters: {{
            casa: '',
            mes: ''
        }},
        medFilters: {{
            casa: 'SESI',
            periodo: 'ANUAL'
        }},
        futurasAberta: false
    }};

    const fmt = v => new Intl.NumberFormat('pt-BR', {{ minimumFractionDigits: 2, maximumFractionDigits: 2 }}).format(v || 0);

    function showToast(msg) {{
        const el = document.getElementById('toastMessage');
        if (!el) return;
        el.textContent = msg;
        el.style.display = 'block';
        setTimeout(() => {{ el.style.display = 'none'; }}, 3200);
    }}

    function switchTab(tabId) {{
        document.querySelectorAll('.tab-view').forEach(v => v.style.display = 'none');
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));

        const target = document.getElementById('view-' + tabId);
        const btn = document.getElementById('tabBtn-' + tabId);
        if (target) target.style.display = 'block';
        if (btn) btn.classList.add('active');
        window.scrollTo({{ top: 0, behavior: 'smooth' }});
    }}

    // =========================================================================
    // RENDERIZADORES DO PAINEL INICIAL
    // =========================================================================
    function renderMainKPIs() {{
        const all = state.tickets;
        const total = all.length;
        const concluidos = all.filter(t => t.status === 'CONCLUIDO').length;
        const execucao = all.filter(t => t.status.includes('EXECUÇÃO') || t.status.includes('EXECUCAO')).length;
        const paralisados = all.filter(t => t.status === 'PARALISADO').length;
        const criticos = all.filter(t => t.dias_abertos !== null && t.dias_abertos > 30).length;
        const totalInvestido = all.reduce((acc, t) => acc + (t.valor || 0), 0);

        document.getElementById('mainKPIs').innerHTML = `
            <div class="kpi-card blue">
                <div class="kpi-label">Total de Chamados</div>
                <div class="kpi-val blue">${{total}}</div>
                <div class="kpi-sub">Cadastrados na base</div>
            </div>
            <div class="kpi-card green">
                <div class="kpi-label">Concluídos</div>
                <div class="kpi-val green">${{concluidos}}</div>
                <div class="kpi-sub">${{((concluidos/total)*100 || 0).toFixed(1)}}% do total</div>
            </div>
            <div class="kpi-card amber">
                <div class="kpi-label">Em Execução</div>
                <div class="kpi-val">${{execucao}}</div>
                <div class="kpi-sub">${{((execucao/total)*100 || 0).toFixed(1)}}% do total</div>
            </div>
            <div class="kpi-card red">
                <div class="kpi-label">Paralisados</div>
                <div class="kpi-val red">${{paralisados}}</div>
                <div class="kpi-sub">Aguardando definição</div>
            </div>
            <div class="kpi-card red">
                <div class="kpi-label">Críticos (>30 dias)</div>
                <div class="kpi-val red">${{criticos}}</div>
                <div class="kpi-sub">Acompanhamento prioritário</div>
            </div>
            <div class="kpi-card green">
                <div class="kpi-label">Investimento Total</div>
                <div class="kpi-val green">R$ ${{fmt(totalInvestido)}}</div>
                <div class="kpi-sub">Comprometido geral</div>
            </div>
        `;
    }}

    function renderContractCards() {{
        const casas = ['SESI', 'SENAI', 'FIBRA'];
        const colors = {{ SESI: '#0d6efd', SENAI: '#ea580c', FIBRA: '#7c3aed' }};
        
        let html = '<div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 16px;">';
        casas.forEach(c => {{
            const teto = (state.contratos[c] && state.contratos[c].teto) || (c === 'FIBRA' ? 80000.0 : 1440000.0);
            const chamadosCasa = state.tickets.filter(t => t.casa === c);
            const liberados = chamadosCasa.filter(t => t.liberado_nfe === 'LIBERADO P/ NFE');
            const totalGasto = liberados.reduce((acc, t) => acc + (t.valor || 0), 0);
            const saldo = teto - totalGasto;
            const pct = Math.min(100, Math.max(0, (totalGasto / teto) * 100));

            html += `
                <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 16px; border-top: 4px solid ${{colors[c]}};">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                        <strong style="color: ${{colors[c]}}; font-size: 14px;">CONTRATO ${{c}}</strong>
                        <span style="font-size: 12px; color: #64748b;">Teto: R$ ${{fmt(teto)}}</span>
                    </div>
                    <div style="font-size: 20px; font-weight: 800; color: #0f172a; margin-bottom: 4px;">
                        Saldo: R$ ${{fmt(saldo)}}
                    </div>
                    <div style="font-size: 12px; color: #64748b; margin-bottom: 8px;">
                        Utilizado: <strong>R$ ${{fmt(totalGasto)}}</strong> (${{pct.toFixed(1)}}%)
                    </div>
                    <div style="height: 8px; background: #e2e8f0; border-radius: 4px; overflow: hidden;">
                        <div style="height: 100%; width: ${{pct}}%; background: ${{colors[c]}};"></div>
                    </div>
                </div>
            `;
        }});
        html += '</div>';
        document.getElementById('contractCards').innerHTML = html;
    }}

    function renderTempoAberto() {{
        const pendentes = state.tickets.filter(t => t.dias_abertos !== null);
        const totalDias = pendentes.reduce((acc, t) => acc + t.dias_abertos, 0);
        const media = pendentes.length ? Math.round(totalDias / pendentes.length) : 0;

        const f1 = pendentes.filter(t => t.dias_abertos <= 30).length;
        const f2 = pendentes.filter(t => t.dias_abertos > 30 && t.dias_abertos <= 60).length;
        const f3 = pendentes.filter(t => t.dias_abertos > 60).length;

        document.getElementById('tempoAbertoContent').innerHTML = `
            <div style="text-align: center; margin-bottom: 16px;">
                <div style="font-size: 32px; font-weight: 800; color: #0f172a;">${{media}} dias</div>
                <div style="font-size: 12px; color: #64748b;">Tempo médio geral dos chamados pendentes</div>
            </div>
            <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 10px; text-align: center;">
                <div style="background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 6px; padding: 10px;">
                    <strong style="color: #15803d; font-size: 16px;">${{f1}}</strong>
                    <div style="font-size: 11px; color: #166534;">0 a 30 dias</div>
                </div>
                <div style="background: #fffbeb; border: 1px solid #fde68a; border-radius: 6px; padding: 10px;">
                    <strong style="color: #b45309; font-size: 16px;">${{f2}}</strong>
                    <div style="font-size: 11px; color: #92400e;">31 a 60 dias</div>
                </div>
                <div style="background: #fef2f2; border: 1px solid #fecaca; border-radius: 6px; padding: 10px;">
                    <strong style="color: #b91c1c; font-size: 16px;">${{f3}}</strong>
                    <div style="font-size: 11px; color: #991b1b;">> 60 dias</div>
                </div>
            </div>
        `;
    }}

    function renderStatusChart() {{
        const ctx = document.getElementById('statusChart');
        if (!ctx) return;
        
        const counts = {{}};
        state.tickets.forEach(t => {{
            counts[t.status] = (counts[t.status] || 0) + 1;
        }});

        const labels = Object.keys(counts);
        const data = Object.values(counts);

        if (window.myStatusChart) window.myStatusChart.destroy();
        window.myStatusChart = new Chart(ctx, {{
            type: 'doughnut',
            data: {{
                labels: labels,
                datasets: [{{
                    data: data,
                    backgroundColor: ['#10b981', '#0d6efd', '#f59e0b', '#ef4444', '#8b5cf6', '#64748b', '#ec4899', '#06b6d4']
                }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{
                    legend: {{ position: 'right', labels: {{ boxWidth: 12, font: {{ size: 11 }} }} }}
                }}
            }}
        }});
    }}

    function renderUnitsGrid() {{
        const unidades = [...new Set(state.tickets.map(t => t.unidade))].sort((a,b) => a.localeCompare(b, 'pt-BR'));
        const sesi = unidades.filter(u => u.includes('SESI'));
        const senai = unidades.filter(u => u.includes('SENAI'));
        const fibra = unidades.filter(u => u.includes('FIBRA') || (!u.includes('SESI') && !u.includes('SENAI')));

        function buildCol(title, list, color) {{
            let h = `<div style="background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 14px; border-top: 4px solid ${{color}};">`;
            h += `<div style="font-size: 13px; font-weight: 800; color: ${{color}}; margin-bottom: 10px;">${{title}} (${{list.length}})</div>`;
            h += `<div style="display: flex; flex-direction: column; gap: 8px; max-height: 420px; overflow-y: auto;">`;
            list.forEach(u => {{
                const chamados = state.tickets.filter(t => t.unidade === u);
                const totalVal = chamados.reduce((s, t) => s + (t.valor || 0), 0);
                h += `
                    <div style="background: #f8fafc; border: 1px solid #f1f5f9; border-radius: 6px; padding: 8px 10px; display: flex; justify-content: space-between; align-items: center;">
                        <span style="font-size: 12px; font-weight: 600; color: #1e293b;">${{u}}</span>
                        <div style="text-align: right;">
                            <span style="font-size: 11px; color: #64748b;">${{chamados.length}} chamados</span><br/>
                            <strong style="font-size: 12px; color: #0f172a;">R$ ${{fmt(totalVal)}}</strong>
                        </div>
                    </div>
                `;
            }});
            h += `</div></div>`;
            return h;
        }}

        document.getElementById('unitsGrid').innerHTML = `
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 16px;">
                ${{buildCol('🔵 UNIDADES SESI', sesi, '#0d6efd')}}
                ${{buildCol('🟠 UNIDADES SENAI', senai, '#ea580c')}}
                ${{buildCol('🟣 UNIDADES FIBRA', fibra, '#7c3aed')}}
            </div>
        `;
    }}

    // =========================================================================
    // RENDERIZADORES DO CONTROLE DE PAGAMENTOS
    // =========================================================================
    function renderPaymentFilters() {{
        const casas = ['', 'SESI', 'SENAI', 'FIBRA'];
        document.getElementById('payCasaFilter').innerHTML = casas.map(c => `
            <button class="filter-chip ${{state.payFilters.casa === c ? (c==='SENAI'?'active-senai':c==='FIBRA'?'active-fibra':'active') : ''}}" onclick="setPayCasa('${{c}}')">
                ${{c === '' ? 'Todas as Entidades (SESI + SENAI + FIBRA)' : (c==='SESI'?'🔵 SESI':c==='SENAI'?'🟠 SENAI':'🟣 FIBRA')}}
            </button>
        `).join('');

        const meses = [...new Set(state.tickets.map(t => t.mes_competencia))].filter(m => m && m !== 'NÃO DEFINIDO').sort();
        let mHtml = `<button class="filter-chip ${{state.payFilters.mes === '' ? 'active' : ''}}" onclick="setPayMes('')">Todos os Meses</button>`;
        meses.forEach(m => {{
            mHtml += `<button class="filter-chip ${{state.payFilters.mes === m ? 'active' : ''}}" onclick="setPayMes('${{m}}')">${{m}}</button>`;
        }});
        document.getElementById('payMesFilter').innerHTML = mHtml;
    }}

    function setPayCasa(c) {{ state.payFilters.casa = c; renderPaymentPanel(); }}
    function setPayMes(m) {{ state.payFilters.mes = m; renderPaymentPanel(); }}

    function renderPaymentPanel() {{
        renderPaymentFilters();
        let list = state.tickets;
        if (state.payFilters.casa) list = list.filter(t => t.casa === state.payFilters.casa);
        if (state.payFilters.mes) list = list.filter(t => t.mes_competencia === state.payFilters.mes);

        const liberados = list.filter(t => t.liberado_nfe === 'LIBERADO P/ NFE');
        const emMedicao = list.filter(t => t.status.includes('EXECUÇÃO') || t.status.includes('EXECUCAO'));
        const paralisados = list.filter(t => t.status === 'PARALISADO');

        const valLib = liberados.reduce((s, t) => s + (t.valor || 0), 0);
        const valMed = emMedicao.reduce((s, t) => s + (t.valor || 0), 0);
        const valPar = paralisados.reduce((s, t) => s + (t.valor || 0), 0);
        const valTot = list.reduce((s, t) => s + (t.valor || 0), 0);

        document.getElementById('payKPIs').innerHTML = `
            <div class="kpi-card green">
                <div class="kpi-label">Liberado para Emitir NFE</div>
                <div class="kpi-val green">R$ ${{fmt(valLib)}}</div>
                <div class="kpi-sub">${{liberados.length}} O.S. aprovadas</div>
            </div>
            <div class="kpi-card blue">
                <div class="kpi-label">Em Medição / Andamento</div>
                <div class="kpi-val blue">R$ ${{fmt(valMed)}}</div>
                <div class="kpi-sub">${{emMedicao.length}} O.S. em execução</div>
            </div>
            <div class="kpi-card red">
                <div class="kpi-label">Bloqueado / Paralisado</div>
                <div class="kpi-val red">R$ ${{fmt(valPar)}}</div>
                <div class="kpi-sub">${{paralisados.length}} O.S. retidas</div>
            </div>
            <div class="kpi-card dark">
                <div class="kpi-label">Total do Filtro</div>
                <div class="kpi-val">R$ ${{fmt(valTot)}}</div>
                <div class="kpi-sub">${{list.length}} chamados</div>
            </div>
        `;

        document.getElementById('payTableBody').innerHTML = list.map(t => `
            <tr>
                <td><strong>#${{t.os}}</strong></td>
                <td>${{t.nr}}</td>
                <td><span class="pill" style="background: ${{t.casa==='SESI'?'#e0f2fe':t.casa==='SENAI'?'#ffedd5':'#f3e8ff'}}; color: ${{t.casa==='SESI'?'#0369a1':t.casa==='SENAI'?'#c2410c':'#7c3aed'}};">${{t.casa}}</span></td>
                <td><strong>${{t.unidade}}</strong></td>
                <td>${{t.descricao || '-'}}</td>
                <td>${{t.mes_competencia}}</td>
                <td><span class="pill concluido">${{t.status}}</span></td>
                <td><span class="pill liberado-nfe">${{t.liberado_nfe}}</span></td>
                <td style="text-align: right; font-weight: 700; color: #10b981;">R$ ${{fmt(t.valor)}}</td>
            </tr>
        `).join('');

        document.getElementById('payFootDesc').textContent = `TOTAL GERAL DO FILTRO (${{list.length}} chamados)`;
        document.getElementById('payFootLib').textContent = `Liberado: R$ ${{fmt(valLib)}}`;
        document.getElementById('payFootTotal').textContent = `R$ ${{fmt(valTot)}}`;

        document.getElementById('payRibbonTxt').textContent = `Entidade: ${{state.payFilters.casa || 'Todas'}} | Mês: ${{state.payFilters.mes || 'Todos'}} | Chamados Aprovados: ${{liberados.length}}`;
        document.getElementById('payRibbonVal').textContent = `Total Liberado: R$ ${{fmt(valLib)}}`;
    }}

    // =========================================================================
    // RENDERIZADORES DO MÓDULO DE MEDIÇÕES
    // =========================================================================
    function renderMedFilters() {{
        const casas = ['SESI', 'SENAI', 'FIBRA'];
        document.getElementById('medCasaFilter').innerHTML = casas.map(c => `
            <button class="filter-chip ${{state.medFilters.casa === c ? (c==='SESI'?'active':c==='SENAI'?'active-senai':'active-fibra') : ''}}" onclick="setMedCasa('${{c}}')">
                ${{c === 'SESI' ? '🔵 SESI (CNPJ SESI)' : c === 'SENAI' ? '🟠 SENAI (CNPJ SENAI)' : '🟣 FIBRA (CNPJ FIBRA)'}}
            </button>
        `).join('');

        const meses = [...new Set(state.tickets.map(t => t.mes_competencia))].filter(m => m && m !== 'NÃO DEFINIDO').sort();
        let mHtml = meses.map(m => `
            <button class="filter-chip ${{state.medFilters.periodo === m ? 'active' : ''}}" onclick="setMedPeriodo('${{m}}')">📅 ${{m}}</button>
        `).join('');
        mHtml += `<button class="filter-chip ${{state.medFilters.periodo === 'ANUAL' ? 'active-dark' : ''}}" onclick="setMedPeriodo('ANUAL')">📊 ACUMULADO ANUAL (Todas as Medições)</button>`;
        document.getElementById('medPeriodoFilter').innerHTML = mHtml;
    }}

    function setMedCasa(c) {{ state.medFilters.casa = c; renderMedicoesPanel(); }}
    function setMedPeriodo(p) {{ state.medFilters.periodo = p; renderMedicoesPanel(); }}

    function toggleFuturasDrawer() {{
        state.futurasAberta = !state.futurasAberta;
        const d = document.getElementById('drawerFuturas');
        const b = document.getElementById('btnToggleFuturas');
        if (d) d.className = 'drawer-futuras' + (state.futurasAberta ? ' open' : '');
        if (b) b.textContent = state.futurasAberta 
            ? '✖️ Ocultar O.S. Futuras' 
            : `👁️ Inspecionar O.S. Futuras (${{document.getElementById('prevCountBadge').textContent}})`;
    }}

    function renderMedicoesPanel() {{
        renderMedFilters();
        const casa = state.medFilters.casa;
        const periodo = state.medFilters.periodo;

        const chamadosCasa = state.tickets.filter(t => t.casa === casa && t.liberado_nfe === 'LIBERADO P/ NFE');
        const totalAnualLiberado = chamadosCasa.reduce((s, t) => s + (t.valor || 0), 0);

        let chamadosPeriodo = chamadosCasa;
        if (periodo !== 'ANUAL') {{
            chamadosPeriodo = chamadosCasa.filter(t => t.mes_competencia === periodo);
        }}
        const totalMedicao = chamadosPeriodo.reduce((s, t) => s + (t.valor || 0), 0);

        const teto = (state.contratos[casa] && state.contratos[casa].teto) || (casa === 'FIBRA' ? 80000.0 : 1440000.0);
        const saldoDisponivel = teto - totalAnualLiberado;
        const pctDisponivel = Math.max(0, (saldoDisponivel / teto) * 100);

        // Previsão futura: O.S. em andamento / orçadas da casa com valor > 0
        const chamadosFuturos = state.tickets.filter(t => t.casa === casa && t.liberado_nfe !== 'LIBERADO P/ NFE' && (t.valor || 0) > 0);
        const totalPrevisto = chamadosFuturos.reduce((s, t) => s + (t.valor || 0), 0);
        const totalComprometido = totalAnualLiberado + totalPrevisto;
        const margemRealLivre = teto - totalComprometido;

        document.getElementById('medTetoBadge').textContent = `Livro Oficial de Medições: Contrato ${{casa}}: R$ ${{fmt(teto)}}`;
        document.getElementById('prevTitle').textContent = `🛡️ GESTÃO DE RISCO: PREVISÃO DE MEDIÇÕES FUTURAS & SALDO ESTIMADO ( ${{casa}} )`;
        document.getElementById('prevCountBadge').textContent = chamadosFuturos.length;

        // 4 KPIs Contratuais
        document.getElementById('medKPIs').innerHTML = `
            <div class="kpi-card green">
                <div class="kpi-label">Valor da Medição Selecionada</div>
                <div class="kpi-val green">R$ ${{fmt(totalMedicao)}}</div>
                <div class="kpi-sub">${{chamadosPeriodo.length}} O.S. liberadas nesta seleção</div>
            </div>
            <div class="kpi-card blue">
                <div class="kpi-label">Acumulado Anual Liberado</div>
                <div class="kpi-val blue">R$ ${{fmt(totalAnualLiberado)}}</div>
                <div class="kpi-sub">${{chamadosCasa.length}} O.S. faturadas no exercício (${{casa}})</div>
            </div>
            <div class="kpi-card dark">
                <div class="kpi-label">Teto Contratual Total</div>
                <div class="kpi-val">R$ ${{fmt(teto)}}</div>
                <div class="kpi-sub">Valor limite homologado</div>
            </div>
            <div class="kpi-card amber">
                <div class="kpi-label">Saldo Disponível em Contrato</div>
                <div class="kpi-val" style="color: #b45309;">R$ ${{fmt(saldoDisponivel)}}</div>
                <div class="kpi-sub">${{pctDisponivel.toFixed(1)}}% disponível no teto homologado</div>
            </div>
        `;

        // 3 KPIs de Previsão de Risco
        document.getElementById('prevKPIs').innerHTML = `
            <div class="kpi-card blue">
                <div class="kpi-label">Previsão em Investimentos (O.S. Futuras)</div>
                <div class="kpi-val blue">R$ ${{fmt(totalPrevisto)}}</div>
                <div class="kpi-sub">${{chamadosFuturos.length}} O.S. orçadas / em execução</div>
            </div>
            <div class="kpi-card amber">
                <div class="kpi-label">Total Geral Comprometido (Medido + Previsto)</div>
                <div class="kpi-val" style="color: #ea580c;">R$ ${{fmt(totalComprometido)}}</div>
                <div class="kpi-sub">${{((totalComprometido/teto)*100).toFixed(1)}}% do teto consumido</div>
            </div>
            <div class="kpi-card green">
                <div class="kpi-label">Saldo Estimado em Contrato (Margem Real Livre)</div>
                <div class="kpi-val green">R$ ${{fmt(margemRealLivre)}}</div>
                <div class="kpi-sub">Margem segura para novos planejamentos</div>
            </div>
        `;

        const pctMed = Math.min(100, (totalAnualLiberado / teto) * 100);
        const pctPrev = Math.min(100 - pctMed, (totalPrevisto / teto) * 100);
        const pctLivre = Math.max(0, 100 - pctMed - pctPrev);

        document.getElementById('prevProgressBar').innerHTML = `
            <div class="progress-bar-seg" style="width: ${{pctMed}}%; background: #0d6efd;" title="Liberado: ${{pctMed.toFixed(1)}}%"></div>
            <div class="progress-bar-seg" style="width: ${{pctPrev}}%; background: #ea580c;" title="Previsto: ${{pctPrev.toFixed(1)}}%"></div>
            <div class="progress-bar-seg" style="width: ${{pctLivre}}%; background: #10b981;" title="Margem Livre: ${{pctLivre.toFixed(1)}}%"></div>
        `;

        const badgeEl = document.getElementById('prevRiskBadge');
        if (margemRealLivre < 0) {{
            badgeEl.textContent = '⛔ ESTOURO CONTRATUAL';
            badgeEl.style.background = '#fee2e2'; badgeEl.style.color = '#b91c1c';
        }} else if ((margemRealLivre / teto) < 0.15) {{
            badgeEl.textContent = '🟡 ATENÇÃO: MARGEM CRÍTICA';
            badgeEl.style.background = '#fef3c7'; badgeEl.style.color = '#b45309';
        }} else {{
            badgeEl.textContent = '🟢 SALDO ESTIMADO SEGURO';
            badgeEl.style.background = '#dcfce7'; badgeEl.style.color = '#15803d';
        }}

        // Renderiza gaveta de O.S. Futuras
        document.getElementById('futurasTableBody').innerHTML = chamadosFuturos.map(t => `
            <tr>
                <td><strong>#${{t.os}}</strong></td>
                <td>${{t.nr}}</td>
                <td><strong>${{t.unidade}}</strong></td>
                <td>${{t.descricao || '-'}}</td>
                <td><span class="pill execucao">${{t.status}}</span></td>
                <td style="text-align: right; font-weight: 700; color: #ea580c;">R$ ${{fmt(t.valor)}}</td>
            </tr>
        `).join('');
        document.getElementById('futurasFootDesc').textContent = `TOTAL DE PREVISÕES FUTURAS (${{chamadosFuturos.length}} O.S.)`;
        document.getElementById('futurasFootVal').textContent = `R$ ${{fmt(totalPrevisto)}}`;

        // Renderiza Tabela Oficial de Medições Liberadas
        document.getElementById('medTableBody').innerHTML = chamadosPeriodo.map(t => `
            <tr>
                <td><strong>#${{t.os}}</strong></td>
                <td>${{t.nr}}</td>
                <td><span class="pill" style="background: ${{t.casa==='SESI'?'#e0f2fe':t.casa==='SENAI'?'#ffedd5':'#f3e8ff'}}; color: ${{t.casa==='SESI'?'#0369a1':t.casa==='SENAI'?'#c2410c':'#7c3aed'}};">${{t.casa}}</span></td>
                <td><strong>${{t.unidade}}</strong></td>
                <td>${{t.descricao || '-'}}</td>
                <td>${{t.mes_competencia}}</td>
                <td><span class="pill concluido">${{t.status}}</span></td>
                <td><span class="pill liberado-nfe">✅ LIBERADO P/ NFE</span></td>
                <td style="text-align: right; font-weight: 700; color: #10b981;">R$ ${{fmt(t.valor)}}</td>
            </tr>
        `).join('');

        document.getElementById('medFootDesc').textContent = `TOTAL GERAL: ${{casa}} | ${{periodo === 'ANUAL' ? 'EXERCÍCIO ANUAL COMPLETO' : periodo}}`;
        document.getElementById('medFootLib').textContent = `Aprovado: R$ ${{fmt(totalMedicao)}}`;
        document.getElementById('medFootTotal').textContent = `R$ ${{fmt(totalMedicao)}}`;
    }}

    // =========================================================================
    // RENDERIZADORES DA LISTA COMPLETA DE CHAMADOS
    // =========================================================================
    function getFilteredChamados() {{
        return state.tickets.filter(t => {{
            if (state.filters.os && !t.os.toString().includes(state.filters.os)) return false;
            if (state.filters.listCasa && t.casa !== state.filters.listCasa) return false;
            if (state.filters.listStatus && t.status !== state.filters.listStatus) return false;
            if (state.filters.listUnit && t.unidade !== state.filters.listUnit) return false;
            return true;
        }});
    }}

    function setListCasa(casa) {{
        state.filters.listCasa = casa;
        if (casa && state.filters.listUnit) {{
            const exists = state.tickets.some(t => t.unidade === state.filters.listUnit && t.casa === casa);
            if (!exists) state.filters.listUnit = '';
        }}
        renderListCasaButtons();
        renderListUnitFilter();
        renderTableChamados();
    }}

    function renderListCasaButtons() {{
        const c = state.filters.listCasa;
        const bAll = document.getElementById('btnCasaAll');
        const bSesi = document.getElementById('btnCasaSesi');
        const bSenai = document.getElementById('btnCasaSenai');
        const bFibra = document.getElementById('btnCasaFibra');

        if (bAll) bAll.className = 'list-casa-btn' + (c === '' ? ' active' : '');
        if (bSesi) bSesi.className = 'list-casa-btn' + (c === 'SESI' ? ' active-sesi' : '');
        if (bSenai) bSenai.className = 'list-casa-btn' + (c === 'SENAI' ? ' active-senai' : '');
        if (bFibra) bFibra.className = 'list-casa-btn' + (c === 'FIBRA' ? ' active-fibra' : '');
    }}

    function renderListStatusFilter() {{
        const statuses = ['Todos', ...new Set(state.tickets.map(t => t.status))].sort();
        document.getElementById('listStatusFilter').innerHTML = statuses.map(s => `
            <button class="filter-chip ${{state.filters.listStatus === (s === 'Todos' ? '' : s) ? 'active' : ''}}" onclick="setListStatus('${{s === 'Todos' ? '' : s}}')">
                ${{s}}
            </button>
        `).join('');
    }}

    function setListStatus(s) {{
        state.filters.listStatus = s;
        renderListStatusFilter();
        renderTableChamados();
    }}

    function renderListUnitFilter() {{
        const select = document.getElementById('unitSelect');
        if (!select) return;

        let base = state.tickets;
        if (state.filters.listCasa) base = base.filter(t => t.casa === state.filters.listCasa);

        const unidades = [...new Set(base.map(t => t.unidade))].sort((a,b) => a.localeCompare(b, 'pt-BR'));
        let h = `<option value="">🏢 Todas as Unidades (${{unidades.length}})</option>`;
        unidades.forEach(u => {{
            const sel = state.filters.listUnit === u ? 'selected' : '';
            h += `<option value="${{u}}" ${{sel}}>${{u}}</option>`;
        }});
        select.innerHTML = h;
    }}

    function filterByOS(v) {{
        state.filters.os = v.trim();
        renderTableChamados();
    }}

    function filterByListUnit(v) {{
        state.filters.listUnit = v;
        renderTableChamados();
    }}

    function renderTableChamados() {{
        const list = getFilteredChamados();
        document.getElementById('ticketCount').textContent = list.length;
        document.getElementById('totalTickets').textContent = state.filters.listCasa 
            ? state.tickets.filter(t => t.casa === state.filters.listCasa).length 
            : state.tickets.length;

        document.getElementById('tableBody').innerHTML = list.map(t => `
            <tr>
                <td><strong>#${{t.os}}</strong></td>
                <td>${{t.nr}}</td>
                <td><strong>${{t.unidade}}</strong></td>
                <td>${{t.descricao || '-'}}</td>
                <td><span class="pill ${{t.status === 'CONCLUIDO' ? 'concluido' : t.status === 'PARALISADO' ? 'paralisado' : 'execucao'}}">${{t.status}}</span></td>
                <td>${{t.data_envio}}</td>
                <td>${{t.entrega_orcamento}}</td>
                <td>
                    ${{t.dias_abertos !== null 
                        ? `<span style="font-weight: 700; color: ${{t.dias_abertos > 30 ? '#ef4444' : '#10b981'}};">${{t.dias_abertos}}d</span>` 
                        : '-'}}
                </td>
                <td style="text-align: right; font-weight: 700; color: #10b981;">R$ ${{fmt(t.valor)}}</td>
            </tr>
        `).join('');
    }}

    // =========================================================================
    // FUNÇÕES DE EXPORTAÇÃO E CÓPIA
    // =========================================================================
    function copiarChamadosEmail() {{
        const list = getFilteredChamados();
        if (!list.length) return showToast('⚠️ Nenhum chamado para copiar.');

        let html = `
            <div style="font-family: Arial, sans-serif; color: #1e293b;">
                <h3 style="color: #0d6efd; margin-bottom: 6px;">📋 LISTA DE CHAMADOS DE MANUTENÇÃO PREDIAL</h3>
                <p style="font-size: 12px; color: #64748b; margin-bottom: 12px;">Total: ${{list.length}} chamados</p>
                <table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; font-size: 11px; width: 100%;">
                    <tr style="background: #0d6efd; color: #ffffff;">
                        <th>O.S</th><th>NR</th><th>Unidade</th><th>Descrição Resumida</th><th>Status</th><th>Data Envio</th><th>Dias</th><th>Valor</th>
                    </tr>
                    ${{list.map(t => `
                        <tr>
                            <td>#${{t.os}}</td><td>${{t.nr}}</td><td>${{t.unidade}}</td><td>${{t.descricao || '-'}}</td>
                            <td>${{t.status}}</td><td>${{t.data_envio}}</td><td>${{t.dias_abertos || '-'}}</td>
                            <td style="text-align: right;">R$ ${{fmt(t.valor)}}</td>
                        </tr>
                    `).join('')}}
                </table>
            </div>
        `;
        navigator.clipboard.write([
            new ClipboardItem({{ 'text/html': new Blob([html], {{ type: 'text/html' }}), 'text/plain': new Blob([html], {{ type: 'text/plain' }}) }})
        ]).then(() => showToast('✅ Lista copiada para colar no E-mail!'));
    }}

    function copiarChamadosExcel() {{
        const list = getFilteredChamados();
        if (!list.length) return showToast('⚠️ Nenhum chamado para copiar.');
        let tsv = "O.S\\tNR\\tUnidade\\tDescrição Resumida\\tStatus\\tData Envio\\tEntrega Orçamento\\tDias Abertos\\tValor\\n";
        list.forEach(t => {{
            tsv += `${{t.os}}\\t${{t.nr}}\\t${{t.unidade}}\\t${{t.descricao || ''}}\\t${{t.status}}\\t${{t.data_envio}}\\t${{t.entrega_orcamento}}\\t${{t.dias_abertos || ''}}\\tR$ ${{fmt(t.valor)}}\\n`;
        }});
        navigator.clipboard.writeText(tsv).then(() => showToast('✅ Lista copiada para o Excel!'));
    }}

    function copiarTabelaEmail() {{
        const list = state.tickets.filter(t => t.liberado_nfe === 'LIBERADO P/ NFE');
        let html = `
            <table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; font-size: 11px; width: 100%; font-family: Arial;">
                <tr style="background: #0d6efd; color: #fff;"><th>O.S</th><th>NR</th><th>Entidade</th><th>Unidade</th><th>Descrição</th><th>Mês</th><th>Valor</th></tr>
                ${{list.map(t => `<tr><td>#${{t.os}}</td><td>${{t.nr}}</td><td>${{t.casa}}</td><td>${{t.unidade}}</td><td>${{t.descricao}}</td><td>${{t.mes_competencia}}</td><td>R$ ${{fmt(t.valor)}}</td></tr>`).join('')}}
            </table>
        `;
        navigator.clipboard.write([
            new ClipboardItem({{ 'text/html': new Blob([html], {{ type: 'text/html' }}) }})
        ]).then(() => showToast('✅ Tabela copiada para o E-mail!'));
    }}

    function copiarTabelaExcel() {{
        let tsv = "O.S\\tNR\\tEntidade\\tUnidade\\tDescrição\\tMês\\tValor\\n";
        state.tickets.filter(t => t.liberado_nfe === 'LIBERADO P/ NFE').forEach(t => {{
            tsv += `${{t.os}}\\t${{t.nr}}\\t${{t.casa}}\\t${{t.unidade}}\\t${{t.descricao}}\\t${{t.mes_competencia}}\\tR$ ${{fmt(t.valor)}}\\n`;
        }});
        navigator.clipboard.writeText(tsv).then(() => showToast('✅ Tabela copiada para o Excel!'));
    }}

    function copiarMedicaoEmail() {{
        const casa = state.medFilters.casa;
        const periodo = state.medFilters.periodo;
        let list = state.tickets.filter(t => t.casa === casa && t.liberado_nfe === 'LIBERADO P/ NFE');
        if (periodo !== 'ANUAL') list = list.filter(t => t.mes_competencia === periodo);
        const total = list.reduce((s, t) => s + (t.valor || 0), 0);

        let html = `
            <div style="font-family: Arial; color: #1e293b;">
                <h3 style="color: #0284c7; margin-bottom: 4px;">📐 BOLETIM DE MEDIÇÃO — ${{casa}} (${{periodo}})</h3>
                <p style="font-size: 12px; margin-bottom: 10px;">Total Liberado: <strong>R$ ${{fmt(total)}}</strong> (${{list.length}} O.S.)</p>
                <table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; font-size: 11px; width: 100%;">
                    <tr style="background: #0284c7; color: #fff;"><th>O.S</th><th>NR</th><th>Unidade</th><th>Descrição</th><th>Mês</th><th>Valor</th></tr>
                    ${{list.map(t => `<tr><td>#${{t.os}}</td><td>${{t.nr}}</td><td>${{t.unidade}}</td><td>${{t.descricao}}</td><td>${{t.mes_competencia}}</td><td>R$ ${{fmt(t.valor)}}</td></tr>`).join('')}}
                </table>
            </div>
        `;
        navigator.clipboard.write([
            new ClipboardItem({{ 'text/html': new Blob([html], {{ type: 'text/html' }}) }})
        ]).then(() => showToast('✅ Medição copiada para o E-mail!'));
    }}

    function copiarMedicaoExcel() {{
        const casa = state.medFilters.casa;
        const periodo = state.medFilters.periodo;
        let list = state.tickets.filter(t => t.casa === casa && t.liberado_nfe === 'LIBERADO P/ NFE');
        if (periodo !== 'ANUAL') list = list.filter(t => t.mes_competencia === periodo);

        let tsv = "O.S\\tNR\\tUnidade\\tDescrição\\tMês Competência\\tStatus\\tValor\\n";
        list.forEach(t => {{
            tsv += `${{t.os}}\\t${{t.nr}}\\t${{t.unidade}}\\t${{t.descricao}}\\t${{t.mes_competencia}}\\tLIBERADO P/ NFE\\tR$ ${{fmt(t.valor)}}\\n`;
        }});
        navigator.clipboard.writeText(tsv).then(() => showToast('✅ Medição copiada para o Excel!'));
    }}

    function copiarFuturasEmail() {{
        const casa = state.medFilters.casa;
        const chamadosFuturos = state.tickets.filter(t => t.casa === casa && t.liberado_nfe !== 'LIBERADO P/ NFE' && (t.valor || 0) > 0);
        const total = chamadosFuturos.reduce((s, t) => s + (t.valor || 0), 0);

        let html = `
            <div style="font-family: Arial; color: #1e293b;">
                <h3 style="color: #ea580c; margin-bottom: 4px;">🛡️ PREVISÃO DE O.S. FUTURAS — ${{casa}}</h3>
                <p style="font-size: 12px; margin-bottom: 10px;">Total Previsionado: <strong>R$ ${{fmt(total)}}</strong> (${{chamadosFuturos.length}} O.S.)</p>
                <table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; font-size: 11px; width: 100%;">
                    <tr style="background: #ea580c; color: #fff;"><th>O.S</th><th>NR</th><th>Unidade</th><th>Descrição</th><th>Status</th><th>Valor Previsionado</th></tr>
                    ${{chamadosFuturos.map(t => `<tr><td>#${{t.os}}</td><td>${{t.nr}}</td><td>${{t.unidade}}</td><td>${{t.descricao}}</td><td>${{t.status}}</td><td>R$ ${{fmt(t.valor)}}</td></tr>`).join('')}}
                </table>
            </div>
        `;
        navigator.clipboard.write([
            new ClipboardItem({{ 'text/html': new Blob([html], {{ type: 'text/html' }}) }})
        ]).then(() => showToast('✅ O.S. Futuras copiadas para o E-mail!'));
    }}

    function copiarFuturasExcel() {{
        const casa = state.medFilters.casa;
        const chamadosFuturos = state.tickets.filter(t => t.casa === casa && t.liberado_nfe !== 'LIBERADO P/ NFE' && (t.valor || 0) > 0);
        let tsv = "O.S\\tNR\\tUnidade\\tDescrição\\tStatus\\tValor Previsionado\\n";
        chamadosFuturos.forEach(t => {{
            tsv += `${{t.os}}\\t${{t.nr}}\\t${{t.unidade}}\\t${{t.descricao}}\\t${{t.status}}\\tR$ ${{fmt(t.valor)}}\\n`;
        }});
        navigator.clipboard.writeText(tsv).then(() => showToast('✅ O.S. Futuras copiadas para o Excel!'));
    }}

    // =========================================================================
    // IMPRESSÃO ISOLADA (SEM PÁGINA EM BRANCO)
    // =========================================================================
    function imprimirPainel(elementId, titulo) {{
        const printSec = document.getElementById('printSection');
        const orig = document.getElementById(elementId);
        if (!printSec || !orig) return;

        const clone = orig.cloneNode(true);
        clone.querySelectorAll('.no-print').forEach(el => el.remove());

        printSec.innerHTML = `
            <div style="margin-bottom: 14px; border-bottom: 2px solid #0f172a; padding-bottom: 8px;">
                <h2 style="font-size: 18px; color: #0f172a; margin: 0;">SESI / SENAI / FIBRA — Manutenção Predial</h2>
                <div style="font-size: 12px; color: #475569;">Relatório Oficial: ${{titulo}} • Emitido em: ${{new Date().toLocaleDateString('pt-BR')}}</div>
            </div>
        `;
        printSec.appendChild(clone);
        window.print();
        setTimeout(() => {{ printSec.innerHTML = ''; }}, 1000);
    }}

    function imprimirFuturas() {{
        const printSec = document.getElementById('printSection');
        const orig = document.getElementById('previsaoBox');
        if (!printSec || !orig) return;

        const clone = orig.cloneNode(true);
        clone.querySelectorAll('.no-print').forEach(el => el.remove());
        const drawer = clone.querySelector('#drawerFuturas');
        if (drawer) drawer.style.display = 'block';

        printSec.innerHTML = `
            <div style="margin-bottom: 14px; border-bottom: 2px solid #0f172a; padding-bottom: 8px;">
                <h2 style="font-size: 18px; color: #0f172a; margin: 0;">SESI / SENAI / FIBRA — Relatório de Previsão de Medições Futuras</h2>
                <div style="font-size: 12px; color: #475569;">Emitido em: ${{new Date().toLocaleDateString('pt-BR')}}</div>
            </div>
        `;
        printSec.appendChild(clone);
        window.print();
        setTimeout(() => {{ printSec.innerHTML = ''; }}, 1000);
    }}

    function imprimirRelatorioCompleto() {{
        const printSec = document.getElementById('printSection');
        if (!printSec) return;
        const clone = document.getElementById('mainContainer').cloneNode(true);
        clone.querySelectorAll('.no-print, .nav-tabs, .shortcut-grid').forEach(el => el.remove());
        clone.querySelectorAll('.tab-view').forEach(v => v.style.display = 'block');

        printSec.innerHTML = `
            <div style="margin-bottom: 14px; border-bottom: 2px solid #0f172a; padding-bottom: 8px;">
                <h2 style="font-size: 18px; color: #0f172a; margin: 0;">SESI / SENAI / FIBRA — Relatório Gerencial Completo</h2>
                <div style="font-size: 12px; color: #475569;">Emitido em: ${{new Date().toLocaleDateString('pt-BR')}}</div>
            </div>
        `;
        printSec.appendChild(clone);
        window.print();
        setTimeout(() => {{ printSec.innerHTML = ''; }}, 1000);
    }}

    // =========================================================================
    // INICIALIZAÇÃO
    // =========================================================================
    function init() {{
        renderMainKPIs();
        renderContractCards();
        renderTempoAberto();
        renderStatusChart();
        renderUnitsGrid();
        renderPaymentPanel();
        renderMedicoesPanel();
        renderListCasaButtons();
        renderListStatusFilter();
        renderListUnitFilter();
        renderTableChamados();
    }}

    init();
</script>
</body>
</html>"""

# Renderização do componente Streamlit com rolagem e altura otimizada
components.html(HTML_TEMPLATE, height=1400, scrolling=True)

# Exportação do dashboard em HTML
st.sidebar.markdown("### 💾 Exportar Dashboard")
st.sidebar.download_button(
    label="Baixar Painel em HTML",
    data=HTML_TEMPLATE,
    file_name=f"painel_manutencao_completo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html",
    mime="text/html"
)
