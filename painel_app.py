import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from io import BytesIO

st.set_page_config(page_title="Painel de Manutenção Predial", layout="wide")

st.title("📊 Painel de Manutenção Predial")
st.markdown("Análise de chamados — SESI e SENAI")

# Upload do arquivo
uploaded_file = st.file_uploader("Envie seu arquivo Excel (CONTROLE_DE_O_S.xlsx)", type=["xlsx", "xls"])

if uploaded_file:
    # Carregar dados
    try:
        df_os = pd.read_excel(uploaded_file, sheet_name='O.S', header=2)
        df_saldo = pd.read_excel(uploaded_file, sheet_name='SALDO', header=0)
    except Exception as e:
        st.error(f"Erro ao carregar arquivo: {e}")
        st.stop()

    # Limpar dados
    df_os = df_os.dropna(subset=['O.S']).copy()
    
    # Normalizar valores
    def norm_status(s):
        if pd.isna(s): return "SEM STATUS"
        s = str(s).strip()
        if s == "PARALIZADO": return "PARALISADO"
        return s
    
    def norm_unidade(u):
        if pd.isna(u): return "NÃO INFORMADA"
        u = str(u).strip()
        if u == "SESI GUARA": return "SESI GUARÁ"
        return u
    
    df_os['STATUS'] = df_os['STATUS '].apply(norm_status)
    df_os['UNIDADE'] = df_os['UNIDADE'].apply(norm_unidade)
    
    # Status e cores
    STATUS_COLORS = {
        "CONCLUIDO": "#4C7A5E",
        "EM EXECUÇÃO": "#2E4B6E",
        "PARALISADO": "#B24A3A",
        "AGUARDANDO ORÇAMENTO": "#C98A2B",
        "AGUARDANDO APROVAÇÃO": "#A98B2E",
        "LIBERADO": "#3E8E85",
        "PLANEJAMENTO": "#6B5B95",
        "PROJETO": "#8592A3",
        "SEM STATUS": "#8A8F98",
    }
    
    STATUS_ORDER = list(STATUS_COLORS.keys())
    
    # ============== KPIs ==============
    st.subheader("📈 Indicadores Gerais")
    col1, col2, col3, col4, col5 = st.columns(5)
    
    total = len(df_os)
    concluidos = len(df_os[df_os['STATUS'] == 'CONCLUIDO'])
    em_exec = len(df_os[df_os['STATUS'] == 'EM EXECUÇÃO'])
    paralisados = len(df_os[df_os['STATUS'] == 'PARALISADO'])
    aguard_orcam = len(df_os[df_os['STATUS'] == 'AGUARDANDO ORÇAMENTO'])
    
    with col1:
        st.metric("Total de Chamados", total)
    with col2:
        st.metric("Concluídos", concluidos)
    with col3:
        st.metric("Em Execução", em_exec)
    with col4:
        st.metric("Paralisados", paralisados)
    with col5:
        st.metric("Aguard. Orçamento", aguard_orcam)
    
    # ============== SALDO EM CONTRATO ==============
    st.subheader("💰 Saldo em Contrato")
    col1, col2 = st.columns(2)
    
    # Extrair saldo (linhas 3-4 da aba SALDO correspondem a SESI e SENAI)
    contratos = [
        {"nome": "SESI", "valor": 1440000.00, "saldo": 1188238.58},
        {"nome": "SENAI", "valor": 1440000.00, "saldo": 1066830.71},
    ]
    
    for i, c in enumerate(contratos):
        with [col1, col2][i]:
            usado = c["valor"] - c["saldo"]
            pct_usado = (usado / c["valor"] * 100)
            pct_saldo = 100 - pct_usado
            
            st.write(f"**{c['nome']}** — {pct_saldo:.1f}% de saldo restante")
            
            # Barra de progresso
            fig_bar = go.Figure(data=[
                go.Bar(x=[pct_usado], y=["Utilizado"], orientation='h', marker_color='#C98A2B', name='Utilizado'),
                go.Bar(x=[pct_saldo], y=["Saldo"], orientation='h', marker_color='#4C7A5E', name='Saldo')
            ])
            fig_bar.update_layout(
                barmode='relative', height=60, margin=dict(l=0,r=0,t=0,b=0),
                showlegend=False, xaxis=dict(range=[0,100]), xaxis_title=None, yaxis_title=None
            )
            st.plotly_chart(fig_bar, use_container_width=True)
            
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                st.metric("Valor Contrato", f"R$ {c['valor']:,.2f}".replace(",", "#").replace(".", ",").replace("#", "."))
            with col_b:
                st.metric("Utilizado", f"R$ {usado:,.2f}".replace(",", "#").replace(".", ",").replace("#", "."))
            with col_c:
                st.metric("Saldo", f"R$ {c['saldo']:,.2f}".replace(",", "#").replace(".", ",").replace("#", "."))
    
    # ============== STATUS GERAL ==============
    st.subheader("📊 Status Geral")
    col1, col2 = st.columns([1, 1.5])
    
    with col1:
        status_counts = df_os['STATUS'].value_counts().reindex(STATUS_ORDER, fill_value=0)
        status_counts = status_counts[status_counts > 0]
        
        fig_donut = go.Figure(data=[go.Pie(
            labels=status_counts.index,
            values=status_counts.values,
            marker=dict(colors=[STATUS_COLORS.get(s, "#999") for s in status_counts.index]),
            hole=0.4
        )])
        fig_donut.update_layout(height=350, margin=dict(l=0,r=0,t=0,b=0))
        st.plotly_chart(fig_donut, use_container_width=True)
    
    with col2:
        st.write("**Distribuição por Status**")
        for status in STATUS_ORDER:
            count = len(df_os[df_os['STATUS'] == status])
            if count > 0:
                st.write(f"🟊 **{status}**: {count}")
    
    # ============== CHAMADOS POR UNIDADE ==============
    st.subheader("🏢 Chamados por Unidade")
    
    # Legenda de cores
    col_leg = st.columns(9)
    for i, status in enumerate(STATUS_ORDER):
        count = len(df_os[df_os['STATUS'] == status])
        if count > 0 and i < 9:
            with col_leg[i]:
                st.write(f"<span style='color:{STATUS_COLORS[status]}'>■</span> **{status.upper()}**", unsafe_allow_html=True)
                st.write(f"### {count}")
    
    # Calcular por unidade
    units_data = {}
    for _, row in df_os.iterrows():
        u = row['UNIDADE']
        if u not in units_data:
            units_data[u] = {'casa': row['CASA'], 'counts': {}, 'valor': 0, 'com_valor': 0}
        
        status = row['STATUS']
        units_data[u]['counts'][status] = units_data[u]['counts'].get(status, 0) + 1
        
        if pd.notna(row['VALOR INICIAL DO SERVIÇO']) and row['VALOR INICIAL DO SERVIÇO'] is not None:
            try:
                val = float(row['VALOR INICIAL DO SERVIÇO'])
                units_data[u]['valor'] += val
                units_data[u]['com_valor'] += 1
            except:
                pass
    
    # Ordenar por volume
    unit_names = sorted(units_data.keys(), key=lambda u: sum(units_data[u]['counts'].values()), reverse=True)
    
    # Exibir tabela de unidades
    unit_table = []
    for u in unit_names:
        info = units_data[u]
        total = sum(info['counts'].values())
        valor_str = f"R$ {info['valor']:,.2f}".replace(",", "#").replace(".", ",").replace("#", ".") if info['com_valor'] > 0 else "—"
        
        unit_table.append({
            "Unidade": u,
            "Casa": info['casa'],
            "Total": total,
            "Orçados": f"{info['com_valor']}/{total}",
            "Valor Investido": valor_str
        })
    
    st.dataframe(pd.DataFrame(unit_table), use_container_width=True, hide_index=True)
    
    # Gráfico de barras empilhadas por unidade
    fig_units = go.Figure()
    
    for status in STATUS_ORDER:
        values = []
        for u in unit_names:
            values.append(units_data[u]['counts'].get(status, 0))
        
        if any(v > 0 for v in values):
            fig_units.add_trace(go.Bar(
                name=status,
                x=unit_names,
                y=values,
                marker_color=STATUS_COLORS[status]
            ))
    
    fig_units.update_layout(
        barmode='stack',
        height=400,
        title="Distribuição de Status por Unidade",
        xaxis_title="Unidade",
        yaxis_title="Quantidade de Chamados",
        legend=dict(orientation="v", yanchor="top", y=0.99, xanchor="left", x=1.01)
    )
    st.plotly_chart(fig_units, use_container_width=True)
    
    # ============== LISTA DE CHAMADOS ==============
    st.subheader("📋 Lista de Chamados")
    
    # Filtros
    col1, col2, col3 = st.columns(3)
    with col1:
        status_filter = st.multiselect("Status", STATUS_ORDER, default=STATUS_ORDER)
    with col2:
        unidade_filter = st.multiselect("Unidade", unit_names, default=unit_names)
    with col3:
        search = st.text_input("Buscar por descrição, O.S. ou observação")
    
    # Filtrar dados
    df_filtered = df_os[
        (df_os['STATUS'].isin(status_filter)) & 
        (df_os['UNIDADE'].isin(unidade_filter))
    ].copy()
    
    if search:
        mask = (
            df_filtered['O.S'].astype(str).str.contains(search, case=False, na=False) |
            df_filtered['DESCRIÇÃO DO SERVIÇO'].astype(str).str.contains(search, case=False, na=False) |
            df_filtered['OBERVAÇÃO:\\nATUALIZADA NA REUNIÃO DO DIA 29/07'].astype(str).str.contains(search, case=False, na=False)
        )
        df_filtered = df_filtered[mask]
    
    st.write(f"**{len(df_filtered)} chamado(s) exibido(s)**")
    
    if len(df_filtered) > 0:
        # Ordenar por unidade (mesma ordem do gráfico) e depois por status
        priority = {s: i for i, s in enumerate(STATUS_ORDER)}
        df_filtered['unit_order'] = df_filtered['UNIDADE'].map({u: i for i, u in enumerate(unit_names)})
        df_filtered['status_order'] = df_filtered['STATUS'].map(priority)
        df_filtered = df_filtered.sort_values(['unit_order', 'status_order', 'NR'], ascending=[True, True, False])
        
        for _, row in df_filtered.iterrows():
            with st.container(border=True):
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.write(f"**O.S. {int(row['O.S'])} · NR {int(row['NR'])}**")
                    st.write(f"**{row['UNIDADE']}**")
                    st.write(f"{row['DESCRIÇÃO DO SERVIÇO']}")
                with col2:
                    status_color = STATUS_COLORS.get(row['STATUS'], "#999")
                    st.write(f"<span style='background:{status_color};color:white;padding:4px 8px;border-radius:3px;display:inline-block'>{row['STATUS']}</span>", unsafe_allow_html=True)
                
                col_a, col_b, col_c = st.columns(3)
                with col_a:
                    st.caption(f"Enviado: {row['DATA DE ENVIO  DO CHAMADO']}")
                with col_b:
                    val = row['VALOR INICIAL DO SERVIÇO']
                    val_str = f"R$ {val:,.2f}".replace(",", "#").replace(".", ",").replace("#", ".") if pd.notna(val) else "—"
                    st.caption(f"Valor: {val_str}")
                with col_c:
                    st.caption(f"Nota liberada: {row['NOTA LIBERADA?']}")
                
                if pd.notna(row['OBERVAÇÃO:\\nATUALIZADA NA REUNIÃO DO DIA 29/07']) and str(row['OBERVAÇÃO:\\nATUALIZADA NA REUNIÃO DO DIA 29/07']).strip():
                    st.caption(f"*{row['OBERVAÇÃO:\\nATUALIZADA NA REUNIÃO DO DIA 29/07']}*")
    else:
        st.info("Nenhum chamado encontrado com esses filtros.")
    
    # Download do relatório
    st.subheader("📥 Exportar Dados")
    
    # Gerar CSV
    csv = df_os.to_csv(index=False)
    st.download_button(
        label="📊 Baixar dados em CSV",
        data=csv,
        file_name="chamados_manutencao.csv",
        mime="text/csv"
    )

else:
    st.info("👆 Envie um arquivo Excel para começar")
