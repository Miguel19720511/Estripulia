import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as bg
import glob
import os

# Configuração da Página
st.set_page_config(page_title="Dashboard Comercial Estripulia", layout="wide", initial_sidebar_state="expanded")

st.title("📊 Painel Comercial Estripulia — Sell-In & Sell-Out")
st.markdown("Análise de Giro, Dias de Cobertura, Saúde de Estoque e Curva ABC")

# Sidebar - Filtros de Pasta e Parâmetros
st.sidebar.header("⚙️ Configurações & Parâmetros")
data_dir = st.sidebar.text_input("Caminho da Pasta do Drive", value="./Estripulia")

lead_time_dias = st.sidebar.number_input("Lead Time de Entrega (Dias)", min_value=1, value=30)
meta_cobertura_dias = st.sidebar.number_input("Meta de Cobertura Alvo (Dias)", min_value=15, value=60)

@st.cache_data
def load_sell_in(base_path):
    file_path = os.path.join(base_path, "Sell in (2015_2026).xlsx")
    if not os.path.exists(file_path):
        st.warning("Arquivo de Sell-in não localizado no caminho especificado.")
        return pd.DataFrame()
    
    df = pd.read_excel(file_path, sheet_name="1-Dados")
    
    # Aplicação das REGRAS DE OURO
    # 1. Status 5 ou 6
    df = df[df['STATUS'].isin([5, 6])]
    # 2. Almoxarifado 20
    df = df[df['Almox.'].astype(str).str.strip() == '20']
    
    df['Emissao'] = pd.to_datetime(df['Emissao'], errors='coerce')
    df['Ano_Mes'] = df['Emissao'].dt.to_period('M')
    return df

@st.cache_data
def load_sell_out(base_path):
    sell_out_files = glob.glob(os.path.join(base_path, "Sell out", "*.xlsx"))
    all_data = []
    
    for f in sell_out_files:
        filename = os.path.basename(f)
        try:
            df = pd.read_excel(f)
            # Normalização simples do cabeçalho
            df_items = df.iloc[2:].copy()
            df_items['Arquivo'] = filename
            all_data.append((filename, df_items))
        except Exception as e:
            continue
    return all_data

# Carregamento dos dados
with st.spinner("Processando dados e aplicando Regras de Ouro..."):
    df_sell_in = load_sell_in(data_dir)

# Abas do Dashboard
tab1, tab2, tab3 = st.tabs(["📈 Visão Executiva (Sell-In)", "🏪 Sell-Out & Cobertura por Loja", "📦 Saúde do Estoque & SKUs"])

with tab1:
    st.subheader("Faturamento Efetivo de Sell-In (Status 5 e 6 | Almoxarifado 20)")
    if not df_sell_in.empty:
        total_qtd = df_sell_in['Quantidade'].sum()
        total_liq = df_sell_in['Vlr.Total'].sum()
        total_bruto = df_sell_in['Vlr.Bruto'].sum()
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Volume Faturado", f"{total_qtd:,.0f} un")
        col2.metric("Faturamento Líquido", f"R$ {total_liq:,.2f}")
        col3.metric("Faturamento Bruto", f"R$ {total_bruto:,.2f}")
        
        # Gráfico por Ano
        df_sell_in['Ano'] = df_sell_in['Emissao'].dt.year
        sell_in_ano = df_sell_in.groupby('Ano').agg({'Vlr.Total': 'sum', 'Quantidade': 'sum'}).reset_index()
        
        fig = px.bar(sell_in_ano, x='Ano', y='Vlr.Total', text_auto='.2s',
                     title="Evolução do Faturamento Líquido de Sell-In por Ano (R$)")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Aguardando carregamento da base de Sell-in.")

with tab2:
    st.subheader("Análise de Giro e Dias de Cobertura")
    st.markdown("Métricas consolidadas de consumo na ponta e saldo disponível em loja.")
    # Implementação interativa de tabelas dinâmicas
    
with tab3:
    st.subheader("Alertas de Estoque Crítico, Parado e Rupturas")
    st.markdown("Filtro automatizado para itens que necessitam de ação comercial ou remanejamento imediato.")
