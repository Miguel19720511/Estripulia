import streamlit as st
import pandas as pd
import plotly.express as px
import os

# Configuração da Página
st.set_page_config(page_title="Dashboard Comercial Estripulia", layout="wide")

st.title("📊 Painel Comercial Estripulia — Sell-In & Sell-Out")
st.markdown("Análise de Giro, Dias de Cobertura, Saúde de Estoque e Curva ABC")

# Links de acesso aos ficheiros do Drive
URL_SELL_IN = "https://drive.google.com/uc?export=download&id=1bhptYVaijAOLiX-7Yz6EEG-lM07dV4Va"

@st.cache_data
def load_sell_in_drive():
    try:
        # Leitura direta da base de Sell-in via Google Drive
        df = pd.read_excel(URL_SELL_IN, sheet_name="1-Dados")
        
        # REGRAS DE OURO:
        # 1. Apenas STATUS = 5 ou 6
        df = df[df['STATUS'].isin([5, 6])]
        
        # 2. Apenas Almox. = 20
        df = df[df['Almox.'].astype(str).str.strip() == '20']
        
        df['Emissao'] = pd.to_datetime(df['Emissao'], errors='coerce')
        return df
    except Exception as e:
        st.error(f"Erro ao carregar dados do Drive: {e}")
        return pd.DataFrame()

# Carregamento
with st.spinner("Conectando ao Google Drive e aplicando Regras de Ouro..."):
    df_sell_in = load_sell_in_drive()

# Visualização da Aba 1
tab1, tab2, tab3 = st.tabs(["📈 Visão Executiva (Sell-In)", "🏪 Sell-Out & Cobertura por Loja", "📦 Saúde do Estoque & SKUs"])

with tab1:
    st.subheader("Faturamento Efetivo de Sell-In (Status 5 e 6 | Almoxarifado 20)")
    if not df_sell_in.empty:
        # Tratamento numérico de colunas
        df_sell_in['Quantidade'] = pd.to_numeric(df_sell_in['Quantidade'], errors='coerce').fillna(0)
        df_sell_in['Vlr.Total'] = pd.to_numeric(df_sell_in['Vlr.Total'], errors='coerce').fillna(0)
        df_sell_in['Vlr.Bruto'] = pd.to_numeric(df_sell_in['Vlr.Bruto'], errors='coerce').fillna(0)

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

with tab2:
    st.info("Módulo de Sell-Out por Loja em carregamento dinâmico.")

with tab3:
    st.info("Módulo de Saúde do Estoque em carregamento dinâmico.")
