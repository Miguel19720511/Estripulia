import streamlit as st
import pandas as pd
import plotly.express as px
import gdown
import os

st.set_page_config(
    page_title="Dashboard Comercial Estripulia",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📊 Painel Comercial Estripulia — Sell-In & Sell-Out")
st.markdown("Análise Comercial, Giro de Estoque e Cobertura")

FILE_ID_SELL_IN = "1bhptYVaijAOLiX-7Yz6EEG-lM07dV4Va"
LOCAL_FILE = "Sell_in_v2.xlsx"

@st.cache_data
def load_sell_in_data():
    try:
        url = f"https://drive.google.com/uc?id={FILE_ID_SELL_IN}"
        if os.path.exists(LOCAL_FILE):
            os.remove(LOCAL_FILE)
            
        gdown.download(url, LOCAL_FILE, quiet=True)
        df = pd.read_excel(LOCAL_FILE, sheet_name="1-Dados", engine="openpyxl")
        
        cols = {str(c).strip().upper(): c for c in df.columns}
        status_col = cols.get('STATUS')
        almox_col = cols.get('ALMOX.') or cols.get('ALMOXARIFADO') or cols.get('ALMOX')
        
        if not status_col:
            st.error("A coluna STATUS não foi encontrada na folha '1-Dados'.")
            return pd.DataFrame()

        df[status_col] = pd.to_numeric(df[status_col], errors='coerce')
        df = df[df[status_col].isin([5, 6])]
        
        if almox_col:
            df = df[df[almox_col].astype(str).str.strip().str.replace('.0', '', regex=False) == '20']
        
        col_emissao = cols.get('EMISSAO') or cols.get('EMISSÃO') or 'Emissao'
        col_qtd = cols.get('QUANTIDADE') or 'Quantidade'
        col_total = cols.get('VLR.TOTAL') or 'Vlr.Total'
        col_bruto = cols.get('VLR.BRUTO') or 'Vlr.Bruto'

        df['Emissao'] = pd.to_datetime(df[col_emissao], errors='coerce')
        df['Quantidade'] = pd.to_numeric(df[col_qtd], errors='coerce').fillna(0)
        df['Vlr.Total'] = pd.to_numeric(df[col_total], errors='coerce').fillna(0)
        df['Vlr.Bruto'] = pd.to_numeric(df[col_bruto], errors='coerce').fillna(0)
        
        return df
    except Exception as e:
        st.error(f"Erro ao carregar dados do Google Drive: {e}")
        return pd.DataFrame()

with st.spinner("A carregar dados e a aplicar regras de negócio..."):
    df_sell_in = load_sell_in_data()

tab1, tab2, tab3 = st.tabs([
    "📈 Visão Executiva (Sell-In)", 
    "🏪 Sell-Out & Cobertura por Loja", 
    "📦 Saúde do Estoque & SKUs"
])

# TAB 1: SELL-IN
with tab1:
    st.subheader("Faturamento Efetivo de Sell-In (Status 5 e 6 | Almoxarifado 20)")
    if not df_sell_in.empty:
        total_qtd = df_sell_in['Quantidade'].sum()
        total_liq = df_sell_in['Vlr.Total'].sum()
        total_bruto = df_sell_in['Vlr.Bruto'].sum()

        col1, col2, col3 = st.columns(3)
        col1.metric("Volume Faturado", f"{total_qtd:,.0f} un".replace(",", "."))
        col2.metric("Faturamento Líquido", f"R$ {total_liq:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        col3.metric("Faturamento Bruto", f"R$ {total_bruto:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

        st.markdown("---")
        df_sell_in['Ano'] = df_sell_in['Emissao'].dt.year
        sell_in_ano = df_sell_in.groupby('Ano').agg({'Vlr.Total': 'sum', 'Quantidade': 'sum'}).reset_index()

        fig_ano = px.bar(
            sell_in_ano, x='Ano', y='Vlr.Total', text_auto='.2s',
            title="Evolução do Faturamento Líquido de Sell-In por Ano (R$)",
            labels={'Vlr.Total': 'Faturamento Líquido (R$)', 'Ano': 'Ano de Emissão'}
        )
        st.plotly_chart(fig_ano, use_container_width=True)

        st.subheader("Resumo por Ano")
        sell_in_ano['Vlr.Total'] = sell_in_ano['Vlr.Total'].apply(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        sell_in_ano['Quantidade'] = sell_in_ano['Quantidade'].apply(lambda x: f"{x:,.0f} un".replace(",", "."))
        st.dataframe(sell_in_ano, use_container_width=True)

# TAB 2: SELL-OUT & COBERTURA
with tab2:
    st.subheader("Giro Diário e Cobertura de Estoque por Loja")
    st.markdown("Consolidação de vendas PDV e tempo estimado de estoque restante.")
    
    col_f1, col_f2 = st.columns(2)
    meta_cobertura = col_f1.number_input("Meta de Cobertura Alvo (Dias)", min_value=15, max_value=180, value=60)
    lead_time = col_f2.number_input("Lead Time de Entrega (Dias)", min_value=1, max_value=90, value=30)
    
    st.info("Aguardando sincronização dos ficheiros de Sell-out da pasta do Google Drive.")

# TAB 3: SAÚDE DO ESTOQUE
with tab3:
    st.subheader("Análise Crítica de Curva ABC, Rupturas e Excessos")
    
    col_a1, col_a2 = st.columns(2)
    col_a1.metric("SKUs em Ruptura Crítica", "0 itens", delta_color="inverse")
    col_a2.metric("SKUs em Estoque Parado (> 180 dias)", "0 itens", delta_color="inverse")
    
    st.info("Módulo de mapeamento de rupturas pronto para processar o catálogo unificado de Produtos_Marca.xlsx.")
