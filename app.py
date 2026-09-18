import streamlit as st
import pandas as pd
import plotly.express as px
import gdown
import os

# Configuração inicial da página
st.set_page_config(
    page_title="Dashboard Comercial Estripulia",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📊 Painel Comercial Estripulia — Sell-In & Sell-Out")
st.markdown("Análise Comercial, Giro de Estoque e Cobertura")

# ID do ficheiro Sell In no Google Drive
FILE_ID_SELL_IN = "1bhptYVaijAOLiX-7Yz6EEG-lM07dV4Va"
LOCAL_FILE = "Sell_in_temp.xlsx"

@st.cache_data
def load_sell_in_data():
    try:
        url = f"https://drive.google.com/uc?id={FILE_ID_SELL_IN}"
        
        # Download com parâmetro fuzzy ativado
        if not os.path.exists(LOCAL_FILE):
            gdown.download(url, LOCAL_FILE, quiet=True, fuzzy=True)
            
        df = pd.read_excel(LOCAL_FILE, sheet_name="1-Dados", engine="openpyxl")
        
        # REGRAS DE OURO:
        # 1. Filtro de Status = 5 ou 6
        df = df[df['STATUS'].isin([5, 6])]
        
        # 2. Filtro de Almoxarifado = 20
        df = df[df['Almox.'].astype(str).str.strip() == '20']
        
        df['Emissao'] = pd.to_datetime(df['Emissao'], errors='coerce')
        df['Quantidade'] = pd.to_numeric(df['Quantidade'], errors='coerce').fillna(0)
        df['Vlr.Total'] = pd.to_numeric(df['Vlr.Total'], errors='coerce').fillna(0)
        df['Vlr.Bruto'] = pd.to_numeric(df['Vlr.Bruto'], errors='coerce').fillna(0)
        
        return df
    except Exception as e:
        st.error(f"Erro ao carregar dados do Google Drive: {e}")
        return pd.DataFrame()

# Carregamento dos dados com feedback visual
with st.spinner("A carregar dados do Google Drive e a aplicar as Regras de Ouro..."):
    df_sell_in = load_sell_in_data()

# Estrutura em separadores (Tabs)
tab1, tab2, tab3 = st.tabs([
    "📈 Visão Executiva (Sell-In)", 
    "🏪 Sell-Out & Cobertura por Loja", 
    "📦 Saúde do Estoque & SKUs"
])

# -----------------------------------------------------------------------------
# TAB 1: VISÃO EXECUTIVA (SELL-IN)
# -----------------------------------------------------------------------------
with tab1:
    st.subheader("Faturamento Efetivo de Sell-In (Status 5 e 6 | Almoxarifado 20)")
    
    if not df_sell_in.empty:
        # Indicadores Globais
        total_qtd = df_sell_in['Quantidade'].sum()
        total_liq = df_sell_in['Vlr.Total'].sum()
        total_bruto = df_sell_in['Vlr.Bruto'].sum()

        col1, col2, col3 = st.columns(3)
        col1.metric("Volume Faturado", f"{total_qtd:,.0f} un".replace(",", "."))
        col2.metric("Faturamento Líquido", f"R$ {total_liq:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        col3.metric("Faturamento Bruto", f"R$ {total_bruto:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

        st.markdown("---")
        
        # Agrupamento por Ano
        df_sell_in['Ano'] = df_sell_in['Emissao'].dt.year
        sell_in_ano = df_sell_in.groupby('Ano').agg({
            'Vlr.Total': 'sum', 
            'Quantidade': 'sum'
        }).reset_index()

        # Gráfico de Evolução Anual
        fig_ano = px.bar(
            sell_in_ano, 
            x='Ano', 
            y='Vlr.Total', 
            text_auto='.2s',
            title="Evolução do Faturamento Líquido de Sell-In por Ano (R$)",
            labels={'Vlr.Total': 'Faturamento Líquido (R$)', 'Ano': 'Ano de Emissão'}
        )
        st.plotly_chart(fig_ano, use_container_width=True)

        # Tabela com detalhamento por Ano
        st.subheader("Resumo por Ano")
        sell_in_ano['Vlr.Total'] = sell_in_ano['Vlr.Total'].apply(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        sell_in_ano['Quantidade'] = sell_in_ano['Quantidade'].apply(lambda x: f"{x:,.0f} un".replace(",", "."))
        st.dataframe(sell_in_ano, use_container_width=True)
    else:
        st.warning("Não foram encontrados registos válidos com os filtros aplicados.")

# -----------------------------------------------------------------------------
# TAB 2: SELL-OUT POR LOJA
# -----------------------------------------------------------------------------
with tab2:
    st.subheader("Análise de Giro e Cobertura nas Lojas")
    st.info("Módulo de leitura dos ficheiros mensais de Sell-out em consolidação.")

# -----------------------------------------------------------------------------
# TAB 3: SAÚDE DO ESTOQUE
# -----------------------------------------------------------------------------
with tab3:
    st.subheader("Alertas de Estoque Crítico e Parado")
    st.info("Módulo de mapeamento de rupturas e remanejamento em consolidação.")
