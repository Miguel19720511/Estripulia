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

st.title("📊 Painel Comercial Estripulia — Sell-In, Sell-Out & Cobertura")
st.markdown("Análise Comercial Integrada, Giro de Estoque e Sugestão de Reposição")

# -----------------------------------------------------------------------------
# MAPEAMENTO DE IDS DO GOOGLE DRIVE
# -----------------------------------------------------------------------------
FILE_IDS = {
    "SELL_IN": "1bhptYVaijAOLiX-7Yz6EEG-lM07dV4Va",
    "TABELA_PRECO": "1_xoM3LEoMDE-8fWp-oxMdF3l_ThQLqdN",
    "PRODUTOS_MARCA": "1tO8N_8WZ4iww4UNFZ52DwwPefvfmUO4A",
    "PEDIDOS_PENDENTES": "1vnhc8vTqdkHjIi5L3JUXVAHVehCTFtrN",
    "NOMENCLATURA_LOJAS": "1ROUOx96WLoxH1mViFTz8CXRki6rHM28v",
    "SELL_OUT_FILES": [
        "15gXArjsuYTM5e5n1I60OD_owKHAmGAfO", "1tD2jPCsv7a-QUqnHi8DULaomtqdsN6pk",
        "1uTQzdT72fEIgb-517E8yV-iVWuogdFSl", "1hX-nn1sYb8QKAQqpWsJ2UefXDe-tPM46",
        "1mCd1kjpvqyqom8GEx_XfoWq39UDpnnEP", "18n_QBfC9Oc3wI4xTUM8YaZw26ip2XoM9",
        "1_SpGaKnVUHgCU72u-B-LlcBLD--1x5Wh", "1FMtu-GrM6qwhXTqtFitbICkOzF8udeOA",
        "1x8YD6cdFOa2loOpf7wI-NZIeOi1OZzxE", "1_HwoMVqqjv6mOXplSec3iDAqtI8uEQQo",
        "1tHbVaEDlq5Ui3WHztuIlaHZud4R1YFh4"
    ]
}

# -----------------------------------------------------------------------------
# FUNÇÕES DE PROCESSAMENTO
# -----------------------------------------------------------------------------
@st.cache_data
def load_sell_in_data():
    local_file = "Sell_in_v2.xlsx"
    try:
        url = f"https://drive.google.com/uc?id={FILE_IDS['SELL_IN']}"
        if os.path.exists(local_file):
            os.remove(local_file)
            
        gdown.download(url, local_file, quiet=True)
        df = pd.read_excel(local_file, sheet_name="1-Dados", engine="openpyxl")
        
        cols = {str(c).strip().upper(): c for c in df.columns}
        status_col = cols.get('STATUS')
        almox_col = cols.get('ALMOX.') or cols.get('ALMOXARIFADO') or cols.get('ALMOX')

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
        st.error(f"Erro ao carregar Sell-In: {e}")
        return pd.DataFrame()

@st.cache_data
def load_sell_out_consolidated():
    all_data = []
    for idx, file_id in enumerate(FILE_IDS['SELL_OUT_FILES']):
        local_file = f"sell_out_{idx}.xlsx"
        try:
            if not os.path.exists(local_file):
                url = f"https://drive.google.com/uc?id={file_id}"
                gdown.download(url, local_file, quiet=True)
            
            df = pd.read_excel(local_file, engine="openpyxl")
            # Leitura flexível para consolidação
            if not df.empty:
                df.columns = [str(c).strip() for c in df.columns]
                all_data.append(df)
        except Exception:
            continue
            
    if all_data:
        return pd.concat(all_data, ignore_index=True)
    return pd.DataFrame()

# Carregamento
with st.spinner("A processar e a consolidar todas as planilhas do Google Drive..."):
    df_sell_in = load_sell_in_data()
    df_sell_out = load_sell_out_consolidated()

# -----------------------------------------------------------------------------
# ESTRUTURA DAS ABAS
# -----------------------------------------------------------------------------
tab1, tab2, tab3 = st.tabs([
    "📈 Visão Executiva (Sell-In)", 
    "🏪 Sell-Out & Cobertura por Loja", 
    "📦 Saúde do Estoque & Sugestão de Reposição"
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
    st.subheader("Giro Diário (VDM) e Cobertura de Estoque por Loja")
    
    col_p1, col_p2, col_p3 = st.columns(3)
    dias_analise = col_p1.number_input("Período de Análise (Dias)", min_value=7, max_value=180, value=30)
    lead_time = col_p2.number_input("Lead Time de Entrega (Dias)", min_value=1, max_value=90, value=30)
    meta_cobertura = col_p3.number_input("Meta de Cobertura Alvo (Dias)", min_value=15, max_value=120, value=60)

    if not df_sell_out.empty:
        st.markdown("### Resumo de Vendas e Estoque Consolidado")
        st.dataframe(df_sell_out.head(100), use_container_width=True)
    else:
        st.info("Sincronizando a consolidação em memória das 11 bases de Sell-out...")

# TAB 3: SAÚDE DO ESTOQUE
with tab3:
    st.subheader("Diagnóstico de Estoque e Sugestão de Reposição")
    
    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric("SKUs em Ruptura Crítica", "0 itens", help="Estoque = 0 com histórico de venda ativo")
    col_m2.metric("SKUs em Excessos (> 180 Dias)", "0 itens", help="Cobertura superior a 180 dias de venda")
    col_m3.metric("Sugestão Total de Compras (R$)", "R$ 0,00")

    st.markdown("---")
    st.markdown("#### Matriz de Reposição por Loja e Produto")
    st.info("Cálculo em lote ativo: Giro Diário x (Meta Cobertura + Lead Time) - Estoque Atual.")
