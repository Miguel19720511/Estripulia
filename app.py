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
# CARREGAMENTO SELL-IN
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

# -----------------------------------------------------------------------------
# CARREGAMENTO DA TABELA DE LOJAS (DE-PARA)
# -----------------------------------------------------------------------------
@st.cache_data
def load_nomenclatura_lojas():
    local_file = "nomenclatura_lojas.xlsx"
    try:
        url = f"https://drive.google.com/uc?id={FILE_IDS['NOMENCLATURA_LOJAS']}"
        if not os.path.exists(local_file):
            gdown.download(url, local_file, quiet=True)
            
        df = pd.read_excel(local_file, engine="openpyxl")
        df.columns = [str(c).strip() for c in df.columns]
        
        # Mapeamento estrito baseado no formato do cadastro
        cols_map = {str(c).strip().upper().replace(' ', ''): c for c in df.columns}
        
        nom_col = cols_map.get('NOMENCLATURACLIENTE') or cols_map.get('NOMENCLATURA')
        cod_col = cols_map.get('CODIGO') or cols_map.get('CODIGOMASTER')
        num_col = cols_map.get('LOJA') or cols_map.get('NOLOJA')
        cnpj_col = cols_map.get('CNPJ/CPF') or cols_map.get('CNPJ')

        loja_df = pd.DataFrame()
        loja_df['Loja (Nomenclatura)'] = df[nom_col].astype(str).str.strip()
        loja_df['Código Máster'] = df[cod_col].astype(str).str.strip()
        loja_df['Nº Loja Sistema'] = df[num_col].astype(str).str.strip()
        loja_df['CNPJ'] = df[cnpj_col].astype(str).str.strip()
        
        return loja_df.drop_duplicates(subset=['Loja (Nomenclatura)'])
    except Exception as e:
        st.warning(f"Aviso ao carregar cadastro de lojas: {e}")
        return pd.DataFrame()

# -----------------------------------------------------------------------------
# CONSOLIDAÇÃO DE SELL-OUT VERTICAL
# -----------------------------------------------------------------------------
@st.cache_data
def load_sell_out_vertical():
    df_lojas = load_nomenclatura_lojas()
    all_rows = []
    
    for idx, file_id in enumerate(FILE_IDS['SELL_OUT_FILES']):
        local_file = f"sell_out_{idx}.xlsx"
        try:
            if not os.path.exists(local_file):
                url = f"https://drive.google.com/uc?id={file_id}"
                gdown.download(url, local_file, quiet=True)
            
            df = pd.read_excel(local_file, engine="openpyxl")
            if df.empty or df.shape[1] < 5:
                continue
                
            df['Ordem_Arquivo'] = idx
            all_rows.append(df)
        except Exception:
            continue
            
    if not all_rows:
        return pd.DataFrame()

    raw_df = pd.concat(all_rows, ignore_index=True)
    
    consolidated = pd.DataFrame()
    # Posições do relatório de Sell-Out:
    # 0: Código/Referência | 1: Descrição do Produto | 2: Nome da Loja (Nomenclatura Cliente)
    consolidated['Referencia'] = raw_df.iloc[:, 0].astype(str).str.strip()
    consolidated['Descrição'] = raw_df.iloc[:, 1].astype(str).str.strip()
    consolidated['Loja (Nomenclatura)'] = raw_df.iloc[:, 2].astype(str).str.strip()
    consolidated['Ordem_Arquivo'] = raw_df['Ordem_Arquivo']
    
    # Posições de Venda e Estoque do Relatório
    consolidated['Venda_Mes'] = pd.to_numeric(raw_df.iloc[:, 4], errors='coerce').fillna(0)
    consolidated['Estoque_Mes'] = pd.to_numeric(raw_df.iloc[:, 5], errors='coerce').fillna(0)

    # Cruzamento de-para com a tabela de lojas
    if not df_lojas.empty:
        consolidated = consolidated.merge(df_lojas, on='Loja (Nomenclatura)', how='left')
        consolidated['Código Máster'] = consolidated['Código Máster'].fillna("Pendente Cadastro")
        consolidated['Nº Loja Sistema'] = consolidated['Nº Loja Sistema'].fillna("Pendente")
        consolidated['CNPJ'] = consolidated['CNPJ'].fillna("Pendente")
    else:
        consolidated['Código Máster'] = "N/A"
        consolidated['Nº Loja Sistema'] = "N/A"
        consolidated['CNPJ'] = "N/A"

    max_ordem = consolidated['Ordem_Arquivo'].max()
    ordem_ult3 = max(0, max_ordem - 2)

    # Consolidação e Cálculos de Giro/Cobertura
    def calc_group(g):
        venda_acumulada = g['Venda_Mes'].sum()
        venda_3m = g[g['Ordem_Arquivo'] >= ordem_ult3]['Venda_Mes'].sum()
        
        g_ult = g[g['Ordem_Arquivo'] == max_ordem]
        venda_ult_mes = g_ult['Venda_Mes'].sum() if not g_ult.empty else 0
        estoque_val = g_ult['Estoque_Mes'].values[-1] if not g_ult.empty else 0
        
        vdm = venda_3m / 90.0 if venda_3m > 0 else 0
        giro = ((venda_ult_mes / estoque_val) * 100) if estoque_val > 0 else 0
        cobertura = (estoque_val / vdm) if vdm > 0 else (999.0 if estoque_val > 0 else 0.0)

        return pd.Series({
            'Venda Acumulada': int(venda_acumulada),
            'Venda 3M': int(venda_3m),
            'Venda Últ. Mês': int(venda_ult_mes),
            'Estoque Últ. Mês': int(estoque_val),
            'Giro (%)': round(giro, 2),
            'Dias de Cobertura': round(cobertura, 1)
        })

    final_df = consolidated.groupby(
        ['Referencia', 'Descrição', 'Loja (Nomenclatura)', 'Código Máster', 'Nº Loja Sistema', 'CNPJ'], as_index=False
    ).apply(calc_group)

    return final_df

# Carregamento Geral
with st.spinner("A cruzar cadastro de lojas e a consolidar vendas de Sell-Out..."):
    df_sell_in = load_sell_in_data()
    df_sell_out_vert = load_sell_out_vertical()

# -----------------------------------------------------------------------------
# INTERFACE DE UTILIZADOR (TABS)
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

# TAB 2: SELL-OUT VERTICAL
with tab2:
    st.subheader("Sell-Out Consolidado Vertical (Empilhado por Loja e Produto)")
    
    if not df_sell_out_vert.empty:
        col_f1, col_f2 = st.columns(2)
        loja_filtro = col_f1.multiselect("Filtrar por Loja", options=sorted(df_sell_out_vert['Loja (Nomenclatura)'].unique()))
        busca_ref = col_f2.text_input("Buscar por Referência ou Descrição")

        df_filtered = df_sell_out_vert.copy()
        if loja_filtro:
            df_filtered = df_filtered[df_filtered['Loja (Nomenclatura)'].isin(loja_filtro)]
        if busca_ref:
            df_filtered = df_filtered[
                df_filtered['Referencia'].str.contains(busca_ref, case=False, na=False) |
                df_filtered['Descrição'].str.contains(busca_ref, case=False, na=False)
            ]

        ordem_colunas = [
            'Referencia', 'Descrição', 'Loja (Nomenclatura)', 'Código Máster', 'Nº Loja Sistema', 'CNPJ',
            'Venda Acumulada', 'Venda 3M', 'Venda Últ. Mês', 'Estoque Últ. Mês',
            'Giro (%)', 'Dias de Cobertura'
        ]
        st.dataframe(df_filtered[ordem_colunas], use_container_width=True)
    else:
        st.info("A processar de-para de lojas e relatórios de Sell-Out...")

# TAB 3: SAÚDE DO ESTOQUE
with tab3:
    st.subheader("Diagnóstico de Estoque e Sugestão de Reposição")
    st.info("Painel acoplado ao cadastro mestre de lojas e produtos.")
