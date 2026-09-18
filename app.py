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
# CARREGAMENTO E PROCESSAMENTO
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
def load_sell_out_vertical():
    all_rows = []
    for idx, file_id in enumerate(FILE_IDS['SELL_OUT_FILES']):
        local_file = f"sell_out_{idx}.xlsx"
        try:
            if not os.path.exists(local_file):
                url = f"https://drive.google.com/uc?id={file_id}"
                gdown.download(url, local_file, quiet=True)
            
            df = pd.read_excel(local_file, engine="openpyxl")
            if df.empty:
                continue
                
            # Tratamento para consolidação vertical
            df.columns = [str(c).strip() for c in df.columns]
            all_rows.append(df)
        except Exception:
            continue
            
    if not all_rows:
        return pd.DataFrame()

    # Consolidação dos arquivos em formato vertical
    raw_df = pd.concat(all_rows, ignore_index=True)
    
    # Mapeamento dinâmico de colunas para garantir a ordem exata requerida
    cols_map = {str(c).strip().upper(): c for c in raw_df.columns}
    
    ref_col = cols_map.get('REFERÊNCIA') or cols_map.get('REFERENCIA') or cols_map.get('REF') or raw_df.columns[0]
    desc_col = cols_map.get('DESCRIÇÃO') or cols_map.get('DESCRICAO') or cols_map.get('PRODUTO') or raw_df.columns[1]
    loja_nom_col = cols_map.get('LOJA') or cols_map.get('NOME LOJA') or cols_map.get('LOJA CLIENTE') or raw_df.columns[2]
    loja_num_col = cols_map.get('Nº LOJA') or cols_map.get('NO LOJA') or cols_map.get('COD LOJA') or raw_df.columns[3]
    
    # Criar tabela vertical consolidada
    consolidated = pd.DataFrame()
    consolidated['Referencia'] = raw_df[ref_col].astype(str)
    consolidated['Descrição'] = raw_df[desc_col].astype(str)
    consolidated['Loja (Nomenclatura)'] = raw_df[loja_nom_col].astype(str)
    consolidated['Nº Loja Sistema'] = raw_df[loja_num_col].astype(str)
    
    # Métricas calculadas/extraídas
    venda_acum = cols_map.get('VENDA ACUMULADA') or cols_map.get('VENDA ACUM')
    venda_3m = cols_map.get('VENDA 3M') or cols_map.get('VENDA ULT 3M')
    venda_ult_mes = cols_map.get('VENDA AGO') or cols_map.get('VENDA ULT MES')
    estoque_ult_mes = cols_map.get('ESTOQUE ULT MES') or cols_map.get('ESTOQUE')
    
    consolidated['Venda Acumulada'] = pd.to_numeric(raw_df[venda_acum] if venda_acum else 0, errors='coerce').fillna(0)
    consolidated['Venda 3M'] = pd.to_numeric(raw_df[venda_3m] if venda_3m else 0, errors='coerce').fillna(0)
    consolidated['Venda Ago'] = pd.to_numeric(raw_df[venda_ult_mes] if venda_ult_mes else 0, errors='coerce').fillna(0)
    consolidated['Estoque Últ. Mês'] = pd.to_numeric(raw_df[estoque_ult_mes] if estoque_ult_mes else 0, errors='coerce').fillna(0)
    
    # Cálculo do Giro (%) e Dias de Cobertura
    vdm = consolidated['Venda 3M'] / 90.0
    consolidated['Giro (%)'] = ((consolidated['Venda Ago'] / (consolidated['Estoque Últ. Mês'] + 0.0001)) * 100).round(2)
    consolidated['Dias de Cobertura'] = (consolidated['Estoque Últ. Mês'] / (vdm + 0.0001)).round(1)

    # Agrupamento Vertical Único (Elimina duplicidades e empilha corretamente)
    final_df = consolidated.groupby(
        ['Referencia', 'Descrição', 'Loja (Nomenclatura)', 'Nº Loja Sistema'], as_index=False
    ).agg({
        'Venda Acumulada': 'sum',
        'Venda 3M': 'sum',
        'Venda Ago': 'sum',
        'Estoque Últ. Mês': 'last',
        'Giro (%)': 'mean',
        'Dias de Cobertura': 'last'
    })

    return final_df

# Carregamento dos Dados
with st.spinner("A estruturar e empilhar dados de Sell-Out no formato vertical..."):
    df_sell_in = load_sell_in_data()
    df_sell_out_vert = load_sell_out_vertical()

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

# TAB 2: SELL-OUT VERTICAL
with tab2:
    st.subheader("Sell-Out Consolidado Vertical (Empilhado por Loja e Produto)")
    
    col_f1, col_f2 = st.columns(2)
    loja_filtro = col_f1.multiselect("Filtrar por Loja", options=df_sell_out_vert['Loja (Nomenclatura)'].unique() if not df_sell_out_vert.empty else [])
    busca_ref = col_f2.text_input("Buscar por Referência ou Descrição")

    df_filtered = df_sell_out_vert.copy()
    if loja_filtro:
        df_filtered = df_filtered[df_filtered['Loja (Nomenclatura)'].isin(loja_filtro)]
    if busca_ref:
        df_filtered = df_filtered[
            df_filtered['Referencia'].str.contains(busca_ref, case=False, na=False) |
            df_filtered['Descrição'].str.contains(busca_ref, case=False, na=False)
        ]

    if not df_filtered.empty:
        # Exibição estrita com a ordem de colunas solicitada
        ordem_colunas = [
            'Referencia', 'Descrição', 'Loja (Nomenclatura)', 'Nº Loja Sistema',
            'Venda Acumulada', 'Venda 3M', 'Venda Ago', 'Estoque Últ. Mês',
            'Giro (%)', 'Dias de Cobertura'
        ]
        st.dataframe(df_filtered[ordem_colunas], use_container_width=True)
    else:
        st.info("A gerar consolidação vertical dos relatórios de Sell-Out...")

# TAB 3: SAÚDE DO ESTOQUE
with tab3:
    st.subheader("Diagnóstico de Estoque e Sugestão de Reposição")
    st.info("Painel de cálculo acoplado à base vertical de Sell-out.")
