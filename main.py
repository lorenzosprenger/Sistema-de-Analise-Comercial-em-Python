# streamlit_app.py
import streamlit as st
import pandas as pd
from service import analyze, prepare_df

st.set_page_config(page_title="Análise Comercial Inteligente", layout="wide")

st.write("\U0001F680 **Sistema de Análise Comercial em Python** - Interface carregada com sucesso!")
st.markdown("---")

st.write("**Passo 1:** Selecione as quatro planilhas abaixo:")

def store_file(label, key):
    file = st.file_uploader(label, type="xlsx", key=key)
    return file

uploaded_faturado = store_file("📄 Planilha faturado.XLSX", "faturado")
uploaded_orcamento = store_file("📄 Planilha orcamento.XLSX", "orcamento")
uploaded_pedidos = store_file("📄 Planilha pedidos.XLSX", "pedidos")
uploaded_estoque = store_file("📄 Planilha estoque.XLSX", "estoque")

if not uploaded_faturado or not uploaded_orcamento or not uploaded_pedidos or not uploaded_estoque:
    st.warning("Aguardando upload das quatro planilhas para prosseguir...")
else:
    st.success("Arquivos carregados! Role a página ou clique no botão para iniciar a análise.")

    if st.button("🔍 Analisar"):
        df_fat = pd.read_excel(uploaded_faturado, header=5)
        df_orc = pd.read_excel(uploaded_orcamento, header=0)
        df_ped = pd.read_excel(uploaded_pedidos, header=0)
        df_estoque = pd.read_excel(uploaded_estoque, header=0)

        df_orc = df_orc.loc[:, ~df_orc.columns.str.contains('^Unnamed')]
        df_ped = df_ped.loc[:, ~df_ped.columns.str.contains('^Unnamed')]

        df_fat = prepare_df(df_fat)
        df_orc = prepare_df(df_orc)
        df_ped = prepare_df(df_ped)

        try:
            insights = analyze(df_fat, df_orc, df_ped)
        except Exception as e:
            st.error(f"Erro na análise: {e}")
            st.stop()

        df_all = pd.concat([
            df_orc.assign(Tipo="Orçamento"),
            df_ped.assign(Tipo="Pedido"),
            df_fat.assign(Tipo="Faturado")
        ], ignore_index=True)
        df_all['Valor_Total'] = df_all['Quantidade'] * df_all['Valor Unitário']
        df_all['AnoMes'] = df_all['Data'].dt.to_period('M').astype(str)

        df_estoque['Mês/Ano'] = pd.to_datetime(df_estoque['Mês/Ano'], format='%m/%Y', errors='coerce')
        ultimo_mes_estoque = df_estoque['Mês/Ano'].max().to_period('M')

        st.session_state['df_fat'] = df_fat
        st.session_state['df_all'] = df_all
        st.session_state['df_estoque'] = df_estoque
        st.session_state['ultimo_mes_estoque'] = ultimo_mes_estoque
        st.session_state['insights'] = insights

if 'df_all' in st.session_state:
    df_fat = st.session_state['df_fat']
    df_all = st.session_state['df_all']
    df_estoque = st.session_state['df_estoque']
    ultimo_mes_estoque = st.session_state['ultimo_mes_estoque']
    insights = st.session_state['insights']

    st.header("🏆 Ranking de Vendedores")
    rev_vendedores = df_all.groupby('Representante')['Valor_Total'].sum().reset_index()
    rev_vendedores = rev_vendedores.sort_values('Valor_Total', ascending=False)
    rev_vendedores['Receita Total'] = rev_vendedores['Valor_Total'].map(lambda x: f"R$ {x:,.2f}")
    st.table(rev_vendedores[['Representante', 'Receita Total']])

    st.header("📋 Clientes por Representante")
    rev_rep_month = df_all.groupby(['Representante', 'AnoMes'])['Valor_Total'].sum().reset_index()
    rev_rep_month['Receita'] = rev_rep_month['Valor_Total'].map(lambda x: f"R$ {x:,.2f}")
    st.dataframe(rev_rep_month[['Representante', 'AnoMes', 'Receita']])

    rev_cli_month = df_all.groupby(['Representante', 'Cliente', 'AnoMes'])['Valor_Total'].sum().reset_index()
    rev_cli_month['Receita'] = rev_cli_month['Valor_Total'].map(lambda x: f"R$ {x:,.2f}")
    st.dataframe(rev_cli_month[['Representante', 'Cliente', 'AnoMes', 'Receita']])

    st.header("📊 Taxas de Conversão & Evolução de Vendas")
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("🔄 Taxas de Conversão")
        st.bar_chart(insights['conversion_rates'].set_index('Etapa'))
    with col2:
        st.subheader("📈 Evolução Mensal")
        evo_qty = insights['sales_evolution'].set_index('Data')['Quantidade']
        evo_val = df_all.groupby(pd.Grouper(key='Data', freq='M'))['Valor_Total'].sum()
        evo_df = pd.concat([evo_qty, evo_val], axis=1)
        evo_df.columns = ['Quantidade', 'Valor Total']
        st.line_chart(evo_df)

    st.header("📈 Análise de Estoque")
    opcao_estoque = st.selectbox("Filtrar estoque por:",
                                 ("Todos (ITECH + Representantes)", "Apenas ITECH", "Apenas Representantes"),
                                 key="filtro_estoque")

    # Filtro aplicado
    if opcao_estoque == "Apenas ITECH":
        df_estoque_filtrado = df_estoque[(df_estoque['Mês/Ano'].dt.to_period('M') == ultimo_mes_estoque) &
                                        (df_estoque['Local'] == 49)]
    elif opcao_estoque == "Apenas Representantes":
        df_estoque_filtrado = df_estoque[(df_estoque['Mês/Ano'].dt.to_period('M') == ultimo_mes_estoque) &
                                        (df_estoque['Local'] != 49)]
    else:
        df_estoque_filtrado = df_estoque[df_estoque['Mês/Ano'].dt.to_period('M') == ultimo_mes_estoque]

    # A partir daqui, sempre use df_estoque_filtrado

    df_estoque_agrupado = (
        df_estoque_filtrado
        .groupby(['Produto', 'Referência', 'Descrição'], as_index=False)['Qtd.Física']
        .sum()
        .rename(columns={'Descrição': 'Descrição do Produto', 'Qtd.Física': 'Estoque'})
    )


    st.dataframe(df_estoque_agrupado)

    st.header("🔧 Consumo Médio de Ferramentas")
    df_prod_mes = df_fat.groupby([
    pd.Grouper(key='Data', freq='M'),
    'Produto', 'Descrição do Produto', 'Cliente'
])['Quantidade'].sum().reset_index()



    consumo_medio = (
        df_prod_mes.groupby(['Cliente', 'Produto', 'Descrição do Produto'])['Quantidade']
        .mean().round(2).reset_index(name='Consumo Médio Mensal')
    )


    consumo_join = consumo_medio.merge(
        df_estoque_agrupado[['Produto', 'Descrição do Produto', 'Referência', 'Estoque']],
        on=['Produto', 'Descrição do Produto'],
        how='left'
    )

    def color_code(row):
        if pd.isna(row['Estoque']): return 'N/A'
        if row['Estoque'] > row['Consumo Médio Mensal']: return f"\U0001F7E2 {row['Estoque']}"
        elif row['Estoque'] == row['Consumo Médio Mensal']: return f"\U0001F7E1 {row['Estoque']}"
        else: return f"\U0001F534 {row['Estoque']}"

    consumo_join['Estoque Situação'] = consumo_join.apply(color_code, axis=1)

    st.subheader("Consumo Médio Mensal por Cliente e Estoque Atual")
    st.dataframe(consumo_join[[
        'Cliente', 'Produto', 'Referência', 'Descrição do Produto',
        'Consumo Médio Mensal', 'Estoque Situação'
    ]])

    st.header("📦 Giro de Produtos no Estoque (últimos 6 meses)")

    # Definir intervalo dos últimos 6 meses
    last_date = df_fat['Data'].max()
    meses = pd.date_range(end=last_date, periods=7, freq='MS')

    # Agrupar vendas por Produto e Mês
    vendas_mes = (
        df_fat[df_fat['Data'] >= meses.min()]
        .groupby([pd.Grouper(key='Data', freq='MS'), 'Produto', 'Descrição do Produto'])['Quantidade']
        .sum()
        .reset_index()
)


    # Criar base com todos produtos do estoque cruzados com os últimos 6 meses
    estoque_produtos = df_estoque_agrupado[['Produto', 'Descrição do Produto', 'Referência']].drop_duplicates()
    meses_df = pd.DataFrame({'Data': meses})
    base = meses_df.merge(estoque_produtos, how='cross')

    # Mescla vendas com produtos + referência
    giro_df = base.merge(vendas_mes, on=['Data', 'Produto', 'Descrição do Produto'], how='left')

    # Mantém a coluna 'Referência' do estoque
    giro_df['Quantidade'] = giro_df['Quantidade'].fillna(0).astype(int)
    giro_df['Mês'] = giro_df['Data'].dt.strftime('%Y-%m')

    # Agora podemos fazer o pivot com Referência inclusa
    tabela_giro = giro_df.pivot_table(
        index=['Produto', 'Referência', 'Descrição do Produto'], 
        columns='Mês', values='Quantidade', fill_value=0
    ).reset_index()

    tabela_giro = tabela_giro.merge(df_estoque_agrupado, on=['Produto', 'Referência', 'Descrição do Produto'], how='left')
    tabela_giro = tabela_giro.rename(columns={'Estoque': 'Estoque Atual'})

    colunas_finais = ['Produto', 'Referência', 'Descrição do Produto', 'Estoque Atual'] + sorted(
        [col for col in tabela_giro.columns if col not in ['Produto', 'Referência', 'Descrição do Produto', 'Estoque Atual']]
    )
    tabela_giro = tabela_giro[colunas_finais]

    st.dataframe(tabela_giro)


    # Identificar produtos sem nenhuma venda
    vendas_totais = tabela_giro.set_index(['Produto', 'Referência', 'Descrição do Produto']).drop(columns='Estoque Atual').sum(axis=1)
    estoque_atual = tabela_giro.set_index(['Produto', 'Referência', 'Descrição do Produto'])['Estoque Atual']

    produtos_sem_venda = vendas_totais[(vendas_totais == 0) & (estoque_atual > 0)].reset_index()
    produtos_sem_venda['Estoque Atual'] = produtos_sem_venda.apply(
        lambda row: estoque_atual.loc[(row['Produto'], row['Referência'], row['Descrição do Produto'])], axis=1
    )

    if not produtos_sem_venda.empty:
        st.warning(f"{len(produtos_sem_venda)} produtos sem nenhuma venda nos últimos 6 meses e ainda em estoque:")
        st.dataframe(produtos_sem_venda[['Produto', 'Referência', 'Descrição do Produto', 'Estoque Atual']])

        st.success("Análise concluída com sucesso!")