import streamlit as st
import pandas as pd
from service import analyze, prepare_df

st.set_page_config(page_title="Análise Comercial Inteligente", layout="wide")

st.write("\U0001F680 **Sistema de Análise Comercial em Python** - Interface carregada com sucesso!")
st.markdown("---")

st.write("**Passo 1:** Selecione as quatro planilhas abaixo:")

uploaded_faturado = st.file_uploader("📄 Planilha faturado.XLSX", type="xlsx")
uploaded_orcamento = st.file_uploader("📄 Planilha orcamento.XLSX", type="xlsx")
uploaded_pedidos = st.file_uploader("📄 Planilha pedidos.XLSX", type="xlsx")
uploaded_estoque = st.file_uploader("📄 Planilha estoque.XLSX", type="xlsx")

if not uploaded_faturado or not uploaded_orcamento or not uploaded_pedidos or not uploaded_estoque:
    st.warning("Aguardando upload das quatro planilhas para prosseguir...")
else:
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
            prepare_df(df_orc).assign(Tipo="Orçamento"),
            prepare_df(df_ped).assign(Tipo="Pedido"),
            df_fat.assign(Tipo="Faturado")
        ], ignore_index=True)
        df_all['Valor_Total'] = df_all['Quantidade'] * df_all['Valor Unitário']
        df_all['AnoMes'] = df_all['Data'].dt.to_period('M').astype(str)

        st.header("🏆 Ranking de Vendedores")
        rev_vendedores = (
            df_all.groupby('Representante')['Valor_Total'].sum().reset_index()
            .sort_values('Valor_Total', ascending=False)
        )
        rev_vendedores['Receita Total'] = rev_vendedores['Valor_Total'].map(lambda x: f"R$ {x:,.2f}")
        st.table(rev_vendedores[['Representante', 'Receita Total']].reset_index(drop=True))

        st.header("📋 Clientes por Representante")
        rev_rep_month = df_all.groupby(['Representante', 'AnoMes'])['Valor_Total'].sum().reset_index()
        rev_rep_month['Receita'] = rev_rep_month['Valor_Total'].map(lambda x: f"R$ {x:,.2f}")
        st.subheader("Receita Total por Representante / Mês")
        st.dataframe(rev_rep_month[['Representante', 'AnoMes', 'Receita']])

        rev_cli_month = df_all.groupby(['Representante', 'Cliente', 'AnoMes'])['Valor_Total'].sum().reset_index()
        rev_cli_month['Receita'] = rev_cli_month['Valor_Total'].map(lambda x: f"R$ {x:,.2f}")
        st.subheader("Receita por Cliente por Mês")
        st.dataframe(rev_cli_month[['Representante', 'Cliente', 'AnoMes', 'Receita']])

        st.header("📊 Taxas de Conversão & Evolução de Vendas")
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("🔄 Taxas de Conversão")
            st.bar_chart(insights['conversion_rates'].set_index('Etapa'))
        with col2:
            st.subheader("📈 Evolução Mensal (Quantidade & Valor)")
            evo_qty = insights['sales_evolution'].set_index('Data')['Quantidade']
            evo_val = df_all.groupby(pd.Grouper(key='Data', freq='M'))['Valor_Total'].sum()
            evo_df = pd.concat([evo_qty, evo_val], axis=1)
            evo_df.columns = ['Quantidade', 'Valor Total']
            st.line_chart(evo_df)

        st.header("📉 Clientes com Redução de Volume")
        last_date = df_fat['Data'].max()
        threshold = last_date - pd.DateOffset(months=6)
        recent = df_fat[df_fat['Data'] >= threshold]
        volumes_all = df_fat.groupby('Cliente')['Quantidade'].sum()
        volumes_recent = recent.groupby('Cliente')['Quantidade'].sum().reindex(volumes_all.index, fill_value=0)
        df_decline = pd.DataFrame({
            'Cliente': volumes_all.index,
            'Antes': volumes_all.values,
            'Agora': volumes_recent.values,
            'Diferença': (volumes_all - volumes_recent).values
        })
        df_decline = df_decline[df_decline['Diferença'] > 0].sort_values('Diferença', ascending=False)

        for _, row in df_decline.iterrows():
            with st.expander(f"{row['Cliente']} - Queda: {row['Diferença']}", expanded=False):
                st.markdown(f"**Antes:** {row['Antes']} | **Agora:** {row['Agora']} | :red[**Diferença:** {row['Diferença']}] ")
                client_data = df_fat[df_fat['Cliente'] == row['Cliente']]
                df_items = (client_data.groupby(['Produto', 'Descrição do Produto'])
                            .agg(Quantidade=('Quantidade','sum'), Ultima_Compra=('Data','max'))
                            .reset_index()
                            .sort_values('Quantidade', ascending=False))
                df_items['Última Compra'] = df_items['Ultima_Compra'].dt.strftime('%Y-%m-%d')
                st.table(df_items[['Produto','Descrição do Produto','Quantidade','Última Compra']])

        st.header("🔧 Consumo Médio de Ferramentas")
        df_prod_mes = df_fat.groupby([pd.Grouper(key='Data', freq='M'), 'Produto', 'Descrição do Produto', 'Cliente'])['Quantidade'].sum().reset_index()
        consumo_medio = (df_prod_mes.groupby(['Cliente', 'Produto', 'Descrição do Produto'])['Quantidade']
                         .mean().round(2).reset_index(name='Consumo Médio Mensal'))

        df_estoque_agrupado = (
        df_estoque.groupby(['Produto', 'Descrição'], as_index=False)['Qtd.Física']
              .sum()
              .rename(columns={'Descrição': 'Descrição do Produto', 'Qtd.Física': 'Estoque'})
)


        consumo_join = consumo_medio.merge(df_estoque_agrupado, on=['Produto', 'Descrição do Produto'], how='left')

        def color_code(row):
            if pd.isna(row['Estoque']): return 'N/A'
            if row['Estoque'] > row['Consumo Médio Mensal']: return f"\U0001F7E2 {row['Estoque']}"
            elif row['Estoque'] == row['Consumo Médio Mensal']: return f"\U0001F7E1 {row['Estoque']}"
            else: return f"\U0001F534 {row['Estoque']}"

        consumo_join['Estoque Situação'] = consumo_join.apply(color_code, axis=1)

        st.subheader("Consumo Médio Mensal por Cliente e Estoque Atual")
        st.dataframe(consumo_join[['Cliente','Produto','Descrição do Produto','Consumo Médio Mensal','Estoque Situação']])

        st.success("Análise concluída com sucesso!")