import pandas as pd
import unicodedata

def normalize_columns(df):
    """
    Normaliza nomes de colunas: remove espaços, converte para lowercase e remove acentuação.
    """
    df = df.copy()
    df.columns = (
        df.columns.str.strip()
                  .str.lower()
                  .map(lambda x: unicodedata.normalize('NFKD', x)
                                    .encode('ascii', 'ignore')
                                    .decode('ascii'))
    )
    return df


def standardize_column_names(df):
    """
    Renomeia colunas normalizadas para os títulos esperados, incluindo o nome do representante.
    """
    rename_map = {
        # Datas
        'data': 'Data',
        'dt.faturam': 'Data',
        # Clientes
        'cliente': 'Cliente',
        'razao social': 'Cliente',
        # Quantidades
        'quantidade': 'Quantidade',
        'qtd.item': 'Quantidade',
        # Valores
        'valor unitario': 'Valor Unitário',
        'vlr.un': 'Valor Unitário',
        # Produtos
        'produto': 'Produto',
        'descricao do produto': 'Descrição do Produto',
        'desc.prod': 'Descrição do Produto',
        # Representante (campo de descrição)
        'desc.repr/prep': 'Representante'
    }
    # Aplica renomeação apenas para colunas presentes
    df = df.rename(columns={col: new for col, new in rename_map.items() if col in df.columns})
    # Remove colunas duplicadas após renomeação
    df = df.loc[:, ~df.columns.duplicated()]
    return df


def prepare_df(df):
    """
    Normaliza colunas, padroniza nomes e converte 'Data' para datetime.
    """
    df = normalize_columns(df)
    df = standardize_column_names(df)
    if 'Data' in df.columns:
        df['Data'] = pd.to_datetime(df['Data'], dayfirst=True, errors='coerce')
    return df


def analyze(df_fat, df_orc, df_ped):
    """
    Executa análise de dados comerciais:
    - Identifica clientes inativos e em queda
    - Calcula taxas de conversão
    - Gera evolução mensal de vendas
    """
    # Preparar DataFrames
    df_fat = prepare_df(df_fat)
    df_orc = prepare_df(df_orc)
    df_ped = prepare_df(df_ped)

    # Valida coluna Data
    if 'Data' not in df_fat.columns or df_fat['Data'].isna().all():
        raise ValueError("A coluna 'Data' não foi detectada no faturado.XLSX.")

    # Período de referência (últimos 6 meses)
    last_date = df_fat['Data'].max()
    threshold = last_date - pd.DateOffset(months=6)
    recent = df_fat[df_fat['Data'] >= threshold]

    # Clientes inativos
    all_clients = pd.Series(df_fat['Cliente'].unique(), name='Cliente')
    inactive_clients = all_clients[~all_clients.isin(recent['Cliente'])].to_frame()

    # Clientes com queda de volume
    volumes = df_fat.groupby('Cliente')['Quantidade'].sum()
    recent_volumes = recent.groupby('Cliente')['Quantidade'].sum().reindex(volumes.index, fill_value=0)
    declining_clients = volumes[volumes > recent_volumes].index.to_series(name='Cliente').to_frame()

    # Taxas de conversão
    conv_orc_to_ped = len(df_ped) / len(df_orc) if len(df_orc) else 0
    conv_ped_to_fat = len(df_fat) / len(df_ped) if len(df_ped) else 0
    conversion_rates = pd.DataFrame({
        'Etapa': ['Orçamento→Pedido', 'Pedido→Venda'],
        'Taxa': [conv_orc_to_ped, conv_ped_to_fat]
    })

    # Evolução mensal de vendas
    sales_evolution = (
        df_fat.groupby(pd.Grouper(key='Data', freq='M'))['Quantidade']
              .sum()
              .reset_index()
    )

    return {
        'inactive_clients': inactive_clients,
        'declining_clients': declining_clients,
        'conversion_rates': conversion_rates,
        'sales_evolution': sales_evolution
    }