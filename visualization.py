import streamlit as st
import matplotlib.pyplot as plt

def plot_sales_evolution(df):
    fig, ax = plt.subplots()
    ax.plot(df['Data'], df['Quantidade'])
    ax.set_xlabel('Data')
    ax.set_ylabel('Quantidade')
    ax.set_title('Evolução de Vendas Mensal')
    st.pyplot(fig)

def plot_conversion_rates(df):
    fig, ax = plt.subplots()
    ax.bar(df['Etapa'], df['Taxa'])
    ax.set_ylim(0, 1)
    ax.set_ylabel('Taxa')
    ax.set_title('Taxas de Conversão')
    st.pyplot(fig)
