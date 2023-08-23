#Importar bibliotecas
import streamlit as st
import pandas as pd
import os

st.set_page_config("Streamlit Components Hub", "🎪", layout="wide")

#Título
st.markdown('## Ofertas de Peças - Rede de Concessionárias do Regional Brasília')

#carregar o dataframe
pasta_atual = os.getcwd()
arquivo = os.path.join(pasta_atual, "OPORTUNIDADES SEM GIRO DEALER.xlsx")
df_pecas = pd.read_excel(arquivo, dtype={"DESENHO": str})

#barra de pesquisa
part_number = st.text_input('Informe o número do desenho: ')

condicao = st.button('Pesquisar')

if condicao == True:
    #Retorno é uma lista das ofertas daquele item
    resultado = df_pecas.loc[df_pecas['DESENHO'] == part_number]

    st.dataframe(resultado)

    #opção de download da lista completa


st.download_button('Download da lista completa', arquivo)
