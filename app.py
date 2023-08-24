#Importar bibliotecas
import streamlit as st
import pandas as pd
import io

import os

st.set_page_config("Fiat Peças BSB", "⚙️")
st.image('./FIAT_LOGO.png', width=100)

#Título
st.markdown('## Ofertas de Peças - Rede de Concessionárias do Regional Brasília')
st.write('Este site é uma plataforma de compartilhamento de ofertas entre concessionárias. As concessionárias participantes podem enviar a lista de peças que desejam ofertar ao seu Consultor de Pós-Vendas do Regional Brasília, dando visibilidade aos seus itens para toda a rede do Regional Brasília, já com o preço de oferta. A ferramenta tem como objetivo aumentar o sell-out das concessionárias, permitindo que itens sejam adquiridos entre as próprias concessionárias, caso a oferta seja conveniente. \nA adesão é livre. Todas que quiserem expor suas ofertas podem participar, bastando enviar a lista ao seu CPV.')
st.error('')

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

st.error('')

output = io.BytesIO()
with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    df_pecas.to_excel(writer, sheet_name='Sheet1', index=False)

st.download_button(
    label="Download da lista completa",
    data=output.getvalue(),
    file_name="OPORTUNIDADES_SEM_GIRO_DEALER.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.error('')
st.write('Este site é uma PoC (Prova de Conceito), não utlize este site oficialmente.')
st.write('Para dúvidas ou sugestões, falar com Marcos Feitosa (marcos.feitosa@stellantis.com) ou Bruno Schmeisck (bruno.schmeisck@stellantis.com).')
