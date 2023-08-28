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

def main():
    # Carregar o dataframe
    pasta_atual = os.getcwd()
    arquivo = os.path.join(pasta_atual, "OPORTUNIDADES SEM GIRO DEALER.xlsx")
    df_pecas = pd.read_excel(arquivo, dtype={"DESENHO": str})

    # Barra de pesquisa
    part_number = st.text_input('Informe o número do desenho: ')

    # Botão de pesquisa
    if st.button('Pesquisar'):
        lista_pecas = df_pecas[df_pecas['DESENHO'] == part_number]
        
        if not resultado.empty:
            st.dataframe(lista_pecas)
            st.download_button('Download da lista de' + part_number, data=resultado.to_excel(index=False), file_name='resultado_' + part_number + '.xlsx')
        else:
            st.warning("Nenhum resultado encontrado para o número de desenho informado.")

    # Botão de download
    st.download_button('Download da lista completa', data=df_pecas, file_name='lista_completa.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == "__main__":
    main()

st.error('')
st.write('Este site é uma PoC (Prova de Conceito), não utlize este site oficialmente.')
st.write('Para dúvidas ou sugestões, falar com Marcos Feitosa (marcos.feitosa@stellantis.com) ou Bruno Schmeisck (bruno.schmeisck@stellantis.com).')
