import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Agrupador de Lançamentos", layout="wide")
st.title("📊 Agrupador de Lançamentos Bancários")

uploaded_file = st.file_uploader("📥 Faça upload do seu arquivo PDF", type=["pdf"])

if uploaded_file:
    bytes_data = uploaded_file.read()
    text = ""
    with pdfplumber.open(BytesIO(bytes_data)) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"

    # Expressão regular para capturar data, descrição e valor
    pattern = re.compile(r"(\d{2}-\d{2}-\d{4})\s+(.+?)\s+\d+\s+R\$ ([\-\d\.,]+)", re.DOTALL)
    records = []
    for date, desc, raw_val in pattern.findall(text):
        # Normaliza e converte valor
        valor = float(raw_val.replace('.', '').replace(',', '.'))
        # Limpa descrição
        desc = desc.replace("\n", " ").strip()
        # Aplica regra especial
        if desc.startswith("Pagamento com Código QR Pix") or desc == "Liberação de dinheiro":
            desc = "Receita por produtos"
        records.append((date, desc, valor))

    # Cria DataFrame e agrupa
    df = pd.DataFrame(records, columns=["Data", "Descrição", "Valor"] )
    grouped = df.groupby(["Data", "Descrição"], as_index=False).sum()
    # Formata valor com duas casas decimais
    grouped["Valor Total (R$)"] = grouped["Valor"].map(lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    result = grouped[["Data", "Descrição", "Valor Total (R$)"]]

    # Exibe e permite download
    st.subheader("📋 Resultado Agrupado")
    st.dataframe(result, use_container_width=True)
    csv = result.to_csv(index=False, sep=';')
    st.download_button(
        label="⬇️ Baixar CSV",
        data=csv,
        file_name="agrupado_lancamentos.csv",
        mime="text/csv"
    )
