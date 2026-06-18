from datetime import datetime
from pathlib import Path

import streamlit as st

from gerar_diagnostico_ml import run_pipeline

st.set_page_config(page_title="Diagnóstico Mercado Livre", layout="wide")

logo = Path(__file__).parent / "FIAT_LOGO.png"
header_cols = st.columns([1, 6])
with header_cols[0]:
    if logo.exists():
        st.image(str(logo), width=90)
with header_cols[1]:
    st.title("🛒 Diagnóstico Mercado Livre")
    st.caption(
        "Faça upload dos relatórios semanais de anúncios e baixe o Excel de "
        "diagnóstico e plano de ação pronto."
    )

st.markdown(
    """
    **Como usar**
    1. Exporte os relatórios semanais de desempenho de publicações do Mercado Livre
       (um arquivo `.xlsx` por semana) **mantendo o nome original** — ele contém o
       período (datas no formato `AAAA_MM_DD`), usado para montar a evolução.
    2. Envie um ou mais relatórios abaixo.
    3. (Opcional) Envie a planilha de projeção de meta e/ou o CSV de benchmark Brasil por PN.
    4. Clique em **Gerar diagnóstico** e baixe o Excel.
    """
)

report_files = st.file_uploader(
    "📥 Relatórios semanais de anúncios (.xlsx) — obrigatório",
    type=["xlsx"],
    accept_multiple_files=True,
)

with st.expander("➕ Arquivos opcionais (projeção de meta e benchmark Brasil)"):
    projection_file = st.file_uploader(
        "Projeção de Sell Out (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=False,
        key="projection",
    )
    market_file = st.file_uploader(
        "Benchmark Brasil por PN (.csv)",
        type=["csv"],
        accept_multiple_files=False,
        key="market",
    )

top_n = st.number_input(
    "Quantidade de itens nas abas de ranking e oportunidade",
    min_value=10,
    max_value=2000,
    value=200,
    step=10,
)

if st.button("⚙️ Gerar diagnóstico", type="primary"):
    if not report_files:
        st.warning("Envie pelo menos um relatório semanal (.xlsx) para continuar.")
        st.stop()

    try:
        with st.spinner("Processando relatórios e montando o Excel..."):
            buffer, kpis = run_pipeline(
                report_files=report_files,
                projection_file=projection_file,
                market_file=market_file,
                top_n=int(top_n),
            )
    except ValueError as exc:
        st.error(f"Não foi possível processar os relatórios: {exc}")
        st.stop()
    except Exception as exc:  # noqa: BLE001 - feedback amigável ao usuário
        st.error(
            "Erro ao processar os arquivos. Verifique se são os relatórios de "
            f"desempenho de publicações exportados do Mercado Livre.\n\nDetalhe: {exc}"
        )
        st.stop()

    st.success("Diagnóstico gerado com sucesso!")

    st.subheader(f"Resumo — período {kpis['periodo']}")
    row1 = st.columns(4)
    row1[0].metric("Receita bruta ML", f"R$ {kpis['receita_total']:,.0f}".replace(",", "."))
    row1[1].metric("Pedidos", f"{kpis['pedidos']:,}".replace(",", "."))
    row1[2].metric("Unidades", f"{kpis['unidades']:,}".replace(",", "."))
    row1[3].metric("Anúncios únicos", f"{kpis['anuncios_unicos']:,}".replace(",", "."))

    row2 = st.columns(4)
    row2[0].metric("Conversão média", f"{kpis['conversao_media'] * 100:.2f}%")
    row2[1].metric("Ticket médio", f"R$ {kpis['ticket_medio']:,.0f}".replace(",", "."))
    row2[2].metric("Semanas analisadas", f"{kpis['semanas']}")
    row2[3].metric("Benchmark PN", "Sim" if kpis["tem_benchmark"] else "Não enviado")

    file_name = f"Diagnostico_Mercado_Livre_{datetime.now():%Y%m%d_%H%M}.xlsx"
    st.download_button(
        label="⬇️ Baixar Excel do diagnóstico",
        data=buffer,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
