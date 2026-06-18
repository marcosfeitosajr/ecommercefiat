# ecommercefiat

Ferramentas internas em [Streamlit](https://streamlit.io/) para apoiar a operação de
e-commerce e o trabalho do dia a dia (Stellantis / Feitozzi).

## Apps disponíveis

### 1. Diagnóstico Mercado Livre — `diagnostico_ml_app.py`
Recebe os relatórios semanais de **desempenho de publicações** do Mercado Livre e gera um
Excel completo de **diagnóstico e plano de ação** (resumo, evolução semanal, metas,
rankings de receita/unidades, gargalos de conversão, tráfego sem venda, inativos com
demanda, vencedores a escalar, base consolidada e, opcionalmente, benchmark Brasil por PN).

O usuário só faz upload dos arquivos e baixa o Excel — sem rodar código.

**Rodar localmente:**
```bash
pip install -r requirements.txt
streamlit run diagnostico_ml_app.py
```

**Entradas:**
- **Obrigatório:** um ou mais relatórios semanais `.xlsx` exportados do Mercado Livre.
  Mantenha o **nome original** do arquivo: ele precisa conter o período no formato
  `AAAA_MM_DD` (ex.: `...2026_03_01_2026_03_07.xlsx`). O período é usado para montar a
  evolução semanal. Cada relatório é lido na primeira aba, com o cabeçalho na linha 6.
- **Opcional:** planilha de **Projeção de Sell Out** (`.xlsx`) — define a meta de dezembro
  usada na aba de metas. Se não enviada, usa valores padrão.
- **Opcional:** **Benchmark Brasil por PN** (`.csv`) — habilita as abas de comparação por
  PN (participação e ticket Florença vs. Brasil). Se não enviado, essas abas são puladas.

### 2. Agrupador de Lançamentos Bancários — `app.py`
Recebe um PDF de extrato e agrupa os lançamentos por data e descrição, exportando CSV.

```bash
streamlit run app.py
```

## Motor de diagnóstico via linha de comando

O `diagnostico_ml_app.py` reutiliza o motor `gerar_diagnostico_ml.py`, que também pode ser
executado direto no terminal a partir de uma pasta de relatórios:

```bash
python gerar_diagnostico_ml.py \
  --reports-dir "Relatório de anúncios" \
  --projection-file "Projeção Sell out.xlsx" \
  --market-csv "Sell out Ecommerce Detalhado por PN.csv" \
  --output "Diagnostico_Mercado_Livre.xlsx"
```

## Roadmap

Estrutura preparada para evoluir o Diagnóstico ML em produto: separação entre motor
(`gerar_diagnostico_ml.py`) e interface (`diagnostico_ml_app.py`) facilita adicionar, no
futuro, login/multiusuário, cobrança e separação de dados por cliente.
