# Painel de fundos — nossos × peers

Um arquivo (`painel.py`) que baixa os dados públicos da CVM, ANBIMA, Banco Central e B3, processa os fundos das listas
`NOSSOS_FUNDOS` e `PEERS` (no topo do arquivo) e grava **`painel_fundos.html`**: um HTML único, que abre com dois cliques, sem servidor.

```bash
pip install pandas numpy requests pyarrow openpyxl plotly fastapi uvicorn
python painel.py              # atualiza tudo e gera painel_fundos.html
python painel.py servidor     # painel antigo em localhost
python painel.py excel        # look-through em Excel
```

**Abas do HTML**
- **Visão geral** — retorno em % do CDI (mês, 3 meses, ano, 12 e 24 meses), vol, spread CDI+ e duration da última carteira aberta, maiores categorias; risco × retorno e spread × duration.
- **Fundo** — alocação por categoria, spread e duration (carteira ou por categoria), retorno mensal × CDI com atribuição, carteira dos últimos meses abertos, melhores ativos em 12 meses (retorno ou contribuição em bps) e mudanças na carteira.
- **Performance de ativos** — todos os ativos de todas as carteiras e meses: retorno, % do CDI, vol, Sharpe, nº de fundos, % do PL, spread, duration e se temos. Clique: retorno total mensal × CDI e curva acumulada, % do PL (o fundo aberto + os nossos, ou nossos + média dos peers), spread e % do PU par, taxa indicativa e % do PU par da ANBIMA.
- **Exposições vs. peers** — papéis e grupos econômicos que os peers têm e nós não.

**Retorno do ativo (comparável ao CDI)**: retorno total do mês = preço + juros/amortização pagos, sem hedge nem resíduo do fundo.
Debêntures: preço do fundo + pagamento pelo PU par do SND. FIDC: rentabilidade oficial da série (CVM). LF/LFSN/CDB: preço (não pagam no meio).
CRI/CRA/NC/CCB: o fluxo não é público; preço abaixo do esperado = pagamento → carrego + curva; queda forte que não volta fica marcada (amortização ou perda?).
Meses com a carteira confidencial usam a última carteira aberta (marcados com *).

Só entra carteira pública: um mês conta como aberto quando a parte confidencial é menor que 5% do PL (a CVM permite esconder posições por até ~90 dias).
Retorno pela cota diária da CVM (confere com o Mais Retorno). A ANBIMA só publica ~5 semanas de marcação diária: o histórico local cresce a cada execução.
Dados em `%LOCALAPPDATA%/cvm_cache` no Windows (fora do OneDrive; ~2–3 GB na primeira vez) ou `dados_painel/` ao lado do arquivo.
