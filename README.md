Gerador de gráfico ponte
========================

Como usar
1. Abra o arquivo index.html no navegador (Chrome ou Edge). Precisa de internet
   para carregar a leitura de Excel (SheetJS) e a fonte Carlito.
2. Importe a planilha (.xlsx ou .csv).
3. Confira as linhas, os degraus marcados e o aviso de fechamento.
4. Baixe o gráfico em PNG ou SVG.

Formato esperado da planilha
- Primeira linha: na coluna A um título qualquer; nas demais, (total) e as empresas.
- Linhas abaixo: cenários (ex.: "6 + 6", "FCST 12/10"), com os valores em R$ mil.
- Colunas que somam outras
  vêm desmarcadas para não contar a receita duas vezes.

Arquivos
- index.html  estrutura da página
- style.css   visual da página
- app.js      leitura da planilha, cálculo e desenho do gráfico

Os dados não saem do seu computador: tudo é processado no navegador.
