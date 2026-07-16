# DASH-COMPRAS

Dashboard estática de acompanhamento operacional da planilha **Notas fiscais sem pedido**.

## Publicação

A publicação principal usa `index.html` para hospedagem estática.

## Arquivos

- `index.html` — dashboard gerada pronta para publicação
- `scripts/gerar-dashboard-notas-sem-pedido.js` — leitura, normalização, deduplicação e geração da dashboard
- `scripts/validate-dashboard.js` — reconcilia fonte, abas, categorias, duplicidades e campos críticos antes da publicação
- `update-dashboard.ps1` — atualização automática com validação obrigatória, commit e push
