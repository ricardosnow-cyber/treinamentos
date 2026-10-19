# Controle de custos de IA

## Rotas
- `GET /api/costs`: devolve até 500 registros recentes e o resumo financeiro.
- `POST /api/costs`: registra um consumo/estimativa. O backend calcula os custos e ignora qualquer total enviado pelo cliente.
- `GET /custos`: abre o painel de custos.

## Tarifas configuradas
- `heygen`: US$ 0,05 por segundo de vídeo.
- `openai-image`: US$ 0,04 por imagem.
- `heygen-ghibli`: combinação de US$ 0,04 por imagem mais US$ 0,05 por segundo de vídeo.

Exemplo de payload para registrar o fluxo combinado:
```json
{
  "provider": "heygen-ghibli",
  "durationSeconds": 30,
  "imageCount": 1,
  "label": "Treinamento NR",
  "prompt": "Descrição do vídeo",
  "requestId": "id-unico-da-geracao",
  "videoId": "id-do-video",
  "status": "estimated"
}
```
O custo calculado para 30 segundos é US$ 1,54. Para evitar duplicidade em tentativas repetidas, envie o mesmo `requestId` para a mesma geração.

## Persistência e limite diário
Os registros ficam em `storage/costs.json`, com retenção dos últimos 5.000 lançamentos. O limite de alerta é US$ 5/dia, calculado no fuso `COST_TIME_ZONE` (padrão `America/Sao_Paulo`). O painel exibe o toast **“Limite diário atingido”** quando o total diário ultrapassa US$ 5. O alerta não bloqueia operações.

## Importante antes de produção
O repositório original não tinha uma rota de geração `/api/heygen-ghibli`, autenticação nem banco de dados. Este commit acrescenta o ledger de custos e sua API; para contabilizar automaticamente vídeos reais, a futura rota de geração deve chamar o registro de custo no backend usando o `requestId` da geração, e atualizar a estimativa quando a duração real for conhecida. O painel não deve ser considerado prova de custos faturados pelo provedor. A API de escrita também deve ser protegida com autenticação/autorização antes de expor o app publicamente; armazenamento em arquivo pode não ser persistente em hospedagens efêmeras.
