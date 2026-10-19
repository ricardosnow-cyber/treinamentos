# Controle de custos e geração HeyGen

## Rotas protegidas
Todas as rotas abaixo exigem HTTP Basic Auth, configurado com `ADMIN_USER` e `ADMIN_PASSWORD` no servidor:
- `GET /custos`: painel de custos e geração de vídeos.
- `GET /api/costs`: até 500 lançamentos recentes e resumo financeiro.
- `POST /api/costs`: registra estimativas de custos (não aceita o custo calculado pelo cliente como fonte de verdade).
- `POST /api/heygen-ghibli` e `POST /api/videos/generate`: envia um roteiro ao HeyGen e registra o custo estimado automaticamente.
- `GET /api/heygen-ghibli/:videoId`: consulta status, link do vídeo e atualiza o custo quando a duração real for disponibilizada.

Sem `ADMIN_USER` e `ADMIN_PASSWORD`, as rotas administrativas respondem HTTP 503. As credenciais e a chave do provedor não devem ser colocadas no HTML, no frontend, no Git ou em variáveis públicas de build. O servidor usa `HEYGEN_API_KEY` no cabeçalho `X-Api-Key`.

## Variáveis de ambiente
Configure no ambiente do servidor (não comite valores reais):
- `ADMIN_USER`: usuário administrativo.
- `ADMIN_PASSWORD`: senha administrativa forte e única.
- `HEYGEN_API_KEY`: chave privada da API HeyGen.
- `HEYGEN_AVATAR_ID`: avatar padrão (opcional se informado por requisição).
- `HEYGEN_VOICE_ID`: voz padrão (opcional se informada por requisição).
- `COST_TIME_ZONE=America/Sao_Paulo`: fuso de agrupamento do limite diário.

A autenticação usa Basic Auth via HTTPS. Não publique este serviço sem HTTPS, credenciais fortes e armazenamento persistente. O Basic Auth é um controle simples para o painel administrativo; para múltiplos usuários ou uma plataforma pública, prefira sessões com MFA ou um provedor de identidade.

## Tarifas e estimativas
- `heygen`: US$ 0,05 por segundo de vídeo.
- `openai-image`: US$ 0,04 por imagem.
- `heygen-ghibli`: combinação de US$ 0,04 por imagem mais US$ 0,05 por segundo de vídeo.

A geração HeyGen estima a duração a partir do roteiro (aprox. 2,5 palavras por segundo) ou usa `estimated_duration`. O valor inicial é estimado; quando o status do HeyGen retornar a duração real, o registro é atualizado. Se a geração falhar, o custo estimado desse lançamento é zerado. Valores são estimativas configuradas e não substituem a fatura real do provedor.

Exemplo:
```json
{
  "requestId": "id-unico-da-geracao",
  "title": "Treinamento NR",
  "script": "Texto da narração do avatar.",
  "avatar_id": "ID_DO_AVATAR",
  "voice_id": "ID_DA_VOZ",
  "aspect_ratio": "16:9",
  "estimated_duration": 30
}
```
Envie `requestId` também no cabeçalho `Idempotency-Key` para que repetições do mesmo pedido possam ser reconhecidas. O servidor não retorna a chave de API ao navegador.

## Limite e persistência
Os lançamentos ficam em `storage/costs.json`, com retenção dos últimos 5.000 registros. O alerta é US$ 5/dia no fuso configurado. O painel mostra “Limite diário atingido” após ultrapassar o valor; o alerta não bloqueia geração. Em hospedagens efêmeras, o arquivo pode ser apagado durante redeploy/restart, portanto configure volume persistente ou banco de dados antes de produção.

## Escopo
O endpoint gera vídeo de avatar HeyGen a partir de um roteiro. Ele não gera automaticamente imagens de fundo via OpenAI; essa integração exige uma etapa separada de geração e hospedagem segura da imagem. A API de geração HeyGen utilizada é `POST https://api.heygen.com/v2/video/generate`; o status é consultado via `GET https://api.heygen.com/v1/video_status.get?video_id=...`.
