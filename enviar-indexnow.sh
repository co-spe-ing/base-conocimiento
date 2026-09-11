#!/usr/bin/env bash
# Avisa a Bing que las páginas son nuevas o cambiaron (IndexNow).
# Ejecutar DESPUÉS de publicar en GitHub Pages, no antes.
set -euo pipefail
CLAVE="d739ebb1888b38b517416c406a7ad8ff"
HOST="co-spe-ing.github.io"
LISTA=$(python3 - <<'EOF'
import json
urls=[l.strip() for l in open('urls.txt') if l.strip()]
print(json.dumps(urls))
EOF
)
curl -sS -X POST "https://api.indexnow.org/IndexNow" \
  -H "Content-Type: application/json; charset=utf-8" \
  -d "{\"host\":\"$HOST\",\"key\":\"$CLAVE\",\"keyLocation\":\"https://$HOST/base-conocimiento/$CLAVE.txt\",\"urlList\":$LISTA}" \
  -w "\nHTTP %{http_code}\n"
