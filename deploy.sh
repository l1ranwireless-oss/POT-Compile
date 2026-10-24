#!/usr/bin/env sh
set -eu

ROOT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
cd "$ROOT_DIR"

if ! command -v docker >/dev/null 2>&1; then
  printf '%s\n' 'Docker tidak ditemukan. Install Docker Engine beserta Docker Compose plugin terlebih dahulu.' >&2
  exit 1
fi

docker compose version >/dev/null
if [ -z "${CMDB_API_TOKEN:-}" ]; then
  printf '%s\n' 'CMDB_API_TOKEN wajib diatur (minimal 32 karakter). Contoh: export CMDB_API_TOKEN="$(openssl rand -hex 32)"' >&2
  exit 1
fi
docker compose up -d --build

printf 'Aplikasi tersedia di http://localhost:%s\n' "${PORT:-8080}"