#!/usr/bin/env sh
set -eu

ROOT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
cd "$ROOT_DIR"

if ! command -v docker >/dev/null 2>&1; then
  printf '%s\n' 'Docker tidak ditemukan. Install Docker Engine beserta Docker Compose plugin terlebih dahulu.' >&2
  exit 1
fi

docker compose version >/dev/null
mkdir -p dist
cp index.html dist/index.html
docker compose up -d --build

printf 'Aplikasi tersedia di http://localhost:%s\n' "${PORT:-8080}"