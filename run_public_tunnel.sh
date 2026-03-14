#!/usr/bin/env bash
set -euo pipefail

if [[ ! -f ".venv/bin/activate" ]]; then
  echo "Missing .venv. Create it first:"
  echo "  python3 -m venv .venv && . .venv/bin/activate && pip install -r requirements.txt"
  exit 1
fi

CLOUDFLARED_BIN="${CLOUDFLARED_BIN:-}"
if [[ -z "${CLOUDFLARED_BIN}" ]]; then
  if [[ -x ".cloudflared/cloudflared" ]]; then
    CLOUDFLARED_BIN=".cloudflared/cloudflared"
  elif command -v cloudflared >/dev/null 2>&1; then
    CLOUDFLARED_BIN="cloudflared"
  fi
fi

if [[ -z "${CLOUDFLARED_BIN}" ]]; then
  echo "Missing 'cloudflared'."
  echo "Either install it system-wide, or download it into: .cloudflared/cloudflared"
  exit 1
fi

source ".venv/bin/activate"

HOST="${HOST:-127.0.0.1}"
PORT="${PORT:-8000}"
export HOST PORT

python3 webapp.py >/tmp/wordd_webapp.log 2>&1 &
APP_PID=$!

cleanup() {
  kill "$APP_PID" >/dev/null 2>&1 || true
  wait "$APP_PID" >/dev/null 2>&1 || true
}
trap cleanup EXIT

sleep 1
echo "Local app: http://${HOST}:${PORT}"
echo "Starting Cloudflare Quick Tunnel (Ctrl+C to stop)..."
"${CLOUDFLARED_BIN}" tunnel --url "http://${HOST}:${PORT}"
