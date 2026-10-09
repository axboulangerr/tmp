#!/bin/sh
set -eu

export PORT="${PORT:-10000}"
export DISPLAY="${DISPLAY:-:99}"
printf 'Launching Drafter API on port %s\n' "$PORT"
command -v Xvfb
Xvfb "$DISPLAY" -screen 0 1440x1100x24 -nolisten tcp -ac &
printf 'Started Xvfb on %s\n' "$DISPLAY"

exec python -m uvicorn api:app --host 0.0.0.0 --port "$PORT" --proxy-headers --forwarded-allow-ips="*"