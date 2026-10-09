#!/bin/sh
set -eu

export PORT="${PORT:-10000}"
printf 'Launching Drafter API on port %s\n' "$PORT"
command -v xauth
command -v xvfb-run

exec xvfb-run -a sh -c 'printf "Virtual display: %s\n" "$DISPLAY"; exec python -m uvicorn api:app --host 0.0.0.0 --port "$PORT" --proxy-headers --forwarded-allow-ips="*"'