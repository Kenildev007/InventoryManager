#!/usr/bin/env bash
# Serves the app locally on http://localhost:$PORT
cd "$(dirname "$0")"
PORT="${PORT:-8080}"
echo "Inventory Console → http://localhost:$PORT"
( sleep 1; open "http://localhost:$PORT" 2>/dev/null || xdg-open "http://localhost:$PORT" 2>/dev/null ) &
python3 -m http.server "$PORT"
