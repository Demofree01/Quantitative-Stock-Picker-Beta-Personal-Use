#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"
VENV="/home/admin/.openclaw/workspace/.venvs/stock_quant"
if [ -x "$VENV/bin/python" ]; then
  PY="$VENV/bin/python"
else
  PY="python3"
fi
if command -v uv >/dev/null 2>&1; then
  uv pip install --python "$PY" 'tushare>=1.4.21' pandas numpy openpyxl PyYAML requests >/dev/null
else
  "$PY" -m ensurepip --upgrade >/dev/null 2>&1 || true
  "$PY" -m pip install -q 'tushare>=1.4.21' pandas numpy openpyxl PyYAML requests >/dev/null
fi
exec "$PY" scripts/tushare_runner.py "$@"
