#!/usr/bin/env bash
set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="${OPENCLAW_STOCK_QUANT_VENV:-$HOME/.openclaw/workspace/.venvs/stock_quant}"
PYTHON_BIN="${OPENCLAW_PYTHON:-$HOME/.local/bin/python3.11}"
UV_BIN="${OPENCLAW_UV:-$HOME/.local/bin/uv}"

cd "$PROJECT_DIR"

if [[ ! -x "$VENV_DIR/bin/python" ]]; then
  mkdir -p "$(dirname "$VENV_DIR")"
  "$UV_BIN" venv "$VENV_DIR" --python "$PYTHON_BIN"
fi

"$UV_BIN" pip install --python "$VENV_DIR/bin/python" -r "$PROJECT_DIR/requirements.txt" >/dev/null

export OPENCLAW_HEADLESS=1
"$VENV_DIR/bin/python" "$PROJECT_DIR/weekly_quant.py"
