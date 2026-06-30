#!/usr/bin/env bash
# Backward-compatible entrypoint. Prefer run_tushare_once.sh for foreground runs.
set -euo pipefail
cd "$(dirname "$0")"
exec bash run_tushare_once.sh "$@"
