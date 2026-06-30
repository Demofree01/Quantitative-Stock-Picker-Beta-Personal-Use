#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"
mkdir -p output/logs
stamp="$(date +%Y%m%d_%H%M%S)"
log="output/logs/tushare_${stamp}.log"
pidfile="output/logs/tushare_latest.pid"
nohup bash run_tushare_once.sh > "$log" 2>&1 &
pid=$!
printf '%s\n' "$pid" > "$pidfile"
printf 'Started Tushare quant run: pid=%s log=%s\n' "$pid" "$log"
