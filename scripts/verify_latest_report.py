#!/usr/bin/env python3
from __future__ import annotations
import glob, os
from datetime import datetime
import pandas as pd

reports = glob.glob('output/weekly_report_*.xlsx')
if not reports:
    raise SystemExit('未找到 output/weekly_report_*.xlsx')
report = max(reports, key=os.path.getmtime)
spot_files = glob.glob('cache/tushare/stock_spot_*.csv')
spot = max(spot_files, key=os.path.getmtime) if spot_files else None
allscore = pd.read_excel(report, sheet_name='全部评分')
hold = pd.read_excel(report, sheet_name='当前持仓检查')
ranks = pd.to_numeric(allscore.get('股票池初筛排名'), errors='coerce')
etf_bad = hold[(hold.get('标的类型').eq('ETF')) & (hold.get('操作建议').astype(str).str.contains('卖|减', na=False))] if '标的类型' in hold and '操作建议' in hold else pd.DataFrame()
print(f'报告: {report} ({datetime.fromtimestamp(os.path.getmtime(report)):%Y-%m-%d %H:%M:%S})')
if spot:
    sp = pd.read_csv(spot)
    print(f'Tushare快照: {spot} 行数={len(sp)} PE_TTM有效={sp.get("PE_TTM").notna().sum() if "PE_TTM" in sp else "?"}')
print(f'全部评分: 行数={len(allscore)} PE_TTM有效={allscore.get("PE_TTM").notna().sum() if "PE_TTM" in allscore else "?"} 排名范围={int(ranks.min()) if ranks.notna().any() else "?"}-{int(ranks.max()) if ranks.notna().any() else "?"}')
print(f'股票池三段命中: 1-80={((ranks>=1)&(ranks<=80)).sum()} 200-220={((ranks>=200)&(ranks<=220)).sum()} 500-520={((ranks>=500)&(ranks<=520)).sum()}')
print(f'ETF错误卖出/减仓建议数: {len(etf_bad)}')
print('Top10:')
cols=[c for c in ['排名','代码','名称','最新价','PE_TTM','估值得分','趋势动量分','成交活跃分','风险分','综合得分','股票池初筛排名','操作建议'] if c in allscore.columns]
print(allscore[cols].head(10).to_string(index=False))
