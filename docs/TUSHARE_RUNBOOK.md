# Tushare 独立运行说明

## 当前口径

- 数据源：Tushare。
- 只取 A 股股票，不取 ETF/基金行情。
- 不取盘中价格，使用 Tushare 盘后 `daily` + `daily_basic`。
- PE_TTM 使用 `daily_basic.pe_ttm`。
- 历史行情优先使用 Tushare `pro_bar`（按 `config.yaml` 的 `adjust: qfq`），失败时回退 `daily`。
- 初始股票池：总市值排名 `1-80` + `200-220` + `500-520`。
- ETF 持仓按长期持有处理：不因未取 ETF 行情给卖出/减仓建议。

## 文件结构

- `weekly_quant.py`：核心选股、打分、报表生成逻辑。
- `scripts/tushare_runner.py`：Tushare 数据适配层；保持核心算法不变，只替换数据获取。
- `run_tushare_once.sh`：前台运行一次，适合调试/人工验收。
- `start_tushare_background.sh`：后台启动一次，日志写入 `output/logs/`，适合“开始运行后不用介入”。
- `scripts/verify_latest_report.py`：验收最新报告，检查 PE_TTM、股票池排名段、ETF 是否误给卖出建议。
- `.tushare_token`：本地 token 文件，已加入 `.gitignore`，不要提交或外发。

## 常用命令

前台跑一次：

```bash
bash run_tushare_once.sh
```

后台启动一次：

```bash
bash start_tushare_background.sh
```

验收最新报告：

```bash
python scripts/verify_latest_report.py
```

## 最近一次验收

2026-06-30 22:16 验收通过：

- Tushare 快照日期：20260630
- A 股快照：4382 只
- PE_TTM 有效：3257 只
- 初始股票池：120 只
- 最终全部评分：50 行，PE_TTM 全有效
- 排名段命中：1-80 段 34 只、200-220 段 8 只、500-520 段 8 只
- ETF 错误卖出/减仓建议数：0
