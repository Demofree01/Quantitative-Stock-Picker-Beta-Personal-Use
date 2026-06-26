# Quantitative Stock Picker for China A-Shares (Personal Beta)

> **AI-assisted production notice**: This project was specified by the user and organized, refactored, debugged, and documented with assistance from OpenClaw/Alice. The output is for personal research and manual decision-making only. It is not investment advice.

## 1. Purpose

This is a weekly quantitative stock screening and rebalancing helper for China A-shares.

- It does not place trades automatically.
- It does not connect to brokerage APIs.
- It does not guarantee returns.
- It generates an Excel report for manual review.

## 2. Features

- Builds a China A-share candidate universe.
- Fetches historical prices and latest prices.
- Scores stocks using momentum, trend, liquidity, risk, valuation, and quality factors.
- Generates Excel sheets such as Top 10 candidates, current holding review, actual buy plan, and rebalance suggestions.
- Reads `holdings.csv` as the manual baseline account file.
- Optionally reads `交割单.xlsx` or legacy `trades.csv` to roll forward current holdings and cash.
- Runs in OpenClaw/headless environments.

## 3. Installation

```bash
python3 -m pip install -r requirements.txt
```

For OpenClaw or OneDrive/rclone-mounted directories, use:

```bash
bash run_openclaw.sh
```

The script uses `~/.openclaw/workspace/.venvs/stock_quant` as a Python 3.11 virtual environment and sets `OPENCLAW_HEADLESS=1`.

## 4. Usage

### 4.1 Prepare holdings

Create `holdings.csv`:

```csv
code,name,asset_type,shares,cost_price
605358,立昂微,STOCK,100,74.712
CASH,现金,CASH,5000,1
```

Columns:

- Stock/ETF row: `code,name,asset_type,shares,cost_price`
- Cash row: `CASH,现金,CASH,cash_amount,1`
- `asset_type` can be `STOCK`, `ETF`, or `CASH`.

### 4.2 Prepare trade ledger (optional)

You may maintain `交割单.xlsx`. The program first reads `holdings.csv`, then rolls holdings and cash forward using the trade ledger, and writes `自动持仓.csv` for review.

If `交割单.xlsx` is missing, the legacy `trades.csv` is still supported.

### 4.3 Run

Normal environment:

```bash
python3 weekly_quant.py
```

OpenClaw/headless environment:

```bash
bash run_openclaw.sh
```

### 4.4 Output

Reports are written to:

```text
output/weekly_report_YYYYMMDD.xlsx
```

Common sheets:

- `系统摘要`
- `调仓建议`
- `实际买入计划`
- `候选买入Top10`
- `当前持仓检查`
- `全部评分`
- `失败列表`

## 5. Current scoring mechanism

### 5.1 Base filters

Before scoring, the program removes stocks that fail basic requirements:

- Exclude ST-like names.
- Listing age must be at least 180 days.
- Stock price must be at least CNY 3.
- 20-day average trading amount must be at least CNY 80 million.
- Historical data must contain at least 120 trading bars.
- 20-day return must not exceed 45%.
- 60-day maximum drawdown must not exceed 35%.
- If the price is below MA60 and MA20 is also below MA60, the stock is removed.
- If float market cap is available, the minimum float market cap is CNY 5 billion. If the data source does not provide this field, the program skips this hard filter.

### 5.2 Weighted score

```text
Total Score = Momentum-Trend Score × 35%
            + Quality Score        × 20%
            + Valuation Score      × 15%
            + Liquidity Score      × 15%
            + Risk Control Score   × 15%
```

Details:

```text
Momentum-Trend Score = Momentum Score × 55% + Trend Score × 45%

Momentum Score = 20-day return percentile  × 30%
               + 60-day return percentile  × 45%
               + 120-day return percentile × 25%

Trend Score = latest price / MA60 percentile × 40%
            + MA20 / MA60 percentile         × 30%
            + latest price / MA120 percentile × 30%

Liquidity Score = 20-day average amount percentile × 70%
                + amount expansion percentile       × 30%

Risk Control Score = inverse 60-day annualized volatility percentile × 50%
                   + inverse 60-day max drawdown percentile          × 50%

Quality Score = ROE score                    × 40%
              + Operating cash flow / net profit score × 35%
              + Deducted net profit growth score       × 25%
```

Valuation score:

- Uses `PE_TTM`.
- Lower PE within the same industry receives a higher score.
- If the industry sample is too small, it falls back to a market-wide PE percentile.
- Missing, non-positive, or abnormal PE is handled with a low/fallback score.
- PEG is currently disabled.

### 5.3 Data availability

Stable fields currently used:

- Code and name.
- Latest price.
- Historical OHLC and trading amount.
- MA20, MA60, MA120.
- 20/60/120-day returns.
- 60-day annualized volatility.
- 60-day maximum drawdown.
- 20/60-day average trading amount.
- Amount expansion ratio.
- Whether price is above MA60/MA120.
- Whether MA20 is above MA60.
- Current holdings, cost, shares, and cash from local files.

Available but not fully stable, so handled defensively:

- Float/total market cap: skipped when unavailable.
- PE_TTM: used in valuation score; missing values receive fallback scores.
- ROE, operating cash flow / net profit, deducted net profit growth: missing values receive neutral quality scores.
- Industry: used for industry-relative PE percentile; missing values fall back to market-wide grouping.

Currently disabled or not core inputs:

- PEG.
- Brokerage API positions.

## 6. Market exposure logic

The program calculates a market vitality score from 0 to 5 based on:

- Ratio of candidates above MA60.
- Ratio of candidates above MA120.
- Ratio with positive 20-day return.
- Ratio with positive 60-day return.
- Ratio with expanded trading amount.

Suggested equity exposure:

- Weak market: about 30%.
- Normal market: about 60%.
- Strong market: about 85%.
- Active market: about 95%.

## 7. Risk disclaimer

This is a personal research tool. Markets are risky, models can fail, and data sources can be incomplete or wrong. Always review the report manually before making any trade.

## 8. Privacy

The following files may contain personal account or trading information and should not be published:

- `holdings.csv`
- `trades.csv`
- `交割单.xlsx`
- `自动持仓.csv`
- `output/`
- `cache/`
