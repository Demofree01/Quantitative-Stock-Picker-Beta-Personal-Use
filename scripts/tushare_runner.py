#!/usr/bin/env python3
"""Standalone Tushare-backed runner for weekly_quant.py.

Data policy:
- A-share only, no ETF/fund market data.
- No intraday quotes; uses Tushare post-market daily + daily_basic.
- PE_TTM comes from Tushare daily_basic.pe_ttm.
- Uses Tushare historical daily bars (qfq via pro_bar when available).
"""
from __future__ import annotations

import json
import os
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Optional

import numpy as np
import pandas as pd
import yaml

try:
    import tushare as ts
except Exception as exc:  # pragma: no cover
    raise RuntimeError("缺少 tushare 包，请先安装：python -m pip install tushare") from exc

PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

import weekly_quant as w

CONFIG_PATH = PROJECT_DIR / "config.yaml"
TOKEN_FILE = PROJECT_DIR / ".tushare_token"
CACHE_DIR = PROJECT_DIR / "cache" / "tushare"
CACHE_DIR.mkdir(parents=True, exist_ok=True)


def read_csv_if_valid(path: Path, required_cols: list[str]) -> Optional[pd.DataFrame]:
    if not path.exists():
        return None
    try:
        df = pd.read_csv(path, dtype={"代码": str})
    except Exception:
        return None
    missing = [col for col in required_cols if col not in df.columns]
    if missing or df.empty:
        return None
    return df


def atomic_write_csv(df: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    df.to_csv(tmp, index=False, encoding="utf-8-sig")
    tmp.replace(path)


def load_token() -> str:
    token = os.environ.get("TUSHARE_TOKEN") or os.environ.get("TS_TOKEN")
    if not token and TOKEN_FILE.exists():
        token = TOKEN_FILE.read_text(encoding="utf-8").strip()
    if not token and CONFIG_PATH.exists():
        cfg = yaml.safe_load(CONFIG_PATH.read_text(encoding="utf-8")) or {}
        token = str((cfg.get("data") or {}).get("tushare_token") or "").strip()
    if not token:
        raise RuntimeError(
            "未找到 Tushare token。请放到环境变量 TUSHARE_TOKEN/TS_TOKEN，"
            "或 config.yaml:data.tushare_token，或项目 .tushare_token 文件。"
        )
    return token


class TushareProvider:
    def __init__(self, token: str):
        ts.set_token(token)
        self.pro = ts.pro_api(token)
        self.trade_date: Optional[str] = None
        self.stock_spot_cache: Optional[pd.DataFrame] = None
        self.history_cache_hits = 0
        self.history_api_fetches = 0
        self.spot_cache_hit = False

    def call(self, name: str, **kwargs) -> pd.DataFrame:
        last_exc = None
        for i in range(3):
            try:
                fn = getattr(self.pro, name)
                return fn(**kwargs)
            except Exception as exc:
                last_exc = exc
                time.sleep(1.5 * (i + 1))
        raise last_exc  # type: ignore[misc]

    def latest_trade_date(self) -> str:
        if self.trade_date:
            return self.trade_date
        # Do not use intraday data. Walk back until post-market daily_basic has rows.
        today = datetime.now()
        for delta in range(0, 14):
            d = (today - timedelta(days=delta)).strftime("%Y%m%d")
            df = self.call("daily_basic", trade_date=d, fields="ts_code,trade_date,close,pe_ttm,total_mv,circ_mv")
            if df is not None and not df.empty:
                self.trade_date = d
                return d
        raise RuntimeError("Tushare 最近14天 daily_basic 均为空，无法确定最新交易日")

    @staticmethod
    def code_from_ts(ts_code: Any) -> str:
        return str(ts_code or "").split(".")[0]

    def stock_spot(self) -> pd.DataFrame:
        if self.stock_spot_cache is not None:
            return self.stock_spot_cache.copy()
        trade_date = self.latest_trade_date()
        cache_path = CACHE_DIR / f"stock_spot_{trade_date}.csv"
        cached = read_csv_if_valid(cache_path, ["代码", "名称", "最新价", "成交额", "成交量", "总市值", "流通市值", "行业", "PE_TTM", "市值排名"])
        if cached is not None and cached["PE_TTM"].notna().sum() > 0:
            self.spot_cache_hit = True
            print(f"Tushare股票快照缓存命中：{trade_date}，{len(cached)}只，PE_TTM有效{cached['PE_TTM'].notna().sum()}只；缓存 {cache_path}")
            self.stock_spot_cache = w.normalize_spot_frame(cached, "STOCK")
            return self.stock_spot_cache.copy()
        daily = self.call(
            "daily",
            trade_date=trade_date,
            fields="ts_code,trade_date,open,high,low,close,vol,amount",
        )
        basic = self.call(
            "daily_basic",
            trade_date=trade_date,
            fields="ts_code,trade_date,close,turnover_rate,volume_ratio,pe,pe_ttm,pb,total_mv,circ_mv",
        )
        stocks = self.call(
            "stock_basic",
            exchange="",
            list_status="L",
            fields="ts_code,symbol,name,area,industry,list_date,market",
        )
        if daily.empty or basic.empty or stocks.empty:
            raise RuntimeError(f"Tushare {trade_date} daily/daily_basic/stock_basic 返回为空")
        df = stocks.merge(daily, on="ts_code", how="inner").merge(
            basic.drop(columns=["trade_date", "close"], errors="ignore"), on="ts_code", how="left"
        )
        df["代码"] = df["ts_code"].map(self.code_from_ts)
        df["名称"] = df["name"].map(w.clean_name)
        df = df[df["代码"].str.startswith(w.STOCK_PREFIXES)].copy()
        df = df[~df["名称"].map(w.is_st_like)].copy()
        df["最新价"] = pd.to_numeric(df["close"], errors="coerce")
        # Tushare vol is hands; amount is thousand CNY. Keep both compatible with existing scoring.
        df["成交量"] = pd.to_numeric(df["vol"], errors="coerce")
        df["成交额"] = pd.to_numeric(df["amount"], errors="coerce") * 1000.0
        df["总市值"] = pd.to_numeric(df["total_mv"], errors="coerce") * 10000.0
        df["流通市值"] = pd.to_numeric(df["circ_mv"], errors="coerce") * 10000.0
        df["上市日期"] = df["list_date"]
        df["行业"] = df["industry"].fillna("未知")
        df["PE_TTM"] = pd.to_numeric(df["pe_ttm"], errors="coerce")
        df["PEG"] = np.nan
        df = df.sort_values("总市值", ascending=False).reset_index(drop=True)
        df["市值排名"] = np.arange(1, len(df) + 1)
        out = df[["代码", "名称", "最新价", "成交额", "成交量", "总市值", "流通市值", "上市日期", "行业", "PE_TTM", "PEG", "市值排名"]].copy()
        if out["PE_TTM"].notna().sum() == 0:
            raise RuntimeError("Tushare daily_basic 未返回有效 PE_TTM")
        atomic_write_csv(out, cache_path)
        print(f"Tushare股票快照：{trade_date}，{len(out)}只，PE_TTM有效{out['PE_TTM'].notna().sum()}只；缓存 {cache_path}")
        self.stock_spot_cache = w.normalize_spot_frame(out, "STOCK")
        return self.stock_spot_cache.copy()

    def history(self, asset_type: str, code: str, cfg: Dict[str, Any]) -> pd.DataFrame:
        if asset_type != "STOCK":
            raise RuntimeError("Tushare独立版不获取ETF/基金行情")
        today = datetime.now()
        start = (today - timedelta(days=int(cfg["data"].get("history_days", 300)) + 30)).strftime("%Y%m%d")
        end = self.latest_trade_date()
        adjust = str(cfg["data"].get("adjust", "qfq"))
        cache_path = CACHE_DIR / "history" / end / f"stock_{code}_{adjust}.csv"
        cached = read_csv_if_valid(cache_path, ["date", "open", "close", "high", "low", "amount", "turnover_amount"])
        if cached is not None:
            cached["date"] = pd.to_datetime(cached["date"], errors="coerce")
            if cached["date"].notna().any() and cached["date"].max().strftime("%Y%m%d") == end:
                self.history_cache_hits += 1
                return cached
        ts_code = f"{code}.SH" if str(code).startswith("6") else f"{code}.SZ"
        try:
            df = ts.pro_bar(ts_code=ts_code, adj=adjust, start_date=start, end_date=end)
        except Exception:
            df = None
        if df is None or df.empty:
            df = self.call("daily", ts_code=ts_code, start_date=start, end_date=end, fields="ts_code,trade_date,open,high,low,close,vol,amount")
        if df is None or df.empty:
            raise RuntimeError(f"Tushare历史行情为空：{code}")
        df = df.copy()
        date_col = "trade_date" if "trade_date" in df.columns else "date"
        df["date"] = pd.to_datetime(df[date_col], format="%Y%m%d", errors="coerce")
        for col in ["open", "close", "high", "low", "vol", "amount"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")
        df = df.dropna(subset=["date", "close"]).sort_values("date").tail(int(cfg["data"].get("history_days", 300))).reset_index(drop=True)
        if "vol" in df.columns:
            volume = df["vol"]
        elif "amount" in df.columns:
            volume = df["amount"]
        else:
            volume = np.nan
        turnover_amount = pd.to_numeric(df.get("amount"), errors="coerce") * 1000.0
        out = pd.DataFrame({
            "date": df["date"],
            "open": df["open"],
            "close": df["close"],
            "high": df["high"],
            "low": df["low"],
            "amount": volume,
            "turnover_amount": turnover_amount,
        }).dropna(subset=["date", "close", "turnover_amount"])
        if len(out) < int(cfg.get("filter", {}).get("min_history_bars", 120)):
            raise RuntimeError(f"Tushare历史行情不足：{code}，仅{len(out)}条")
        self.history_api_fetches += 1
        atomic_write_csv(out, cache_path)
        return out


def main() -> None:
    token = load_token()
    provider = TushareProvider(token)

    # Patch the existing strategy engine's data layer only. Keep scoring/report logic unchanged.
    w.HISTORY_CACHE_READ_ENABLED = False
    w.download_stock_spot = provider.stock_spot
    w.download_etf_spot = lambda: pd.DataFrame(columns=["代码", "名称", "最新价", "成交额", "成交量"])
    w.build_history_frame = provider.history

    original_build_holdings_check = w.build_holdings_check

    def build_holdings_check_tushare(scored, holdings, cfg, failures):
        df = original_build_holdings_check(scored, holdings, cfg, failures)
        asset_col = "标的类型" if "标的类型" in df.columns else "asset_type" if "asset_type" in df.columns else None
        if asset_col:
            etf_mask = df[asset_col].astype(str).str.upper().isin(["ETF", "基金", "场内基金"])
            if etf_mask.any():
                for col in ["操作建议", "action"]:
                    if col in df.columns:
                        df.loc[etf_mask, col] = "长期持有"
                for col in ["建议卖出股数", "sell_shares"]:
                    if col in df.columns:
                        df.loc[etf_mask, col] = 0
                for col in ["建议交易金额", "trade_amount"]:
                    if col in df.columns:
                        df.loc[etf_mask, col] = 0
                for col in ["建议执行", "execution"]:
                    if col in df.columns:
                        df.loc[etf_mask, col] = "不操作"
                for col in ["建议理由", "reason"]:
                    if col in df.columns:
                        df.loc[etf_mask, col] = "ETF为个人长线持仓；Tushare独立版按要求不获取ETF行情，不给短期卖出/减仓建议。"
        return df

    w.build_holdings_check = build_holdings_check_tushare

    original_read_config = w.load_config

    def read_config_tushare() -> Dict[str, Any]:
        cfg = original_read_config()
        cfg.setdefault("data", {})
        cfg["data"]["include_etf"] = False
        cfg["data"]["stock_history_source"] = "tushare"
        cfg["data"]["require_fresh_price_on_weekday"] = True
        cfg["data"]["require_live_spot_price_on_weekday"] = False
        cfg["data"]["cache_history"] = False
        return cfg

    w.load_config = read_config_tushare
    w.main()

    metadata = {
        "run_finished_at": datetime.now().isoformat(timespec="seconds"),
        "trade_date": provider.trade_date,
        "spot_cache_hit": provider.spot_cache_hit,
        "history_cache_hits": provider.history_cache_hits,
        "history_api_fetches": provider.history_api_fetches,
        "data_policy": "Tushare daily + daily_basic; A-share only; no intraday; no ETF quotes",
    }
    meta_path = PROJECT_DIR / "output" / "logs" / "latest_tushare_metadata.json"
    meta_path.parent.mkdir(parents=True, exist_ok=True)
    meta_path.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Tushare运行元数据：{metadata}")


if __name__ == "__main__":
    main()
