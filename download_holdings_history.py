#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
一次性下载所有合约的完整历史前20多空持仓数据到本地 CSV。
运行一次后，App 启动和 Tab 6 均为秒出。

注意：Sina API 仅覆盖 2025-04-15 至今的数据。
      更早的日期接口无数据，自动跳过。

用法：python download_holdings_history.py
"""

import sys
import time
import pandas as pd
from pathlib import Path

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
FUTURES_DIR = DATA_DIR / "futures"
HOLDINGS_DIR = DATA_DIR / "holdings"
HOLDINGS_DIR.mkdir(exist_ok=True)

# Sina API 最早有数据的日期
DATA_CUTOFF = "20250415"

# 所有可能的生猪合约
ALL_MONTHS = ["01", "03", "05", "07", "09", "11"]
ALL_CONTRACTS = []
for y in range(21, 28):
    for m in ALL_MONTHS:
        c = f"LH{y}{m}"
        if "LH2109" <= c <= "LH2705":
            ALL_CONTRACTS.append(c)


def get_trading_dates(ct: str):
    """从期货 CSV 获取合约在 DATA_CUTOFF 之后的交易日"""
    f = FUTURES_DIR / f"{ct}.csv"
    if not f.exists():
        return []
    try:
        df = pd.read_csv(f)
        df["date"] = pd.to_datetime(df["date"])
        df = df[df["date"] >= DATA_CUTOFF]
        return sorted(df["date"].dt.strftime("%Y%m%d").unique())
    except Exception:
        return []


def fetch_one(ct: str, date_str: str) -> bool:
    """拉取单日持仓数据并缓存，返回是否成功"""
    cache_file = HOLDINGS_DIR / f"{ct}_{date_str}.csv"
    if cache_file.exists():
        return True

    import akshare as ak

    try:
        df_vol = ak.futures_hold_pos_sina(symbol="成交量", contract=ct, date=date_str)
        df_long = ak.futures_hold_pos_sina(symbol="多单持仓", contract=ct, date=date_str)
        df_short = ak.futures_hold_pos_sina(symbol="空单持仓", contract=ct, date=date_str)

        def _norm(df, val_col, chg_col):
            out = pd.DataFrame()
            for c in df.columns:
                if "会员" in str(c) or "简称" in str(c):
                    out["company"] = df[c].astype(str).str.strip()
                    break
            for c in df.columns:
                cs = str(c)
                if "名次" in cs or "会员" in cs or "简称" in cs or "增减" in cs or "比上" in cs:
                    continue
                out[val_col] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
                break
            for c in df.columns:
                cs = str(c)
                if "增减" in cs or "比上" in cs:
                    out[chg_col] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
                    break
            if chg_col not in out.columns:
                out[chg_col] = 0
            return out[out["company"].notna() & (out["company"] != "")]

        vol_df = _norm(df_vol, "volume", "volume_chg")
        long_df = _norm(df_long, "long", "long_chg")
        short_df = _norm(df_short, "short", "short_chg")

        merged = long_df.merge(short_df, on="company", how="outer")
        merged = merged.merge(vol_df, on="company", how="outer")
        merged = merged.fillna(0)
        for col in ["long", "long_chg", "short", "short_chg", "volume", "volume_chg"]:
            if col in merged.columns:
                merged[col] = merged[col].astype(int)

        if merged.empty or merged["long"].sum() == 0:
            return False

        merged = merged.sort_values("long", ascending=False).reset_index(drop=True)
        merged.to_csv(cache_file, index=False)

        generic = HOLDINGS_DIR / f"{ct}.csv"
        meta = HOLDINGS_DIR / f"{ct}_meta.txt"
        merged.to_csv(generic, index=False)
        meta.write_text(date_str)
        return True

    except Exception:
        return False


def main():
    print("=" * 60)
    print("  生猪期货前20多空持仓 — 历史数据批量下载")
    print(f"  数据起始日期: {DATA_CUTOFF}")
    print("=" * 60)

    contracts = [c for c in ALL_CONTRACTS if (FUTURES_DIR / f"{c}.csv").exists()]
    print(f"\n发现 {len(contracts)} 个有期货数据的合约")

    total_new = 0
    total_miss = 0

    for ct in contracts:
        dates = get_trading_dates(ct)
        if not dates:
            continue

        missing = [d for d in dates if not (HOLDINGS_DIR / f"{ct}_{d}.csv").exists()]
        if not missing:
            print(f"  {ct}: 已全部缓存 ({len(dates)} 天)")
            continue

        print(f"  {ct}: 下载中 ({len(missing)} 天) ...", end=" ", flush=True)

        for i, d in enumerate(missing):
            ok = fetch_one(ct, d)
            if ok:
                total_new += 1
            else:
                total_miss += 1
            if i < len(missing) - 1:
                time.sleep(0.1)

        print(f"完成, 成功 {len(missing)-sum(1 for d in missing if not (HOLDINGS_DIR / f'{ct}_{d}.csv').exists())}/{len(missing)}")

    print(f"\n{'=' * 60}")
    print(f"完成！新下载 {total_new} 天, 失败/无数据 {total_miss} 天")
    print(f"数据目录: {HOLDINGS_DIR}")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
