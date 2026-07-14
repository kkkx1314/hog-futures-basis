#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Batch download all historical top-20 position data for LH futures.
Run once, then the app loads instantly from local CSV cache.

Usage: python download_holdings_history.py
"""

import time
import pandas as pd
from pathlib import Path

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
FUTURES_DIR = DATA_DIR / "futures"
HOLDINGS_DIR = DATA_DIR / "holdings"
HOLDINGS_DIR.mkdir(exist_ok=True)

ALL_MONTHS = ["01", "03", "05", "07", "09", "11"]
ALL_CONTRACTS = []
for y in range(21, 28):
    for m in ALL_MONTHS:
        c = f"LH{y}{m}"
        if "LH2109" <= c <= "LH2705":
            ALL_CONTRACTS.append(c)


def get_trading_dates(ct: str):
    """Read futures CSV and return all trading dates for the contract."""
    f = FUTURES_DIR / f"{ct}.csv"
    if not f.exists():
        return []
    try:
        df = pd.read_csv(f)
        df["date"] = pd.to_datetime(df["date"])
        return sorted(df["date"].dt.strftime("%Y%m%d").unique())
    except Exception:
        return []


def fetch_one(ct: str, date_str: str) -> bool:
    """Fetch and cache one day of position data. Returns True on success."""
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
    print("  LH Futures Top-20 Position Data - Batch Download")
    print("=" * 60)

    contracts = [c for c in ALL_CONTRACTS if (FUTURES_DIR / f"{c}.csv").exists()]
    print(f"\nFound {len(contracts)} contracts with futures data")

    total_new = 0
    total_miss = 0

    for ct in contracts:
        dates = get_trading_dates(ct)
        if not dates:
            continue

        missing = [d for d in dates if not (HOLDINGS_DIR / f"{ct}_{d}.csv").exists()]
        if not missing:
            print(f"  {ct}: all {len(dates)} days cached")
            continue

        print(f"  {ct}: downloading {len(missing)} days ...", end=" ", flush=True)
        ct_ok = 0
        for i, d in enumerate(missing):
            if fetch_one(ct, d):
                ct_ok += 1
                total_new += 1
            else:
                total_miss += 1
            # no sleep - API handles rapid requests fine

        print(f"done ({ct_ok}/{len(missing)} ok)")

    print(f"\n{'=' * 60}")
    print(f"Done! New: {total_new}, Failed: {total_miss}")
    print(f"Data dir: {HOLDINGS_DIR}")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
