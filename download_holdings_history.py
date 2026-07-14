#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Robust batch download with timeout. Run from sentiment_platform directory."""

import signal
import pandas as pd
from pathlib import Path
import akshare as ak

BASE_DIR = Path(__file__).parent
FUTURES_DIR = BASE_DIR / "data" / "futures"
HOLDINGS_DIR = BASE_DIR / "data" / "holdings"
HOLDINGS_DIR.mkdir(exist_ok=True)

TIMEOUT = 15  # seconds per API call


class TimeoutError(Exception):
    pass


def with_timeout(seconds):
    def decorator(func):
        def wrapper(*args, **kwargs):
            result = [None]
            exception = [None]

            def target():
                try:
                    result[0] = func(*args, **kwargs)
                except Exception as e:
                    exception[0] = e

            import threading
            t = threading.Thread(target=target)
            t.daemon = True
            t.start()
            t.join(seconds)
            if t.is_alive():
                raise TimeoutError(f"Timed out after {seconds}s")
            if exception[0]:
                raise exception[0]
            return result[0]
        return wrapper
    return decorator


@with_timeout(TIMEOUT)
def call_api(symbol, contract, date_str):
    return ak.futures_hold_pos_sina(symbol=symbol, contract=contract, date=date_str)


def fetch_one(ct: str, date_str: str) -> bool:
    cache_file = HOLDINGS_DIR / f"{ct}_{date_str}.csv"
    if cache_file.exists():
        return True

    try:
        df_vol = call_api("成交量", ct, date_str)
        df_long = call_api("多单持仓", ct, date_str)
        df_short = call_api("空单持仓", ct, date_str)

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
        merged.to_csv(generic, index=False)
        (HOLDINGS_DIR / f"{ct}_meta.txt").write_text(date_str)
        return True
    except Exception:
        return False


def main():
    # Read all contracts with futures data
    contracts = sorted([f.stem for f in FUTURES_DIR.glob("LH*.csv")])
    print(f"Found {len(contracts)} contracts")

    for ct in contracts:
        df = pd.read_csv(FUTURES_DIR / f"{ct}.csv")
        df["date"] = pd.to_datetime(df["date"])
        dates = sorted(df["date"].dt.strftime("%Y%m%d").unique())
        missing = [d for d in dates if not (HOLDINGS_DIR / f"{ct}_{d}.csv").exists()]

        if not missing:
            print(f"  {ct}: {len(dates)} cached, skip")
            continue

        print(f"  {ct}: {len(dates)} total, {len(missing)} to fetch ...", end=" ", flush=True)
        ok = 0
        for d in missing:
            if fetch_one(ct, d):
                ok += 1
        print(f"{ok}/{len(missing)} ok ({len(dates)-len(missing)+ok}/{len(dates)} total)")

    # Final stats
    total_files = len(list(HOLDINGS_DIR.glob("LH*_*.csv")))
    print(f"\nDone! Total cached files: {total_files}")


if __name__ == "__main__":
    main()
