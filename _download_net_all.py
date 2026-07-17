#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""独立脚本：批量下载所有合约的前20净持仓聚合数据。
不依赖 Streamlit / basis_app，直接调 akshare + 本地 CSV 缓存。
"""

import sys
import time
import re
from pathlib import Path
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

# Windows GBK 编码兼容
if sys.platform == "win32":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

import pandas as pd
import akshare as ak

BASE_DIR = Path(__file__).parent
FUTURES_DIR = BASE_DIR / "data" / "futures"
HOLDINGS_DIR = BASE_DIR / "data" / "holdings"
HOLDINGS_DIR.mkdir(parents=True, exist_ok=True)

ALL_MONTHS = ["01", "03", "05", "07", "09", "11"]


def _build_contracts():
    cts = []
    for y in range(21, 28):
        for m in ALL_MONTHS:
            c = f"LH{y}{m}"
            if "LH2109" <= c <= "LH2705":
                cts.append(c)
    return cts


ALL_CONTRACTS = _build_contracts()


def _csv_path(ct: str) -> Path:
    return FUTURES_DIR / f"{ct}.csv"


def _net_agg_path(ct: str) -> Path:
    return HOLDINGS_DIR / f"{ct}_net_agg.csv"


def _norm_sina(df, val_col, chg_col):
    """标准化新浪返回的 DataFrame"""
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


def fetch_holdings_for_date(ct: str, date_str: str):
    """拉取指定合约在指定日期的真实持仓数据"""
    cache_file = HOLDINGS_DIR / f"{ct}_{date_str}.csv"
    if cache_file.exists():
        try:
            df = pd.read_csv(cache_file)
            if "company" in df.columns and "long" in df.columns and "short" in df.columns:
                return df
        except Exception:
            pass

    for attempt in range(2):
        try:
            df_vol = ak.futures_hold_pos_sina(symbol="成交量", contract=ct, date=date_str)
            df_long = ak.futures_hold_pos_sina(symbol="多单持仓", contract=ct, date=date_str)
            df_short = ak.futures_hold_pos_sina(symbol="空单持仓", contract=ct, date=date_str)

            vol_df = _norm_sina(df_vol, "volume", "volume_chg")
            long_df = _norm_sina(df_long, "long", "long_chg")
            short_df = _norm_sina(df_short, "short", "short_chg")

            merged = long_df.merge(short_df, on="company", how="outer")
            merged = merged.merge(vol_df, on="company", how="outer")
            merged = merged.fillna(0)
            for col in ["long", "long_chg", "short", "short_chg", "volume", "volume_chg"]:
                if col in merged.columns:
                    merged[col] = merged[col].astype(int)

            if merged.empty or merged["long"].sum() == 0:
                return None  # 数据未发布

            merged = merged.sort_values("long", ascending=False).reset_index(drop=True)
            merged.to_csv(cache_file, index=False)
            return merged
        except Exception:
            if attempt < 1:
                time.sleep(0.5)
    return None


def download_net_agg(ct: str, max_workers: int = 5):
    """下载单个合约的全量净持仓聚合数据"""
    cp = _csv_path(ct)
    if not cp.exists():
        print(f"  {ct}: 无期货CSV，跳过")
        return 0

    # 读取期货交易日列表
    try:
        fut_df = pd.read_csv(cp, usecols=["date"])
        fut_df["date"] = pd.to_datetime(fut_df["date"])
    except Exception:
        print(f"  {ct}: 读取期货CSV失败")
        return 0

    all_trading_dates = sorted(fut_df["date"].unique(), reverse=True)
    # 转为 YYYYMMDD 字符串列表
    date_strs = [d.strftime("%Y%m%d") for d in all_trading_dates]

    # 检查已有聚合数据
    agg_path = _net_agg_path(ct)
    existing_dates = set()
    agg_df = None
    if agg_path.exists():
        try:
            agg_df = pd.read_csv(agg_path)
            agg_df["date"] = pd.to_datetime(agg_df["date"])
            existing_dates = set(agg_df["date"].dt.strftime("%Y%m%d"))
        except Exception:
            agg_df = None

    # 筛选待下载日期
    pending = [d for d in date_strs if d not in existing_dates]
    if not pending:
        print(f"  {ct}: 已是最新 ({len(existing_dates)}条)")
        return 0

    total_pending = len(pending)
    print(f"  {ct}: 需下载 {total_pending} 个交易日…", end="", flush=True)

    fetched = 0
    # 并行拉取每个交易日
    def _fetch_one(ds):
        h = fetch_holdings_for_date(ct, ds)
        if h is not None and not h.empty:
            return ds, int(h["long"].sum() - h["short"].sum())
        return ds, None

    results = {}
    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        future_map = {ex.submit(_fetch_one, ds): ds for ds in pending}
        for future in as_completed(future_map):
            ds, net = future.result()
            if net is not None:
                results[ds] = net
                fetched += 1

    # 合并写入聚合文件
    if fetched > 0:
        new_rows = [{"date": pd.to_datetime(ds), "net_position": net}
                    for ds, net in results.items()]
        new_df = pd.DataFrame(new_rows)
        if agg_df is not None and not agg_df.empty:
            agg_df = agg_df[~agg_df["date"].isin(new_df["date"])]
            agg_df = pd.concat([agg_df, new_df], ignore_index=True)
        else:
            agg_df = new_df
        agg_df = agg_df.sort_values("date").reset_index(drop=True)
        agg_df.to_csv(agg_path, index=False)
        print(f" OK: +{fetched}条 (总计{len(agg_df)}条)")
    else:
        print(f" WARN: 0条获取成功")

    return fetched


def main():
    # 1. 扫描需要下载的合约
    all_with_futures = [c for c in ALL_CONTRACTS if _csv_path(c).exists()]
    def _is_missing(c):
        p = _net_agg_path(c)
        if not p.exists():
            return True
        try:
            df = pd.read_csv(p)
            return df.empty
        except Exception:
            return True

    missing = [c for c in all_with_futures if _is_missing(c)]
    # 已有聚合的做增量更新
    existing = [c for c in all_with_futures if not _is_missing(c)]

    print(f"{'='*60}")
    print(f"Total contracts: {len(ALL_CONTRACTS)}")
    print(f"With futures CSV: {len(all_with_futures)}")
    print(f"Missing net agg: {len(missing)}")
    print(f"Existing (incremental): {len(existing)}")
    print(f"{'='*60}")

    # 2. Full download for missing contracts (parallel across contracts)
    if missing:
        print(f"\n>>> Full download for {len(missing)} missing contracts (parallel)")
        start = time.time()

        def _download_one(ct):
            n = download_net_agg(ct, max_workers=3)
            return ct, n

        with ThreadPoolExecutor(max_workers=3) as ex:
            future_map = {ex.submit(_download_one, ct): ct for ct in missing}
            done = 0
            for future in as_completed(future_map):
                ct, n = future.result()
                done += 1
                elapsed = time.time() - start
                eta = (elapsed / done) * (len(missing) - done) if done > 0 else 0
                print(f"  [{done:2d}/{len(missing)}] {ct} done (+{n}条) | elapsed={elapsed/60:.1f}min ETA={eta/60:.1f}min")

        print(f"DONE: Full download completed in {((time.time()-start)/60):.1f} minutes")

    # 3. Incremental update for existing contracts (also parallel)
    if existing:
        # Only update contracts that are not too old (API might not have data)
        # Contracts from 2023+ are more likely to have data
        recent = [c for c in existing if c >= "LH2301"]
        old = [c for c in existing if c < "LH2301"]
        if old:
            print(f"\n>>> Skipping {len(old)} old contracts (pre-2023, API data may be unavailable): {', '.join(old)}")

        if recent:
            print(f"\n>>> Incremental update for {len(recent)} recent contracts (parallel)")
            start = time.time()

            def _update_one(ct):
                n = download_net_agg(ct, max_workers=2)
                return ct, n

            with ThreadPoolExecutor(max_workers=3) as ex:
                future_map = {ex.submit(_update_one, ct): ct for ct in recent}
                done = 0
                for future in as_completed(future_map):
                    ct, n = future.result()
                    done += 1
                    print(f"  [{done:2d}/{len(recent)}] {ct} done (+{n}条)")

            print(f"DONE: Incremental update completed in {((time.time()-start)/60):.1f} minutes")

    # 4. Final stats
    ok = sum(1 for c in all_with_futures if _net_agg_path(c).exists())
    print(f"\n{'='*60}")
    print(f"ALL DONE! {ok}/{len(all_with_futures)} contracts ready")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
