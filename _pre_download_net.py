#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
一键预下载所有合约的净持仓完整历史数据。
用法: python _pre_download_net.py

逻辑与 _download_all.py（期货数据）完全一致：
- 遍历 ALL_CONTRACTS 所有合约
- 对缺失的合约全量下载所有交易日净持仓
- 已有完整数据的合约秒过
- 并行 4 线程下载
"""
import sys, io, time
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

from pathlib import Path
from collections import defaultdict

# 直接复用 basis_app 中的函数（绕开 Streamlit）
BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR))

import pandas as pd
import akshare as ak
from concurrent.futures import ThreadPoolExecutor, as_completed

# ═══════════════════════════════════════════
# 合约列表（与 basis_app 一致）
# ═══════════════════════════════════════════
ALL_MONTHS = ["01", "03", "05", "07", "09", "11"]
def _build_contracts():
    cts = []
    for y in range(21, 28):
        for m in ALL_MONTHS:
            c = f"LH{y}{m}"
            if "LH2109" <= c <= "LH2705":
                cts.append(c)
    # 加上 LH2707
    cts.append("LH2707")
    return cts

ALL_CONTRACTS = _build_contracts()
FUTURES_DIR = BASE_DIR / "data" / "futures"
HOLDINGS_DIR = BASE_DIR / "data" / "holdings"
HOLDINGS_DIR.mkdir(parents=True, exist_ok=True)

PARALLEL = 4  # 净持仓 API 较慢，4 线程即可


def _csv_path(ct):
    return FUTURES_DIR / f"{ct}.csv"


def _net_agg_path(ct):
    return HOLDINGS_DIR / f"{ct}_net_agg.csv"


def _fetch_holdings(ct, date_str):
    """拉取单日持仓，返回 net = long_total - short_total"""
    for attempt in range(2):
        try:
            df_long = ak.futures_hold_pos_sina(symbol="多单持仓", contract=ct, date=date_str)
            df_short = ak.futures_hold_pos_sina(symbol="空单持仓", contract=ct, date=date_str)

            def _extract(df, col_name):
                for c in df.columns:
                    cs = str(c)
                    if "名次" in cs or "会员" in cs or "简称" in cs or "增减" in cs or "比上" in cs:
                        continue
                    return pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int).sum()
                return 0

            long_sum = _extract(df_long, "long")
            short_sum = _extract(df_short, "short")
            if long_sum > 0 and short_sum > 0:
                return int(long_sum - short_sum)
            return None
        except Exception:
            if attempt < 1:
                time.sleep(0.5)
    return None


def sync_one(ct, force_full=False):
    """同步单个合约净持仓（与 sync_net_holdings 逻辑一致）"""
    cp = _csv_path(ct)
    if not cp.exists():
        return ct, False, "无期货CSV"

    fut_df = pd.read_csv(cp, usecols=["date"])
    fut_df["date"] = pd.to_datetime(fut_df["date"])
    all_dates = sorted(fut_df["date"].unique())
    all_date_strs = [d.strftime("%Y%m%d") for d in all_dates]

    agg_path = _net_agg_path(ct)

    # ── 增量模式 ──
    if agg_path.exists() and not force_full:
        agg_df = pd.read_csv(agg_path)
        agg_df["date"] = pd.to_datetime(agg_df["date"])
        cached = set(agg_df["date"].dt.strftime("%Y%m%d"))
        pending = [d for d in all_date_strs if d not in cached]
        if not pending:
            return ct, True, f"已是最新 ({len(agg_df)}条)"

        fetched = 0
        for ds in pending:
            net = _fetch_holdings(ct, ds)
            if net is not None:
                new_row = pd.DataFrame([{"date": pd.to_datetime(ds), "net_position": net}])
                agg_df = agg_df[agg_df["date"] != pd.to_datetime(ds)]
                agg_df = pd.concat([agg_df, new_row], ignore_index=True)
                fetched += 1
        if fetched > 0:
            agg_df = agg_df.sort_values("date").reset_index(drop=True)
            agg_df.to_csv(agg_path, index=False)
        return ct, True, f"增量+{fetched}条 (共{len(agg_df)}条)"

    # ── 全量模式 ──
    cached = set()
    agg_df = None
    if agg_path.exists():
        agg_df = pd.read_csv(agg_path)
        agg_df["date"] = pd.to_datetime(agg_df["date"])
        cached = set(agg_df["date"].dt.strftime("%Y%m%d"))

    pending = [d for d in all_date_strs if d not in cached]
    if not pending:
        return ct, True, f"已是最新 ({len(agg_df) if agg_df is not None else 0}条)"

    fetched = 0

    def _fetch_one(ds):
        return ds, _fetch_holdings(ct, ds)

    with ThreadPoolExecutor(max_workers=5) as ex:
        future_map = {ex.submit(_fetch_one, ds): ds for ds in pending}
        for future in as_completed(future_map):
            ds, net = future.result()
            if net is not None:
                new_row = pd.DataFrame([{"date": pd.to_datetime(ds), "net_position": net}])
                if agg_df is not None and not agg_df.empty:
                    agg_df = agg_df[agg_df["date"] != pd.to_datetime(ds)]
                    agg_df = pd.concat([agg_df, new_row], ignore_index=True)
                else:
                    agg_df = new_row
                fetched += 1

    if agg_df is not None and not agg_df.empty:
        agg_df = agg_df.sort_values("date").reset_index(drop=True)
        agg_df.to_csv(agg_path, index=False)

    return ct, True, f"全量{fetched}条/{len(pending)}天 (共{len(agg_df) if agg_df is not None else 0}条)"


def main():
    # 找出有期货 CSV 但缺净持仓聚合的合约
    with_futures = [c for c in ALL_CONTRACTS if _csv_path(c).exists()]
    missing = [c for c in with_futures if not _net_agg_path(c).exists()]
    existing = [c for c in with_futures if _net_agg_path(c).exists()]

    print(f"{'='*60}")
    print(f"合约总数: {len(ALL_CONTRACTS)}")
    print(f"有期货数据: {len(with_futures)}")
    print(f"缺净持仓: {len(missing)}")
    print(f"已有净持仓: {len(existing)}")
    print(f"{'='*60}")

    # ── 全量下载缺失的 ──
    if missing:
        print(f"\n>>> 全量下载 {len(missing)} 个缺失合约 (并行 {PARALLEL} 线程)...")
        start = time.time()

        with ThreadPoolExecutor(max_workers=PARALLEL) as ex:
            future_map = {ex.submit(sync_one, ct, True): ct for ct in missing}
            done = 0
            for future in as_completed(future_map):
                ct, ok, msg = future.result()
                done += 1
                elapsed = time.time() - start
                eta = (elapsed / done) * (len(missing) - done) if done > 0 else 0
                print(f"  [{done:2d}/{len(missing)}] {ct}: {msg} | {elapsed/60:.1f}m ETA {eta/60:.1f}m")

        print(f"DONE: {((time.time()-start)/60):.1f} 分钟")

    # ── 增量更新已有的 ──
    if existing:
        print(f"\n>>> 增量更新 {len(existing)} 个已有合约...")
        for i, ct in enumerate(existing):
            _, ok, msg = sync_one(ct, force_full=False)
            if i < 5 or "已是最新" not in msg:
                print(f"  [{i+1}/{len(existing)}] {ct}: {msg}")

    # ── 统计 ──
    total_ok = sum(1 for c in with_futures if _net_agg_path(c).exists())
    print(f"\n{'='*60}")
    print(f"完成: {total_ok}/{len(with_futures)} 个合约已有净持仓数据")
    print(f"数据目录: {HOLDINGS_DIR}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
