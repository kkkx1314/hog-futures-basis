#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
后台批量下载所有生猪期货合约数据到本地 CSV。
直接运行即可，不依赖 Streamlit。
"""
import os
import sys
import time
import requests

# Windows GBK 编码兼容
if sys.platform == "win32":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

# ── 配置 ──
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
FUTURES_DIR = DATA_DIR / "futures"
DATA_DIR.mkdir(exist_ok=True)
FUTURES_DIR.mkdir(exist_ok=True)

ALL_MONTHS = ["01", "03", "05", "07", "09", "11"]
PARALLEL_WORKERS = 5

def _build_contracts():
    cts = []
    for y in range(21, 28):
        for m in ALL_MONTHS:
            c = f"LH{y}{m}"
            if "LH2109" <= c <= "LH2705":
                cts.append(c)
    return cts

ALL_CONTRACTS = _build_contracts()


def _download_one(ct: str) -> tuple:
    """下载单个合约全量历史数据，返回 (合约, 成功, 消息, 条数)"""
    today = datetime.now()
    start = (today - timedelta(days=365 * 6)).strftime("%Y%m%d")
    end = today.strftime("%Y%m%d")

    # ── akshare ──
    try:
        import akshare as ak
        df = ak.futures_zh_daily_sina(symbol=ct)
        if df is not None and not df.empty:
            df["date"] = pd.to_datetime(df["date"])
            df = df.sort_values("date")
            df = df[(df["date"] >= start) & (df["date"] <= end)]
            if not df.empty:
                csv_path = FUTURES_DIR / f"{ct}.csv"
                df.to_csv(csv_path, index=False)
                return (ct, True, "akshare", len(df))
    except Exception as e:
        pass  # 静默，尝试兜底

    # ── eastmoney 兜底 ──
    try:
        r = requests.get(
            "https://push2his.eastmoney.com/api/qt/stock/kline/get",
            params={
                "secid": f"114.{ct}",
                "fields1": "f1,f2,f3,f4,f5,f6",
                "fields2": "f51,f52,f53,f54,f55,f56,f57",
                "klt": "101",
                "fqt": "1",
                "end": "20500101",
                "lmt": "3000",
            },
            timeout=15,
            headers={
                "User-Agent": "Mozilla/5.0",
                "Referer": "https://quote.eastmoney.com/",
            },
        )
        if r.status_code == 200:
            d = r.json()
            if d and d.get("data") and d["data"].get("klines"):
                recs = []
                for k in d["data"]["klines"]:
                    p = k.split(",")
                    if len(p) >= 7:
                        recs.append({
                            "date": pd.to_datetime(p[0]),
                            "open": float(p[1]),
                            "close": float(p[2]),
                            "high": float(p[3]),
                            "low": float(p[4]),
                            "volume": int(float(p[5])),
                            "settle": float(p[2]),
                            "hold": 0,
                        })
                if recs:
                    df = pd.DataFrame(recs).sort_values("date")
                    csv_path = FUTURES_DIR / f"{ct}.csv"
                    df.to_csv(csv_path, index=False)
                    return (ct, True, "eastmoney", len(df))
    except Exception:
        pass

    return (ct, False, "无数据", 0)


def main():
    # 1. 检查已有哪些
    existing = set()
    for f in FUTURES_DIR.glob("LH*.csv"):
        existing.add(f.stem)

    missing = [c for c in ALL_CONTRACTS if c not in existing]

    if not missing:
        print("✅ 所有合约数据已就绪！")
        print(f"   📦 {len(existing)} 个合约：{'、'.join(sorted(existing))}")
        return

    print(f"📡 共 {len(ALL_CONTRACTS)} 个合约，已有 {len(existing)} 个，需下载 {len(missing)} 个")
    print(f"   缺失：{'、'.join(missing)}")
    print(f"   🔧 并行线程数：{PARALLEL_WORKERS}")
    print()

    # 2. 并行下载
    results = []
    start_time = time.time()

    with ThreadPoolExecutor(max_workers=PARALLEL_WORKERS) as executor:
        futures = {executor.submit(_download_one, ct): ct for ct in missing}
        done = 0
        for future in as_completed(futures):
            ct, ok, source, count = future.result()
            done += 1
            status = "✅" if ok else "❌"
            print(f"  [{done:2d}/{len(missing)}] {status} {ct}  {source}  {count}条")
            results.append((ct, ok, source, count))

    elapsed = time.time() - start_time

    # 3. 汇总
    success = sum(1 for _, ok, _, _ in results if ok)
    failed = sum(1 for _, ok, _, _ in results if not ok)
    total_rows = sum(c for _, _, _, c in results if c)

    print()
    print(f"{'='*60}")
    print(f"📊 下载完成！耗时 {elapsed:.1f} 秒")
    print(f"   ✅ 成功：{success} 个合约")
    if failed:
        failed_list = [ct for ct, ok, _, _ in results if not ok]
        print(f"   ❌ 失败：{failed} 个合约（{'、'.join(failed_list)}）")
    print(f"   📋 共 {total_rows:,} 条数据")
    print(f"   📁 数据目录：{FUTURES_DIR}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
