#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
从钢联 Excel 导入前20多空净持仓历史数据，写入 data/holdings/{ct}_net_agg.csv。
只读取"前20多空净持仓" sheet，覆盖 2021-2026 所有合约。
"""
import sys, io
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime
from collections import defaultdict

BASE_DIR = Path(__file__).parent
HOLDINGS_DIR = BASE_DIR / "data" / "holdings"
HOLDINGS_DIR.mkdir(parents=True, exist_ok=True)

EXCEL_PATH = Path(r"D:\CC\Desktop\平台数据\7.期货市场（日）.xlsx")

# ═══════════════════════════════════════════
# 1. 读取 Excel
# ═══════════════════════════════════════════
print(f"📖 读取: {EXCEL_PATH}")
df_raw = pd.read_excel(EXCEL_PATH, sheet_name="前20多空净持仓", header=None, skiprows=4)

# 列布局（每个月份占一列组，各组结构相同）:
#   col+0: 日期
#   col+1: 多单持仓合计
#   col+2: 净持仓（多-空，钢联计算）
#   col+3: 空单持仓合计
MONTH_COLUMNS = {
    "01": 0,   # cols 0-3
    "05": 13,  # cols 13-16
    "09": 19,  # cols 19-22
    "11": 25,  # cols 25-28
}
# 注: 03合约标记为"停-"，07合约无此sheet数据

# ═══════════════════════════════════════════
# 2. 解析
# ═══════════════════════════════════════════
def parse_month(df, base_col: int, month: str) -> list:
    """解析某个月份的数据，返回 [(contract_code, date, net_position), ...]"""
    records = []
    dates = pd.to_datetime(df.iloc[:, base_col], errors="coerce")

    for i in range(len(df)):
        d = dates.iloc[i]
        if pd.isna(d):
            continue
        d = d.date()

        long_val = df.iloc[i, base_col + 1]
        short_val = df.iloc[i, base_col + 3]

        # 跳过无效行
        if pd.isna(long_val) or pd.isna(short_val):
            continue
        try:
            lv = int(float(long_val))
            sv = int(float(short_val))
        except (ValueError, TypeError):
            continue

        if lv == 0 and sv == 0:
            continue

        # 日期 → 合约映射
        # 规则: 如果当前月份 <= 合约月份，合约年份 = 当前年份；否则合约年份 = 当前年份 + 1
        m = int(month)
        if d.month <= m:
            contract_year = d.year
        else:
            contract_year = d.year + 1

        ct = f"LH{str(contract_year)[2:]}{month}"
        records.append((ct, d, lv - sv))  # net = long - short

    return records


all_records = []
for month, base_col in MONTH_COLUMNS.items():
    recs = parse_month(df_raw, base_col, month)
    print(f"  {month}月合约: {len(recs)} 条")
    all_records.extend(recs)

print(f"\n📊 总计: {len(all_records)} 条记录")

# ═══════════════════════════════════════════
# 3. 按合约分组
# ═══════════════════════════════════════════
by_contract = defaultdict(list)
for ct, d, net in all_records:
    by_contract[ct].append((d, net))

# ═══════════════════════════════════════════
# 4. 与现有数据合并，写入文件
# ═══════════════════════════════════════════
updated = 0
new_files = 0

for ct, recs in sorted(by_contract.items()):
    # 新数据
    new_df = pd.DataFrame(recs, columns=["date", "net_position"])
    new_df["date"] = pd.to_datetime(new_df["date"])
    new_df = new_df.sort_values("date").drop_duplicates(subset=["date"])

    agg_path = HOLDINGS_DIR / f"{ct}_net_agg.csv"

    # 合并已有数据（如有）
    if agg_path.exists():
        try:
            old_df = pd.read_csv(agg_path)
            old_df["date"] = pd.to_datetime(old_df["date"])
            # Excel 数据优先（更权威），API 数据作为补充
            combined = pd.concat([old_df, new_df], ignore_index=True)
            combined = combined.drop_duplicates(subset=["date"], keep="last")  # Excel wins
            combined = combined.sort_values("date")
        except Exception:
            combined = new_df
    else:
        combined = new_df
        new_files += 1

    combined.to_csv(agg_path, index=False)
    updated += 1

    yrs = sorted(combined["date"].dt.year.unique())
    dmin = combined["date"].min().strftime("%Y-%m-%d")
    dmax = combined["date"].max().strftime("%Y-%m-%d")
    print(f"  {ct}: {len(combined)}条, 年份={yrs}, {dmin} ~ {dmax}")

print(f"\n✅ 完成: {updated} 个合约已更新 (其中 {new_files} 个新文件)")
print(f"📁 数据目录: {HOLDINGS_DIR}")
