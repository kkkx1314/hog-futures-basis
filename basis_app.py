#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
生猪期货基差分析平台 (Streamlit)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
无侧边栏设计——所有控件位于各 Tab 内部。
Tab 1 — 当日基差分布（柱状图 + 四指标卡片）
Tab 2 — 单合约基差走势（区域色板 + 汇总指标固定色）
Tab 3 — 合约基差比较（同比 / 交易日对齐，颜色按合约年份）
Tab 4 — 合约价差比较（月份选择，颜色按合约年份）
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
核心设计原则：
  • 颜色始终按 **合约年份** 固定，不按数据实际发生的交易年份
  • 跨年数据强制切分为独立 trace（12月与次年1月之间不连线）
  • 基差 = 现货(元/公斤) × 1000 - (期货收盘价 + 升贴水)
  • 全国均价升贴水强制为 0
  • 所有日期显示为中文格式
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import requests
import time
import shutil
import re
from datetime import datetime, timedelta
from pathlib import Path
from collections import defaultdict
from typing import Dict, List, Tuple, Optional
import warnings

warnings.filterwarnings("ignore")

# ══════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="生猪基差分析平台",
    page_icon="🐷",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ══════════════════════════════════════════════════════════════
# 路径 & 目录
# ══════════════════════════════════════════════════════════════
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
FUTURES_DIR = DATA_DIR / "futures"
DATA_DIR.mkdir(exist_ok=True)
FUTURES_DIR.mkdir(exist_ok=True)
# 现货数据路径：自动扫描桌面取最新文件，否则用项目内缓存
def _find_latest_spot() -> Path:
    """在桌面和项目目录中寻找最新的涌益咨询 Excel"""
    candidates = []
    # 1. 扫描桌面
    desktop = Path(r"D:\CC\Desktop")
    if desktop.exists():
        for f in desktop.glob("*涌益咨询日度数据*.xlsx"):
            candidates.append((f.stat().st_mtime, f))
        for f in desktop.glob("*涌益咨询*.xlsx"):
            if f not in [c[1] for c in candidates]:
                candidates.append((f.stat().st_mtime, f))
    # 2. 项目内备份
    local = DATA_DIR / "涌益咨询日度数据.xlsx"
    if local.exists():
        candidates.append((local.stat().st_mtime, local))
    # 3. 按修改时间降序，取最新
    if candidates:
        candidates.sort(key=lambda x: x[0], reverse=True)
        return candidates[0][1]
    # 4. 兜底
    return Path(r"D:\CC\Desktop\2026年6月29日涌益咨询日度数据.xlsx")

SPOT_PATH = _find_latest_spot()

# ══════════════════════════════════════════════════════════════
# 中文日期
# ══════════════════════════════════════════════════════════════
def _cn(d) -> str:
    if isinstance(d, pd.Timestamp):
        return f"{d.year}年{d.month:02d}月{d.day:02d}日"
    if isinstance(d, datetime):
        return f"{d.year}年{d.month:02d}月{d.day:02d}日"
    return str(d)

def _cn_md(d) -> str:
    if isinstance(d, pd.Timestamp):
        return f"{d.month:02d}月{d.day:02d}日"
    if isinstance(d, datetime):
        return f"{d.month:02d}月{d.day:02d}日"
    return str(d)

# ══════════════════════════════════════════════════════════════
# 合约列表
# ══════════════════════════════════════════════════════════════
ALL_MONTHS = ["01", "03", "05", "07", "09", "11"]

def _build_contracts() -> List[str]:
    cts = []
    for y in range(21, 28):
        for m in ALL_MONTHS:
            c = f"LH{y}{m}"
            if "LH2109" <= c <= "LH2705":
                cts.append(c)
    return cts

ALL_CONTRACTS = _build_contracts()

def ct_display(c: str) -> str:
    return f"生猪20{c[2:4]}年{c[4:6]}月（{c}）"

def ct_month(c: str) -> str:
    return c[4:6]

def ct_year(c: str) -> str:
    return f"20{c[2:4]}"

# ══════════════════════════════════════════════════════════════
# 年份配色
# ══════════════════════════════════════════════════════════════
YEAR_COLORS = {
    "2021": "#9B59B6", "2022": "#F1C40F", "2023": "#3498DB",
    "2024": "#2C3E50", "2025": "#27AE60", "2026": "#E74C3C", "2027": "#E67E22",
}
FALLBACK_COLOR = "#95A5A6"
AVG_LINE_COLOR = "#95A5A6"
AVG_LINE_WIDTH = 1
AVG_LINE_DASH = "dash"

# Tab 2 汇总指标固定颜色
SUMMARY_COLORS = {
    "全国均价": "#2C3E50",
    "最大基差": "#E74C3C",
    "最小基差": "#3498DB",
    "基差平均值": "#9B59B6",
}

# Tab 2 区域色板
REGION_PALETTE = [
    "#1F78B4", "#33A02C", "#FF7F00", "#6A3D9A", "#B15928",
    "#FB9A99", "#A6CEE3", "#B2DF8A", "#FDBF6F", "#CAB2D6",
    "#FFFF99", "#8DD3C7", "#BEBADA", "#80B1D3", "#FCCDE5",
    "#BC80BD", "#CCEBC5", "#FFED6F", "#B3E5FC", "#FF8A80",
    "#EA80FC", "#FFD180",
]

# ══════════════════════════════════════════════════════════════
# 辅助：合约代码提取 & 颜色
# ══════════════════════════════════════════════════════════════
def _extract_contract_code(label: str) -> str:
    m = re.search(r"LH\d{4}", label)
    return m.group(0) if m else ""

def _contract_color_from_label(label: str) -> str:
    """颜色按合约年份固定（非交易年份）"""
    ct = _extract_contract_code(label)
    if ct:
        return YEAR_COLORS.get(ct_year(ct), FALLBACK_COLOR)
    m = re.search(r"\b(\d{2})(0[13579])\b", label)
    if m:
        year = f"20{m.group(1)}"
        return YEAR_COLORS.get(year, FALLBACK_COLOR)
    return FALLBACK_COLOR

# ══════════════════════════════════════════════════════════════
# 升贴水 V1~V4
# ══════════════════════════════════════════════════════════════
PREMIUM_V1 = {"河南":0,"江苏":500,"浙江":1500,"安徽":100,"山东":-200,"湖北":500}
PREMIUM_V2 = {**PREMIUM_V1,"河北":-300,"陕西":-300,"山西":-300,"辽宁":-500,
              "内蒙古":-600,"湖南":1300,"江西":1400,"重庆":1400,"四川":1500}
PREMIUM_V3 = {
    "河南":0,"浙江":1100,"广东":600,"江苏":500,"福建":400,"安徽":300,
    "江西":100,"湖南":100,"湖北":0,"山东":0,"河北":-100,"四川":-200,
    "重庆":-200,"广西":-200,"陕西":-300,"山西":-400,"辽宁":-700,
    "内蒙古":-800,"吉林":-800,"黑龙江":-1000,"贵州":-1000,"云南":-1400,
}
PREMIUM_V4 = {
    "浙江":900,"福建":500,"广东":500,"江苏":500,"安徽":200,"山东":200,
    "河北":100,"湖南":100,"江西":100,"河南":0,"湖北":0,"陕西":0,
    "四川":-100,"重庆":-100,"山西":-100,"广西":-200,"辽宁":-300,
    "内蒙古":-300,"吉林":-300,"贵州":-300,"黑龙江":-500,"云南":-600,
}

def get_version(ct: str) -> Tuple[str, Dict]:
    try:
        n = int(ct[2:])
    except ValueError:
        return ("V4", PREMIUM_V4)
    if n <= 2203: return ("V1", PREMIUM_V1)
    if n <= 2303: return ("V2", PREMIUM_V2)
    if n <= 2503: return ("V3", PREMIUM_V3)
    return ("V4", PREMIUM_V4)

def get_premium(ct: str, region: str) -> int:
    return get_version(ct)[1].get(region, 0)

def get_regions(ct: str) -> List[str]:
    return list(get_version(ct)[1].keys())

# ══════════════════════════════════════════════════════════════
# 区域标准化
# ══════════════════════════════════════════════════════════════
_REGION_ALIAS = {
    "黑龙江省":"黑龙江","吉林省":"吉林","辽宁省":"辽宁","河北省":"河北",
    "河南省":"河南","山东省":"山东","山西省":"山西","湖北省":"湖北",
    "湖南省":"湖南","江苏省":"江苏","安徽省":"安徽","浙江省":"浙江",
    "福建省":"福建","江西省":"江西","广东省":"广东","广西壮族自治区":"广西",
    "广西省":"广西","四川省":"四川","重庆市":"重庆","陕西省":"陕西",
    "云南省":"云南","贵州省":"贵州","内蒙古自治区":"内蒙古",
    "内蒙古（东部）":"内蒙古","内蒙古东部":"内蒙古",
}
_STANDARD_REGIONS = {"黑龙江","吉林","辽宁","河北","河南","山东","山西","湖北",
                     "湖南","江苏","安徽","浙江","福建","江西","广东","广西",
                     "四川","重庆","陕西","云南","贵州","内蒙古"}

def norm_region(name: str) -> str:
    name = str(name).strip()
    if name in ("全国均价","全国","全国平均"):
        return "全国均价"
    if name in _REGION_ALIAS:
        return _REGION_ALIAS[name]
    if name in _STANDARD_REGIONS:
        return name
    for full, short in _REGION_ALIAS.items():
        if short in name or name in full:
            return short
    return ""

# ══════════════════════════════════════════════════════════════
# 现货加载
# ══════════════════════════════════════════════════════════════
@st.cache_data(ttl=3600)
def load_spot(path_str: str) -> Tuple[Dict[str, pd.DataFrame], str]:
    path = Path(path_str)
    if not path.exists():
        return {}, f"❌ 文件不存在：{path}"
    try:
        xls = pd.ExcelFile(path)
        raw: Dict[str, Dict[pd.Timestamp, float]] = {}
        if len(xls.sheet_names) > 0:
            try:
                df = pd.read_excel(xls, sheet_name=0, header=None)
                for reg, dp in _parse_wide(df).items():
                    s = norm_region(reg)
                    if s and s not in raw: raw[s] = dp
            except Exception: pass
        if len(xls.sheet_names) > 2:
            try:
                df = pd.read_excel(xls, sheet_name=2, header=None)
                for reg, dp in _parse_long(df).items():
                    s = norm_region(reg)
                    if s and s not in raw: raw[s] = dp
            except Exception: pass
        result = {}
        all_dates = set()
        for reg, dp in raw.items():
            dates = sorted(dp.keys())
            all_dates.update(dates)
            result[reg] = pd.DataFrame({
                "date": pd.to_datetime(dates),
                "price": [float(dp[d]) for d in dates],
            }).sort_values("date").reset_index(drop=True)
        msg = f"已加载 {len(result)} 个区域"
        if all_dates:
            msg += f"，日期 {_cn(min(all_dates))} ~ {_cn(max(all_dates))}"
        return result, msg
    except Exception as e:
        return {}, f"❌ 加载失败：{e}"

def _parse_wide(df):
    data = defaultdict(dict)
    if df.empty: return dict(data)
    row0 = df.iloc[0]
    date_cols = []
    for col in range(2, len(row0)):
        v = row0.iloc[col]
        if pd.notna(v):
            try: date_cols.append((col, pd.to_datetime(v)))
            except Exception: pass
    dt_to_avg = {dt: dc+2 for dc, dt in date_cols if dc+2 < len(row0)}
    for ridx in range(2, len(df)):
        reg = str(df.iloc[ridx,0]).strip() if pd.notna(df.iloc[ridx,0]) else ""
        if not reg or reg == "nan": continue
        for dt, ac in dt_to_avg.items():
            if ac >= len(df.columns): continue
            v = df.iloc[ridx, ac]
            if pd.isna(v): continue
            try:
                p = float(v)
                if 0 < p < 100: data[reg][dt] = p
            except Exception: pass
    return dict(data)

def _parse_long(df):
    data = defaultdict(dict)
    if df.empty: return dict(data)
    hdr = df.iloc[0]
    col_reg = {}
    for col in range(1, len(hdr)):
        n = str(hdr.iloc[col]).strip()
        if n and n.lower() != "nan" and n != "日期": col_reg[col] = n
    for ridx in range(1, len(df)):
        cd = df.iloc[ridx, 0]
        if pd.isna(cd): continue
        try: dt = pd.to_datetime(cd)
        except Exception: continue
        for col, rn in col_reg.items():
            v = df.iloc[ridx, col]
            if pd.isna(v): continue
            try:
                p = float(v)
                if 0 < p < 100: data[rn][dt] = p
            except Exception: pass
    return dict(data)

# ══════════════════════════════════════════════════════════════
# 期货加载 & CSV 缓存
# ══════════════════════════════════════════════════════════════
def _csv_path(ct: str) -> Path:
    return FUTURES_DIR / f"{ct}.csv"

def _cache_fresh(ct: str) -> bool:
    p = _csv_path(ct)
    if not p.exists(): return False
    try:
        df = pd.read_csv(p)
        return "date" in df.columns and not df.empty and pd.to_datetime(df["date"].max()).date() >= datetime.now().date()
    except Exception: return False

def get_cached_contracts() -> List[str]:
    if not FUTURES_DIR.exists(): return []
    return sorted(f.stem for f in FUTURES_DIR.glob("LH*.csv"))

def get_latest_futures_date() -> Optional[str]:
    latest = None
    for f in FUTURES_DIR.glob("LH*.csv"):
        try:
            df = pd.read_csv(f)
            if "date" not in df.columns or df.empty: continue
            d = pd.to_datetime(df["date"].max())
            if latest is None or d > latest: latest = d
        except Exception: pass
    return _cn(latest) if latest else None

def load_futures(ct: str, force: bool = False) -> Tuple[Optional[pd.DataFrame], str]:
    cp = _csv_path(ct)
    if not force and _cache_fresh(ct):
        try:
            df = pd.read_csv(cp); df["date"] = pd.to_datetime(df["date"])
            return df.sort_values("date").reset_index(drop=True), "📁 本地缓存"
        except Exception: pass
    if not force and cp.exists():
        try:
            old = pd.read_csv(cp); old["date"] = pd.to_datetime(old["date"])
            start = (old["date"].max()+timedelta(days=1)).strftime("%Y%m%d")
            end = datetime.now().strftime("%Y%m%d")
            new = _download_futures(ct, start, end)
            if new is not None and not new.empty:
                c = pd.concat([old,new],ignore_index=True).drop_duplicates(subset=["date"]).sort_values("date")
                c.to_csv(cp,index=False); return c.reset_index(drop=True), "🔄 增量更新"
            return old.sort_values("date").reset_index(drop=True), "📁 本地缓存（已最新）"
        except Exception: pass
    try:
        start = (datetime.now()-timedelta(days=365*5)).strftime("%Y%m%d")
        end = datetime.now().strftime("%Y%m%d")
        df = _download_futures(ct, start, end)
        if df is not None and not df.empty:
            df.to_csv(cp,index=False); return df.sort_values("date").reset_index(drop=True), "🌐 网络下载"
    except Exception: pass
    if cp.exists():
        try:
            df = pd.read_csv(cp); df["date"] = pd.to_datetime(df["date"])
            return df.sort_values("date").reset_index(drop=True), "⚠️ 旧缓存（网络失败）"
        except Exception: pass
    return None, "❌ 获取失败"

def _download_futures(ct: str, sd: str, ed: str) -> Optional[pd.DataFrame]:
    for i in range(3):
        try:
            import akshare as ak
            df = ak.futures_zh_daily_sina(symbol=ct)
            if df is not None and not df.empty:
                df["date"] = pd.to_datetime(df["date"]); df = df.sort_values("date")
                return df[(df["date"]>=sd)&(df["date"]<=ed)].reset_index(drop=True)
        except Exception:
            if i < 2: time.sleep(1)
    for i in range(3):
        try:
            r = requests.get("https://push2his.eastmoney.com/api/qt/stock/kline/get",
                params={"secid":f"114.{ct}","fields1":"f1,f2,f3,f4,f5,f6","fields2":"f51,f52,f53,f54,f55,f56,f57","klt":"101","fqt":"1","end":"20500101","lmt":"3000"},
                timeout=15, headers={"User-Agent":"Mozilla/5.0","Referer":"https://quote.eastmoney.com/"})
            if r.status_code != 200: continue
            d = r.json()
            if not d or not d.get("data") or not d["data"].get("klines"): continue
            recs = []
            for k in d["data"]["klines"]:
                p = k.split(",")
                if len(p) >= 7:
                    recs.append({"date":pd.to_datetime(p[0]),"open":float(p[1]),"close":float(p[2]),
                                 "high":float(p[3]),"low":float(p[4]),"volume":int(float(p[5])),"settle":float(p[2]),"hold":0})
            if recs: return pd.DataFrame(recs).sort_values("date").reset_index(drop=True)
        except Exception:
            if i < 2: time.sleep(1)
    return None

# ══════════════════════════════════════════════════════════════
# 基差计算
# ══════════════════════════════════════════════════════════════
def _to_ton(p: float) -> float:
    return float(p) * 1000

def calc_basis(ct: str, region: str, spot_df: pd.DataFrame, fut_df: pd.DataFrame) -> Optional[pd.DataFrame]:
    """basis = 现货(元/吨) − (期货收盘价 + 升贴水)"""
    if fut_df is None or fut_df.empty or spot_df is None or spot_df.empty: return None
    pm = get_premium(ct, region)
    fi = fut_df.set_index("date"); si = spot_df.set_index("date")
    common = si.index.intersection(fi.index)
    if len(common) == 0: return None
    recs = []
    for dt in sorted(common):
        sp = float(si.loc[dt,"price"]); fc = float(fi.loc[dt,"close"])
        recs.append({"date":dt,"basis":int(round(_to_ton(sp)-(fc+pm))),"spot_price":sp,"futures_close":fc,"premium":pm})
    return pd.DataFrame(recs).sort_values("date").reset_index(drop=True)

def calc_national_basis(spot_dict: dict, fut_df: pd.DataFrame) -> Optional[pd.DataFrame]:
    """全国均价基差，升贴水强制为 0"""
    if fut_df is None or fut_df.empty: return None
    fi = fut_df.set_index("date")
    national_spot = spot_dict.get("全国均价")
    recs = []
    for dt in sorted(fi.index):
        if national_spot is not None:
            row = national_spot[national_spot["date"]==dt]
            if not row.empty: avg_sp = float(row["price"].iloc[0])
            else: continue
        else:
            prices = [float(df[df["date"]==dt]["price"].iloc[0]) for _,df in spot_dict.items() if not df[df["date"]==dt].empty]
            if prices: avg_sp = np.mean(prices)
            else: continue
        fc = float(fi.loc[dt,"close"])
        recs.append({"date":dt,"basis":int(round(_to_ton(avg_sp)-fc)),"spot_price":avg_sp,"futures_close":fc,"premium":0})
    return pd.DataFrame(recs).sort_values("date").reset_index(drop=True) if recs else None

def get_summary_series(ct: str, spot_dict: dict, fut_df: pd.DataFrame, regions: List[str]):
    """四个汇总指标序列"""
    if fut_df is None or fut_df.empty: return None,None,None,None
    valid = [r for r in regions if r in spot_dict]
    if not valid: return None,None,None,None
    fi = fut_df.set_index("date")
    na_recs, max_recs, min_recs, avg_recs = [], [], [], []
    for dt in sorted(fi.index):
        fc = float(fi.loc[dt,"close"])
        day = []
        for reg in valid:
            row = spot_dict[reg][spot_dict[reg]["date"]==dt]
            if row.empty: continue
            sp = float(row["price"].iloc[0]); pm = get_premium(ct, reg)
            day.append(int(round(_to_ton(sp)-(fc+pm))))
        if day:
            max_recs.append({"date":dt,"basis":max(day)})
            min_recs.append({"date":dt,"basis":min(day)})
            avg_recs.append({"date":dt,"basis":int(round(np.mean(day)))})
        prices = [float(spot_dict[r][spot_dict[r]["date"]==dt]["price"].iloc[0]) for r in valid if not spot_dict[r][spot_dict[r]["date"]==dt].empty]
        if prices: na_recs.append({"date":dt,"basis":int(round(np.mean(prices)*1000-fc))})
    def _df(r): return pd.DataFrame(r).sort_values("date").reset_index(drop=True) if r else None
    return _df(na_recs), _df(max_recs), _df(min_recs), _df(avg_recs)

def compute_snapshot(ct: str, spot_dict: dict, fut_df: pd.DataFrame, target_date, regions: List[str]) -> dict:
    """单日极值快照"""
    row = fut_df[fut_df["date"]==target_date]
    if row.empty: return {}
    fc = float(row["close"].iloc[0])
    items = []
    for reg in regions:
        if reg not in spot_dict: continue
        r = spot_dict[reg][spot_dict[reg]["date"]==target_date]
        if r.empty: continue
        sp = float(r["price"].iloc[0]); pm = get_premium(ct, reg)
        items.append((reg, int(round(_to_ton(sp)-(fc+pm))), sp))
    if not items: return {}
    items.sort(key=lambda x: x[1], reverse=True)
    national_spot = spot_dict.get("全国均价")
    if national_spot is not None:
        nr = national_spot[national_spot["date"]==target_date]
        na_basis = int(round(_to_ton(float(nr["price"].iloc[0]))-fc)) if not nr.empty else int(round(_to_ton(np.mean([x[2] for x in items]))-fc))
    else:
        na_basis = int(round(_to_ton(np.mean([x[2] for x in items]))-fc))
    return {
        "max_region": items[0][0], "max_basis": items[0][1],
        "min_region": items[-1][0], "min_basis": items[-1][1],
        "avg_basis": int(round(np.mean([x[1] for x in items]))),
        "national_avg": na_basis,
        "range": items[0][1]-items[-1][1],
        "futures_close": fc,
    }

# ══════════════════════════════════════════════════════════════
# _doy_to_date — 修复: 加入 source_year 正确适配闰年/非闰年
# ══════════════════════════════════════════════════════════════
def _is_leap(year: int) -> bool:
    return (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0)


def _doy_to_date(doy: int, source_year: int = None) -> pd.Timestamp:
    """day-of-year → 固定参考日期(2020年)。
    source_year 用于修正闰年/非闰年转换：非闰年的 doy>59 需+1 补偿2020年的2月29日。"""
    if source_year is not None and not _is_leap(source_year) and doy > 59:
        doy += 1
    return pd.Timestamp("2020-01-01") + pd.Timedelta(days=int(doy)-1)

def _make_trace_label(ct: str, trade_year, item_label: str) -> str:
    cy = ct_year(ct); ty = str(trade_year)
    if ty != cy: return f"{ct}({ty}) {item_label}"
    return f"{ct} {item_label}"

# ══════════════════════════════════════════════════════════════
# 图表
# ══════════════════════════════════════════════════════════════
def fig_distribution(recs: list, ct: str, target_date, data_date: str = "") -> go.Figure:
    """柱状图：四个汇总指标用橙色(#FF8C00)高亮，实际区域用蓝色(#1f77b4)"""
    if not recs: return go.Figure()
    df = pd.DataFrame(recs).sort_values("basis", ascending=True)
    # ★ 指标用橙色高亮，区域用蓝色
    clrs = ["#FF8C00" if r.get("is_indicator") else "#1f77b4" for _, r in df.iterrows()]
    title = f"{ct} 当日基差分布（{_cn(pd.to_datetime(target_date))}）"
    if data_date: title += f"<br><sup>📡 期货数据来源：akshare，数据日期：{data_date}</sup>"
    fig = go.Figure()
    fig.add_trace(go.Bar(x=df["region"], y=df["basis"], marker_color=clrs,
        text=[f"{v:+,}" for v in df["basis"]], textposition="outside", textfont=dict(size=11),
        hovertemplate="<b>%{x}</b><br>基差：%{y:+,}元/吨<br>现货：%{customdata[0]:.0f}元/公斤<br>期货：%{customdata[1]:.0f}元/吨<br>升贴水：%{customdata[2]:+d}元/吨<extra></extra>",
        customdata=df[["spot_price","futures_close","premium"]].values))
    fig.add_hline(y=0, line_dash="solid", line_color="gray", opacity=0.5)
    fig.update_layout(title=title, xaxis_title="区域", yaxis_title="基差（元/吨）",
        template="plotly_white", height=500, margin=dict(t=80,b=60,l=60,r=40), showlegend=False)
    fig.update_xaxes(tickangle=45)
    return fig

def fig_trend(basis_dict: Dict[str, pd.DataFrame], ct: str, data_date: str = "") -> go.Figure:
    if not basis_dict: return go.Figure()
    fig = go.Figure(); ri = 0
    for label, df in basis_dict.items():
        if df is None or df.empty: continue
        if label in SUMMARY_COLORS:
            c, w, d = SUMMARY_COLORS[label], 3, "solid"
        else:
            c, w, d = REGION_PALETTE[ri % len(REGION_PALETTE)], 2, "solid"; ri += 1
        fig.add_trace(go.Scatter(x=df["date"], y=df["basis"],
            mode="lines+markers" if len(df)<60 else "lines", name=label,
            line=dict(color=c, width=w, dash=d), marker=dict(size=3),
            hovertemplate=f"<b>{label}</b><br>%{{customdata}}<br>基差：%{{y:+,}}元/吨<extra></extra>",
            customdata=[_cn(d) for d in df["date"]]))
    fig.add_hline(y=0, line_dash="dash", line_color="gray", opacity=0.5)
    title = f"{ct} 基差走势"
    if data_date: title += f"<br><sup>📡 期货数据来源：akshare，数据日期：{data_date}</sup>"
    fig.update_layout(title=title, xaxis_title="日期", yaxis_title="基差（元/吨）",
        template="plotly_white", height=550, hovermode="x unified",
        legend=dict(orientation="h", y=1.02, x=0), margin=dict(t=80,b=40,l=60,r=40))
    fig.update_xaxes(rangeslider_visible=True, tickformat="%Y年%m月")
    return fig

def fig_calendar_comparison(series: Dict[str, pd.DataFrame], tmon: str, data_date: str = "") -> go.Figure:
    if not series: return go.Figure()
    fig = go.Figure()
    for label, df in series.items():
        if df is None or df.empty: continue
        if "历史均值" in label:
            c, w, d = AVG_LINE_COLOR, AVG_LINE_WIDTH, AVG_LINE_DASH
        else:
            c, w, d = _contract_color_from_label(label), 2, "solid"
        fig.add_trace(go.Scatter(x=df["plot_date"], y=df["basis"], mode="lines", name=label,
            line=dict(color=c, width=w, dash=d),
            hovertemplate=f"<b>{label}</b><br>%{{customdata}}<br>基差：%{{y:+,}}元/吨<extra></extra>",
            customdata=[_cn_md(r["plot_date"]) for _,r in df.iterrows()]))
    fig.add_hline(y=0, line_dash="dash", line_color="gray", opacity=0.5)
    title = f"{tmon}月合约基差季节图（同比 — 自然日对齐）"
    if data_date: title += f"<br><sup>📡 期货数据来源：akshare，数据日期：{data_date}</sup>"
    fig.update_layout(title=title, xaxis_title="日期（月-日）", yaxis_title="基差（元/吨）",
        template="plotly_white", height=550, hovermode="x unified",
        legend=dict(orientation="h", y=1.02, x=0), margin=dict(t=80,b=40,l=60,r=40))
    fig.update_xaxes(tickformat="%m-%d", dtick="M1", range=["2020-01-01","2020-12-31"])
    return fig

def fig_delivery_comparison(series: Dict[str, pd.DataFrame], data_date: str = "") -> go.Figure:
    if not series: return go.Figure()
    fig = go.Figure()
    for label, df in series.items():
        if df is None or df.empty: continue
        if "历史均值" in label: c, w, d = AVG_LINE_COLOR, AVG_LINE_WIDTH, AVG_LINE_DASH
        else: c, w, d = _contract_color_from_label(label), 2, "solid"
        fig.add_trace(go.Scatter(x=df["days"], y=df["basis"], mode="lines", name=label,
            line=dict(color=c, width=w, dash=d),
            hovertemplate=f"<b>{label}</b><br>距交割：%{{x}}天<br>基差：%{{y:+,}}元/吨<extra></extra>"))
    fig.add_hline(y=0, line_dash="dash", line_color="gray", opacity=0.5)
    title = "合约基差比较 — 距离交易日对齐"
    if data_date: title += f"<br><sup>📡 期货数据来源：akshare，数据日期：{data_date}</sup>"
    fig.update_layout(title=title, xaxis_title="距交割日天数", yaxis_title="基差（元/吨）",
        template="plotly_white", height=550, hovermode="x unified",
        legend=dict(orientation="h", y=1.02, x=0), margin=dict(t=80,b=40,l=60,r=40))
    fig.update_xaxes(autorange="reversed")
    return fig

def fig_spread_season(data: Dict[str, pd.DataFrame], ma: str, mb: str, data_date: str = "") -> go.Figure:
    if not data: return go.Figure()
    fig = go.Figure()
    for label, df in data.items():
        if df is None or df.empty: continue
        if "历史均值" in label: c, w, d = AVG_LINE_COLOR, AVG_LINE_WIDTH, AVG_LINE_DASH
        else: c, w, d = _contract_color_from_label(label), 2, "solid"
        fig.add_trace(go.Scatter(x=df["plot_date"], y=df["spread"], mode="lines", name=label,
            line=dict(color=c, width=w, dash=d),
            hovertemplate=f"<b>{label}</b><br>%{{customdata}}<br>价差：%{{y:+,}}元/吨<extra></extra>",
            # ★ 修复：直接使用已修正的 plot_date
            customdata=[_cn_md(r["plot_date"]) for _,r in df.iterrows()]))
    fig.add_hline(y=0, line_dash="dash", line_color="gray", opacity=0.5)
    title = f"{ma}月 − {mb}月 合约价差季节图"
    if data_date: title += f"<br><sup>📡 期货数据来源：akshare，数据日期：{data_date}</sup>"
    fig.update_layout(title=title, xaxis_title="日期（月-日）", yaxis_title="价差（元/吨）",
        template="plotly_white", height=550, hovermode="x unified",
        legend=dict(orientation="h", y=1.02, x=0), margin=dict(t=80,b=40,l=60,r=40))
    fig.update_xaxes(tickformat="%m-%d", dtick="M1", range=["2020-01-01","2020-12-31"])
    return fig

# ══════════════════════════════════════════════════════════════
# Tab 1：当日基差分布
# ══════════════════════════════════════════════════════════════
def tab1():
    st.subheader("📊 当日基差分布")

    # 预先加载可用的合约列表
    cached = set(get_cached_contracts())
    all_ct = sorted(ALL_CONTRACTS)
    spot_dict, spot_msg = load_spot(str(SPOT_PATH))
    fut_update_date = get_latest_futures_date()

    # 左列：控件
    col_ctrl, col_chart = st.columns([1, 3.5])

    with col_ctrl:
        ct = st.selectbox("📋 合约选择", options=all_ct, index=all_ct.index("LH2609") if "LH2609" in all_ct else 0,
                          format_func=lambda c: f"{ct_display(c)} {'📁' if c in cached else '🆕'}", key="t1_ct")
        ver, _ = get_version(ct)
        vregions = get_regions(ct)
        st.caption(f"升贴水版本：**{ver}**（{len(vregions)}个区域）")

        with st.spinner("加载期货…"):
            fut_df, fut_src = load_futures(ct)
        today = datetime.now().date()
        latest = fut_df["date"].max().date() if (fut_df is not None and not fut_df.empty) else today
        st.caption(f"{fut_src}，{len(fut_df) if fut_df is not None else 0}个交易日")

        sel_date = st.date_input("📅 选择日期", value=latest, max_value=today, key="t1_date")

        # 区域 + 四个汇总指标 — 默认全选
        available_regions = [r for r in vregions if r in spot_dict] or vregions
        region_opts = list(available_regions) + ["─── 汇总指标 ───", "📊 全国均价基差", "🔴 最大基差", "🟢 最小基差", "🟣 基差平均值"]
        defaults = list(available_regions) + ["📊 全国均价基差", "🔴 最大基差", "🟢 最小基差", "🟣 基差平均值"]
        sel_items = st.multiselect("🗺️ 地区与指标", options=region_opts,
            default=[x for x in defaults if x in region_opts], key="t1_items")

    # 右列：图表
    with col_chart:
        if fut_df is None or fut_df.empty:
            st.error("❌ 期货数据不可用"); return

        fds = sorted(fut_df["date"].unique())
        td = pd.to_datetime(sel_date)
        actual_td = td
        if td not in fds:
            nearby = [d for d in fds if d <= td]
            if nearby: actual_td = nearby[-1]; st.info(f"ℹ️ {_cn(td)} 非交易日，已使用 {_cn(actual_td)}")

        row = fut_df[fut_df["date"] == actual_td]
        if row.empty: st.error("❌ 无期货数据"); return
        fc = float(row["close"].iloc[0])

        # 快照
        snap = compute_snapshot(ct, spot_dict, fut_df, actual_td, available_regions)

        # ── 四个指标卡片（带阴影 + 数值着色） ──
        if snap:
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.markdown(f"""<div class="metric-card">
                    <div class="mlabel">🔴 最大基差（{snap.get('max_region','')}）</div>
                    <div class="mvalue" style="color:#E74C3C;">{snap['max_basis']:+,}</div>
                    <div class="munit">元/吨</div>
                </div>""", unsafe_allow_html=True)
            with c2:
                st.markdown(f"""<div class="metric-card">
                    <div class="mlabel">🟢 最小基差（{snap.get('min_region','')}）</div>
                    <div class="mvalue" style="color:#3498DB;">{snap['min_basis']:+,}</div>
                    <div class="munit">元/吨</div>
                </div>""", unsafe_allow_html=True)
            with c3:
                st.markdown(f"""<div class="metric-card">
                    <div class="mlabel">🟣 基差平均值</div>
                    <div class="mvalue" style="color:#9B59B6;">{snap['avg_basis']:+,}</div>
                    <div class="munit">元/吨</div>
                </div>""", unsafe_allow_html=True)
            with c4:
                st.markdown(f"""<div class="metric-card">
                    <div class="mlabel">⚫ 全国均价基差</div>
                    <div class="mvalue" style="color:#2C3E50;">{snap['national_avg']:+,}</div>
                    <div class="munit">元/吨</div>
                </div>""", unsafe_allow_html=True)
            st.caption(f"📅 数据日期：{_cn(actual_td)}")
        else:
            st.warning("⚠️ 所选日期无可用数据")

        # 构建柱状图数据
        if not sel_items: st.warning("⚠️ 请选择至少一个指标"); return

        na_df, max_df, min_df, avg_df = get_summary_series(ct, spot_dict, fut_df, available_regions)
        recs = []
        INDICATOR_NAMES = {"全国均价", "最大基差", "最小基差", "基差平均值"}
        for raw in sel_items:
            is_ind = False
            if raw in available_regions:
                if raw in spot_dict:
                    r = spot_dict[raw][spot_dict[raw]["date"]==actual_td]
                    if not r.empty:
                        sp = float(r["price"].iloc[0]); pm = get_premium(ct, raw)
                        recs.append({"region":raw,"basis":int(round(_to_ton(sp)-(fc+pm))),"spot_price":sp,"futures_close":fc,"premium":pm,"is_indicator":False})
            elif "全国均价" in raw and na_df is not None:
                r = na_df[na_df["date"]==actual_td]
                if not r.empty: recs.append({"region":"全国均价","basis":int(r["basis"].iloc[0]),"spot_price":0,"futures_close":fc,"premium":0,"is_indicator":True})
            elif "最大基差" in raw and max_df is not None:
                r = max_df[max_df["date"]==actual_td]
                if not r.empty: recs.append({"region":"最大基差","basis":int(r["basis"].iloc[0]),"spot_price":0,"futures_close":fc,"premium":0,"is_indicator":True})
            elif "最小基差" in raw and min_df is not None:
                r = min_df[min_df["date"]==actual_td]
                if not r.empty: recs.append({"region":"最小基差","basis":int(r["basis"].iloc[0]),"spot_price":0,"futures_close":fc,"premium":0,"is_indicator":True})
            elif "基差平均值" in raw and avg_df is not None:
                r = avg_df[avg_df["date"]==actual_td]
                if not r.empty: recs.append({"region":"基差平均值","basis":int(r["basis"].iloc[0]),"spot_price":0,"futures_close":fc,"premium":0,"is_indicator":True})

        if recs:
            st.plotly_chart(fig_distribution(recs, ct, actual_td, fut_update_date or ""), use_container_width=True)

            with st.expander("📋 数据明细表"):
                tbl = pd.DataFrame(recs).sort_values("basis", ascending=False)
                tbl["基差（元/吨）"] = tbl["basis"].apply(lambda x: f"{x:+,}")
                tbl["现货（元/公斤）"] = tbl["spot_price"].apply(lambda x: f"{int(round(x))}" if x > 0 else "—")
                tbl["期货（元/吨）"] = tbl["futures_close"].apply(lambda x: f"{int(round(x))}")
                display_cols = ["region","基差（元/吨）","现货（元/公斤）","期货（元/吨）"]
                st.dataframe(tbl[display_cols].rename(columns={"region":"区域/指标"}), use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════
# Tab 2：单合约基差走势
# ══════════════════════════════════════════════════════════════
def tab2():
    st.subheader("📈 单合约基差走势")

    cached = set(get_cached_contracts())
    all_ct = sorted(ALL_CONTRACTS)
    spot_dict, spot_msg = load_spot(str(SPOT_PATH))
    fut_update_date = get_latest_futures_date()

    col_ctrl, col_chart = st.columns([1, 3.5])

    with col_ctrl:
        ct = st.selectbox("📋 合约选择", options=all_ct, index=all_ct.index("LH2609") if "LH2609" in all_ct else 0,
                          format_func=lambda c: f"{ct_display(c)} {'📁' if c in cached else '🆕'}", key="t2_ct")
        ver, _ = get_version(ct); vregions = get_regions(ct)
        st.caption(f"升贴水：**{ver}**（{len(vregions)}个区域）")

        with st.spinner("加载期货…"):
            fut_df, fut_src = load_futures(ct)
        st.caption(f"{fut_src}，{len(fut_df) if fut_df is not None else 0}个交易日")

        # 区域多选 + 汇总指标 — 默认：河南、四川、广东、江苏、全国均价基差
        available_regions = [r for r in vregions if r in spot_dict] or vregions
        region_opts = list(available_regions) + ["─── 汇总指标 ───", "📊 全国均价基差", "🔴 最大基差", "🟢 最小基差", "🟣 基差平均值"]
        tab2_defaults = ["河南","四川","广东","江苏","📊 全国均价基差"]
        defaults = [x for x in tab2_defaults if x in region_opts]
        sel_items = st.multiselect("🗺️ 地区与指标", options=region_opts, default=defaults, key="t2_items")

    with col_chart:
        if fut_df is None or fut_df.empty:
            st.error("❌ 期货数据不可用"); return
        if not sel_items:
            st.warning("⚠️ 请选择至少一个区域或指标"); return
        if len(spot_dict) <= 1 and available_regions == vregions:
            st.warning("⚠️ 现货数据不足，走势图仅显示散点")

        na_df, max_df, min_df, avg_df = get_summary_series(ct, spot_dict, fut_df, available_regions)
        basis_dict = {}
        for raw in sel_items:
            if raw in available_regions:
                if raw in spot_dict:
                    df = calc_basis(ct, raw, spot_dict[raw], fut_df)
                    if df is not None and not df.empty:
                        basis_dict[f"{raw}（升贴水{get_premium(ct,raw):+d}）"] = df
            elif "全国均价" in raw:
                if na_df is not None and not na_df.empty: basis_dict["全国均价"] = na_df
            elif "最大基差" in raw:
                if max_df is not None and not max_df.empty: basis_dict["最大基差"] = max_df
            elif "最小基差" in raw:
                if min_df is not None and not min_df.empty: basis_dict["最小基差"] = min_df
            elif "基差平均值" in raw:
                if avg_df is not None and not avg_df.empty: basis_dict["基差平均值"] = avg_df

        if not basis_dict:
            st.warning("⚠️ 无可用数据"); return

        st.plotly_chart(fig_trend(basis_dict, ct, fut_update_date or ""), use_container_width=True)

        with st.expander("📋 基差统计表"):
            stats = []
            for label, df in basis_dict.items():
                if df.empty: continue
                stats.append({"区域/指标":label,"最新":f"{df['basis'].iloc[-1]:+,}","均值":f"{int(round(df['basis'].mean())):+,}",
                    "最大":f"{df['basis'].max():+,}","最小":f"{df['basis'].min():+,}","标准差":f"{int(round(df['basis'].std())):,}","数据点":len(df)})
            st.dataframe(pd.DataFrame(stats), use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════
# Tab 3：合约基差比较
# ══════════════════════════════════════════════════════════════
def tab3():
    st.subheader("🔄 合约基差比较")

    cached = set(get_cached_contracts())
    all_ct = sorted(ALL_CONTRACTS)
    spot_dict, spot_msg = load_spot(str(SPOT_PATH))
    fut_update_date = get_latest_futures_date()

    col_ctrl, col_chart = st.columns([1, 3.5])

    with col_ctrl:
        # 合约多选
        default_t3 = [c for c in ["LH2509","LH2609"] if c in all_ct] or [all_ct[-1]]
        contracts = st.multiselect("📋 合约选择（多选）", options=all_ct, default=default_t3,
            format_func=lambda c: f"{ct_display(c)} {'📁' if c in cached else '🆕'}", key="t3_ct")
        if not contracts: contracts = [all_ct[-1]]

        mode = st.selectbox("🔄 比较模式", options=["同比（自然日对齐）","距离交易日对齐"], key="t3_mode")

        # ★ 包含所有实际区域 + 四个指标，多选，默认"全国均价基差"
        ref_ct = contracts[0]; ref_regions = get_regions(ref_ct)
        available_regions = [r for r in ref_regions if r in spot_dict] or ref_regions
        item_opts = list(available_regions) + ["─── 汇总指标 ───", "📊 全国均价基差", "🔴 最大基差", "🟢 最小基差", "🟣 基差平均值"]
        sel_items = st.multiselect("📐 地区与指标", options=item_opts,
            default=["📊 全国均价基差"], key="t3_items")
        if not sel_items: sel_items = ["📊 全国均价基差"]

    with col_chart:
        if len(contracts) < 1:
            st.info("ℹ️ 请选择至少 1 个合约"); return

        if mode == "同比（自然日对齐）":
            _tab3_calendar(contracts, spot_dict, ref_regions, sel_items, fut_update_date or "")
        else:
            _tab3_delivery(contracts, spot_dict, ref_regions, available_regions, sel_items, fut_update_date or "")

def _tab3_calendar(contracts, spot_dict, ref_regions, sel_items, data_date):
    """同比模式：按 item 分别生成 traces"""
    tmon = ct_month(contracts[0])
    same_month = [c for c in ALL_CONTRACTS if ct_month(c)==tmon]
    avail = []
    for c in same_month:
        df, _ = load_futures(c)
        if df is not None and not df.empty: avail.append(c)
    if len(avail) < 1:
        st.warning(f"⚠️ {tmon}月合约暂无可用的历史数据"); return
    st.info(f"📌 {tmon}月合约，共 {len(avail)} 个可用：{'、'.join(avail)}")

    series: Dict[str, pd.DataFrame] = {}
    # ★ 修复：用 (month, day) 聚合，避免闰年/非闰年 doy 混乱
    md_collectors: Dict[str, Dict[Tuple[int, int], List[int]]] = defaultdict(lambda: defaultdict(list))

    for sel_item in sel_items:
        is_national = "全国均价" in sel_item
        is_max = "最大基差" in sel_item
        is_min = "最小基差" in sel_item
        is_avg = "基差平均值" in sel_item
        is_region = sel_item in ref_regions

        if is_national: item_short = "全国均价"
        elif is_max: item_short = "最大"
        elif is_min: item_short = "最小"
        elif is_avg: item_short = "均值"
        else: item_short = sel_item

        for c in avail:
            fut_df, _ = load_futures(c)
            if fut_df is None or fut_df.empty: continue
            fut_df = fut_df.sort_values("date").reset_index(drop=True)
            df_basis = None
            if is_region and sel_item in spot_dict:
                df_basis = calc_basis(c, sel_item, spot_dict[sel_item], fut_df)
            elif is_national:
                df_basis = calc_national_basis(spot_dict, fut_df)
            elif is_max:
                _, mx, _, _ = get_summary_series(c, spot_dict, fut_df, ref_regions); df_basis = mx
            elif is_min:
                _, _, mn, _ = get_summary_series(c, spot_dict, fut_df, ref_regions); df_basis = mn
            elif is_avg:
                _, _, _, av = get_summary_series(c, spot_dict, fut_df, ref_regions); df_basis = av
            if df_basis is None or df_basis.empty: continue
            df_basis["year"] = df_basis["date"].dt.year
            df_basis["doy"] = df_basis["date"].dt.dayofyear
            # ★ 修复：传入 source_year，正确处理闰年/非闰年
            df_basis["plot_date"] = df_basis.apply(
                lambda r: _doy_to_date(int(r["doy"]), int(r["year"])), axis=1)
            for yr, grp in df_basis.groupby("year"):
                grp = grp.sort_values("doy").copy()
                label = _make_trace_label(c, yr, item_short)
                series[label] = grp
                for _, row in grp.iterrows():
                    md_collectors[item_short][(row["date"].month, row["date"].day)].append(row["basis"])

    if not series: st.warning("⚠️ 无可用数据"); return

    # ★ 修复：每个 item 一条历史均值线，用 (month, day) 构造正确的 plot_date
    for item_short, mdc in md_collectors.items():
        avg_rows = [{"doy": m*100+d, "basis": int(round(np.mean(v))),
                      "plot_date": pd.Timestamp(year=2020, month=m, day=d)}
                    for (m, d), v in sorted(mdc.items()) if v]
        if avg_rows:
            series[f"历史均值-{item_short}"] = pd.DataFrame(avg_rows).sort_values("doy")

    st.plotly_chart(fig_calendar_comparison(series, tmon, data_date), use_container_width=True)

def _tab3_delivery(contracts, spot_dict, ref_regions, available_regions, sel_items, data_date):
    """交易日对齐模式：按 item 分别生成 traces"""
    series: Dict[str, pd.DataFrame] = {}

    for sel_item in sel_items:
        is_national = "全国均价" in sel_item
        is_max = "最大基差" in sel_item
        is_min = "最小基差" in sel_item
        is_avg = "基差平均值" in sel_item
        is_region = sel_item in ref_regions

        if is_national: item_short = "全国均价"
        elif is_max: item_short = "最大"
        elif is_min: item_short = "最小"
        elif is_avg: item_short = "均值"
        else: item_short = sel_item

        for c in contracts:
            fut_df, _ = load_futures(c)
            if fut_df is None or fut_df.empty:
                st.warning(f"⚠️ {c} 加载失败"); continue
            try:
                yr = int(f"20{c[2:4]}"); mo = int(c[4:6])
                delivery_day = pd.Timestamp(year=yr, month=mo, day=15)
            except Exception: continue
            df_basis = None
            if is_region and sel_item in spot_dict:
                df_basis = calc_basis(c, sel_item, spot_dict[sel_item], fut_df)
            elif is_national:
                df_basis = calc_national_basis(spot_dict, fut_df)
            elif is_max:
                _, mx, _, _ = get_summary_series(c, spot_dict, fut_df, ref_regions); df_basis = mx
            elif is_min:
                _, _, mn, _ = get_summary_series(c, spot_dict, fut_df, ref_regions); df_basis = mn
            elif is_avg:
                _, _, _, av = get_summary_series(c, spot_dict, fut_df, ref_regions); df_basis = av
            if df_basis is None or df_basis.empty: continue
            df_basis["days"] = (delivery_day - df_basis["date"]).dt.days
            df_basis = df_basis[df_basis["days"] >= 0].copy()
            series[f"{c} {item_short}"] = df_basis

    if not series: st.warning("⚠️ 无可用数据"); return
    st.plotly_chart(fig_delivery_comparison(series, data_date), use_container_width=True)

# ══════════════════════════════════════════════════════════════
# Tab 4：合约价差比较
# ══════════════════════════════════════════════════════════════
def tab4():
    st.subheader("📉 合约价差比较")

    cached = set(get_cached_contracts())
    fut_update_date = get_latest_futures_date()

    col_ctrl, col_chart = st.columns([1, 3.5])

    with col_ctrl:
        months = ["01","03","05","07","09","11"]
        ma = st.selectbox("合约 A 月份", months, index=months.index("09"), format_func=lambda m: f"{m}月", key="t4_ma")
        mb = st.selectbox("合约 B 月份", months, index=months.index("07"), format_func=lambda m: f"{m}月", key="t4_mb")

    with col_chart:
        if ma == mb: st.warning("⚠️ 请选择不同的月份"); return

        valid_years, skipped_years = [], []
        for y in range(21, 28):
            ca, cb = f"LH{y}{ma}", f"LH{y}{mb}"
            if ca not in ALL_CONTRACTS or cb not in ALL_CONTRACTS: continue
            ca_ok = ca in cached or _csv_path(ca).exists()
            cb_ok = cb in cached or _csv_path(cb).exists()
            if ca_ok and cb_ok: valid_years.append(y)
            elif ca_ok or cb_ok: skipped_years.append(y)

        if not valid_years:
            for y in range(21, 28):
                ca, cb = f"LH{y}{ma}", f"LH{y}{mb}"
                if ca not in ALL_CONTRACTS or cb not in ALL_CONTRACTS: continue
                dfa, _ = load_futures(ca); dfb, _ = load_futures(cb)
                if dfa is not None and not dfa.empty and dfb is not None and not dfb.empty: valid_years.append(y)
                elif (dfa is not None and not dfa.empty) or (dfb is not None and not dfb.empty): skipped_years.append(y)

        if not valid_years:
            st.warning(f"⚠️ 暂无同时存在 {ma}月 和 {mb}月 合约的年份数据"); return

        info = f"✅ {len(valid_years)} 个有效年份：{'、'.join('20'+str(y) for y in sorted(valid_years))}"
        if skipped_years: info += f" ｜ ⚠️ 跳过：{'、'.join('20'+str(y) for y in sorted(skipped_years))}"
        st.info(info)

        spreads, failed = {}, []
        # ★ 修复：用 (month, day) 聚合，正确处理闰年/非闰年
        spread_collector = defaultdict(list)  # key: (month, day) tuple

        for y in valid_years:
            ca, cb = f"LH{y}{ma}", f"LH{y}{mb}"
            dfa, _ = load_futures(ca); dfb, _ = load_futures(cb)
            if dfa is None or dfa.empty or dfb is None or dfb.empty: failed.append(y); continue
            ac = dfa.set_index("date")["close"]; bc = dfb.set_index("date")["close"]
            cm = ac.index.intersection(bc.index)
            if len(cm) == 0: failed.append(y); continue
            sv = bc[cm] - ac[cm]; doy = cm.dayofyear
            # ★ 修复：plot_date 直接从实际日期构造，不依赖 doy
            df_sp = pd.DataFrame({"date":cm,"spread":[int(round(v)) for v in sv.values],
                "day_of_year":doy,
                "plot_date":[pd.Timestamp(year=2020, month=d.month, day=d.day) for d in cm],
                "trade_year":cm.year}).sort_values("date")
            contract_pair_year = f"20{y:02d}"
            for trade_yr, grp in df_sp.groupby("trade_year"):
                ty_str = str(trade_yr)
                label = f"{cb[2:]}-{ca[2:]}({ty_str})" if ty_str != contract_pair_year else f"{cb[2:]}-{ca[2:]}"
                spreads[label] = grp.sort_values("day_of_year")
                for _, row in grp.iterrows():
                    d = row["date"]
                    spread_collector[(d.month, d.day)].append(row["spread"])

        if failed: st.warning(f"⚠️ 计算失败：{'、'.join('20'+str(y) for y in failed)}")
        if not spreads: st.warning("⚠️ 无法计算价差"); return

        # ★ 修复：用 (month, day) 构造均值 plot_date
        avg_rows = [{"day_of_year": m*100+d, "spread": int(round(np.mean(v))),
                      "plot_date": pd.Timestamp(year=2020, month=m, day=d)}
                    for (m, d), v in sorted(spread_collector.items()) if v]
        if avg_rows: spreads["历史均值"] = pd.DataFrame(avg_rows).sort_values("day_of_year")

        st.plotly_chart(fig_spread_season(spreads, ma, mb, fut_update_date or ""), use_container_width=True)

        with st.expander("📋 价差统计表"):
            stats = []
            for label, df in spreads.items():
                if df.empty or "历史均值" in label: continue
                stats.append({"合约对":label,"均值":f"{int(round(df['spread'].mean())):+,}",
                    "最大":f"{df['spread'].max():+,}","最小":f"{df['spread'].min():+,}",
                    "标准差":f"{int(round(df['spread'].std())):,}","数据点":len(df)})
            st.dataframe(pd.DataFrame(stats), use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════
# 主入口
# ══════════════════════════════════════════════════════════════
def main():
    # ── 全局 CSS ──
    st.markdown("""<style>
    /* 指标卡片 */
    .metric-card {
        background: #f8f9fa; border-radius: 12px; padding: 14px 8px;
        text-align: center; border: 1px solid #e9ecef; margin: 2px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.06); transition: box-shadow 0.2s;
    }
    .metric-card:hover { box-shadow: 0 4px 12px rgba(0,0,0,0.10); }
    .metric-card .mlabel { font-size: 12px; color: #6c757d; margin-bottom: 4px; }
    .metric-card .mvalue { font-size: 24px; font-weight: 700; margin-bottom: 2px; }
    .metric-card .munit  { font-size: 12px; color: #adb5bd; }
    /* 底部信息 */
    .footer-info { font-size: 13px; color: #95a5a6; text-align: center; margin-top: 6px; }
    .footer-info b { color: #7f8c8d; }
    </style>
    """, unsafe_allow_html=True)

    # ── 标题区（自定义 HTML/CSS） ──
    st.markdown("""
    <div style="text-align: center; padding: 1rem 0 0.5rem 0;">
        <h1 style="font-size: 2.8rem; font-weight: 700; color: #1a1a2e; margin-bottom: 0.2rem;">
            🐷 生猪期货基差分析平台
        </h1>
        <p style="font-size: 1rem; color: #666; letter-spacing: 1px;">
            数据来源：涌益咨询现货数据 ｜ 大连商品交易所期货数据 ｜ 大商所升贴水公告
        </p>
    </div>
    <hr style="border: none; border-top: 1px solid #e9ecef; margin: 0.5rem 0 1.5rem 0;">
    """, unsafe_allow_html=True)

    # ── 四个 Tab ──
    t1, t2, t3, t4 = st.tabs(["📊 当日基差分布", "📈 单合约基差走势", "🔄 合约基差比较", "📉 合约价差比较"])

    with t1: tab1()
    with t2: tab2()
    with t3: tab3()
    with t4: tab4()

    # ── 页面底部信息 ──
    st.markdown("---")
    fut_update_date = get_latest_futures_date() or "加载中…"
    cached = get_cached_contracts()
    cache_str = f"<b>{'、'.join(cached)}</b>" if cached else "暂无"
    st.markdown(f"""
    <p class="footer-info">
        📅 期货数据更新日期：<b>{fut_update_date}</b> &nbsp;｜&nbsp;
        📦 已缓存合约：<b>{len(cached)}</b> 个（{cache_str}）
    </p>
    """, unsafe_allow_html=True)

    # ── 操作按钮 ──
    c1, c2, c3 = st.columns([1, 1, 8])
    with c1:
        if st.button("🔄 刷新数据", use_container_width=True, key="main_refresh"):
            st.cache_data.clear(); st.rerun()
    with c2:
        if st.button("🗑️ 清除缓存", use_container_width=True, key="main_clear"):
            st.cache_data.clear()
            if FUTURES_DIR.exists(): shutil.rmtree(FUTURES_DIR); FUTURES_DIR.mkdir()
            st.rerun()

    st.caption("⚠️ 免责声明：本平台数据仅供参考，不构成任何投资建议。投资有风险，入市需谨慎。")

if __name__ == "__main__":
    main()
