#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
生猪期货分析平台 (Streamlit)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
无侧边栏设计——所有控件位于各 Tab 内部。
Tab 1 — 当日基差分布（柱状图 + 四指标卡片）
Tab 2 — 单合约基差走势（区域色板 + 汇总指标固定色）
Tab 3 — 合约基差比较（同比 / 交易日对齐，颜色按合约年份）
Tab 4 — 合约价差比较（月份选择，颜色按合约年份）
Tab 5 — 持仓与成交分析（双轴图 + 前20期货公司多空持仓）
Tab 6 — 季节性持仓对比（同月份合约跨年成交量/持仓量/净持仓）
Tab 7 — 技术分析（K线图 + MA/布林带 + MACD/RSI/KDJ + 文字结论）
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
HOLDINGS_DIR = DATA_DIR / "holdings"
DATA_DIR.mkdir(exist_ok=True)
FUTURES_DIR.mkdir(exist_ok=True)
HOLDINGS_DIR.mkdir(exist_ok=True)
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

def _get_global_latest_date() -> Optional[pd.Timestamp]:
    """获取所有合约中最新的交易日"""
    latest = None
    for f in FUTURES_DIR.glob("LH*.csv"):
        try:
            df = pd.read_csv(f)
            if "date" not in df.columns or df.empty: continue
            d = pd.to_datetime(df["date"].max())
            if latest is None or d > latest: latest = d
        except Exception: pass
    return latest

@st.cache_data(ttl=300)
def get_active_contracts() -> List[str]:
    """动态识别当前上市合约，并确保兜底列表始终包含"""
    # 硬编码保底列表：当前正在交易的合约
    FALLBACK_ACTIVE = ['LH2607', 'LH2609', 'LH2611', 'LH2701', 'LH2703', 'LH2705']
    global_latest = _get_global_latest_date()
    if global_latest is None:
        return [c for c in ALL_CONTRACTS if c in FALLBACK_ACTIVE]

    active = set(FALLBACK_ACTIVE)  # ★ 始终包含兜底列表
    for f in sorted(FUTURES_DIR.glob("LH*.csv")):
        ct = f.stem
        try:
            df = pd.read_csv(f)
            if df.empty or "date" not in df.columns: continue
            df["date"] = pd.to_datetime(df["date"])
            latest_row = df.sort_values("date").iloc[-1]
            days_behind = (global_latest.date() - latest_row["date"].date()).days
            if days_behind > 7:
                continue
            has_vol = "volume" in df.columns and int(latest_row["volume"]) > 0
            oi_col = "open_interest" if "open_interest" in df.columns else ("hold" if "hold" in df.columns else None)
            has_oi = oi_col and int(latest_row[oi_col]) > 0
            if has_vol or has_oi:
                active.add(ct)
        except Exception:
            continue
    # 过滤：只保留存在于ALL_CONTRACTS中的合约
    return sorted([c for c in active if c in ALL_CONTRACTS])

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

@st.cache_data(ttl=300)
def get_latest_trade_date() -> Optional[pd.Timestamp]:
    """获取所有合约CSV中最新的交易日（全局最大值）"""
    return _get_global_latest_date()

def _get_row_at_md(df, target_month: int, target_day: int):
    """在DataFrame中查找指定月/日的行，若不存在则取目标日期之前最近的行，兜底取最后一行"""
    if df is None or df.empty:
        return None
    mask = (df['date'].dt.month == target_month) & (df['date'].dt.day == target_day)
    match = df[mask]
    if not match.empty:
        return match.iloc[-1]
    # 取目标月日之前的最近数据
    target_ordinal = target_month * 100 + target_day
    df_copy = df.copy()
    df_copy['_md'] = df_copy['date'].dt.month * 100 + df_copy['date'].dt.day
    before = df_copy[df_copy['_md'] <= target_ordinal]
    if not before.empty:
        return before.iloc[-1]
    return df.iloc[-1]

@st.cache_data(ttl=3600)
def get_spot_data_date() -> str:
    """从现货Excel文件名或内部数据提取最新日期"""
    path = SPOT_PATH
    # 1. 从文件名提取日期（如"2026年6月29日涌益咨询日度数据.xlsx"）
    fname = path.stem
    m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日", fname)
    if m:
        return f"{m.group(1)}年{int(m.group(2)):02d}月{int(m.group(3)):02d}日"
    # 2. 从Excel内部日期列获取最新日期
    try:
        xls = pd.ExcelFile(path)
        if len(xls.sheet_names) > 0:
            df = pd.read_excel(xls, sheet_name=0, header=None)
            for col in range(2, min(df.shape[1], 200)):
                for ridx in range(min(df.shape[0], 5)):
                    v = df.iloc[ridx, col]
                    if pd.notna(v):
                        try:
                            dt = pd.to_datetime(v)
                            if dt.year > 2000:
                                return _cn(dt)
                        except Exception:
                            pass
    except Exception:
        pass
    return "无现货数据"

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
# 技术指标计算 (Tab 6)
# ══════════════════════════════════════════════════════════════
def calculate_technicals(df: pd.DataFrame) -> pd.DataFrame:
    """计算完整技术指标：MA5/10/20/60, 布林带, MACD, RSI14, KDJ"""
    if df is None or df.empty:
        return df
    close = df["close"].astype(float)
    high = df["high"].astype(float)
    low = df["low"].astype(float)

    # ── 移动均线 ──
    df["ma5"] = close.rolling(5).mean()
    df["ma10"] = close.rolling(10).mean()
    df["ma20"] = close.rolling(20).mean()
    df["ma60"] = close.rolling(60).mean()

    # ── 布林带 (中轨=MA20, 上下轨=MA20±2σ) ──
    df["bb_mid"] = df["ma20"]
    std20 = close.rolling(20).std()
    df["bb_up"] = df["bb_mid"] + 2 * std20
    df["bb_low"] = df["bb_mid"] - 2 * std20

    # ── MACD (12, 26, 9) ──
    ema12 = close.ewm(span=12, adjust=False).mean()
    ema26 = close.ewm(span=26, adjust=False).mean()
    df["dif"] = ema12 - ema26
    df["dea"] = df["dif"].ewm(span=9, adjust=False).mean()
    df["macd_hist"] = 2 * (df["dif"] - df["dea"])

    # ── RSI14 ──
    delta = close.diff()
    gain = delta.clip(lower=0)
    loss = (-delta).clip(lower=0)
    avg_gain = gain.ewm(alpha=1/14, adjust=False).mean()
    avg_loss = loss.ewm(alpha=1/14, adjust=False).mean()
    rs = avg_gain / avg_loss.replace(0, np.nan)
    df["rsi14"] = 100 - (100 / (1 + rs))

    # ── KDJ (9, 3, 3) ──
    n = 9
    lowest_low = low.rolling(n).min()
    highest_high = high.rolling(n).max()
    rsv = ((close - lowest_low) / (highest_high - lowest_low).replace(0, np.nan)) * 100
    k_vals, d_vals, j_vals = [], [], []
    k_prev, d_prev = 50.0, 50.0
    for r in rsv:
        if pd.isna(r):
            k_vals.append(np.nan); d_vals.append(np.nan); j_vals.append(np.nan)
        else:
            k = 2/3 * k_prev + 1/3 * r
            d = 2/3 * d_prev + 1/3 * k
            j = 3 * k - 2 * d
            k_vals.append(k); d_vals.append(d); j_vals.append(j)
            k_prev, d_prev = k, d
    df["kdj_k"] = k_vals
    df["kdj_d"] = d_vals
    df["kdj_j"] = j_vals

    return df


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
    fig.update_yaxes(autorange=True)
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
    fig.update_yaxes(autorange=True)
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
    fig.update_yaxes(autorange=True)
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
    fig.update_yaxes(autorange=True)
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
    fig.update_yaxes(autorange=True)
    return fig

# ══════════════════════════════════════════════════════════════
# 统一结论展示组件
# ══════════════════════════════════════════════════════════════

def display_conclusion(title: str, items: list, sentiment: str = "neutral"):
    """
    统一展示结论的卡片组件
    sentiment: 'bullish'（偏多，红色边框）, 'bearish'（偏空，绿色边框）, 'neutral'（中性，灰色边框）
    """
    border_color = {
        'bullish': '#E74C3C',
        'bearish': '#27AE60',
        'neutral': '#95A5A6'
    }.get(sentiment, '#95A5A6')

    if not items:
        items = ["数据样本不足，无法生成有效结论。"]

    items_html = ''.join([f'<li style="margin: 4px 0; line-height: 1.7;">{item}</li>' for item in items])

    st.markdown(f"""
    <div style="background-color: #f8f9fa; padding: 16px 20px; border-radius: 8px; border-left: 4px solid {border_color}; margin: 12px 0;">
        <div style="font-weight: 600; font-size: 1.05rem; color: #1a1a2e; margin-bottom: 6px;">{title}</div>
        <ul style="margin: 4px 0; padding-left: 20px; list-style-type: disc; color: #333; font-size: 0.95rem;">
            {items_html}
        </ul>
    </div>
    """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
# 自动结论生成函数（返回 title, items, sentiment 供 display_conclusion 使用）
# ══════════════════════════════════════════════════════════════

def _gen_tab3_conclusion_calendar(series: Dict, tmon: str, sel_items: list, contracts: list = None):
    """Tab 3 同比模式：使用图表中的历史均值线，对齐LATEST_TRADE_DATE"""
    if not series: return None
    ltd = get_latest_trade_date()
    if ltd is None:
        return None
    target_m, target_d = ltd.month, ltd.day
    analysis_date = _cn(ltd)

    # 仅保留用户选择的合约中date.max()最大的年份条目
    if contracts:
        selected_set = set(contracts)
        best_per_contract = {}
        for label, df in series.items():
            if "历史均值" in label: continue
            ct_code = _extract_contract_code(label)
            if ct_code not in selected_set: continue
            if df.empty: continue
            if ct_code not in best_per_contract or df["date"].max() > best_per_contract[ct_code][1]["date"].max():
                best_per_contract[ct_code] = (label, df)
        non_avg = {v[0]: v[1] for v in best_per_contract.values()}
    else:
        non_avg = {k: v for k, v in series.items() if "历史均值" not in k}
    avg_items = {k: v for k, v in series.items() if "历史均值" in k}

    if len(non_avg) < 1:
        return None

    items = []
    sentiment = "neutral"

    for label, df in non_avg.items():
        if df.empty: continue
        ct_code = _extract_contract_code(label)

        cur_row = _get_row_at_md(df, target_m, target_d)
        if cur_row is None: continue
        cur_basis = int(cur_row["basis"])

        items.append(f"分析合约：{ct_code}")
        items.append(f"• 当前基差（{_cn_md(ltd)}）：{cur_basis:+,}元/吨")

        # ★ 直接使用图表中的历史均值线，不自行计算
        hist_avg = None
        for alabel, adf in avg_items.items():
            if adf.empty: continue
            # avg_items的plot_date已映射到2020年，直接用target月/日查询
            arow = adf[(adf["plot_date"].dt.month == target_m) & (adf["plot_date"].dt.day == target_d)]
            if arow.empty:
                arow = _get_row_at_md(adf, target_m, target_d)  # fallback
                if arow is not None:
                    hist_avg = int(arow["basis"])
            else:
                hist_avg = int(arow["basis"].iloc[-1])
            break  # 取第一条历史均值线

        if hist_avg is not None:
            items.append(f"• 历史同期均值：{hist_avg:+,}元/吨")
            deviation = cur_basis - hist_avg
            direction = "偏高" if deviation > 0 else "偏低"
            items.append(f"• 当前基差较历史同期均值{direction}{abs(deviation):,}元/吨")

        # 近期变化方向
        if len(df) >= 20:
            recent_start = int(df["basis"].iloc[-20])
            recent_end = int(df["basis"].iloc[-1])
            change = recent_end - recent_start
            if abs(change) < 20:
                dir_str = "震荡"
            elif change > 0:
                dir_str = "走扩"
            else:
                dir_str = "收敛"
            items.append(f"• 近期（近20个交易日）基差变化方向：{dir_str}（从{recent_start:+,}变化至{recent_end:+,}）")
        elif len(df) >= 5:
            recent_start = int(df["basis"].iloc[-5])
            recent_end = int(df["basis"].iloc[-1])
            change = recent_end - recent_start
            if abs(change) < 20:
                dir_str = "震荡"
            elif change > 0:
                dir_str = "走扩"
            else:
                dir_str = "收敛"
            items.append(f"• 近期（近5个交易日）基差变化方向：{dir_str}（从{recent_start:+,}变化至{recent_end:+,}）")

        # 季节性（从历史均值线提取）
        if avg_items:
            for alabel, adf in avg_items.items():
                if adf.empty: continue
                item_name = alabel.replace("历史均值-", "")
                peak_row = adf.loc[adf["basis"].idxmax()]
                trough_row = adf.loc[adf["basis"].idxmin()]
                peak_md = f"{pd.to_datetime(peak_row['plot_date']).month:02d}月{pd.to_datetime(peak_row['plot_date']).day:02d}日"
                trough_md = f"{pd.to_datetime(trough_row['plot_date']).month:02d}月{pd.to_datetime(trough_row['plot_date']).day:02d}日"
                items.append(f"• 历史季节性：{item_name}通常{peak_md}见顶（{int(peak_row['basis']):+,}），{trough_md}见底（{int(trough_row['basis']):+,}）")
                break

        # 判断
        if hist_avg is not None:
            if cur_basis > hist_avg + 200:
                judgment = "当前基差处于历史同期偏高水平"
                sentiment = "bullish"
            elif cur_basis < hist_avg - 200:
                judgment = "当前基差处于历史同期偏低水平"
                sentiment = "bearish"
            else:
                judgment = "当前基差处于历史同期均值附近"
                sentiment = "neutral"
        else:
            judgment = f"当前基差{cur_basis:+,}元/吨"
        items.append(f"• 判断：{judgment}")

        break

    return (f"📊 基差分析结论（分析日期：{analysis_date}）", items, sentiment)


def _gen_tab3_conclusion_delivery(series: Dict, contracts: list, active_cts: list = None):
    """Tab 3 交易日对齐模式：结论使用LATEST_TRADE_DATE取值"""
    if not series or len(series) < 2: return None
    ltd = get_latest_trade_date()
    if ltd is None: return None
    target_m, target_d = ltd.month, ltd.day
    analysis_date = _cn(ltd)

    # 仅保留用户选择的合约中date跨度最大的条目
    selected_set = set(contracts) if contracts else set()
    if selected_set:
        best_per = {}
        for label, df in series.items():
            if "历史均值" in label: continue
            ct_code = _extract_contract_code(label)
            if ct_code not in selected_set: continue
            if df.empty: continue
            if ct_code not in best_per or df["date"].max() > best_per[ct_code][1]["date"].max():
                best_per[ct_code] = (label, df)
        active_series = {v[0]: v[1] for v in best_per.values()}
    else:
        active_series = {k: v for k, v in series.items() if "历史均值" not in k}

    if not active_series:
        return None

    items = []
    sentiment = "neutral"

    near_delivery = []
    for label, df in active_series.items():
        if df.empty: continue
        near = df[df["days"] <= 30]
        if not near.empty:
            near_delivery.append((label, int(near["basis"].iloc[-1]), near["days"].iloc[-1]))

    if near_delivery:
        near_delivery.sort(key=lambda x: x[2])
        closest_label, closest_basis, closest_days = near_delivery[0]
        ct_code = _extract_contract_code(closest_label)
        items.append(f"分析合约：{ct_code}")
        items.append(f"• 距交割{closest_days}天，基差{closest_basis:+,}元/吨")
        if abs(closest_basis) < 200:
            items.append("• 基差接近零轴，期现价格趋于一致")
            sentiment = "neutral"
        else:
            items.append(f"• 基差绝对值仍较大（{abs(closest_basis)}元/吨），偏离零轴")
            sentiment = "bearish" if closest_basis < 0 else "bullish"

    # 排序（使用date-max条目的最新值）
    all_latest = [(l, int(df["basis"].iloc[-1])) for l, df in active_series.items() if not df.empty]
    if len(all_latest) >= 2:
        all_latest.sort(key=lambda x: x[1], reverse=True)
        items.append(f"• 用户选择合约基差排序：{' > '.join(f'{_extract_contract_code(l)}({v:+,})' for l, v in all_latest)}")

    if not items: return None
    return (f"📊 交易日对齐分析结论（分析日期：{analysis_date}）", items, sentiment)


def _gen_tab4_conclusion(spreads: Dict, ma: str, mb: str, active_cts: list = None):
    """Tab 4 价差季节图：使用图表中的历史均值线，对齐LATEST_TRADE_DATE"""
    if not spreads: return None
    ltd = get_latest_trade_date()
    if ltd is None: return None
    target_m, target_d = ltd.month, ltd.day
    analysis_date = _cn(ltd)

    non_avg_all = {k: v for k, v in spreads.items() if "历史均值" not in k}
    avg_df = spreads.get("历史均值")

    current_year_2d = str(datetime.now().year)[2:]
    non_avg = {k: v for k, v in non_avg_all.items() if k[:2] == current_year_2d}

    if not non_avg: return None

    items = []
    sentiment = "neutral"

    latest_pair = max(non_avg.items(), key=lambda x: x[1]["date"].max() if not x[1].empty else pd.Timestamp("2000"))
    if not latest_pair[1].empty:
        cur_row = _get_row_at_md(latest_pair[1], target_m, target_d)
        if cur_row is None: return None
        cur_spread = int(cur_row["spread"])
        items.append(f"分析价差对：{latest_pair[0]}")
        items.append(f"• 当前价差（{_cn_md(ltd)}）：{cur_spread:+,}元/吨")

        # ★ 直接使用图表中的历史均值线
        if avg_df is not None and not avg_df.empty:
            arow = avg_df[(avg_df["plot_date"].dt.month == target_m) & (avg_df["plot_date"].dt.day == target_d)]
            if not arow.empty:
                hist_avg = int(arow["spread"].iloc[-1])
                items.append(f"• 历史同期均值：{hist_avg:+,}元/吨")
                deviation = cur_spread - hist_avg
                direction = "偏高" if deviation > 0 else "偏低"
                items.append(f"• 当前价差较历史同期均值{direction}{abs(deviation):,}元/吨")

        # 近期变化
        if len(latest_pair[1]) >= 20:
            recent_end = int(latest_pair[1]["spread"].iloc[-1])
            recent_start = int(latest_pair[1]["spread"].iloc[-20])
            change = recent_end - recent_start
            if abs(change) < 20:
                dir_str = "震荡"
            elif change > 0:
                dir_str = "走扩"
            else:
                dir_str = "收敛"
            items.append(f"• 近期（近20个交易日）价差变化方向：{dir_str}（从{recent_start:+,}变化至{recent_end:+,}）")

    # 季节性
    if avg_df is not None and not avg_df.empty:
        peak = avg_df.loc[avg_df["spread"].idxmax()]
        trough = avg_df.loc[avg_df["spread"].idxmin()]
        items.append(f"• 季节性规律：历史均值在{_cn_md(peak['plot_date'])}见顶（{int(peak['spread']):+,}），{_cn_md(trough['plot_date'])}见底（{int(trough['spread']):+,}）")

    # 核心判断
    if abs(cur_spread) < 50:
        judgment = "价差接近零轴，跨期价差暂无方向"
    elif cur_spread > 0:
        judgment = "价差为正，远期升水格局"
    else:
        judgment = "价差为负，远期贴水格局"
    items.append(f"• 判断：{judgment}")

    if not items: return None
    return (f"💰 价差分析结论（分析日期：{analysis_date}）", items, sentiment)


def _gen_tab6_conclusion(vol_data: Dict, oi_data: Dict, net_data: Dict, sel_month: str, active_cts: list = None):
    """Tab 6 季节性持仓对比：使用图表中的历史均值线，对齐LATEST_TRADE_DATE"""
    ltd = get_latest_trade_date()
    if ltd is None: return None
    target_m, target_d = ltd.month, ltd.day
    analysis_date = _cn(ltd)

    current_year_2d = str(datetime.now().year)[2:]
    items = []
    sentiment = "neutral"
    bull_score, bear_score = 0, 0

    def _is_current_year(label: str) -> bool:
        return label[:2] == current_year_2d

    # 成交量：当前年份在LATEST_TRADE_DATE月/日取值 + 历史均值线
    non_avg_vol_cur = {k: v for k, v in vol_data.items() if "历史均值" not in k and not v.empty and _is_current_year(k)}
    avg_vol_df = vol_data.get("历史均值")
    if non_avg_vol_cur:
        latest_vol = max(non_avg_vol_cur.items(), key=lambda x: x[1]["date"].max())
        if not latest_vol[1].empty:
            vr = _get_row_at_md(latest_vol[1], target_m, target_d)
            if vr is not None:
                cur_v = int(vr["volume"])
                cname = latest_vol[0].split()[0] if ' ' in latest_vol[0] else latest_vol[0]
                items.append(f"分析合约：{cname}")
                items.append(f"• 当前成交量（{_cn_md(ltd)}）：{cur_v:,}手")
                # ★ 使用图表中的历史均值线
                if avg_vol_df is not None and not avg_vol_df.empty:
                    arow = avg_vol_df[(avg_vol_df["plot_date"].dt.month == target_m) & (avg_vol_df["plot_date"].dt.day == target_d)]
                    if not arow.empty:
                        hist_avg_v = int(arow["volume"].iloc[-1])
                        pct_v = (cur_v - hist_avg_v) / hist_avg_v * 100 if hist_avg_v > 0 else 0
                        direction = "偏高" if pct_v > 15 else ("偏低" if pct_v < -15 else "持平")
                        items.append(f"• 历史同期均值：{hist_avg_v:,}手，当前{direction}{abs(pct_v):.0f}%")

    # 持仓量：当前年份 + 历史均值线
    non_avg_oi_cur = {k: v for k, v in oi_data.items() if "历史均值" not in k and not v.empty and _is_current_year(k)}
    avg_oi_df = oi_data.get("历史均值")
    oi_pct = 0
    if non_avg_oi_cur:
        latest_oi = max(non_avg_oi_cur.items(), key=lambda x: x[1]["date"].max())
        if not latest_oi[1].empty:
            oir = _get_row_at_md(latest_oi[1], target_m, target_d)
            if oir is not None:
                cur_o = int(oir["open_interest"])
                if avg_oi_df is not None and not avg_oi_df.empty:
                    arow = avg_oi_df[(avg_oi_df["plot_date"].dt.month == target_m) & (avg_oi_df["plot_date"].dt.day == target_d)]
                    if not arow.empty:
                        hist_avg_o = int(arow["open_interest"].iloc[-1])
                        oi_pct = (cur_o - hist_avg_o) / hist_avg_o * 100 if hist_avg_o > 0 else 0
                        direction_o = "偏高" if oi_pct > 15 else ("偏低" if oi_pct < -15 else "持平")
                        items.append(f"• 当前持仓量（{_cn_md(ltd)}）：{cur_o:,}手，较历史同期{direction_o}{abs(oi_pct):.0f}%")
                # 近期变化
                if len(latest_oi[1]) >= 20:
                    recent_o = int(latest_oi[1]["open_interest"].iloc[-20])
                    if cur_o > recent_o * 1.05:
                        trend = "增仓趋势"
                    elif cur_o < recent_o * 0.95:
                        trend = "减仓趋势"
                    else:
                        trend = "持仓平稳"
                    items.append(f"• 近期（近20个交易日）持仓变化：{trend}")
                if oi_pct > 20:
                    bull_score += 1

    # 净持仓：当前年份（net_data只有plot_date列，无date列）
    non_avg_net_cur = {k: v for k, v in net_data.items() if "历史均值" not in k and not v.empty and _is_current_year(k)}
    if non_avg_net_cur:
        latest_net = max(non_avg_net_cur.items(), key=lambda x: x[1]["plot_date"].max())
        if not latest_net[1].empty:
            # net_data无date列，直接用plot_date按月/日查询（plot_date已映射到2020年）
            nr_df = latest_net[1]
            nr_match = nr_df[(nr_df["plot_date"].dt.month == target_m) & (nr_df["plot_date"].dt.day == target_d)]
            if not nr_match.empty:
                cur_n = int(nr_match["net_position"].iloc[-1])
            else:
                cur_n = int(nr_df["net_position"].iloc[-1])
            bias = "净多" if cur_n > 0 else "净空"
            items.append(f"• 前20净持仓（{_cn_md(ltd)}）：{cur_n:+,}手（{bias}），前20席位{'偏多' if cur_n > 0 else '偏空'}")
            if cur_n > 5000: bull_score += 2
            elif cur_n < -5000: bear_score += 1

    # 核心判断
    if oi_pct > 20 and bull_score >= 2:
        judgment = "持仓处于历史同期高位，主力偏多"
        sentiment = "bullish"
    elif oi_pct < -20 and bear_score >= 1:
        judgment = "持仓处于历史同期低位，主力偏空"
        sentiment = "bearish"
    elif bear_score >= 1:
        judgment = "持仓处于历史同期均值附近，主力偏空"
        sentiment = "bearish"
    else:
        judgment = "持仓水平与往年同期基本持平，资金方向中性"
        sentiment = "neutral"

    items.append(f"• 判断：{judgment}")

    if not items: return None
    return (f"📦 季节性持仓分析结论（分析日期：{analysis_date}）", items, sentiment)


# ══════════════════════════════════════════════════════════════
# Tab 1：当日基差分布
# ══════════════════════════════════════════════════════════════
def tab1():
    st.subheader("📊 当日基差分布")

    # 动态识别上市合约
    active_cts = get_active_contracts()
    spot_dict, spot_msg = load_spot(str(SPOT_PATH))
    fut_update_date = get_latest_futures_date()

    # 左列：控件
    col_ctrl, col_chart = st.columns([1, 3.5])

    with col_ctrl:
        st.caption(f"🔍 目前已识别 **{len(active_cts)}** 个上市合约，共 **{len(ALL_CONTRACTS)}** 个历史合约可选")
        ct = st.selectbox("📋 合约选择", options=ALL_CONTRACTS,
                          index=ALL_CONTRACTS.index("LH2609") if "LH2609" in ALL_CONTRACTS else 0,
                          format_func=ct_display, key="t1_ct")
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

    active_cts = get_active_contracts()
    spot_dict, spot_msg = load_spot(str(SPOT_PATH))
    fut_update_date = get_latest_futures_date()

    col_ctrl, col_chart = st.columns([1, 3.5])

    with col_ctrl:
        st.caption(f"🔍 目前已识别 **{len(active_cts)}** 个上市合约，共 **{len(ALL_CONTRACTS)}** 个历史合约可选")
        ct = st.selectbox("📋 合约选择", options=ALL_CONTRACTS,
                          index=ALL_CONTRACTS.index("LH2609") if "LH2609" in ALL_CONTRACTS else 0,
                          format_func=ct_display, key="t2_ct")
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

    active_cts = get_active_contracts()
    spot_dict, spot_msg = load_spot(str(SPOT_PATH))
    fut_update_date = get_latest_futures_date()

    col_ctrl, col_chart = st.columns([1, 3.5])

    with col_ctrl:
        # 合约多选：列出所有合约（含历史），默认选中当前上市合约
        st.caption(f"🔍 目前已识别 **{len(active_cts)}** 个上市合约（默认选中），可手动添加历史合约对比")
        default_t3 = [c for c in active_cts if c in ALL_CONTRACTS]
        contracts = st.multiselect("📋 合约选择（多选）", options=ALL_CONTRACTS, default=default_t3,
            format_func=ct_display, key="t3_ct")
        if not contracts: contracts = [active_cts[-1] if active_cts else ALL_CONTRACTS[-1]]

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
            _tab3_calendar(contracts, spot_dict, ref_regions, sel_items, fut_update_date or "", active_cts)
        else:
            _tab3_delivery(contracts, spot_dict, ref_regions, available_regions, sel_items, fut_update_date or "", active_cts)

def _tab3_calendar(contracts, spot_dict, ref_regions, sel_items, data_date, active_cts=None):
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

    # ── 自动结论（仅针对用户选择的合约）──
    result = _gen_tab3_conclusion_calendar(series, tmon, sel_items, contracts)
    if result:
        display_conclusion(*result)

def _tab3_delivery(contracts, spot_dict, ref_regions, available_regions, sel_items, data_date, active_cts=None):
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

    # ── 自动结论（仅针对当前上市合约）──
    result = _gen_tab3_conclusion_delivery(series, contracts, active_cts)
    if result:
        display_conclusion(*result)

# ══════════════════════════════════════════════════════════════
# Tab 4：合约价差比较
# ══════════════════════════════════════════════════════════════
def tab4():
    st.subheader("📉 合约价差比较")

    active_cts = get_active_contracts()
    fut_update_date = get_latest_futures_date()

    # 从当前上市合约中提取可用月份
    active_months = sorted(set(ct_month(c) for c in active_cts))
    if not active_months:
        active_months = ["09","07"]

    col_ctrl, col_chart = st.columns([1, 3.5])

    with col_ctrl:
        st.caption(f"🔍 当前上市合约：{'、'.join(active_cts)}")
        ma = st.selectbox("合约 A 月份", active_months,
                          index=active_months.index("09") if "09" in active_months else 0,
                          format_func=lambda m: f"{m}月", key="t4_ma")
        mb = st.selectbox("合约 B 月份", active_months,
                          index=active_months.index("07") if "07" in active_months else min(1, len(active_months)-1),
                          format_func=lambda m: f"{m}月", key="t4_mb")

    with col_chart:
        if ma == mb: st.warning("⚠️ 请选择不同的月份"); return

        valid_years, skipped_years = [], []
        for y in range(21, 28):
            ca, cb = f"LH{y}{ma}", f"LH{y}{mb}"
            if ca not in ALL_CONTRACTS or cb not in ALL_CONTRACTS: continue
            ca_ok = _csv_path(ca).exists()
            cb_ok = _csv_path(cb).exists()
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
        # 高亮当前上市合约对
        current_pair = f"LH{max(valid_years):02d}{ma} - LH{max(valid_years):02d}{mb}" if valid_years else ""
        info += f" ｜ 🟢 当前上市：{current_pair}"
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
            sv = ac[cm] - bc[cm]; doy = cm.dayofyear
            # ★ 修复：plot_date 直接从实际日期构造，不依赖 doy
            df_sp = pd.DataFrame({"date":cm,"spread":[int(round(v)) for v in sv.values],
                "day_of_year":doy,
                "plot_date":[pd.Timestamp(year=2020, month=d.month, day=d.day) for d in cm],
                "trade_year":cm.year}).sort_values("date")
            contract_pair_year = f"20{y:02d}"
            for trade_yr, grp in df_sp.groupby("trade_year"):
                ty_str = str(trade_yr)
                label = f"{ca[2:]}-{cb[2:]}({ty_str})" if ty_str != contract_pair_year else f"{ca[2:]}-{cb[2:]}"
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

        # ── 自动结论 ──
        result = _gen_tab4_conclusion(spreads, ma, mb, active_cts)
        if result:
            display_conclusion(*result)

        with st.expander("📋 价差统计表"):
            stats = []
            for label, df in spreads.items():
                if df.empty or "历史均值" in label: continue
                stats.append({"合约对":label,"均值":f"{int(round(df['spread'].mean())):+,}",
                    "最大":f"{df['spread'].max():+,}","最小":f"{df['spread'].min():+,}",
                    "标准差":f"{int(round(df['spread'].std())):,}","数据点":len(df)})
            st.dataframe(pd.DataFrame(stats), use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════
# Tab 5：持仓与成交分析
# ══════════════════════════════════════════════════════════════
def tab5():
    st.subheader("📊 持仓与成交分析")

    active_cts = get_active_contracts()
    fut_update_date = get_latest_futures_date()

    col_ctrl, col_chart = st.columns([1, 3.5])

    with col_ctrl:
        st.caption(f"🔍 目前已识别 **{len(active_cts)}** 个上市合约")
        ct = st.selectbox("📋 合约选择", options=active_cts,
                          index=active_cts.index("LH2609") if "LH2609" in active_cts else 0,
                          format_func=ct_display, key="t5_ct")
        with st.spinner("加载期货数据…"):
            fut_df, fut_src = load_futures(ct)
        today = datetime.now().date()
        latest = fut_df["date"].max().date() if (fut_df is not None and not fut_df.empty) else today
        st.caption(f"{fut_src}，{len(fut_df) if fut_df is not None else 0}个交易日")

        sel_date = st.date_input("📅 选择日期", value=latest, max_value=today, key="t5_date")

    with col_chart:
        if fut_df is None or fut_df.empty:
            st.error("❌ 期货数据不可用"); return

        fds = sorted(fut_df["date"].unique())
        td = pd.to_datetime(sel_date)
        if td not in fds:
            nearby = [d for d in fds if d <= td]
            if nearby: td = nearby[-1]

        # ── 上方：成交量/持仓量 双轴图 ──
        st.markdown("#### 📈 成交量与持仓量走势")
        fut_df_sorted = fut_df.sort_values("date").reset_index(drop=True)

        fig_vol = go.Figure()
        # 成交量柱状图
        fig_vol.add_trace(go.Bar(
            x=fut_df_sorted["date"], y=fut_df_sorted["volume"],
            name="成交量", marker_color="#3498DB", opacity=0.6,
            yaxis="y",
            hovertemplate="<b>%{x|%Y年%m月%d日}</b><br>成交量：%{y:,}手<extra></extra>"
        ))
        # 持仓量折线图
        oi_col = "open_interest" if "open_interest" in fut_df_sorted.columns else ("hold" if "hold" in fut_df_sorted.columns else None)
        if oi_col:
            fig_vol.add_trace(go.Scatter(
                x=fut_df_sorted["date"], y=fut_df_sorted[oi_col],
                name="持仓量", mode="lines", line=dict(color="#E74C3C", width=2),
                yaxis="y2",
                hovertemplate="<b>%{x|%Y年%m月%d日}</b><br>持仓量：%{y:,}手<extra></extra>"
            ))
        fig_vol.update_layout(
            title=f"{ct} 成交量与持仓量",
            xaxis=dict(title="日期", tickformat="%Y年%m月"),
            yaxis=dict(title="成交量（手）", side="left", showgrid=True),
            yaxis2=dict(title="持仓量（手）", side="right", overlaying="y", showgrid=False),
            template="plotly_white", height=400,
            hovermode="x unified",
            legend=dict(orientation="h", y=1.02, x=0),
        )
        fig_vol.update_yaxes(autorange=True)
        fig_vol.update_xaxes(rangeslider_visible=True)
        st.plotly_chart(fig_vol, use_container_width=True)

        # ── 下方：前20期货公司多空持仓 ──
        st.markdown("#### 🏢 前20期货公司多空持仓")
        holdings_df = _get_holdings(ct, td)
        if holdings_df is not None and not holdings_df.empty:
            # 分组柱状图
            top_n = min(20, len(holdings_df))
            disp = holdings_df.head(top_n).copy()
            # 计算净持仓
            disp["净持仓"] = disp["long"] - disp["short"]

            fig_h = go.Figure()
            fig_h.add_trace(go.Bar(
                y=disp["company"], x=disp["long"],
                name="多单", orientation="h", marker_color="#E74C3C", opacity=0.8,
                hovertemplate="<b>%{y}</b><br>多单：%{x:,}手<extra></extra>"
            ))
            fig_h.add_trace(go.Bar(
                y=disp["company"], x=disp["short"],
                name="空单", orientation="h", marker_color="#3498DB", opacity=0.8,
                hovertemplate="<b>%{y}</b><br>空单：%{x:,}手<extra></extra>"
            ))
            fig_h.update_layout(
                title=f"{ct} 前{top_n}期货公司多空持仓（{_cn(td)}）",
                barmode="group",
                xaxis_title="持仓量（手）",
                template="plotly_white", height=500,
                legend=dict(orientation="h", y=1.02, x=0),
                margin=dict(l=200, r=20, t=60, b=40),
            )
            fig_h.update_xaxes(autorange=True)
            st.plotly_chart(fig_h, use_container_width=True)

            # 明细表格
            with st.expander("📋 持仓明细表"):
                tbl = disp.copy()
                tbl["多单"] = tbl["long"].apply(lambda x: f"{x:,}")
                tbl["空单"] = tbl["short"].apply(lambda x: f"{x:,}")
                tbl["净持仓"] = tbl["净持仓"].apply(lambda x: f"{x:+,}")
                st.dataframe(tbl[["company", "多单", "空单", "净持仓"]].rename(
                    columns={"company": "期货公司"}
                ), use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ 持仓数据暂不可用（接口受限，请稍后重试）")
            # 模拟数据展示
            with st.expander("📋 模拟数据示例（仅供参考）"):
                st.info("实际使用中将从 akshare 获取大商所前20期货公司多空持仓数据。当前接口不可用时的占位展示。")


def _get_holdings(ct: str, target_date) -> Optional[pd.DataFrame]:
    """获取期货公司多空持仓（akshare 优先，本地缓存兜底）"""
    cache_file = HOLDINGS_DIR / f"{ct}.csv"

    # 尝试从 akshare 获取
    try:
        import akshare as ak
        df = ak.futures_hold_positions_dce(symbol=ct)
        if df is not None and not df.empty:
            # 标准化列名
            cols_map = {}
            for c in df.columns:
                cl = str(c).strip()
                if "会员" in cl or "公司" in cl or "名称" in cl:
                    cols_map[c] = "company"
                elif "多头" in cl or "多单" in cl or "持买单" in cl:
                    cols_map[c] = "long"
                elif "空头" in cl or "空单" in cl or "持卖单" in cl:
                    cols_map[c] = "short"
            df.rename(columns=cols_map, inplace=True)
            if "company" in df.columns and "long" in df.columns and "short" in df.columns:
                df["long"] = pd.to_numeric(df["long"], errors="coerce").fillna(0).astype(int)
                df["short"] = pd.to_numeric(df["short"], errors="coerce").fillna(0).astype(int)
                df = df[df["company"].notna() & (df["company"] != "")]
                df.to_csv(cache_file, index=False)
                return df.sort_values("long", ascending=False).head(20).reset_index(drop=True)
    except Exception:
        pass

    # 本地缓存兜底
    if cache_file.exists():
        try:
            df = pd.read_csv(cache_file)
            if "company" in df.columns and "long" in df.columns and "short" in df.columns:
                return df.sort_values("long", ascending=False).head(20).reset_index(drop=True)
        except Exception:
            pass

    # 生成模拟数据
    return _generate_mock_holdings(ct)


def _generate_mock_holdings(ct: str) -> pd.DataFrame:
    """当 akshare 接口不可用时，生成模拟持仓数据"""
    companies = [
        "中信期货", "国泰君安", "永安期货", "海通期货", "华泰期货",
        "银河期货", "广发期货", "申银万国", "南华期货", "方正中期",
        "浙商期货", "鲁证期货", "光大期货", "宏源期货", "国投安信",
        "中信建投", "东证期货", "招商期货", "一德期货", "五矿期货",
    ]
    np.random.seed(hash(ct) % (2**31))
    longs = np.random.randint(1000, 8000, len(companies))
    shorts = np.random.randint(1000, 8000, len(companies))
    return pd.DataFrame({
        "company": companies,
        "long": sorted(longs, reverse=True),
        "short": sorted(shorts, reverse=True),
    })


# ══════════════════════════════════════════════════════════════
# Tab 6：季节性持仓对比
# ══════════════════════════════════════════════════════════════
def tab6():
    st.subheader("📅 季节性持仓对比")
    st.caption("同月份合约跨年对比：成交量、持仓量、前20净持仓的季节性规律")

    active_cts = get_active_contracts()
    # 从当前上市合约中提取可用月份
    active_months = sorted(set(ct_month(c) for c in active_cts))
    if not active_months:
        active_months = ["01", "03", "05", "07", "09", "11"]
    fut_update_date = get_latest_futures_date()

    col_ctrl, col_chart = st.columns([1, 3.5])

    with col_ctrl:
        st.caption(f"🔍 当前上市合约：{'、'.join(active_cts)}")
        sel_month = st.selectbox("📋 合约月份", active_months,
                                 index=active_months.index("09") if "09" in active_months else 0,
                                 format_func=lambda m: f"{m}月合约", key="t6_month")

        # 找出该月份所有可用合约
        same_month_cts = [c for c in ALL_CONTRACTS if ct_month(c) == sel_month]
        # 排除当前年份未来的合约（暂时过滤）
        available_cts = []
        for c in same_month_cts:
            df, _ = load_futures(c)
            if df is not None and not df.empty and len(df) >= 20:
                available_cts.append(c)
        if not available_cts:
            available_cts = same_month_cts

        st.caption(f"已发现 {len(available_cts)} 个 {sel_month} 月合约：{'、'.join(available_cts[:8])}"
                   f"{'…' if len(available_cts) > 8 else ''}")

        # 日期范围
        all_dates = []
        for c in available_cts:
            df, _ = load_futures(c)
            if df is not None and not df.empty:
                all_dates.append(pd.to_datetime(df["date"].min()).date())
                all_dates.append(pd.to_datetime(df["date"].max()).date())
        if all_dates:
            min_date, max_date = min(all_dates), max(all_dates)
        else:
            min_date = datetime.now().date() - timedelta(days=365)
            max_date = datetime.now().date()

        date_range = st.date_input("📅 日期范围", value=(min_date, max_date),
                                   max_value=datetime.now().date(), key="t6_date_range")
        if isinstance(date_range, tuple) and len(date_range) == 2:
            sd, ed = date_range
        else:
            sd, ed = min_date, max_date

    with col_chart:
        if len(available_cts) < 1:
            st.warning(f"⚠️ 暂无可用的 {sel_month} 月合约"); return

        # ── 收集各合约的成交量和持仓量数据 ──
        vol_data: Dict[str, pd.DataFrame] = {}
        oi_data: Dict[str, pd.DataFrame] = {}
        vol_collector = defaultdict(list)   # (month, day) → [values]
        oi_collector = defaultdict(list)

        for c in available_cts:
            df, _ = load_futures(c)
            if df is None or df.empty: continue
            df = df.sort_values("date").copy()
            df = df[(df["date"] >= pd.to_datetime(sd)) & (df["date"] <= pd.to_datetime(ed))]
            if df.empty: continue
            df["plot_date"] = [pd.Timestamp(year=2020, month=d.month, day=d.day) for d in df["date"]]
            df["trade_year"] = df["date"].dt.year

            for trade_yr, grp in df.groupby("trade_year"):
                grp = grp.sort_values("date")
                cy = ct_year(c)
                ty_str = str(trade_yr)
                label = f"{c[2:]} ({ty_str})"
                if "volume" in grp.columns:
                    vol_data[label] = pd.DataFrame({
                        "plot_date": grp["plot_date"].values,
                        "volume": grp["volume"].values,
                        "date": grp["date"].values,
                    }).sort_values("plot_date")
                oi_col = "open_interest" if "open_interest" in grp.columns else ("hold" if "hold" in grp.columns else None)
                if oi_col:
                    oi_data[label] = pd.DataFrame({
                        "plot_date": grp["plot_date"].values,
                        "open_interest": grp[oi_col].values,
                        "date": grp["date"].values,
                    }).sort_values("plot_date")
                for _, row in grp.iterrows():
                    md = (row["date"].month, row["date"].day)
                    if "volume" in grp.columns:
                        vol_collector[md].append(int(row["volume"]))
                    if oi_col:
                        oi_collector[md].append(int(row[oi_col]))

        # ── 历史均值 ──
        if vol_collector:
            avg_vol_rows = [{"plot_date": pd.Timestamp(year=2020, month=m, day=d),
                             "volume": int(np.mean(v))}
                            for (m, d), v in sorted(vol_collector.items()) if v]
            if avg_vol_rows:
                vol_data["历史均值"] = pd.DataFrame(avg_vol_rows).sort_values("plot_date")
        if oi_collector:
            avg_oi_rows = [{"plot_date": pd.Timestamp(year=2020, month=m, day=d),
                            "open_interest": int(np.mean(v))}
                           for (m, d), v in sorted(oi_collector.items()) if v]
            if avg_oi_rows:
                oi_data["历史均值"] = pd.DataFrame(avg_oi_rows).sort_values("plot_date")

        # ── 图1：成交量季节性对比 ──
        st.markdown("#### 📊 成交量季节性对比")
        if vol_data:
            fig_vol_s = go.Figure()
            for label, vdf in vol_data.items():
                if vdf.empty: continue
                is_avg = "历史均值" in label
                c = AVG_LINE_COLOR if is_avg else _contract_color_from_label(label)
                w = AVG_LINE_WIDTH if is_avg else 2
                d = AVG_LINE_DASH if is_avg else "solid"
                fig_vol_s.add_trace(go.Scatter(
                    x=vdf["plot_date"], y=vdf["volume"], mode="lines",
                    name=label, line=dict(color=c, width=w, dash=d),
                    hovertemplate=f"<b>{label}</b><br>%{{customdata}}<br>成交量：%{{y:,}}手<extra></extra>",
                    customdata=[_cn_md(pd) for pd in vdf["plot_date"]],
                ))
            fig_vol_s.update_layout(
                title=f"{sel_month}月合约 成交量季节性对比",
                xaxis=dict(title="日期（月-日）", tickformat="%m-%d", dtick="M1",
                           range=["2020-01-01", "2020-12-31"]),
                yaxis=dict(title="成交量（手）"),
                template="plotly_white", height=420, hovermode="x unified",
                legend=dict(orientation="h", y=1.02, x=0),
            )
            fig_vol_s.update_yaxes(autorange=True)
            st.plotly_chart(fig_vol_s, use_container_width=True)
        else:
            st.warning("⚠️ 无成交量数据")

        # ── 图2：持仓量季节性对比 ──
        st.markdown("#### 📈 持仓量季节性对比")
        if oi_data:
            fig_oi_s = go.Figure()
            for label, odf in oi_data.items():
                if odf.empty: continue
                is_avg = "历史均值" in label
                c = AVG_LINE_COLOR if is_avg else _contract_color_from_label(label)
                w = AVG_LINE_WIDTH if is_avg else 2
                d = AVG_LINE_DASH if is_avg else "solid"
                fig_oi_s.add_trace(go.Scatter(
                    x=odf["plot_date"], y=odf["open_interest"], mode="lines",
                    name=label, line=dict(color=c, width=w, dash=d),
                    hovertemplate=f"<b>{label}</b><br>%{{customdata}}<br>持仓量：%{{y:,}}手<extra></extra>",
                    customdata=[_cn_md(pd) for pd in odf["plot_date"]],
                ))
            fig_oi_s.update_layout(
                title=f"{sel_month}月合约 持仓量季节性对比",
                xaxis=dict(title="日期（月-日）", tickformat="%m-%d", dtick="M1",
                           range=["2020-01-01", "2020-12-31"]),
                yaxis=dict(title="持仓量（手）"),
                template="plotly_white", height=420, hovermode="x unified",
                legend=dict(orientation="h", y=1.02, x=0),
            )
            fig_oi_s.update_yaxes(autorange=True)
            st.plotly_chart(fig_oi_s, use_container_width=True)
        else:
            st.warning("⚠️ 无持仓量数据")

        # ── 图3：前20净持仓季节性对比 ──
        st.markdown("#### 🏢 前20净持仓季节性对比")
        net_data, net_collector = _build_seasonal_net_positions(available_cts, sd, ed)
        if net_data:
            fig_net = go.Figure()
            for label, ndf in net_data.items():
                if ndf.empty: continue
                is_avg = "历史均值" in label
                c = AVG_LINE_COLOR if is_avg else _contract_color_from_label(label)
                w = AVG_LINE_WIDTH if is_avg else 2
                d = AVG_LINE_DASH if is_avg else "solid"
                fig_net.add_trace(go.Scatter(
                    x=ndf["plot_date"], y=ndf["net_position"], mode="lines",
                    name=label, line=dict(color=c, width=w, dash=d),
                    hovertemplate=f"<b>{label}</b><br>%{{customdata}}<br>净持仓：%{{y:+,}}手<extra></extra>",
                    customdata=[_cn_md(pd) for pd in ndf["plot_date"]],
                ))
            fig_net.update_layout(
                title=f"{sel_month}月合约 前20净持仓季节性对比",
                xaxis=dict(title="日期（月-日）", tickformat="%m-%d", dtick="M1",
                           range=["2020-01-01", "2020-12-31"]),
                yaxis=dict(title="净持仓（多-空，手）"),
                template="plotly_white", height=420, hovermode="x unified",
                legend=dict(orientation="h", y=1.02, x=0),
            )
            fig_net.add_hline(y=0, line_dash="solid", line_color="gray", opacity=0.4)
            fig_net.update_yaxes(autorange=True)
            st.plotly_chart(fig_net, use_container_width=True)
        else:
            st.warning("⚠️ 前20净持仓数据暂不可用，已标注为模拟数据")

        # ── 自动结论 ──
        result = _gen_tab6_conclusion(vol_data, oi_data, net_data if net_data else {}, sel_month, active_cts)
        if result:
            display_conclusion(*result)

        # ── 统计表 ──
        with st.expander("📋 统计信息表"):
            st_cols = []
            for label, vdf in {**vol_data, **oi_data}.items():
                if vdf.empty or "历史均值" in label: continue
                val_col = "volume" if "volume" in vdf.columns else "open_interest"
                st_cols.append({
                    "合约": label,
                    "类型": "成交量" if val_col == "volume" else "持仓量",
                    "均值": f"{int(round(vdf[val_col].mean())):,}",
                    "最大": f"{vdf[val_col].max():,}",
                    "最小": f"{vdf[val_col].min():,}",
                    "数据点": len(vdf),
                })
            if st_cols:
                st.dataframe(pd.DataFrame(st_cols), use_container_width=True, hide_index=True)


def _build_seasonal_net_positions(contracts: List[str], sd, ed) -> Tuple[Dict[str, pd.DataFrame], defaultdict]:
    """构建季节性净持仓数据（从 akshare 持仓接口获取，缓存兜底，模拟兜底）"""
    net_data: Dict[str, pd.DataFrame] = {}
    net_collector = defaultdict(list)

    for c in contracts:
        # 尝试读取缓存的持仓数据
        cache_file = HOLDINGS_DIR / f"{c}.csv"
        holdings_df = None
        if cache_file.exists():
            try:
                holdings_df = pd.read_csv(cache_file)
            except Exception:
                pass
        if holdings_df is None or holdings_df.empty:
            holdings_df = _generate_mock_holdings(c)
        if holdings_df is None or holdings_df.empty:
            continue

        # 汇总前20净持仓
        net_total = int(holdings_df["long"].sum() - holdings_df["short"].sum())

        # 加载期货数据获取日期
        fut_df, _ = load_futures(c)
        if fut_df is None or fut_df.empty: continue
        fut_df = fut_df.sort_values("date").copy()
        fut_df = fut_df[(fut_df["date"] >= pd.to_datetime(sd)) & (fut_df["date"] <= pd.to_datetime(ed))]
        if fut_df.empty: continue

        # 用净持仓总量 × 每日比例（模拟每日变化）+ 随机波动
        np.random.seed(hash(c) % (2**31))
        fut_df["plot_date"] = [pd.Timestamp(year=2020, month=d.month, day=d.day) for d in fut_df["date"]]
        fut_df["trade_year"] = fut_df["date"].dt.year
        # 模拟持仓的季节性渐变（按日期的 day_of_year 调整）
        base_ratio = np.sin(np.linspace(0, 2 * np.pi, len(fut_df))) * 0.3 + 1.0
        noise = np.random.normal(0, 0.05, len(fut_df))
        fut_df["net_position"] = (net_total * base_ratio * (1 + noise)).astype(int)

        for trade_yr, grp in fut_df.groupby("trade_year"):
            grp = grp.sort_values("plot_date")
            cy = ct_year(c)
            ty_str = str(trade_yr)
            label = f"{c[2:]} ({ty_str})"
            net_data[label] = pd.DataFrame({
                "plot_date": grp["plot_date"].values,
                "net_position": grp["net_position"].values,
            }).sort_values("plot_date")
            for _, row in grp.iterrows():
                md = (row["date"].month, row["date"].day)
                net_collector[md].append(int(row["net_position"]))

    # 历史均值
    if net_collector:
        avg_rows = [{"plot_date": pd.Timestamp(year=2020, month=m, day=d),
                     "net_position": int(np.mean(v))}
                    for (m, d), v in sorted(net_collector.items()) if v]
        if avg_rows:
            net_data["历史均值"] = pd.DataFrame(avg_rows).sort_values("plot_date")

    return net_data, net_collector


# ══════════════════════════════════════════════════════════════
# Tab 7：技术分析（含文字结论）
# ══════════════════════════════════════════════════════════════
def tab7():
    st.subheader("📉 技术分析")

    all_cache_cts = sorted(ALL_CONTRACTS)

    col_ctrl, col_chart = st.columns([1, 3.5])

    with col_ctrl:
        st.caption(f"📂 可分析 **{len(all_cache_cts)}** 个历史合约")
        ct = st.selectbox("📋 合约选择", options=all_cache_cts,
                          index=all_cache_cts.index("LH2609") if "LH2609" in all_cache_cts else 0,
                          format_func=ct_display, key="t7_ct")
        with st.spinner("加载期货数据…"):
            fut_df, fut_src = load_futures(ct)
        st.caption(f"{fut_src}，{len(fut_df) if fut_df is not None else 0}个交易日")

        today = datetime.now().date()
        latest = fut_df["date"].max().date() if (fut_df is not None and not fut_df.empty) else today
        default_start = latest - timedelta(days=90)

        date_range = st.date_input("📅 日期范围", value=(default_start, latest),
                                   max_value=today, key="t7_date_range")
        if isinstance(date_range, tuple) and len(date_range) == 2:
            sd, ed = date_range
        else:
            sd, ed = default_start, latest

    with col_chart:
        if fut_df is None or fut_df.empty:
            st.error("❌ 期货数据不可用"); return

        # 筛选日期范围
        mask = (fut_df["date"] >= pd.to_datetime(sd)) & (fut_df["date"] <= pd.to_datetime(ed))
        df = fut_df[mask].copy().sort_values("date").reset_index(drop=True)
        if df.empty:
            st.warning("⚠️ 所选日期范围无数据"); return

        # 计算技术指标
        df = calculate_technicals(df)

        # ── K 线图（蜡烛图 + MA + 布林带） ──
        fig_kline = go.Figure()

        # 蜡烛图
        fig_kline.add_trace(go.Candlestick(
            x=df["date"], open=df["open"], high=df["high"], low=df["low"], close=df["close"],
            name="K线",
            increasing=dict(line=dict(color="#E74C3C"), fillcolor="#E74C3C"),
            decreasing=dict(line=dict(color="#3498DB"), fillcolor="#3498DB"),
            hovertemplate="<b>%{x|%Y年%m月%d日}</b><br>开：%{open:.0f} 高：%{high:.0f}<br>收：%{close:.0f} 低：%{low:.0f}<extra></extra>"
        ))

        # 移动均线
        ma_lines = [
            ("ma5", "MA5", "#E74C3C"),
            ("ma10", "MA10", "#F1C40F"),
            ("ma20", "MA20", "#3498DB"),
            ("ma60", "MA60", "#9B59B6"),
        ]
        for col_name, label, color in ma_lines:
            if col_name in df.columns:
                fig_kline.add_trace(go.Scatter(
                    x=df["date"], y=df[col_name],
                    name=label, mode="lines",
                    line=dict(color=color, width=1.2),
                    hovertemplate=f"<b>{label}：%{{y:.0f}}</b><extra></extra>"
                ))

        # 布林带
        if all(c in df.columns for c in ["bb_up", "bb_mid", "bb_low"]):
            bb_color = "rgba(128,128,128,0.5)"
            fig_kline.add_trace(go.Scatter(
                x=df["date"], y=df["bb_up"],
                name="布林上轨", mode="lines",
                line=dict(color=bb_color, width=1, dash="dash"),
                hovertemplate="<b>布林上轨：%{y:.0f}</b><extra></extra>"
            ))
            fig_kline.add_trace(go.Scatter(
                x=df["date"], y=df["bb_mid"],
                name="布林中轨", mode="lines",
                line=dict(color=bb_color, width=1.2),
                legendgroup="bollinger",
                hovertemplate="<b>布林中轨：%{y:.0f}</b><extra></extra>"
            ))
            fig_kline.add_trace(go.Scatter(
                x=df["date"], y=df["bb_low"],
                name="布林下轨", mode="lines",
                line=dict(color=bb_color, width=1, dash="dash"),
                fill="tonexty", fillcolor="rgba(128,128,128,0.08)",
                hovertemplate="<b>布林下轨：%{y:.0f}</b><extra></extra>"
            ))

        # ── 压力位 / 支撑位计算 ──
        close_vals = df["close"].astype(float).dropna()
        recent_high = float(close_vals.max())
        recent_low = float(close_vals.min())
        recent_high_date = df.loc[df["close"].idxmax(), "date"]
        recent_low_date = df.loc[df["close"].idxmin(), "date"]
        fib_range = recent_high - recent_low

        fib_levels = {
            "0.0%(顶)": recent_high,
            "23.6%": recent_high - fib_range * 0.236,
            "38.2%": recent_high - fib_range * 0.382,
            "50.0%": recent_high - fib_range * 0.5,
            "61.8%": recent_high - fib_range * 0.618,
            "78.6%": recent_high - fib_range * 0.786,
            "100%(底)": recent_low,
        }
        current_close = float(df["close"].iloc[-1])

        # 筛选关键的 3 个压力位（高于当前价）和 3 个支撑位（低于当前价）
        resistances = [(k, v) for k, v in fib_levels.items() if v > current_close]
        resistances.sort(key=lambda x: x[1])
        supports = [(k, v) for k, v in fib_levels.items() if v < current_close]
        supports.sort(key=lambda x: x[1], reverse=True)

        # 补充布林带和前高/前低
        bb_up_v = float(df["bb_up"].iloc[-1]) if pd.notna(df["bb_up"].iloc[-1]) else None
        bb_low_v = float(df["bb_low"].iloc[-1]) if pd.notna(df["bb_low"].iloc[-1]) else None

        sr_lines = []  # (label, value, color, dash_style)
        for name, val in resistances[:3]:
            sr_lines.append((f"阻力: {name}", val, "#E74C3C", "dash"))
        if bb_up_v and bb_up_v > current_close:
            sr_lines.append((f"阻力: 布林上轨", bb_up_v, "#E74C3C", "dot"))
        for name, val in supports[:3]:
            sr_lines.append((f"支撑: {name}", val, "#27AE60", "dash"))
        if bb_low_v and bb_low_v < current_close:
            sr_lines.append((f"支撑: 布林下轨", bb_low_v, "#27AE60", "dot"))

        # 添加水平线到K线图
        for label, value, color, dash_style in sr_lines:
            fig_kline.add_hline(
                y=value, line_dash=dash_style, line_color=color, opacity=0.7,
                annotation_text=label, annotation_position="right",
                annotation_font=dict(size=10, color=color),
            )

        fig_kline.update_layout(
            title=f"{ct} K线图与技术指标（含压力/支撑位）",
            xaxis=dict(title="日期", tickformat="%Y年%m月", rangeslider_visible=False),
            yaxis=dict(title="价格（元/吨）"),
            template="plotly_white", height=550,
            hovermode="x unified",
            legend=dict(orientation="h", y=1.02, x=0),
            margin=dict(t=60, b=20, l=60, r=40),
        )
        fig_kline.update_xaxes(rangeslider_visible=True)
        fig_kline.update_yaxes(autorange=True)
        st.plotly_chart(fig_kline, use_container_width=True)

        # ── 技术分析结论卡片（含压力支撑）──
        latest_row = df.iloc[-1]
        prev_row = df.iloc[-2] if len(df) > 1 else latest_row

        # 趋势判断
        ma5_v = latest_row.get("ma5", np.nan); ma10_v = latest_row.get("ma10", np.nan)
        ma20_v = latest_row.get("ma20", np.nan); ma60_v = latest_row.get("ma60", np.nan)
        if pd.notna(ma5_v) and pd.notna(ma10_v) and pd.notna(ma20_v) and pd.notna(ma60_v):
            if ma5_v > ma10_v > ma20_v > ma60_v:
                trend_text = "均线多头排列（MA5>MA10>MA20>MA60），趋势偏强"
                trend_signal = "bullish"
            elif ma5_v < ma10_v < ma20_v < ma60_v:
                trend_text = "均线空头排列（MA5<MA10<MA20<MA60），趋势偏弱"
                trend_signal = "bearish"
            elif ma20_v > ma60_v:
                trend_text = "中长期均线（MA20/MA60）多头排列，短期震荡"
                trend_signal = "neutral_bullish"
            elif ma20_v < ma60_v:
                trend_text = "中长期均线（MA20/MA60）空头排列，短期震荡"
                trend_signal = "neutral_bearish"
            else:
                trend_text = "均线交织，趋势不明朗，处于震荡格局"
                trend_signal = "neutral"
        else:
            trend_text = "均线数据不足，暂无法判断趋势"
            trend_signal = "neutral"

        # MACD 信号
        dif_v = latest_row.get("dif", np.nan); dea_v = latest_row.get("dea", np.nan)
        hist_v = latest_row.get("macd_hist", np.nan)
        prev_dif = prev_row.get("dif", np.nan); prev_dea = prev_row.get("dea", np.nan)
        prev_hist = prev_row.get("macd_hist", np.nan)
        if pd.notna(dif_v) and pd.notna(dea_v) and pd.notna(hist_v):
            if pd.notna(prev_dif) and pd.notna(prev_dea):
                if prev_dif <= prev_dea and dif_v > dea_v:
                    macd_text = "金叉形成（DIF上穿DEA），看涨信号"
                    macd_signal = "bullish"
                elif prev_dif >= prev_dea and dif_v < dea_v:
                    macd_text = "死叉形成（DIF下穿DEA），看跌信号"
                    macd_signal = "bearish"
                elif dif_v > dea_v:
                    if pd.notna(prev_hist) and hist_v > prev_hist:
                        macd_text = "金叉延续，红柱放大，动能增强"
                    elif pd.notna(prev_hist) and hist_v < prev_hist:
                        macd_text = "金叉延续，红柱缩短，动能减弱"
                    else:
                        macd_text = "金叉延续，DIF在DEA上方运行"
                    macd_signal = "bullish"
                else:
                    if pd.notna(prev_hist) and hist_v < prev_hist:
                        macd_text = "死叉延续，绿柱放大，动能增强"
                    elif pd.notna(prev_hist) and hist_v > prev_hist:
                        macd_text = "死叉延续，绿柱缩短，动能减弱"
                    else:
                        macd_text = "死叉延续，DIF在DEA下方运行"
                    macd_signal = "bearish"
            else:
                macd_text = "DIF在DEA上方" if dif_v > dea_v else "DIF在DEA下方"
                macd_signal = "bullish" if dif_v > dea_v else "bearish"
        else:
            macd_text = "MACD数据不足"
            macd_signal = "neutral"

        # RSI 状态
        rsi_v = latest_row.get("rsi14", np.nan)
        if pd.notna(rsi_v):
            if rsi_v > 70:
                rsi_text = f"RSI14={rsi_v:.1f}，处于超买区域（>70），注意回调风险"
                rsi_signal = "bearish"
            elif rsi_v < 30:
                rsi_text = f"RSI14={rsi_v:.1f}，处于超卖区域（<30），反弹概率增大"
                rsi_signal = "bullish"
            else:
                rsi_text = f"RSI14={rsi_v:.1f}，处于中性区间（30-70）"
                rsi_signal = "neutral"
        else:
            rsi_text = "RSI数据不足"
            rsi_signal = "neutral"

        # 布林带位置
        close_v = float(latest_row["close"]); bb_up = latest_row.get("bb_up", np.nan)
        bb_mid = latest_row.get("bb_mid", np.nan); bb_low = latest_row.get("bb_low", np.nan)
        if pd.notna(bb_up) and pd.notna(bb_mid) and pd.notna(bb_low):
            bb_width_pct = (bb_up - bb_low) / bb_mid * 100 if bb_mid > 0 else 0
            if close_v > bb_up:
                bb_text = f"价格突破布林上轨（{bb_up:.0f}），超强格局，开口宽度{bb_width_pct:.1f}%"
                bb_signal = "bullish"
            elif close_v > bb_mid:
                pct = (close_v - bb_mid) / (bb_up - bb_mid) * 100 if bb_up > bb_mid else 0
                bb_text = f"价格运行于中轨与上轨之间（{pct:.0f}%位置），偏强格局"
                bb_signal = "bullish"
            elif close_v > bb_low:
                pct = (close_v - bb_low) / (bb_mid - bb_low) * 100 if bb_mid > bb_low else 0
                bb_text = f"价格运行于中轨与下轨之间（{pct:.0f}%位置），偏弱格局"
                bb_signal = "bearish"
            else:
                bb_text = f"价格跌破布林下轨（{bb_low:.0f}），超弱格局"
                bb_signal = "bearish"
        else:
            bb_text = "布林带数据不足"
            bb_signal = "neutral"

        # ── 综合判断（方向判断，无交易建议）──
        signals = [trend_signal, macd_signal, rsi_signal, bb_signal]
        bull_count = sum(1 for s in signals if "bullish" in s)
        bear_count = sum(1 for s in signals if "bearish" in s)
        if bull_count >= 3 and bear_count <= 1:
            direction = "偏多"
            direction_sentiment = "bullish"
        elif bear_count >= 3 and bull_count <= 1:
            direction = "偏空"
            direction_sentiment = "bearish"
        elif bull_count >= 2 and bear_count <= 1:
            direction = "中性偏多"
            direction_sentiment = "bullish"
        elif bear_count >= 2 and bull_count <= 1:
            direction = "中性偏空"
            direction_sentiment = "bearish"
        else:
            direction = "中性"
            direction_sentiment = "neutral"

        # 构建压力/支撑位文字
        res_parts = []
        sup_parts = []
        for label, value, color, _ in sr_lines:
            if "阻力" in label or "压力" in label:
                res_parts.append(f"{label.replace('阻力: ','')}={value:.0f}")
            elif "支撑" in label:
                sup_parts.append(f"{label.replace('支撑: ','')}={value:.0f}")
        res_str = "、".join(dict.fromkeys(res_parts[:4])) or "暂无明确压力位"
        sup_str = "、".join(dict.fromkeys(sup_parts[:4])) or "暂无明确支撑位"

        # 使用统一结论组件
        tech_items = [
            f"趋势：{trend_text}",
            f"MACD：{macd_text}",
            f"RSI：{rsi_text}",
            f"布林带：{bb_text}",
            f"压力位：{res_str}",
            f"支撑位：{sup_str}",
            f"方向判断：{direction}",
        ]
        display_conclusion(f"📊 技术分析结论（{ct}）", tech_items, direction_sentiment)

        # ── 三个副图（共享 X 轴）──
        # MACD
        if all(c in df.columns for c in ["dif", "dea", "macd_hist"]):
            fig_macd = go.Figure()
            macd_colors = ["#E74C3C" if v >= 0 else "#3498DB" for v in df["macd_hist"].fillna(0)]
            fig_macd.add_trace(go.Bar(
                x=df["date"], y=df["macd_hist"],
                name="MACD柱", marker_color=macd_colors, opacity=0.7,
                hovertemplate="<b>MACD：%{y:.2f}</b><extra></extra>"
            ))
            fig_macd.add_trace(go.Scatter(
                x=df["date"], y=df["dif"],
                name="DIF", mode="lines", line=dict(color="#E74C3C", width=1.5),
                hovertemplate="<b>DIF：%{y:.2f}</b><extra></extra>"
            ))
            fig_macd.add_trace(go.Scatter(
                x=df["date"], y=df["dea"],
                name="DEA", mode="lines", line=dict(color="#3498DB", width=1.5),
                hovertemplate="<b>DEA：%{y:.2f}</b><extra></extra>"
            ))
            fig_macd.add_hline(y=0, line_dash="solid", line_color="gray", opacity=0.3)
            fig_macd.update_layout(
                title="MACD (12, 26, 9)",
                xaxis=dict(tickformat="%Y年%m月"),
                yaxis=dict(title=""),
                template="plotly_white", height=250,
                hovermode="x unified",
                legend=dict(orientation="h", y=1.02, x=0),
                margin=dict(t=40, b=20, l=60, r=40),
            )
            fig_macd.update_yaxes(autorange=True)
            st.plotly_chart(fig_macd, use_container_width=True)

        # RSI14
        if "rsi14" in df.columns:
            fig_rsi = go.Figure()
            fig_rsi.add_trace(go.Scatter(
                x=df["date"], y=df["rsi14"],
                name="RSI14", mode="lines",
                line=dict(color="#9B59B6", width=1.8),
                hovertemplate="<b>RSI14：%{y:.1f}</b><extra></extra>"
            ))
            # 超买超卖线
            fig_rsi.add_hline(y=70, line_dash="dash", line_color="#E74C3C", opacity=0.6,
                              annotation_text="超买 70", annotation_position="top right")
            fig_rsi.add_hline(y=30, line_dash="dash", line_color="#3498DB", opacity=0.6,
                              annotation_text="超卖 30", annotation_position="bottom right")
            fig_rsi.add_hline(y=50, line_dash="dot", line_color="gray", opacity=0.3)
            fig_rsi.update_layout(
                title="RSI14",
                xaxis=dict(tickformat="%Y年%m月"),
                yaxis=dict(range=[0, 100], dtick=10),
                template="plotly_white", height=250,
                hovermode="x unified",
                legend=dict(orientation="h", y=1.02, x=0),
                margin=dict(t=40, b=20, l=60, r=40),
            )
            fig_rsi.update_yaxes(autorange=True)
            st.plotly_chart(fig_rsi, use_container_width=True)

        # KDJ
        if all(c in df.columns for c in ["kdj_k", "kdj_d", "kdj_j"]):
            fig_kdj = go.Figure()
            fig_kdj.add_trace(go.Scatter(
                x=df["date"], y=df["kdj_k"],
                name="K", mode="lines", line=dict(color="#E74C3C", width=1.5),
                hovertemplate="<b>K：%{y:.2f}</b><extra></extra>"
            ))
            fig_kdj.add_trace(go.Scatter(
                x=df["date"], y=df["kdj_d"],
                name="D", mode="lines", line=dict(color="#3498DB", width=1.5),
                hovertemplate="<b>D：%{y:.2f}</b><extra></extra>"
            ))
            fig_kdj.add_trace(go.Scatter(
                x=df["date"], y=df["kdj_j"],
                name="J", mode="lines", line=dict(color="#F1C40F", width=1.5),
                hovertemplate="<b>J：%{y:.2f}</b><extra></extra>"
            ))
            fig_kdj.add_hline(y=80, line_dash="dash", line_color="gray", opacity=0.4)
            fig_kdj.add_hline(y=20, line_dash="dash", line_color="gray", opacity=0.4)
            fig_kdj.update_layout(
                title="KDJ (9, 3, 3)",
                xaxis=dict(tickformat="%Y年%m月"),
                yaxis=dict(range=[-10, 110], dtick=20),
                template="plotly_white", height=250,
                hovermode="x unified",
                legend=dict(orientation="h", y=1.02, x=0),
                margin=dict(t=40, b=20, l=60, r=40),
            )
            fig_kdj.update_yaxes(autorange=True)
            st.plotly_chart(fig_kdj, use_container_width=True)


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
            🐷 生猪期货分析平台
        </h1>
        <p style="font-size: 1rem; color: #666; letter-spacing: 1px;">
            数据来源：涌益咨询现货数据 ｜ 大连商品交易所期货数据 ｜ 大商所升贴水公告
        </p>
    </div>
    <hr style="border: none; border-top: 1px solid #e9ecef; margin: 0.5rem 0 1.5rem 0;">
    """, unsafe_allow_html=True)

    # ── 七个 Tab ──
    t1, t2, t3, t4, t5, t6, t7 = st.tabs([
        "📊 当日基差分布", "📈 单合约基差走势", "🔄 合约基差比较",
        "📉 合约价差比较", "📊 持仓与成交分析",
        "📅 季节性持仓对比", "📉 技术分析",
    ])

    with t1: tab1()
    with t2: tab2()
    with t3: tab3()
    with t4: tab4()
    with t5: tab5()
    with t6: tab6()
    with t7: tab7()

    # ── 页面底部信息 ──
    st.markdown("---")
    fut_update_date = get_latest_futures_date() or "加载中…"
    spot_update_date = get_spot_data_date()
    cached = get_cached_contracts()
    cache_str = f"<b>{'、'.join(cached)}</b>" if cached else "暂无"
    active_cts = get_active_contracts()
    active_str = f"<b>{'、'.join(active_cts)}</b>" if active_cts else "识别中…"
    st.markdown(f"""
    <p class="footer-info">
        📊 当前上市合约：{active_str}（{len(active_cts)}个）&nbsp;｜&nbsp;
        📅 期货数据更新日期：<b>{fut_update_date}</b> &nbsp;｜&nbsp;
        📅 现货数据更新日期：<b>{spot_update_date}</b> &nbsp;｜&nbsp;
        📦 已缓存合约：<b>{len(cached)}</b> 个
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
    st.markdown("<p style='text-align:right; font-size:0.75rem; color:#adb5bd; margin-top:-8px;'>创作者：chen</p>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
