#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
生猪期货分析平台 (Streamlit)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
无侧边栏设计——所有控件位于各 Tab 内部。
Tab 1 — 每日期货分析日报（新增：HTML 日报 + PDF/Markdown 下载）
Tab 2 — 当日基差分布（柱状图 + 四指标卡片）
Tab 3 — 单合约基差走势（区域色板 + 汇总指标固定色）
Tab 4 — 合约基差比较（同比 / 交易日对齐，颜色按合约年份，含现货+期货走势图）
Tab 5 — 合约价差比较（月份选择，颜色按合约年份）
Tab 6 — 持仓与成交分析（双轴图 + 前20期货公司多空持仓）
Tab 7 — 季节性持仓对比（同月份合约跨年成交量/持仓量/净持仓）
Tab 8 — 技术分析（K线图 + MA/布林带 + MACD/RSI/KDJ + 文字结论）
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
from concurrent.futures import ThreadPoolExecutor, as_completed
from io import BytesIO, StringIO
import warnings

warnings.filterwarnings("ignore")

# ── PDF 导出支持 ──
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.colors import HexColor
    from reportlab.lib.units import mm
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                     TableStyle)
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    HAS_REPORTLAB = True
    # 注册中文字体
    try:
        pdfmetrics.registerFont(TTFont('SimHei', 'C:/Windows/Fonts/simhei.ttf'))
        CN_FONT = 'SimHei'
    except Exception:
        try:
            pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
            CN_FONT = 'STSong-Light'
        except Exception:
            CN_FONT = 'Helvetica'
except ImportError:
    HAS_REPORTLAB = False

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
AVG_LINE_COLOR = "#6B7280"
AVG_LINE_WIDTH = 0.8
AVG_LINE_DASH = "dot"
# 后台并行下载线程数
_sync_parallel_workers = 5

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
# 正指 / 反指 期货公司
# ══════════════════════════════════════════════════════════════
ZHENGZHI_COMPANIES = {"国泰君安", "中粮期货", "东证期货"}
FANZHI_COMPANIES = {"东方财富", "徽商期货", "平安期货"}

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
# ══════════════════════════════════════════════════════════════
# 期货数据：本地文件 → 内存缓存（纯读取，零网络）
# ══════════════════════════════════════════════════════════════

def _csv_path(ct: str) -> Path:
    return FUTURES_DIR / f"{ct}.csv"


def _contract_status(ct: str) -> str:
    """判断合约状态：'active'（仍在交易）, 'expired'（已到期）, 'no_data'（无本地文件）
    使用合约交割月+数据新鲜度双重判断，避免数据延迟导致误判为到期。"""
    cp = _csv_path(ct)
    if not cp.exists():
        return "no_data"
    try:
        df = pd.read_csv(cp, usecols=["date"])
        if df.empty:
            return "no_data"
        max_date = pd.to_datetime(df["date"].max()).date()
        days_behind = (datetime.now().date() - max_date).days

        # 数据在 7 天内 → 活跃
        if days_behind <= 7:
            return "active"

        # 判断合约是否真正到期：交割月份距今超过 2 个月
        try:
            ct_year = int(f"20{ct[2:4]}")
            ct_month = int(ct[4:6])
            delivery_date = datetime(ct_year, ct_month, 15).date()
            months_since_delivery = (datetime.now().date().year - delivery_date.year) * 12 + \
                                     (datetime.now().date().month - delivery_date.month)
            # 交割后超过 2 个月且数据超过 7 天 → 真正到期
            if months_since_delivery > 2:
                return "expired"
        except Exception:
            pass

        # 交割月未过或刚过 → 数据只是延迟，仍视为活跃
        return "active"
    except Exception:
        return "no_data"


def get_cached_contracts() -> List[str]:
    if not FUTURES_DIR.exists(): return []
    return sorted(f.stem for f in FUTURES_DIR.glob("LH*.csv"))

def _get_global_latest_date() -> Optional[pd.Timestamp]:
    """获取所有合约中最新的交易日"""
    latest = None
    for f in FUTURES_DIR.glob("LH*.csv"):
        try:
            df = pd.read_csv(f, usecols=["date"])
            if df.empty: continue
            d = pd.to_datetime(df["date"].max())
            if latest is None or d > latest: latest = d
        except Exception: pass
    return latest

@st.cache_data(ttl=300)
def get_active_contracts() -> List[str]:
    """动态识别当前上市合约，并确保兜底列表始终包含"""
    # 硬编码保底列表：当前正在交易的合约
    FALLBACK_ACTIVE = ['LH2609', 'LH2611', 'LH2701', 'LH2703', 'LH2705', 'LH2707']
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


def get_main_contract() -> str:
    """返回主力合约（活跃合约中近20日成交量最高的）。
    兜底返回活跃合约列表中最后一个，或硬编码 LH2609。"""
    active = get_active_contracts()
    best_ct, best_vol = None, 0
    for ct in active:
        cp = _csv_path(ct)
        if not cp.exists():
            continue
        try:
            df = pd.read_csv(cp, usecols=["volume"])
            if df.empty or "volume" not in df.columns:
                continue
            recent_vol = df["volume"].tail(20).mean()
            if recent_vol > best_vol:
                best_vol, best_ct = recent_vol, ct
        except Exception:
            continue
    return best_ct or (active[-1] if active else "LH2609")


def get_latest_futures_date() -> Optional[str]:
    latest = None
    for f in FUTURES_DIR.glob("LH*.csv"):
        try:
            df = pd.read_csv(f, usecols=["date"])
            if df.empty: continue
            d = pd.to_datetime(df["date"].max())
            if latest is None or d > latest: latest = d
        except Exception: pass
    return _cn(latest) if latest else None

@st.cache_data(ttl=300)
def get_latest_trade_date() -> Optional[pd.Timestamp]:
    """获取所有合约CSV中最新的交易日（全局最大值）"""
    return _get_global_latest_date()


# ── 读取：优先本地 CSV，缺失时惰性同步 ──

@st.cache_data(ttl=3600)
def load_futures(ct: str) -> Tuple[Optional[pd.DataFrame], str]:
    """读取期货数据。本地有 CSV → 直接返回；本地无 → 惰性同步这一个合约。"""
    cp = _csv_path(ct)

    # 本地有数据 → 直接读，不联网
    if cp.exists():
        try:
            df = pd.read_csv(cp)
            if "date" not in df.columns or df.empty:
                return None, "❌ 数据为空"
            df["date"] = pd.to_datetime(df["date"])
            return df.sort_values("date").reset_index(drop=True), "📁 本地缓存"
        except Exception:
            return None, "❌ 读取失败"

    # 本地无数据 → 惰性同步这一个合约（仅首次触发，缓存后不再走此分支）
    ok, msg = sync_futures(ct, force_full=True)
    if ok and cp.exists():
        try:
            df = pd.read_csv(cp)
            df["date"] = pd.to_datetime(df["date"])
            return df.sort_values("date").reset_index(drop=True), f"🌐 首次同步"
        except Exception:
            return None, "❌ 同步后读取失败"
    return None, f"❌ 同步失败：{msg}"


# ── 同步：全量 / 增量下载，写入本地 CSV ──

def sync_futures(ct: str, force_full: bool = False) -> Tuple[bool, str]:
    """同步合约期货数据到本地 CSV。
    - 已到期合约（停更 7 天+）跳过
    - 本地有数据 → 增量下载最新
    - 本地无数据 → 全量下载
    返回 (成功与否, 状态信息)"""
    status = _contract_status(ct)

    # 已到期合约永远跳过（除非 force_full）
    if status == "expired" and not force_full:
        return True, "⏭️ 已到期，跳过"

    cp = _csv_path(ct)
    today = datetime.now()

    # ── 增量更新（本地有数据 且 不强制全量）──
    if cp.exists() and not force_full:
        try:
            old = pd.read_csv(cp)
            old["date"] = pd.to_datetime(old["date"])
            start_dt = old["date"].max().date() + timedelta(days=1)
            if start_dt > today.date():
                return True, "📁 已是最新"
            new = _download_futures(ct, start_dt.strftime("%Y%m%d"), today.strftime("%Y%m%d"))
            if new is not None and not new.empty:
                combined = pd.concat([old, new], ignore_index=True)
                combined = combined.drop_duplicates(subset=["date"]).sort_values("date")
                combined.to_csv(cp, index=False)
                return True, f"🔄 增量 +{len(new)}条"
            return True, "📁 已是最新"
        except Exception as e:
            return False, f"❌ 增量失败：{e}"

    # ── 全量下载 ──
    try:
        start = (today - timedelta(days=365 * 6)).strftime("%Y%m%d")
        end = today.strftime("%Y%m%d")
        df = _download_futures(ct, start, end)
        if df is not None and not df.empty:
            df.to_csv(cp, index=False)
            # 清除该合约的 load_futures 缓存，让它下次读到最新数据
            load_futures.clear(ct)
            return True, f"🌐 全量 {len(df)}条"
        return False, "❌ 网络无数据"
    except Exception as e:
        return False, f"❌ 下载失败：{e}"


def sync_active_contracts(silent: bool = True) -> Dict[str, str]:
    """启动时同步所有活跃合约（增量为主）。返回 {合约: 状态}。"""
    results = {}
    active = get_active_contracts()
    for ct in active:
        ok, msg = sync_futures(ct)
        results[ct] = msg
    return results


def sync_all_contracts(max_workers: int = 5, progress_callback=None) -> Dict[str, str]:
    """后台下载 ALL_CONTRACTS 中所有缺失的合约期货数据（并行）。

    只下载本地无 CSV 的合约，已有数据的合约秒过。
    max_workers: 并行下载线程数。
    progress_callback: 可选，签名 (current, total, contract, status) -> None。
    返回 {合约: 状态信息}。
    """
    missing = [c for c in ALL_CONTRACTS if not _csv_path(c).exists()]
    if not missing:
        return {}

    results = {}
    total = len(missing)

    def _download_one(ct):
        ok, msg = sync_futures(ct, force_full=True)
        return ct, ok, msg

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_map = {executor.submit(_download_one, ct): ct for ct in missing}
        completed = 0
        for future in as_completed(future_map):
            ct, ok, msg = future.result()
            results[ct] = msg
            completed += 1
            if progress_callback:
                progress_callback(completed, total, ct, msg)
            # 清除该合约的 load_futures 缓存，确保后续读取最新数据
            if ok:
                load_futures.clear(ct)

    return results


def _download_futures(ct: str, sd: str, ed: str) -> Optional[pd.DataFrame]:
    """底层网络下载（akshare → eastmoney 兜底，减少重试避免卡顿）"""
    # akshare：只试 1 次
    try:
        import akshare as ak
        df = ak.futures_zh_daily_sina(symbol=ct)
        if df is not None and not df.empty:
            df["date"] = pd.to_datetime(df["date"]); df = df.sort_values("date")
            return df[(df["date"]>=sd)&(df["date"]<=ed)].reset_index(drop=True)
    except Exception:
        pass
    # eastmoney 兜底：只试 1 次
    try:
        r = requests.get("https://push2his.eastmoney.com/api/qt/stock/kline/get",
            params={"secid":f"114.{ct}","fields1":"f1,f2,f3,f4,f5,f6","fields2":"f51,f52,f53,f54,f55,f56,f57","klt":"101","fqt":"1","end":"20500101","lmt":"3000"},
            timeout=8, headers={"User-Agent":"Mozilla/5.0","Referer":"https://quote.eastmoney.com/"})
        if r.status_code == 200:
            d = r.json()
            if d and d.get("data") and d["data"].get("klines"):
                recs = []
                for k in d["data"]["klines"]:
                    p = k.split(",")
                    if len(p) >= 7:
                        recs.append({"date":pd.to_datetime(p[0]),"open":float(p[1]),"close":float(p[2]),
                                     "high":float(p[3]),"low":float(p[4]),"volume":int(float(p[5])),"settle":float(p[2]),"hold":0})
                if recs: return pd.DataFrame(recs).sort_values("date").reset_index(drop=True)
    except Exception:
        pass
    return None


def _get_row_at_md(df, target_month: int, target_day: int):
    """在DataFrame中查找指定月/日的行，若不存在则取目标日期之前最近的行，兜底取最后一行"""
    if df is None or df.empty:
        return None
    mask = (df['date'].dt.month == target_month) & (df['date'].dt.day == target_day)
    match = df[mask]
    if not match.empty:
        return match.iloc[-1]
    target_ordinal = target_month * 100 + target_day
    df_copy = df.copy()
    df_copy['_md'] = df_copy['date'].dt.month * 100 + df_copy['date'].dt.day
    before = df_copy[df_copy['_md'] <= target_ordinal]
    if not before.empty:
        return before.iloc[-1]
    return df.iloc[-1]


def get_spot_data_date() -> str:
    """从现货Excel文件名或内部数据提取最新日期"""

    def _date_from_filename(fname: str) -> Optional[str]:
        m = re.search(r"(\d{4})[年\-/](\d{1,2})[月\-/](\d{1,2})", fname)
        if m:
            return f"{m.group(1)}年{int(m.group(2)):02d}月{int(m.group(3)):02d}日"
        return None

    def _date_from_excel(f: Path) -> Optional[str]:
        try:
            xls = pd.ExcelFile(f)
            latest_dt = None
            for sheet_name in xls.sheet_names[:3]:
                try:
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                    for col in df.columns:
                        s = pd.to_datetime(df[col], errors="coerce")
                        valid = s.dropna()
                        if not valid.empty:
                            col_max = valid.max()
                            if pd.Timestamp("2020-01-01") <= col_max <= pd.Timestamp("2030-12-31"):
                                if latest_dt is None or col_max > latest_dt:
                                    latest_dt = col_max
                except Exception:
                    pass
            if latest_dt is not None:
                return _cn(latest_dt)
        except Exception:
            pass
        return None

    # 1. 优先使用 SPOT_PATH
    spot_path = SPOT_PATH
    if spot_path.exists():
        d = _date_from_filename(spot_path.name)
        if d: return d
        d = _date_from_excel(spot_path)
        if d: return d

    # 2. 扫描桌面和项目目录
    candidates = []
    desktop = Path(r"D:\CC\Desktop")
    for search_dir in [desktop, DATA_DIR]:
        if not search_dir.exists(): continue
        try:
            for f in search_dir.iterdir():
                if not f.is_file() or f.suffix not in ('.xlsx', '.xls'): continue
                m = re.search(r"(\d{4})[年\-/](\d{1,2})[月\-/](\d{1,2})", f.name)
                if m:
                    candidates.append((f.stat().st_mtime, f, m))
        except Exception:
            continue

    if candidates:
        candidates.sort(key=lambda x: x[0], reverse=True)
        _, best_file, best_match = candidates[0]
        return f"{best_match.group(1)}年{int(best_match.group(2)):02d}月{int(best_match.group(3)):02d}日"

    # 3. 从内容提取日期
    for search_dir in [desktop, DATA_DIR]:
        if not search_dir.exists(): continue
        try:
            for f in sorted(search_dir.glob("*.xlsx"), key=lambda x: x.stat().st_mtime, reverse=True):
                d = _date_from_excel(f)
                if d: return d
        except Exception:
            continue

    return "无现货数据"


# ══════════════════════════════════════════════════════════════
# 基差计算
# ══════════════════════════════════════════════════════════════
def _to_ton(p: float) -> float:
    return float(p) * 1000

@st.cache_data(ttl=3600, show_spinner=False)
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
def calculate_technicals(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    """计算完整技术指标：MA5/10/20/60, 布林带, MACD, RSI14, KDJ。
    返回 (df, warnings) — warnings 为数据不足等提示列表。"""
    warnings_list: List[str] = []
    if df is None or df.empty:
        return df, ["数据为空，无法计算技术指标"]

    n_rows = len(df)
    try:
        close = df["close"].astype(float)
        high = df["high"].astype(float)
        low = df["low"].astype(float)
    except KeyError as e:
        return df, [f"缺少必要列：{e}"]

    # ── 数据量检查 ──
    if n_rows < 20:
        warnings_list.append(f"数据仅 {n_rows} 个交易日，MA20/布林带/MACD 可能不完整")
    if n_rows < 60:
        warnings_list.append(f"数据仅 {n_rows} 个交易日（<60），MA60 不可用，请扩大日期范围")

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

    # ── KDJ (9, 3, 3) — 向量化，避免 Python 循环 ──
    n = 9
    lowest_low = low.rolling(n).min()
    highest_high = high.rolling(n).max()
    rsv = ((close - lowest_low) / (highest_high - lowest_low).replace(0, np.nan)) * 100
    # ewm(alpha=1/3, adjust=False) 等价于 K = 2/3*K_prev + 1/3*RSV
    k_series = rsv.ewm(alpha=1/3, adjust=False).mean()
    # 前 n-1 个值为 NaN，手动回退到 50
    k_series = k_series.fillna(50.0)
    d_series = k_series.ewm(alpha=1/3, adjust=False).mean()
    j_series = 3 * k_series - 2 * d_series
    df["kdj_k"] = k_series
    df["kdj_d"] = d_series
    df["kdj_j"] = j_series

    return df, warnings_list


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
        hovertemplate="<b>%{x}</b><br>基差：%{y:+,}元/吨<br>现货：%{customdata[0]:.2f}元/公斤<br>期货：%{customdata[1]:.0f}元/吨<br>升贴水：%{customdata[2]:+d}元/吨<extra></extra>",
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
        # ★ 从 label 提取合约代码用于点击联动（格式如 "2409 (2024) 全国均价" → LH2409）
        ct_code = "LH" + label[:4] if len(label) >= 4 else ""
        # 用 lines+markers 确保有点可点击（marker 极小透明，不影响视觉）
        fig.add_trace(go.Scatter(x=df["plot_date"], y=df["basis"],
            mode="lines+markers", name=label,
            line=dict(color=c, width=w, dash=d),
            marker=dict(size=4, opacity=0.01, color=c),
            hovertemplate=f"<b>{label}</b><br>%{{customdata[1]}}<br>基差：%{{y:+,}}元/吨<extra></extra>",
            customdata=[[ct_code, _cn_md(r["plot_date"])] for _,r in df.iterrows()]))
    fig.add_hline(y=0, line_dash="dash", line_color="gray", opacity=0.5)
    title = f"{tmon}月合约基差季节图（同比 — 自然日对齐）"
    if data_date: title += f"<br><sup>📡 期货数据来源：akshare，数据日期：{data_date}</sup>"
    fig.update_layout(title=title, xaxis_title="日期（月-日）", yaxis_title="基差（元/吨）",
        template="plotly_white", height=550, hovermode="x unified",
        legend=dict(orientation="h", y=1.02, x=0), margin=dict(t=80,b=40,l=60,r=40))
    fig.update_xaxes(tickformat="%m-%d", dtick="M1", range=["2020-01-01","2020-12-31"])
    fig.update_yaxes(autorange=True)
    return fig


def _make_spot_futures_chart(ct: str, spot_dict: dict) -> go.Figure:
    """双轴图：左轴现货价格（元/公斤），右轴期货收盘价（元/吨）。
    用于 Tab3 点击合约线时联动展示。"""
    fut_df, _ = load_futures(ct)
    if fut_df is None or fut_df.empty:
        return go.Figure()

    df = fut_df.sort_values("date").reset_index(drop=True)
    fig = go.Figure()

    # 期货收盘价（右轴，红色）
    fig.add_trace(go.Scatter(
        x=df["date"], y=df["close"],
        name=f"{ct} 期货收盘价", mode="lines",
        line=dict(color="#E74C3C", width=2),
        yaxis="y2",
        hovertemplate="<b>%{x|%Y年%m月%d日}</b><br>期货：%{y:.0f}元/吨<extra></extra>"
    ))

    # 现货 — 全国均价（左轴，绿色）
    national_spot = spot_dict.get("全国均价")
    if national_spot is not None and "date" in national_spot.columns:
        merged = df[["date", "close"]].merge(
            national_spot[["date", "price"]], on="date", how="inner"
        )
        if not merged.empty:
            fig.add_trace(go.Scatter(
                x=merged["date"], y=merged["price"],
                name="现货（全国均价）", mode="lines",
                line=dict(color="#27AE60", width=2),
                yaxis="y",
                hovertemplate="<b>%{x|%Y年%m月%d日}</b><br>现货：%{y:.2f}元/公斤<extra></extra>"
            ))

    fig.update_layout(
        title=f"{ct} 现货与期货价格走势",
        xaxis=dict(title="日期", tickformat="%Y年%m月"),
        yaxis=dict(title="现货价格（元/公斤）", side="left", showgrid=True),
        yaxis2=dict(title="期货收盘价（元/吨）", side="right", overlaying="y", showgrid=False),
        template="plotly_white", height=350,
        hovermode="x unified",
        legend=dict(orientation="h", y=1.02, x=0),
        margin=dict(t=50, b=30, l=60, r=60),
    )
    fig.update_xaxes(rangeslider_visible=True)
    fig.update_yaxes(autorange=True)
    return fig


def _resolve_selected_item_for_chart(sel_items, spot_dict, ref_regions):
    """解析用户选择的区域/指标，用于现货+期货走势图联动。
    返回 {'type': 'region'|'indicator', 'label': str, 'title': str, 'region_name': str}"""
    if not sel_items:
        return {"type": "indicator", "label": "全国均价", "title": "全国均价", "region_name": "全国均价"}
    first = sel_items[0]
    # 检查是否为汇总指标
    if "全国均价" in first:
        return {"type": "indicator", "label": "全国均价", "title": "全国均价基差", "region_name": "全国均价"}
    if "最大基差" in first:
        return {"type": "indicator", "label": "最大基差", "title": "最大基差", "region_name": "最大基差"}
    if "最小基差" in first:
        return {"type": "indicator", "label": "最小基差", "title": "最小基差", "region_name": "最小基差"}
    if "基差平均值" in first:
        return {"type": "indicator", "label": "基差平均值", "title": "基差平均值", "region_name": "基差平均值"}
    if "───" in first:
        # Separator - use next item or fallback
        for item in sel_items[1:]:
            if "───" not in item:
                return _resolve_selected_item_for_chart([item], spot_dict, ref_regions)
        return {"type": "indicator", "label": "全国均价", "title": "全国均价", "region_name": "全国均价"}
    # Check for region
    if first in ref_regions:
        region_name = first if first in spot_dict else "全国均价"
        return {"type": "region", "label": first, "title": f"{first}现货", "region_name": region_name}
    return {"type": "indicator", "label": "全国均价", "title": "全国均价", "region_name": "全国均价"}


def _make_spot_futures_chart_with_item(ct: str, spot_dict: dict, sel_item: dict) -> go.Figure:
    """根据用户选择的区域/指标，生成现货+期货双轴走势图。
    sel_item 来自 _resolve_selected_item_for_chart"""
    import pandas as pd

    fut_df, _ = load_futures(ct)
    if fut_df is None or fut_df.empty:
        return go.Figure()

    df = fut_df.sort_values("date").reset_index(drop=True)
    fig = go.Figure()

    # 期货收盘价（右轴）
    fig.add_trace(go.Scatter(
        x=df["date"], y=df["close"],
        name=f"{ct} 期货收盘价", mode="lines",
        line=dict(color="#E74C3C", width=2),
        yaxis="y2",
        hovertemplate="<b>%{x|%Y年%m月%d日}</b><br>期货：%{y:.0f}元/吨<extra></extra>"
    ))

    item_type = sel_item.get("type", "indicator")
    region_name = sel_item.get("region_name", "全国均价")

    if item_type == "region":
        # 指定区域的现货价格
        spot_df = spot_dict.get(region_name)
        if spot_df is not None and "date" in spot_df.columns:
            merged = df[["date", "close"]].merge(
                spot_df[["date", "price"]], on="date", how="inner"
            )
            if not merged.empty:
                fig.add_trace(go.Scatter(
                    x=merged["date"], y=merged["price"],
                    name=f"现货（{region_name}）", mode="lines",
                    line=dict(color="#27AE60", width=2),
                    yaxis="y",
                    hovertemplate="<b>%{x|%Y年%m月%d日}</b><br>现货：%{y:.2f}元/公斤<extra></extra>"
                ))
    else:
        # 汇总指标：全国均价
        national_spot = spot_dict.get("全国均价")
        if national_spot is not None and "date" in national_spot.columns:
            merged = df[["date", "close"]].merge(
                national_spot[["date", "price"]], on="date", how="inner"
            )
            if not merged.empty:
                fig.add_trace(go.Scatter(
                    x=merged["date"], y=merged["price"],
                    name=f"现货（{sel_item.get('title', '全国均价')}）", mode="lines",
                    line=dict(color="#27AE60", width=2),
                    yaxis="y",
                    hovertemplate="<b>%{x|%Y年%m月%d日}</b><br>现货：%{y:.2f}元/公斤<extra></extra>"
                ))

    fig.update_layout(
        title=f"{ct} 现货与期货价格走势 — {sel_item.get('title', '全国均价')}",
        xaxis=dict(title="日期", tickformat="%Y年%m月"),
        yaxis=dict(title="现货价格（元/公斤）", side="left", showgrid=True),
        yaxis2=dict(title="期货收盘价（元/吨）", side="right", overlaying="y", showgrid=False),
        template="plotly_white", height=350,
        hovermode="x unified",
        legend=dict(orientation="h", y=1.02, x=0),
        margin=dict(t=50, b=30, l=60, r=60),
    )
    fig.update_xaxes(rangeslider_visible=True)
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

def _compute_y_padding(all_values: list, padding_pct: float = 0.08):
    """根据数据范围计算 Y 轴显示范围，上下各留 padding_pct 百分比边距。
    确保零轴始终可见（若数据跨越正负）。"""
    if not all_values:
        return None
    ymin, ymax = min(all_values), max(all_values)
    # 若数据全为正或全为负，向零轴方向扩展
    if ymin > 0:
        ymin = 0
    if ymax < 0:
        ymax = 0
    data_range = ymax - ymin
    if data_range == 0:
        data_range = max(abs(ymax), 100)
    pad = data_range * padding_pct
    return [ymin - pad, ymax + pad]

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
                tbl["现货（元/公斤）"] = tbl["spot_price"].apply(lambda x: f"{x:.2f}" if x > 0 else "—")
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
        # ★ 默认选中主力合约，自动对比所有同月历史合约
        main_ct = get_main_contract()
        st.caption(f"🔍 主力合约：**{ct_display(main_ct)}** ｜ 已自动加载所有同月历史合约进行对比")
        default_t3 = [main_ct] if main_ct in ALL_CONTRACTS else [c for c in active_cts if c in ALL_CONTRACTS]
        contracts = st.multiselect("📋 合约选择（多选）", options=ALL_CONTRACTS, default=default_t3,
            format_func=ct_display, key="t3_ct")
        if not contracts: contracts = [main_ct if main_ct in ALL_CONTRACTS else active_cts[-1] if active_cts else ALL_CONTRACTS[-1]]

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
    skipped = []
    for c in same_month:
        df, _ = load_futures(c)
        if df is not None and not df.empty:
            avail.append(c)
        else:
            skipped.append(c)
    if len(avail) < 1:
        st.warning(f"⚠️ {tmon}月合约暂无可用的历史数据"); return
    info_msg = f"📌 {tmon}月合约，共 {len(avail)} 个可用：{'、'.join(avail)}"
    if skipped:
        info_msg += f"  ｜ ⚠️ 无数据：{'、'.join(skipped)}"
    st.info(info_msg)

    series: Dict[str, pd.DataFrame] = {}
    # ★ 使用缓存批量计算基差序列（性能优化 — 避免重复 load_futures / calc_basis）
    shash = _spot_hash(spot_dict)
    series = _build_calendar_series_cached(
        tuple(avail), tuple(sel_items), shash, tuple(ref_regions)
    )

    if not series: st.warning("⚠️ 无可用数据"); return

    fig = fig_calendar_comparison(series, tmon, data_date)
    # ★ 点击联动：选中某条合约线时展示该合约的现货+期货价格双轴图
    sel_event = st.plotly_chart(fig, use_container_width=True, on_select="rerun",
                                selection_mode="points", key="t3_calendar")

    _clicked_ct = None
    if sel_event is not None:
        # Streamlit 1.55+ 返回 PlotlySelection；兼容多种访问路径
        pts = None
        try:
            pts = sel_event.selection.points
        except AttributeError:
            try:
                pts = sel_event.get("selection", {}).get("points", [])
            except Exception:
                pass

        if pts:
            pt = pts[0]
            # 优先从 customdata 取合约代码
            cd = pt.get("customdata", None) if isinstance(pt, dict) else getattr(pt, "customdata", None)
            if cd and isinstance(cd, (list, tuple)) and len(cd) >= 1:
                _clicked_ct = str(cd[0]) if cd[0] else None
            # Fallback: 从 curve_number 反查
            if not _clicked_ct:
                cn = pt.get("curve_number", None) if isinstance(pt, dict) else getattr(pt, "curve_number", None)
                labels = [l for l in series.keys() if "历史均值" not in l]
                if cn is not None and 0 <= cn < len(labels):
                    lbl = labels[cn]
                    _clicked_ct = "LH" + lbl[:4] if len(lbl) >= 4 else None

    if _clicked_ct:
        st.markdown(f"---")
        st.markdown(f"### 📈 {_clicked_ct} 现货与期货价格走势")
        fig_sf = _make_spot_futures_chart(_clicked_ct, spot_dict)
        if fig_sf.data:
            st.plotly_chart(fig_sf, use_container_width=True)
        else:
            st.caption(f"⚠️ 无法加载 {_clicked_ct} 的现货/期货数据")
    else:
        # ★ 默认显示：响应 sel_items 选择的区域/指标
        default_ct = contracts[0] if contracts else get_main_contract()
        sel_label = _resolve_selected_item_for_chart(sel_items, spot_dict, ref_regions)
        st.markdown(f"---")
        st.markdown(f"### 📈 {default_ct} 现货与期货价格走势 — {sel_label['title']}（点击上方合约线可切换）")
        fig_sf = _make_spot_futures_chart_with_item(default_ct, spot_dict, sel_label)
        if fig_sf.data:
            st.plotly_chart(fig_sf, use_container_width=True)
        else:
            st.caption(f"⚠️ 无法加载 {default_ct} 的现货/期货数据")

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
            # ★ 使用 load_futures 而非仅检查 CSV 存在 —— 缺失数据自动下载，不跳过任何年份
            dfa, _ = load_futures(ca)
            dfb, _ = load_futures(cb)
            ca_ok = dfa is not None and not dfa.empty
            cb_ok = dfb is not None and not dfb.empty
            if ca_ok and cb_ok: valid_years.append(y)
            elif ca_ok or cb_ok: skipped_years.append(y)

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
        td_fallback = False
        if td not in fds:
            nearby = [d for d in fds if d <= td]
            if nearby:
                td = nearby[-1]
                td_fallback = True

        # ── 所选日期非交易日提示 ──
        if td_fallback:
            st.info(
                f"📅 **{_cn(pd.to_datetime(sel_date))}** 为非交易日，"
                f"已自动切换至最近交易日 **{_cn(td)}**"
            )

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
        col_title, col_legend = st.columns([3, 1])
        with col_title:
            st.markdown("#### 🏢 前20期货公司多空持仓")
        with col_legend:
            st.markdown(
                '<div style="display:flex;align-items:center;gap:16px;padding-top:4px;font-size:0.9rem;">'
                '<span style="display:inline-block;width:14px;height:14px;background:#E74C3C;'
                'border-radius:2px;vertical-align:middle;margin-right:4px;"></span> 多单'
                '<span style="display:inline-block;width:14px;height:14px;background:#3498DB;'
                'border-radius:2px;vertical-align:middle;margin-right:4px;margin-left:12px;"></span> 空单'
                '</div>',
                unsafe_allow_html=True,
            )

        # 获取前一个交易日
        prev_td = None
        fds_sorted = sorted(fds)
        td_idx = fds_sorted.index(td) if td in fds_sorted else -1
        if td_idx > 0:
            prev_td = fds_sorted[td_idx - 1]
        elif td not in fds_sorted:
            # td 不在 fds 中（已回退到 nearby），取其前一个
            for i, d in enumerate(fds_sorted):
                if d >= td:
                    if i > 0:
                        prev_td = fds_sorted[i - 1]
                    break

        holdings_df, holdings_actual_date, holdings_source = _get_holdings(ct, td, return_meta=True)
        if holdings_df is not None and not holdings_df.empty:
            # ── 检查数据日期是否匹配 ──
            holdings_date_mismatch = False
            holdings_actual_dt = None
            try:
                holdings_actual_dt = pd.to_datetime(holdings_actual_date, format="%Y%m%d")
                if holdings_actual_dt.date() != td.date():
                    holdings_date_mismatch = True
            except Exception:
                pass

            # ── 日期不匹配 / 数据源提示 ──
            if holdings_date_mismatch:
                sel_date_str = _cn(td)
                actual_date_str = _cn(holdings_actual_dt) if holdings_actual_dt is not None else holdings_actual_date
                days_behind = (td.date() - holdings_actual_dt.date()).days if holdings_actual_dt else 999

                if holdings_source in ("akshare", "akshare_fallback"):
                    # ★ API 正常，但所选日期数据尚未发布（大商所 T+1，正常现象）
                    st.info(
                        f"📡 **{sel_date_str}** 的持仓排名数据尚未发布"
                        f"（大商所通常 T+1 更新），"
                        f"当前显示的是最新可用数据 **{actual_date_str}**"
                        f"（{days_behind}天前，来自新浪财经）。"
                    )
                else:
                    # ★ API 完全失败，使用的是本地缓存
                    error_key = f"{ct}_{td.strftime('%Y%m%d')}"
                    api_errors = st.session_state.get("_holdings_api_errors", {})
                    error_detail = api_errors.get(error_key, "")
                    error_hint = ""
                    if error_detail:
                        if "timeout" in error_detail.lower() or "timed out" in error_detail.lower():
                            error_hint = "（接口请求超时，可能是网络波动或新浪服务器繁忙）"
                        elif "connection" in error_detail.lower() or "refused" in error_detail.lower():
                            error_hint = "（接口连接失败，请检查网络或稍后重试）"
                        elif "数据尚未发布" in error_detail:
                            error_hint = "（所选日期及近 5 个交易日数据均未发布，大商所通常 T+1 更新）"
                        else:
                            error_hint = f"（接口异常：{error_detail[:80]}）"
                    st.warning(
                        f"⚠️ **{sel_date_str}** 暂无前20期货公司多空持仓数据，"
                        f"当前显示的是最近可用数据 **{actual_date_str}**。"
                        f"（数据来源：{holdings_source}）{error_hint}"
                    )
            elif holdings_source == "unavailable":
                st.warning("⚠️ 前20期货公司多空持仓数据暂不可用，API 和本地缓存均无数据。")

            # ── 分组柱状图 ──
            top_n = min(20, len(holdings_df))
            disp = holdings_df.head(top_n).copy()
            # 计算净持仓
            disp["净持仓"] = disp["long"] - disp["short"]

            # ── 日变化：优先用 API 返回的增减列（long_chg / short_chg） ──
            api_long_chg = "long_chg" in disp.columns
            api_short_chg = "short_chg" in disp.columns
            if api_long_chg:
                disp["long_change"] = disp["long_chg"].fillna(0).astype(int)
            else:
                disp["long_change"] = 0
            if api_short_chg:
                disp["short_change"] = disp["short_chg"].fillna(0).astype(int)
            else:
                disp["short_change"] = 0

            if not api_long_chg or not api_short_chg:
                # 回退：从前一日数据推算（仅对 API 未提供的列）
                prev_holdings = None
                if prev_td is not None:
                    prev_holdings = _get_holdings(ct, prev_td)
                if prev_holdings is not None and not prev_holdings.empty:
                    prev_map = {r["company"]: r for _, r in prev_holdings.iterrows()}
                    for i, row in disp.iterrows():
                        co = row["company"]
                        if co in prev_map:
                            if not api_long_chg:
                                disp.at[i, "long_change"] = int(row["long"]) - int(prev_map[co]["long"])
                            if not api_short_chg:
                                disp.at[i, "short_change"] = int(row["short"]) - int(prev_map[co]["short"])

            # ── 添加正指/反指标签 ──
            def _tag_company(name: str) -> str:
                if name in ZHENGZHI_COMPANIES:
                    return f"{name} 🟢正"
                if name in FANZHI_COMPANIES:
                    return f"{name} 🔴反"
                return name

            disp["company_label"] = disp["company"].apply(_tag_company)

            # ── 日变化文字（加多/减多/加空/减空） ──
            def _change_text(lc: int, sc: int) -> str:
                parts = []
                if lc > 0:
                    parts.append(f"加多+{lc}")
                elif lc < 0:
                    parts.append(f"减多{lc}")
                if sc > 0:
                    parts.append(f"加空+{sc}")
                elif sc < 0:
                    parts.append(f"减空{sc}")
                return "  ".join(parts) if parts else ""

            disp["change_label"] = disp.apply(
                lambda r: _change_text(int(r["long_change"]), int(r["short_change"])), axis=1)

            # ── 构建柱状图 ──
            fig_h = go.Figure()

            # 多单 — 文字标在柱右侧（与空单同侧）
            fig_h.add_trace(go.Bar(
                y=disp["company_label"], x=disp["long"],
                name="多单", orientation="h",
                marker_color="#E74C3C", opacity=0.85,
                text=[f"{v:,}" for v in disp["long"]],
                textposition="outside",
                textfont=dict(color="#E74C3C", size=13),
                customdata=disp[["short", "long_change", "short_change", "change_label"]].values,
                hovertemplate=(
                    "<b>%{y}</b><br>"
                    "多单：%{x:,}手<br>"
                    "%{customdata[3]}"
                    "<extra></extra>"
                ),
            ))

            # 空单 — 文字标在柱右侧
            fig_h.add_trace(go.Bar(
                y=disp["company_label"], x=disp["short"],
                name="空单", orientation="h",
                marker_color="#3498DB", opacity=0.85,
                text=[f"{v:,}" for v in disp["short"]],
                textposition="outside",
                textfont=dict(color="#3498DB", size=13),
                customdata=disp[["long", "long_change", "short_change", "change_label"]].values,
                hovertemplate=(
                    "<b>%{y}</b><br>"
                    "空单：%{x:,}手<br>"
                    "%{customdata[3]}"
                    "<extra></extra>"
                ),
            ))

            # ── 日变化标注（图表右侧，远离柱体） ──
            annotations = []
            max_x = max(disp["long"].max(), disp["short"].max())
            for i, (_, row) in enumerate(disp.iterrows()):
                cl = row["change_label"]
                if cl:
                    # 颜色：加多为红色，减多为绿色（加空为蓝色，减空为浅蓝）
                    annotations.append(dict(
                        x=max_x * 1.22,
                        y=row["company_label"],
                        text=f"<span style='font-size:10px;color:#555'>{cl}</span>",
                        showarrow=False,
                        xanchor="left",
                        yanchor="middle",
                    ))

            # ── 构建标题，标注数据来源和日期 ──
            title_date = holdings_actual_dt if holdings_actual_dt is not None else td
            title_badge = ""
            if holdings_source == "akshare_fallback":
                title_badge = (
                    f"<br><sup>📡 所选日期（{_cn(td)}）数据未发布，"
                    f"当前显示最新可用数据（{_cn(title_date)}，新浪财经）</sup>"
                )
            elif holdings_source in ("generic_cache", "date_cache"):
                title_badge = (
                    f"<br><sup>⚠️ 所选日期（{_cn(td)}）无持仓数据，"
                    f"当前显示最近可用数据（{_cn(title_date)}）</sup>"
                )
            elif holdings_source == "unavailable":
                title_badge = f"<br><sup>⚠️ 数据不可用，API 和缓存均无数据</sup>"

            fig_h.update_layout(
                title=(
                    f"{ct} 前{top_n}期货公司多空持仓（{_cn(title_date)}）"
                    + title_badge
                    + (f"<br><sup>📐 日变化对比：{_cn(prev_td)} → {_cn(title_date)}" if prev_td is not None else "")
                    + ("（API直接数据）</sup>" if (api_long_chg or api_short_chg) else "</sup>")
                ),
                barmode="group",
                bargap=0.30,
                bargroupgap=0.12,
                xaxis_title="持仓量（手）",
                template="plotly_white", height=max(500, top_n * 32),
                showlegend=False,
                margin=dict(l=170, r=200, t=80, b=30),
                annotations=annotations,
            )
            fig_h.update_xaxes(autorange=True)
            fig_h.update_yaxes(autorange="reversed")
            st.plotly_chart(fig_h, use_container_width=True)

            # ── 图例说明 ──
            st.caption(
                "🟢 **正指**（国泰君安、中粮期货、东证期货）—— 操作方向即市场方向 ｜ "
                "🔴 **反指**（东方财富、徽商期货、平安期货）—— 操作方向与市场相反"
            )

            # ── 正指/反指 动向结论 ──
            def _analyze_direction(name: str, lc: int, sc: int, is_zhengzhi: bool) -> str:
                """根据多空变化判断方向意图。
                正指和反指均直接描述实际持仓行为方向，不做反转。
                返回 (tag, name, action, intent, detail)"""
                parts = []
                if lc > 0: parts.append(f"加多+{lc:,}")
                elif lc < 0: parts.append(f"减多{lc:,}")
                if sc > 0: parts.append(f"加空+{sc:,}")
                elif sc < 0: parts.append(f"减空{sc:,}")
                action = "、".join(parts) if parts else "持仓不变"

                if lc > 0 and sc < 0:
                    intent = "看多"; detail = "加多减空"
                elif lc < 0 and sc > 0:
                    intent = "看空"; detail = "减多加空"
                elif lc > 0 and sc > 0:
                    intent = "多空分歧"; detail = "双向加仓"
                elif lc < 0 and sc < 0:
                    intent = "观望离场"; detail = "双向减仓"
                else:
                    intent = "中性"; detail = "持仓变化不大"

                tag = "🟢正指" if is_zhengzhi else "🔴反指"
                # 市场含义：正指方向=市场方向，反指方向=市场反向
                if is_zhengzhi:
                    market = "利多" if intent == "看多" else ("利空" if intent == "看空" else intent)
                else:
                    market = "利空" if intent == "看多" else ("利多" if intent == "看空" else intent)
                return f"{tag} {name}：{action} → {intent}（{market}）"

            zhengzhi_found = []
            fanzhi_found = []
            for _, row in disp.iterrows():
                co = row["company"]
                lc = int(row["long_change"])
                sc = int(row["short_change"])
                if co in ZHENGZHI_COMPANIES:
                    zhengzhi_found.append((co, lc, sc))
                elif co in FANZHI_COMPANIES:
                    fanzhi_found.append((co, lc, sc))

            if zhengzhi_found or fanzhi_found:
                conclusion_items = []
                zz_bull = zz_bear = fz_bull = fz_bear = 0
                zz_dir = fz_dir = ""

                # 逐行展示每个席位
                for co, lc, sc in zhengzhi_found:
                    line = _analyze_direction(co, lc, sc, True)
                    conclusion_items.append(f"• {line}")
                    if "看多" in line: zz_bull += 1
                    elif "看空" in line: zz_bear += 1
                for co, lc, sc in fanzhi_found:
                    line = _analyze_direction(co, lc, sc, False)
                    conclusion_items.append(f"• {line}")
                    if "看多" in line: fz_bull += 1
                    elif "看空" in line: fz_bear += 1

                # 方向汇总
                zz_dir = "看多" if zz_bull > zz_bear else ("看空" if zz_bear > zz_bull else "分歧")
                fz_dir = "看多" if fz_bull > fz_bear else ("看空" if fz_bear > fz_bull else "分歧")

                # 市场含义：正指方向=市场方向，反指方向相反
                zz_market = "利多" if zz_dir == "看多" else ("利空" if zz_dir == "看空" else "分歧")
                fz_market = "利空" if fz_dir == "看多" else ("利多" if fz_dir == "看空" else "分歧")

                # 综合评分：正指看多+1，反指看空+1（反指看空=市场利多）
                bull_score = zz_bull + fz_bear
                bear_score = zz_bear + fz_bull

                if bull_score > bear_score:
                    overall_sentiment = "bullish"
                    judgment = "🐂 偏多"
                elif bear_score > bull_score:
                    overall_sentiment = "bearish"
                    judgment = "🐻 偏空"
                else:
                    overall_sentiment = "neutral"
                    judgment = "⚖️ 方向分歧"

                conclusion_items.append(
                    f"• 综合判断：{judgment} — "
                    f"正指{zz_dir}（{zz_market}），反指{fz_dir}（{fz_market}）"
                )

                display_conclusion("🔍 关键席位动向分析", conclusion_items, overall_sentiment)

            # ── 明细表格（含日变化列） ──
            with st.expander("📋 持仓明细表"):
                tbl = disp.copy()
                tbl["多单"] = tbl["long"].apply(lambda x: f"{x:,}")
                tbl["空单"] = tbl["short"].apply(lambda x: f"{x:,}")
                tbl["净持仓"] = tbl["净持仓"].apply(lambda x: f"{x:+,}")
                tbl["多单变化"] = tbl["long_change"].apply(
                    lambda x: f"加多+{x:,}" if x > 0 else (f"减多{x:,}" if x < 0 else "—"))
                tbl["空单变化"] = tbl["short_change"].apply(
                    lambda x: f"加空+{x:,}" if x > 0 else (f"减空{x:,}" if x < 0 else "—"))
                col_order = ["company", "多单", "空单", "净持仓"]
                has_changes = (disp["long_change"] != 0).any() or (disp["short_change"] != 0).any()
                if has_changes:
                    col_order += ["多单变化", "空单变化"]
                st.dataframe(
                    tbl[col_order].rename(columns={"company": "期货公司"}),
                    use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ 持仓数据暂不可用，API 和本地缓存均无数据，请检查网络后重试。")


def _get_holdings(ct: str, target_date, return_meta: bool = False):
    """获取期货公司多空持仓（新浪财经 → akshare futures_hold_pos_sina，按日期区分）

    新浪财经接口分三张表返回：成交量排名、多单持仓排名、空单持仓排名。
    三张表按公司名合并后返回统一的 company/long/short/volume + 各变化量。

    当 return_meta=True 时返回 (DataFrame, actual_date_str, source_label)，
    actual_date_str 为 YYYYMMDD 格式，表示数据实际所属日期。
    source_label 表示数据来源：'akshare' / 'akshare_fallback' / 'date_cache' / 'generic_cache' / 'unavailable'
    """
    # 将 target_date 标准化为 YYYYMMDD 字符串
    if isinstance(target_date, pd.Timestamp):
        date_str = target_date.strftime("%Y%m%d")
    elif isinstance(target_date, datetime):
        date_str = target_date.strftime("%Y%m%d")
    elif hasattr(target_date, "strftime"):
        date_str = target_date.strftime("%Y%m%d")
    else:
        date_str = str(target_date).replace("-", "")[:8]

    date_cache_file = HOLDINGS_DIR / f"{ct}_{date_str}.csv"
    generic_cache_file = HOLDINGS_DIR / f"{ct}.csv"
    # 通用缓存的日期元数据文件（记录实际数据日期）
    generic_meta_file = HOLDINGS_DIR / f"{ct}_meta.txt"

    # ── 尝试从新浪财经获取（带重试 + 交易日回退） ──
    last_error = None
    _EMPTY_SENTINEL = "__EMPTY__"  # 哨兵：API 成功但数据未发布（0 rows）

    def _try_fetch_for_date(try_date_str: str):
        """尝试对指定日期拉取持仓数据。

        返回:
          (merged_df, "ok")  — 成功获取到数据
          (None, _EMPTY_SENTINEL) — API 调用成功但返回 0 行（数据尚未发布）
        异常:
          抛出原始异常 — 网络错误 / API 故障
        """
        import akshare as ak

        # 分三次调用，分别取 成交量 / 多单持仓 / 空单持仓
        df_vol = ak.futures_hold_pos_sina(symbol="成交量", contract=ct, date=try_date_str)
        df_long = ak.futures_hold_pos_sina(symbol="多单持仓", contract=ct, date=try_date_str)
        df_short = ak.futures_hold_pos_sina(symbol="空单持仓", contract=ct, date=try_date_str)

        # 标准化列名
        # 新浪返回列: 名次, 会员简称, {成交量|多单持仓|空单持仓}, 比上交易增减
        def _norm_sina(df: pd.DataFrame, val_col: str, chg_col: str) -> pd.DataFrame:
            """将新浪表标准化为 company / {val_col} / {chg_col}"""
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

        vol_df = _norm_sina(df_vol, "volume", "volume_chg")
        long_df = _norm_sina(df_long, "long", "long_chg")
        short_df = _norm_sina(df_short, "short", "short_chg")

        # 三表按 company 做 outer join，缺失填 0
        merged = long_df.merge(short_df, on="company", how="outer")
        merged = merged.merge(vol_df, on="company", how="outer")
        merged = merged.fillna(0)
        for col in ["long", "long_chg", "short", "short_chg", "volume", "volume_chg"]:
            if col in merged.columns:
                merged[col] = merged[col].astype(int)

        # ★ 区分：API 成功但数据未发布（0 rows） vs 真正的 API 异常
        if merged.empty:
            return None, _EMPTY_SENTINEL

        # 按多单持仓降序排列（主力多单排名）
        merged = merged.sort_values("long", ascending=False).reset_index(drop=True)
        return merged, "ok"

    # ★ 构建回退日期列表：只包含交易日（周一至周五），跳过周末
    try_dates = [date_str]
    try:
        base_dt = datetime.strptime(date_str, "%Y%m%d")
        offset = 1
        while len(try_dates) < 6:  # 最多 6 个候选日期
            fallback_dt = base_dt - timedelta(days=offset)
            offset += 1
            # 跳过周末（周六=5, 周日=6）
            if fallback_dt.weekday() >= 5:
                continue
            fallback_str = fallback_dt.strftime("%Y%m%d")
            if fallback_str not in try_dates:
                try_dates.append(fallback_str)
    except Exception:
        pass

    for try_date in try_dates:
        for attempt in range(3):  # 每个日期最多重试 3 次
            try:
                merged, status = _try_fetch_for_date(try_date)
                if status == _EMPTY_SENTINEL:
                    # ★ API 正常但数据未发布 → 不重试，直接尝试上一个日期
                    last_error = f"{try_date}: 数据尚未发布（0 rows）"
                    break  # 跳出重试循环，进入下一个日期
                # 成功：写入缓存
                merged.to_csv(HOLDINGS_DIR / f"{ct}_{try_date}.csv", index=False)
                merged.to_csv(generic_cache_file, index=False)
                generic_meta_file.write_text(try_date)
                source = "akshare" if try_date == date_str else "akshare_fallback"
                if return_meta:
                    return merged.head(20), try_date, source
                return merged.head(20)
            except Exception as e:
                last_error = f"{try_date}: {type(e).__name__}: {str(e)[:120]}"
                if attempt < 2:
                    time.sleep(1)  # 重试前等待 1 秒

    # ★ 所有 API 尝试都失败了，记录最后的错误（供调试）
    if last_error:
        st.session_state.setdefault("_holdings_api_errors", {})[f"{ct}_{date_str}"] = last_error

    # ── 日期缓存兜底 ──
    if date_cache_file.exists():
        try:
            df = pd.read_csv(date_cache_file)
            if "company" in df.columns and "long" in df.columns and "short" in df.columns:
                result = df.sort_values("long", ascending=False).head(20).reset_index(drop=True)
                if return_meta:
                    return result, date_str, "date_cache"
                return result
        except Exception:
            pass

    # ── 通用缓存兜底 ──
    if generic_cache_file.exists():
        try:
            df = pd.read_csv(generic_cache_file)
            if "company" in df.columns and "long" in df.columns and "short" in df.columns:
                # 尝试读取通用缓存的实际日期
                generic_date = date_str  # 默认用请求日期
                if generic_meta_file.exists():
                    try:
                        generic_date = generic_meta_file.read_text().strip()[:8]
                    except Exception:
                        pass
                result = df.sort_values("long", ascending=False).head(20).reset_index(drop=True)
                if return_meta:
                    return result, generic_date, "generic_cache"
                return result
        except Exception:
            pass

    # ── 所有数据源均失败 ──
    if return_meta:
        return None, date_str, "unavailable"
    return None

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

        # ★ 找出该月份所有合约 —— 不跳过任何往年同期合约
        # 使用 load_futures 逐一加载（已有 CSV 秒读，缺失的自动下载）
        same_month_cts = [c for c in ALL_CONTRACTS if ct_month(c) == sel_month]
        available_cts = []
        skipped_cts = []
        for c in same_month_cts:
            df, _ = load_futures(c)
            if df is not None and not df.empty:
                available_cts.append(c)
            else:
                skipped_cts.append(c)

        if not available_cts:
            st.warning(f"⚠️ 暂无可用的 {sel_month} 月合约数据"); return

        st.caption(f"✅ 已加载 {len(available_cts)} 个 {sel_month} 月合约：{'、'.join(available_cts)}"
                   + (f"  ｜ ⚠️ 跳过：{'、'.join(skipped_cts)}" if skipped_cts else ""))

        # 日期范围（轻量：只读 CSV 首尾行）
        all_dates = []
        for c in available_cts:
            cp = _csv_path(c)
            if not cp.exists(): continue
            try:
                # 只读第一行和最后一行，不加载整个文件
                df = pd.read_csv(cp, usecols=["date"])
                if df.empty: continue
                dates = pd.to_datetime(df["date"])
                all_dates.append(dates.min().date())
                all_dates.append(dates.max().date())
            except Exception:
                continue
        if all_dates:
            data_min, data_max = min(all_dates), max(all_dates)
        else:
            data_min = datetime.now().date() - timedelta(days=365)
            data_max = datetime.now().date()

        date_range = st.date_input("📅 日期范围", value=(data_min, data_max),
                                   min_value=data_min, max_value=data_max, key="t6_date_range")
        if isinstance(date_range, tuple) and len(date_range) == 2:
            sd, ed = date_range
        else:
            sd, ed = data_min, data_max

    with col_chart:
        if len(available_cts) < 1:
            st.warning(f"⚠️ 暂无可用的 {sel_month} 月合约"); return

        # ── 收集各合约的成交量和持仓量数据（缓存加速）──
        vol_data, oi_data, net_data_full = _build_vol_oi_seasonal_cached(
            tuple(available_cts), str(sd), str(ed), 0  # spot_hash_key=0, Tab6 不依赖现货数据
        )

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
            # ★ 根据数据范围调整纵轴
            vol_all_vals = []
            for vdf in vol_data.values():
                if not vdf.empty and "volume" in vdf.columns:
                    vol_all_vals.extend(vdf["volume"].tolist())
            vol_yrange = _compute_y_padding(vol_all_vals) if vol_all_vals else None
            if vol_yrange:
                fig_vol_s.update_yaxes(range=vol_yrange)
            else:
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
            # ★ 根据数据范围调整纵轴
            oi_all_vals = []
            for odf in oi_data.values():
                if not odf.empty and "open_interest" in odf.columns:
                    oi_all_vals.extend(odf["open_interest"].tolist())
            oi_yrange = _compute_y_padding(oi_all_vals) if oi_all_vals else None
            if oi_yrange:
                fig_oi_s.update_yaxes(range=oi_yrange)
            else:
                fig_oi_s.update_yaxes(autorange=True)
            st.plotly_chart(fig_oi_s, use_container_width=True)
        else:
            st.warning("⚠️ 无持仓量数据")

        # ── 图3：前20净持仓季节性对比 ──
        # net_data_full 已由 _build_vol_oi_seasonal_cached 缓存返回，只需按日期范围筛选
        st.markdown("#### 🏢 前20净持仓季节性对比")

        net_data: Dict[str, pd.DataFrame] = {}
        if net_data_full:
            ref_start = pd.Timestamp(year=2020, month=sd.month, day=sd.day)
            ref_end = pd.Timestamp(year=2020, month=ed.month, day=ed.day)
            for label, ndf in net_data_full.items():
                if ndf.empty: continue
                mask = (ndf["plot_date"] >= ref_start) & (ndf["plot_date"] <= ref_end)
                fdf = ndf[mask]
                if not fdf.empty:
                    net_data[label] = fdf

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
            # ★ 根据数据范围调整纵轴（净持仓数据可能正负跨度大）
            net_all_vals = []
            for ndf in net_data.values():
                if not ndf.empty and "net_position" in ndf.columns:
                    net_all_vals.extend(ndf["net_position"].tolist())
            net_yrange = _compute_y_padding(net_all_vals) if net_all_vals else None
            if net_yrange:
                fig_net.update_yaxes(range=net_yrange)
            else:
                fig_net.update_yaxes(autorange=True)
            st.plotly_chart(fig_net, use_container_width=True)
        else:
            st.warning("⚠️ 前20净持仓数据暂不可用，请先同步数据后再查看")

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


def _fetch_exact_holdings(ct: str, date_str: str, use_fallback: bool = True) -> Optional[pd.DataFrame]:
    """拉取指定合约在指定日期的真实持仓数据（支持交易日回退）

    成功返回合并后的 DataFrame，失败返回 None。
    结果缓存到日期精确文件。

    use_fallback=True 时：若目标日期数据未发布，自动向前查找最近 10 个交易日。
    """
    # ── 辅助：标准化新浪表 ──
    def _norm_sina(df, val_col, chg_col):
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

    # ── 构建回退日期列表 ──
    try_dates = [date_str]
    if use_fallback:
        try:
            base_dt = datetime.strptime(date_str, "%Y%m%d")
            offset = 1
            while len(try_dates) < 11:  # 当前日期 + 最多往前 10 个交易日
                fallback_dt = base_dt - timedelta(days=offset)
                offset += 1
                if fallback_dt.weekday() >= 5:  # 跳过周末
                    continue
                try_dates.append(fallback_dt.strftime("%Y%m%d"))
        except Exception:
            pass

    # ── 逐个尝试日期 ──
    for try_date in try_dates:
        cache_file = HOLDINGS_DIR / f"{ct}_{try_date}.csv"
        generic_cache_file = HOLDINGS_DIR / f"{ct}.csv"
        generic_meta_file = HOLDINGS_DIR / f"{ct}_meta.txt"

        # 先查缓存
        if cache_file.exists():
            try:
                df = pd.read_csv(cache_file)
                if "company" in df.columns and "long" in df.columns and "short" in df.columns:
                    return df
            except Exception:
                pass

        # 从新浪 API 拉取（每个日期最多 2 次重试）
        for attempt in range(2):
            try:
                import akshare as ak

                df_vol = ak.futures_hold_pos_sina(symbol="成交量", contract=ct, date=try_date)
                df_long = ak.futures_hold_pos_sina(symbol="多单持仓", contract=ct, date=try_date)
                df_short = ak.futures_hold_pos_sina(symbol="空单持仓", contract=ct, date=try_date)

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
                    break  # 数据未发布，跳到上一个日期

                merged = merged.sort_values("long", ascending=False).reset_index(drop=True)

                # 缓存（以请求的 try_date 为 key）
                merged.to_csv(cache_file, index=False)
                merged.to_csv(generic_cache_file, index=False)
                generic_meta_file.write_text(try_date)
                return merged

            except Exception:
                if attempt < 1:
                    time.sleep(0.5)  # 重试前短暂等待
        # 如果这个日期数据未发布（empty），继续尝试上一个日期
        # 如果是网络错误，也继续尝试上一个日期

    return None


# ══════════════════════════════════════════════════════════════
# 前20净持仓 聚合缓存 — 与期货价格数据完全一致的缓存逻辑
# ══════════════════════════════════════════════════════════════

def _net_agg_path(ct: str) -> Path:
    """每个合约一个聚合文件：date, net_position"""
    return HOLDINGS_DIR / f"{ct}_net_agg.csv"


def _load_aggregated_net(ct: str) -> Optional[pd.DataFrame]:
    """读取合约的聚合净持仓缓存文件。不存在返回 None。"""
    p = _net_agg_path(ct)
    if not p.exists():
        return None
    try:
        df = pd.read_csv(p)
        if "date" not in df.columns or "net_position" not in df.columns:
            return None
        df["date"] = pd.to_datetime(df["date"])
        return df.sort_values("date").reset_index(drop=True)
    except Exception:
        return None


def _migrate_existing_to_aggregated(ct: str) -> int:
    """一次性迁移：扫描所有 date-specific CSV，计算 net，写入聚合文件。
    返回迁移的日期数。已存在聚合文件则跳过。"""
    agg_path = _net_agg_path(ct)
    if agg_path.exists():
        return 0

    rows = []
    for f in HOLDINGS_DIR.glob(f"{ct}_????????.csv"):
        date_str = f.stem[len(ct)+1:]
        if len(date_str) != 8:
            continue
        try:
            df = pd.read_csv(f)
            if "company" in df.columns and "long" in df.columns and "short" in df.columns:
                net = int(df["long"].sum() - df["short"].sum())
                rows.append({"date": date_str, "net_position": net})
        except Exception:
            continue

    if rows:
        agg_df = pd.DataFrame(rows)
        agg_df["date"] = pd.to_datetime(agg_df["date"])
        agg_df = agg_df.sort_values("date").reset_index(drop=True)
        agg_df.to_csv(agg_path, index=False)
        return len(rows)
    return 0


def _download_net_single(ct: str, date_str: str) -> Optional[int]:
    """拉取单日净持仓：前20多单合计 - 前20空单合计。失败返回 None。"""
    h = _fetch_exact_holdings(ct, date_str, use_fallback=False)
    if h is not None and not h.empty:
        return int(h["long"].sum() - h["short"].sum())
    return None


def sync_net_holdings(ct: str, force_full: bool = False) -> Tuple[bool, str]:
    """同步单个合约的净持仓数据（与 sync_futures 逻辑完全一致）。

    - force_full=True: 获取该合约全部交易日列表，逐日拉取持仓数据
    - force_full=False: 只拉取最新缺失的交易日
    返回 (成功与否, 状态信息)
    """
    cp = _csv_path(ct)
    if not cp.exists():
        return False, "❌ 无期货 CSV"

    try:
        fut_df = pd.read_csv(cp, usecols=["date"])
        if fut_df.empty:
            return False, "❌ 期货数据为空"
        fut_df["date"] = pd.to_datetime(fut_df["date"])
    except Exception as e:
        return False, f"❌ 读取期货CSV失败：{e}"

    agg_path = _net_agg_path(ct)
    today = datetime.now()

    # ── 增量更新（本地有聚合文件 且 不强制全量）──
    if agg_path.exists() and not force_full:
        try:
            agg_df = pd.read_csv(agg_path)
            agg_df["date"] = pd.to_datetime(agg_df["date"])
            latest_cached = agg_df["date"].max()
            latest_futures = fut_df["date"].max()

            # ★ 如果缓存日期 < 期货最新日期，拉取所有缺失日期
            if latest_cached >= latest_futures:
                return True, "📁 已是最新"

            pending = [d for d in sorted(fut_df["date"].unique()) if d > latest_cached]
            if not pending:
                return True, "📁 已是最新"

            fetched = 0
            for dt in pending:
                ds = dt.strftime("%Y%m%d")
                net = _download_net_single(ct, ds)
                if net is not None:
                    new_row = pd.DataFrame([{"date": pd.to_datetime(ds), "net_position": net}])
                    agg_df = agg_df[agg_df["date"] != pd.to_datetime(ds)]
                    agg_df = pd.concat([agg_df, new_row], ignore_index=True)
                    fetched += 1
                # 如果API返回None（数据未发布），不中断，继续尝试后续日期

            if fetched > 0:
                agg_df = agg_df.sort_values("date").reset_index(drop=True)
                agg_df.to_csv(agg_path, index=False)
                _build_seasonal_net_positions.clear()
            return True, f"🔄 增量 +{fetched}条" if fetched else "📁 暂无新数据可拉取"
        except Exception as e:
            return False, f"❌ 增量失败：{e}"

    # ── 全量下载 ──
    try:
        # 尝试迁移旧 CSV
        _migrate_existing_to_aggregated(ct)
        agg_df = _load_aggregated_net(ct)
        cached_set = set()
        if agg_df is not None and not agg_df.empty:
            cached_set = set(agg_df["date"].dt.strftime("%Y%m%d"))

        all_dates = sorted(fut_df["date"].unique())
        pending = [d.strftime("%Y%m%d") for d in all_dates
                   if d.strftime("%Y%m%d") not in cached_set]

        if not pending:
            return True, "📁 已是最新"

        fetched = 0
        # 全量模式：并行下载（5 线程）
        def _fetch_one(ds):
            net = _download_net_single(ct, ds)
            return ds, net

        with ThreadPoolExecutor(max_workers=5) as executor:
            future_map = {executor.submit(_fetch_one, ds): ds for ds in pending}
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
            # 清除缓存
            _build_seasonal_net_positions.clear()
        return True, f"🌐 全量 {fetched}条/{len(pending)}天"
    except Exception as e:
        return False, f"❌ 下载失败：{e}"


def sync_all_net_holdings(max_workers: int = 4, progress_callback=None) -> Dict[str, str]:
    """预下载 ALL_CONTRACTS 中所有合约的净持仓完整历史数据（并行）。
    与 sync_all_contracts 逻辑完全一致。

    max_workers: 并行下载线程数（净持仓 API 较慢，建议 3-4）
    progress_callback: 可选，签名 (current, total, contract, status) -> None
    返回 {合约: 状态信息}
    """
    missing = [c for c in ALL_CONTRACTS if not _net_agg_path(c).exists()
               or _load_aggregated_net(c) is None]

    if not missing:
        return {}

    results = {}
    total = len(missing)

    def _download_one(ct):
        ok, msg = sync_net_holdings(ct, force_full=True)
        return ct, ok, msg

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_map = {executor.submit(_download_one, ct): ct for ct in missing}
        completed = 0
        for future in as_completed(future_map):
            ct, ok, msg = future.result()
            results[ct] = msg
            completed += 1
            if progress_callback:
                progress_callback(completed, total, ct, msg)

    return results


def _ensure_net_cache(ct: str) -> Optional[pd.DataFrame]:
    """读取净持仓缓存，自动增量补全缺失日期。

    - 本地有聚合文件 → 增量检查并补全缺失交易日
    - 本地无聚合文件 → 全量下载
    """
    agg_df = _load_aggregated_net(ct)
    if agg_df is not None and not agg_df.empty:
        # ★ 增量检查：补全最新缺失的交易日
        sync_net_holdings(ct, force_full=False)
        return _load_aggregated_net(ct)

    # ── 尝试迁移旧 CSV ──
    _migrate_existing_to_aggregated(ct)
    agg_df = _load_aggregated_net(ct)
    if agg_df is not None and not agg_df.empty:
        sync_net_holdings(ct, force_full=False)
        return _load_aggregated_net(ct)

    # ── 无缓存，全量下载 ──
    sync_net_holdings(ct, force_full=True)
    return _load_aggregated_net(ct)


@st.cache_data(ttl=3600, show_spinner=False)
def _build_seasonal_net_positions(contracts: Tuple[str, ...], sel_month: str) -> Tuple[Dict[str, pd.DataFrame], defaultdict]:
    """构建季节性净持仓数据 — 和其他板块一样：有本地文件直接读，没有就当场下载。"""
    net_data: Dict[str, pd.DataFrame] = {}
    net_collector = defaultdict(list)

    for c in contracts:
        # 与 load_futures 完全一致：有本地文件秒读，缺失则当场下载
        agg_df = _ensure_net_cache(c)
        if agg_df is None or agg_df.empty:
            continue

        agg_df["plot_date"] = [pd.Timestamp(year=2020, month=d.month, day=d.day)
                               for d in agg_df["date"]]
        agg_df["trade_year"] = agg_df["date"].dt.year

        for trade_yr, grp in agg_df.groupby("trade_year"):
            grp = grp.sort_values("plot_date")
            ty_str = str(trade_yr)
            label = f"{c[2:]} ({ty_str})"
            net_data[label] = pd.DataFrame({
                "plot_date": grp["plot_date"].values,
                "net_position": grp["net_position"].values,
            }).sort_values("plot_date")
            for _, row in grp.iterrows():
                md = (row["date"].month, row["date"].day)
                v = row["net_position"]
                if v is not None and not pd.isna(v):
                    net_collector[md].append(int(v))

    if net_collector:
        avg_rows = [{"plot_date": pd.Timestamp(year=2020, month=m, day=d),
                     "net_position": int(np.mean(v))}
                    for (m, d), v in sorted(net_collector.items()) if v]
        if avg_rows:
            net_data["历史均值"] = pd.DataFrame(avg_rows).sort_values("plot_date")

    return net_data, net_collector


# ══════════════════════════════════════════════════════════════
# 性能优化：缓存批量计算
# ══════════════════════════════════════════════════════════════

def _spot_hash(spot_dict: dict) -> int:
    """生成 spot_dict 的轻量哈希，用于缓存 key"""
    items = []
    for k in sorted(spot_dict.keys()):
        df = spot_dict[k]
        items.append((k, len(df), float(df["price"].sum()) if not df.empty else 0.0))
    return hash(tuple(items))


@st.cache_data(ttl=1800, show_spinner=False)
def _build_calendar_series_cached(
    avail_tuple: Tuple[str, ...],
    sel_items_tuple: Tuple[str, ...],
    spot_hash_key: int,
    ref_regions_tuple: Tuple[str, ...],
) -> Dict[str, pd.DataFrame]:
    """缓存 Tab3 同比模式的 series 构建（最耗时的部分）。
    输入全部转为可哈希类型，spot_dict 通过 spot_hash_key 标识版本。"""
    # 重新加载 spot_dict（从缓存 key 匹配确保一致性）
    spot_dict, _ = load_spot(str(SPOT_PATH))
    avail = list(avail_tuple)
    sel_items = list(sel_items_tuple)
    ref_regions = list(ref_regions_tuple)

    series: Dict[str, pd.DataFrame] = {}
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
            df_basis["plot_date"] = df_basis.apply(
                lambda r: _doy_to_date(int(r["doy"]), int(r["year"])), axis=1)
            for yr, grp in df_basis.groupby("year"):
                grp = grp.sort_values("doy").copy()
                label = _make_trace_label(c, yr, item_short)
                series[label] = grp
                for _, row in grp.iterrows():
                    md_collectors[item_short][(row["date"].month, row["date"].day)].append(row["basis"])

    # 历史均值线
    for item_short, mdc in md_collectors.items():
        avg_rows = [{"doy": m*100+d, "basis": int(round(np.mean(v))),
                      "plot_date": pd.Timestamp(year=2020, month=m, day=d)}
                    for (m, d), v in sorted(mdc.items()) if v]
        if avg_rows:
            series[f"历史均值-{item_short}"] = pd.DataFrame(avg_rows).sort_values("doy")

    return series


@st.cache_data(ttl=1800, show_spinner=False)
def _build_vol_oi_seasonal_cached(
    contracts_tuple: Tuple[str, ...],
    sd_str: str,
    ed_str: str,
    spot_hash_key: int,
) -> Tuple[Dict[str, pd.DataFrame], Dict[str, pd.DataFrame], Dict[str, pd.DataFrame], Dict[str, pd.DataFrame]]:
    """缓存 Tab6 的成交量/持仓量季节性数据构建。
    返回 (vol_data, oi_data, net_data, all_available_cts_data)"""
    contracts = list(contracts_tuple)
    sd = pd.to_datetime(sd_str)
    ed = pd.to_datetime(ed_str)
    sel_month = ct_month(contracts[0]) if contracts else "09"

    vol_data: Dict[str, pd.DataFrame] = {}
    oi_data: Dict[str, pd.DataFrame] = {}
    vol_collector = defaultdict(list)
    oi_collector = defaultdict(list)

    for c in contracts:
        df, _ = load_futures(c)
        if df is None or df.empty: continue
        df = df.sort_values("date").copy()
        df = df[(df["date"] >= sd) & (df["date"] <= ed)]
        if df.empty: continue
        df["plot_date"] = [pd.Timestamp(year=2020, month=d.month, day=d.day) for d in df["date"]]
        df["trade_year"] = df["date"].dt.year

        for trade_yr, grp in df.groupby("trade_year"):
            grp = grp.sort_values("date")
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

    # 历史均值
    avg_vol = pd.DataFrame()
    avg_oi = pd.DataFrame()
    if vol_collector:
        avg_vol_rows = [{"plot_date": pd.Timestamp(year=2020, month=m, day=d),
                         "volume": int(np.mean(v))}
                        for (m, d), v in sorted(vol_collector.items()) if v]
        if avg_vol_rows:
            avg_vol = pd.DataFrame(avg_vol_rows).sort_values("plot_date")
            vol_data["历史均值"] = avg_vol
    if oi_collector:
        avg_oi_rows = [{"plot_date": pd.Timestamp(year=2020, month=m, day=d),
                        "open_interest": int(np.mean(v))}
                       for (m, d), v in sorted(oi_collector.items()) if v]
        if avg_oi_rows:
            avg_oi = pd.DataFrame(avg_oi_rows).sort_values("plot_date")
            oi_data["历史均值"] = avg_oi

    # 净持仓数据
    net_data_full, _ = _build_seasonal_net_positions(tuple(contracts), sel_month)

    return vol_data, oi_data, net_data_full


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
        df, tech_warnings = calculate_technicals(df)
        if tech_warnings:
            for w in tech_warnings:
                st.warning(f"⚠️ {w}")

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
# Tab 1：每日期货分析日报
# ══════════════════════════════════════════════════════════════

def _analyze_basis_historical(main_ct, spot_dict, fut_df, ltd, regions, snap) -> dict:
    """基差历史同期对比 — 直接复用 Tab4 的 _build_calendar_series_cached 数据管道。
    从同一套缓存的series中提取同月同日值，确保100%一致。"""
    result = {}
    if fut_df is None or fut_df.empty:
        return result

    target_m, target_d = ltd.month, ltd.day
    tmon = ct_month(main_ct)
    same_month = [c for c in ALL_CONTRACTS if ct_month(c) == tmon]
    avail = [c for c in same_month if load_futures(c)[0] is not None]
    if not avail:
        return result

    # ★ 用 Tab4 完全相同的缓存函数生成所有series
    sel_items_list = ["📊 全国均价基差", "🔴 最大基差", "🟢 最小基差", "🟣 基差平均值"]
    shash = _spot_hash(spot_dict)
    all_series = _build_calendar_series_cached(
        tuple(avail), tuple(sel_items_list), shash, tuple(regions)
    )

    indicators = {
        "全国均价基差": ("全国均价", snap.get("national_avg", 0)),
        "最大基差": ("最大", snap.get("max_basis", 0)),
        "最小基差": ("最小", snap.get("min_basis", 0)),
        "基差均值": ("均值", snap.get("avg_basis", 0)),
    }

    for ind_name, (item_short, cur) in indicators.items():
        # 从series中收集所有同月同日的值（与Tab4 md_collector完全一致）
        hist_vals = []
        for label, sdf in all_series.items():
            if "历史均值" in label or sdf.empty:
                continue
            if item_short not in label:
                continue
            # sdf有plot_date列(已映射到2020年)，直接用target_m/target_d筛选
            row = sdf[(sdf["plot_date"].dt.month == target_m) & (sdf["plot_date"].dt.day == target_d)]
            if not row.empty:
                hist_vals.append(int(row["basis"].iloc[-1]))

        if hist_vals:
            hist_avg = int(np.mean(hist_vals))
            # ★ 直接用偏离幅度判断（更直观）
            if hist_avg != 0:
                deviation = (cur - hist_avg) / abs(hist_avg) * 100
            else:
                deviation = 0
            if deviation > 30:
                level = f"明显高于历史同期均值（偏离+{deviation:.0f}%）"
            elif deviation > 10:
                level = f"高于历史同期均值（偏离+{deviation:.0f}%）"
            elif deviation < -30:
                level = f"明显低于历史同期均值（偏离{deviation:.0f}%）"
            elif deviation < -10:
                level = f"低于历史同期均值（偏离{deviation:.0f}%）"
            else:
                level = f"处于历史同期均值附近（偏离{deviation:+.0f}%）"
            result[ind_name] = {"current": cur, "hist_avg": hist_avg, "level": level}
        else:
            result[ind_name] = {"current": cur, "hist_avg": None, "level": "暂无历史同期数据"}

    return result


def _build_basis_analysis_html(snap, basis_enhanced, na_basis, max_region, max_basis, min_region, min_basis):
    """构建基差分析板块HTML，含历史分位"""
    if not basis_enhanced:
        return f"""全国均价基差 <b>{na_basis:+,}元/吨</b>。<br>
最大基差：<b>{max_region}</b>（{max_basis:+,}元/吨）｜ 最小基差：<b>{min_region}</b>（{min_basis:+,}元/吨）。"""

    be = basis_enhanced
    parts = []

    # 全国均价基差
    na = be.get("全国均价基差", {})
    if na and na.get("hist_avg") is not None:
        na_cur = na.get("current", 0)
        na_hist = na.get("hist_avg", 0)
        na_level = na.get("level", "")
        parts.append(f"全国均价基差 <b>{na_cur:+,}元/吨</b>，"
                   f"历史同期均值 {na_hist:+,}元/吨，{na_level}")

    # 最大基差
    mx = be.get("最大基差", {})
    if mx and mx.get("hist_avg") is not None:
        mx_region = snap.get("max_region", "") if snap else ""
        mx_cur = mx.get("current", 0)
        mx_hist = mx.get("hist_avg", 0)
        mx_level = mx.get("level", "")
        parts.append(f"最大基差（{mx_region}）<b>{mx_cur:+,}元/吨</b>，"
                   f"历史同期均值 {mx_hist:+,}元/吨，{mx_level}")

    # 最小基差
    mn = be.get("最小基差", {})
    if mn and mn.get("hist_avg") is not None:
        mn_region = snap.get("min_region", "") if snap else ""
        mn_cur = mn.get("current", 0)
        mn_hist = mn.get("hist_avg", 0)
        mn_level = mn.get("level", "")
        parts.append(f"最小基差（{mn_region}）<b>{mn_cur:+,}元/吨</b>，"
                   f"历史同期均值 {mn_hist:+,}元/吨，{mn_level}")

    # 基差均值
    avg = be.get("基差均值", {})
    if avg and avg.get("hist_avg") is not None:
        avg_cur = avg.get("current", 0)
        avg_hist = avg.get("hist_avg", 0)
        avg_level = avg.get("level", "")
        parts.append(f"基差均值 <b>{avg_cur:+,}元/吨</b>，"
                   f"历史同期均值 {avg_hist:+,}元/吨，{avg_level}")

    return "<br>".join(f"• {p}" for p in parts) if parts else "基差数据暂不可用"


def _analyze_vol_oi_momentum(fut_df, ltd) -> dict:
    """分析成交量和持仓量的动能情况，含历史同期趋势复盘"""
    if fut_df is None or fut_df.empty:
        return {"available": False, "vol_text": "数据不足", "oi_text": "数据不足",
                "oi_decline": "数据不足", "vol_trend": "数据不足",
                "oi_hist_review": "数据不足"}
    try:
        df = fut_df.sort_values("date").reset_index(drop=True)
        close = df["close"].astype(float)
        vol = df["volume"].astype(float)
        oi_col = "open_interest" if "open_interest" in df.columns else ("hold" if "hold" in df.columns else None)
        if oi_col is None:
            return {"available": False, "vol_text": "数据不足", "oi_text": "数据不足",
                    "oi_decline": "数据不足", "vol_trend": "数据不足",
                    "oi_hist_review": "数据不足"}
        oi = df[oi_col].astype(float)

        cur_vol = float(vol.iloc[-1])
        cur_oi = float(oi.iloc[-1])
        cur_close = float(close.iloc[-1])

        # ── 成交量 ──
        vol_5 = float(vol.tail(5).mean())
        vol_20 = float(vol.tail(20).mean()) if len(vol) >= 20 else vol_5
        if vol_5 > vol_20 * 1.2:
            vol_trend = "放量（5日均量>20日均量20%）"
        elif vol_5 < vol_20 * 0.8:
            vol_trend = "缩量（5日均量<20日均量20%）"
        else:
            vol_trend = "量能平稳"
        close_5_ago = float(close.iloc[-6]) if len(close) >= 6 else cur_close
        price_dir = "上涨" if cur_close > close_5_ago else ("下跌" if cur_close < close_5_ago else "持平")
        vol_text = f"成交量 {cur_vol:,.0f}手，{vol_trend}，近5日价格{price_dir}"

        # ── 持仓量当前 ──
        oi_5_ago = float(oi.iloc[-6]) if len(oi) >= 6 else cur_oi
        oi_20_ago = float(oi.iloc[-21]) if len(oi) >= 21 else oi_5_ago
        oi_chg_5 = (cur_oi - oi_5_ago) / oi_5_ago * 100 if oi_5_ago > 0 else 0
        oi_chg_20 = (cur_oi - oi_20_ago) / oi_20_ago * 100 if oi_20_ago > 0 else 0
        if oi_chg_5 > 3: oi_trend = "增仓明显"
        elif oi_chg_5 < -3: oi_trend = "减仓明显"
        else: oi_trend = "持仓平稳"
        oi_text = f"持仓量 {cur_oi:,.0f}手，近5日{oi_trend}（{oi_chg_5:+.1f}%），近20日{oi_chg_20:+.1f}%"

        # ── 持仓峰值回顾 ──
        oi_peak_idx = int(oi.idxmax())
        oi_peak_date = df["date"].iloc[oi_peak_idx]
        oi_peak_val = float(oi.iloc[oi_peak_idx])
        days_since_peak = (pd.Timestamp(ltd) - oi_peak_date).days
        if days_since_peak > 5 and cur_oi < oi_peak_val * 0.95:
            oi_decline = f"持仓量自{_cn(oi_peak_date)}见顶（{oi_peak_val:,.0f}手）后下滑，回落{(1-cur_oi/oi_peak_val)*100:.1f}%（{days_since_peak}天）"
        elif days_since_peak <= 5:
            oi_decline = f"持仓量处于近期高位（峰值{_cn(oi_peak_date)}，{oi_peak_val:,.0f}手），尚无下滑"
        else:
            oi_decline = f"持仓量较峰值{_cn(oi_peak_date)}（{oi_peak_val:,.0f}手）回落{(1-cur_oi/oi_peak_val)*100:.1f}%"

        return {"available": True, "vol_text": vol_text, "oi_text": oi_text,
                "oi_decline": oi_decline, "vol_trend": vol_trend,
                "cur_vol": cur_vol, "cur_oi": cur_oi}
    except Exception:
        return {"available": False, "vol_text": "计算异常", "oi_text": "计算异常",
                "oi_decline": "计算异常", "vol_trend": "计算异常"}


def _build_daily_report_html(main_ct: str, fut_df, spot_dict, ltd, prev_td,
                              snap, holdings_analysis, key_spread_info,
                              trend_direction, sr_lines_info, chart_images=None,
                              basis_enhanced=None, vol_oi_analysis=None,
                              spread_date_str=None) -> str:
    """构建日报 HTML 内容"""
    cn_date = _cn(ltd)
    cn_prev = _cn(prev_td) if prev_td else "前一交易日"

    # 当日行情
    if fut_df is not None and not fut_df.empty:
        row_today = fut_df[fut_df["date"] == ltd]
        row_prev = fut_df[fut_df["date"] == prev_td] if prev_td else None
        row_t = row_today.iloc[-1] if not row_today.empty else None
        if row_t is not None:
            close_v = float(row_t["close"])
            high_v = float(row_t["high"]) if pd.notna(row_t.get("high")) else close_v
            low_v = float(row_t["low"]) if pd.notna(row_t.get("low")) else close_v
            vol_v = int(row_t["volume"]) if pd.notna(row_t.get("volume")) else 0
            oi_col = "open_interest" if "open_interest" in row_t.index else ("hold" if "hold" in row_t.index else None)
            oi_v = int(row_t[oi_col]) if oi_col and pd.notna(row_t.get(oi_col)) else 0
            prev_close = float(row_prev["close"].iloc[-1]) if row_prev is not None and not row_prev.empty else close_v
            chg = close_v - prev_close
            chg_pct = chg / prev_close * 100 if prev_close != 0 else 0
            chg_sign = "+" if chg >= 0 else ""
        else:
            close_v, high_v, low_v, vol_v, oi_v, chg, chg_pct, chg_sign = 0, 0, 0, 0, 0, 0, 0, ""
    else:
        close_v, high_v, low_v, vol_v, oi_v, chg, chg_pct, chg_sign = 0, 0, 0, 0, 0, 0, 0, ""

    chg_color = "#E74C3C" if chg >= 0 else "#27AE60"
    chg_class = "up" if chg >= 0 else "down"

    # 基差
    na_basis = snap.get("national_avg", 0) if snap else 0
    max_region = snap.get("max_region", "—") if snap else "—"
    max_basis = snap.get("max_basis", 0) if snap else 0
    min_region = snap.get("min_region", "—") if snap else "—"
    min_basis = snap.get("min_basis", 0) if snap else 0

    if na_basis > 500: basis_judge = "偏高"
    elif na_basis < -500: basis_judge = "偏低"
    else: basis_judge = "中性"

    # 价差
    spread_text = key_spread_info if key_spread_info else "暂无价差数据"

    # 持仓（使用分析结果，明确标注数据日期）
    ha = holdings_analysis or {}
    # 持仓日期统一为中文格式
    ha_date_cn = ha.get('data_date', '')
    if ha_date_cn and len(ha_date_cn) == 8 and ha_date_cn.isdigit():
        try:
            ha_date_cn = _cn(pd.Timestamp(datetime.strptime(ha_date_cn, "%Y%m%d")))
        except Exception:
            pass
    if ha.get("available"):
        pos_section = f"""
        <div class="grid2">
        <div class="kv"><span class="k">前20多单合计</span><span class="v">{ha['total_long']:,} 手</span></div>
        <div class="kv"><span class="k">前20空单合计</span><span class="v">{ha['total_short']:,} 手</span></div>
        </div>
        <p style="margin-top:8px;">
        净持仓：<b>{ha['net_pos']:+,}</b>手 → <span class="tag tag-{'bull' if ha['net_judge']=='偏多' else ('bear' if ha['net_judge']=='偏空' else 'neutral')}">{ha['net_judge']}</span><br>
        {ha['zhengzhi_summary']}<br>
        {ha['fanzhi_summary']}<br>
        📌 综合：<b>{ha['overall_judge']}</b>
        </p>"""
    else:
        pos_section = f"<p>⚠️ 前20持仓数据未更新（当日大商所持仓排名尚未发布，通常T+1更新）</p>"

    # 技术
    tech_summary = trend_direction if trend_direction else "震荡"
    res_str = "、".join(sr_lines_info.get("resistances", ["暂无"])) if sr_lines_info else "暂无"
    sup_str = "、".join(sr_lines_info.get("supports", ["暂无"])) if sr_lines_info else "暂无"

    # 一句话总结
    bull_score = bear_score = 0
    if chg > 0: bull_score += 1
    elif chg < 0: bear_score += 1
    if na_basis > 300: bull_score += 1
    elif na_basis < -300: bear_score += 1
    if ha.get("net_judge") == "偏多": bull_score += 1
    elif ha.get("net_judge") == "偏空": bear_score += 1
    if tech_summary in ("多头", "偏多"): bull_score += 1
    elif tech_summary in ("空头", "偏空"): bear_score += 1

    if bull_score > bear_score:
        overall = "🐂 市场整体偏多，可关注回调机会"
    elif bear_score > bull_score:
        overall = "🐻 市场整体偏空，建议观望为主"
    else:
        overall = "⚖️ 市场多空交织，短期方向不明，建议观望"

    # ── 持仓量与成交量分析 HTML ──
    vo = vol_oi_analysis or {}
    if vo.get("available"):
        vol_oi_html = f"""
        <p style="font-size:0.92rem;line-height:1.8;">
        • 成交量：{vo.get('vol_text', '—')}<br>
        • 持仓量：{vo.get('oi_text', '—')}<br>
        • 持仓峰值：{vo.get('oi_decline', '—')}
        </p>"""
    else:
        vol_oi_html = "<p>⚠️ 成交量/持仓量数据不足</p>"

    # 图表嵌入
    chart_html = ""
    if chart_images:
        if "basis_comparison" in chart_images:
            chart_html += f'''<div class="card chart-card">
<h2><span class="icon">📊</span>主力合约基差季节对比</h2>
<img src="data:image/png;base64,{chart_images['basis_comparison']}" style="width:100%;max-width:860px;" alt="基差对比">
</div>'''
        if "spread_trend" in chart_images:
            chart_html += f'''<div class="card chart-card">
<h2><span class="icon">📉</span>价差走势图</h2>
<img src="data:image/png;base64,{chart_images['spread_trend']}" style="width:100%;max-width:860px;" alt="价差走势">
</div>'''

    # ── 构建基差分析HTML（含历史分位）──
    basis_analysis_html = _build_basis_analysis_html(snap, basis_enhanced, na_basis, max_region, max_basis, min_region, min_basis)

    # ── 构建 HTML ──
    html = f"""
<!DOCTYPE html>
<html lang="zh-CN">
<head><meta charset="utf-8"><title>生猪期货每日分析报告</title>
<style>
body {{ font-family: 'Microsoft YaHei', 'SimHei', sans-serif; background: #f0f2f5; padding: 16px; color: #2c3e50; }}
.report {{ max-width: 900px; margin: 0 auto; }}
.header {{ background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%); color: #fff; padding: 20px 28px; border-radius: 10px 10px 0 0; }}
.header h1 {{ font-size: 1.5rem; margin: 0 0 4px 0; letter-spacing: 2px; }}
.header .sub {{ font-size: 0.82rem; opacity: 0.8; line-height: 1.6; }}
.card {{ background: #fff; border-radius: 8px; padding: 16px 20px; margin: 10px 0; box-shadow: 0 1px 3px rgba(0,0,0,0.05); border-left: 3px solid #e0e0e0; }}
.card h2 {{ font-size: 1.05rem; margin: 0 0 8px 0; padding-bottom: 6px; border-bottom: 1px solid #f0f0f0; color: #34495e; }}
.card h2 .icon {{ margin-right: 5px; }}
.grid2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }}
.kv {{ display: flex; justify-content: space-between; padding: 3px 0; border-bottom: 1px dotted #f0f0f0; }}
.kv .k {{ color: #999; font-size: 0.9rem; }}
.kv .v {{ font-weight: 600; }}
.up {{ color: #E74C3C; }}
.down {{ color: #27AE60; }}
.tag {{ display: inline-block; padding: 1px 8px; border-radius: 3px; font-size: 0.82rem; font-weight: 600; }}
.tag-bull {{ background: #fde8e8; color: #E74C3C; }}
.tag-bear {{ background: #e8f5e9; color: #27AE60; }}
.tag-neutral {{ background: #e8eaf6; color: #5c6bc0; }}
.conclusion {{ background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: #fff; padding: 14px 20px; border-radius: 8px; margin: 12px 0; font-size: 1rem; text-align: center; }}
.source {{ text-align: center; color: #bbb; font-size: 0.75rem; margin-top: 14px; }}
</style></head>
<body><div class="report">

<div class="header">
<h1>🐷 生猪期货每日分析报告</h1>
<div class="sub">报告日期：{cn_date} ｜ 主力合约：{main_ct}</div>
<div class="sub" style="font-size:0.75rem;margin-top:4px;">
期货数据截止：{_cn(fut_df['date'].max()) if fut_df is not None else '—'} ｜
持仓数据截止：{ha.get('data_date', '—') if ha else '—'} ｜
数据来源：涌益咨询现货 & 大商所期货 & akshare持仓排名
</div>
</div>

<!-- 1. 市场概况 -->
<div class="card">
<h2><span class="icon">📊</span>市场概况 — {main_ct}（{cn_date}）</h2>
<div class="grid2">
<div class="kv"><span class="k">收盘价</span><span class="v">{close_v:.0f} 元/吨</span></div>
<div class="kv"><span class="k">涨跌幅</span><span class="v {chg_class}">{chg_sign}{chg:.0f} ({chg_sign}{chg_pct:.2f}%)</span></div>
<div class="kv"><span class="k">最高价</span><span class="v">{high_v:.0f} 元/吨</span></div>
<div class="kv"><span class="k">最低价</span><span class="v">{low_v:.0f} 元/吨</span></div>
<div class="kv"><span class="k">成交量</span><span class="v">{vol_v:,} 手</span></div>
<div class="kv"><span class="k">持仓量</span><span class="v">{oi_v:,} 手</span></div>
</div>
<p style="margin-top:10px;font-size:0.92rem;color:#666;">
📝 {main_ct} 当日收于 <b>{close_v:.0f}</b> 元/吨，较前日（{cn_prev}）<b class="{chg_class}">{chg_sign}{chg:.0f} ({chg_sign}{chg_pct:.2f}%)</b>，成交量{vol_v:,}手。
</p>
</div>

<!-- 2. 基差分析 -->
<div class="card">
<h2><span class="icon">📐</span>基差分析（{cn_date}）</h2>
<p style="font-size:0.92rem;line-height:1.9;">
{basis_analysis_html}
</p>
</div>

<!-- 3. 价差分析 -->
<div class="card">
<h2><span class="icon">💰</span>价差分析（{spread_date_str if spread_date_str else cn_date}）</h2>
<p>{spread_text}</p>
</div>

<!-- 4. 持仓量与成交量分析 -->
<div class="card">
<h2><span class="icon">📊</span>持仓量与成交量分析（{cn_date}）</h2>
{vol_oi_html}
</div>

<!-- 5. 前20净持仓分析 -->
<div class="card">
<h2><span class="icon">🏢</span>前20净持仓分析（{ha_date_cn if ha_date_cn else '—'}）</h2>
{pos_section}
</div>

<!-- 6. 技术分析 -->
<div class="card">
<h2><span class="icon">📉</span>技术分析（{cn_date}）</h2>
<p style="font-size:0.92rem;line-height:1.8;">
{tech_summary.replace(' | ', '<br>• ')}
</p>
<p style="margin-top:4px;font-size:0.88rem;color:#666;">
压力位：{res_str} ｜ 支撑位：{sup_str}
</p>
</div>

<!-- 图表 -->
{chart_html}

<!-- 7. 综合结论 -->
<div class="conclusion">{overall}</div>

<div class="source">⚠️ 免责声明：本报告仅供参考，不构成任何投资建议。投资有风险，入市需谨慎。<br>数据来源：涌益咨询现货 ｜ 大连商品交易所期货数据 ｜ 前20持仓数据来自大商所/新浪财经</div>

</div></body></html>"""
    return html


def _build_daily_report_md(main_ct, fut_df, spot_dict, ltd, prev_td,
                            snap, holdings_analysis, key_spread_info,
                            trend_direction, sr_lines_info, basis_enhanced=None,
                            vol_oi_analysis=None) -> str:
    """构建日报 Markdown 内容"""
    cn_date = _cn(ltd)
    cn_prev = _cn(prev_td) if prev_td else "前一交易日"

    row_today = fut_df[fut_df["date"] == ltd]
    row_prev = fut_df[fut_df["date"] == prev_td] if prev_td else None
    if not row_today.empty:
        row_t = row_today.iloc[-1]
        close_v = float(row_t["close"])
        high_v = float(row_t.get("high", close_v))
        low_v = float(row_t.get("low", close_v))
        vol_v = int(row_t.get("volume", 0))
        oi_col = "open_interest" if "open_interest" in fut_df.columns else ("hold" if "hold" in fut_df.columns else None)
        oi_v = int(row_t[oi_col]) if oi_col and pd.notna(row_t.get(oi_col)) else 0
        prev_close = float(row_prev["close"].iloc[-1]) if row_prev is not None and not row_prev.empty else close_v
        chg = close_v - prev_close
        chg_pct = chg / prev_close * 100 if prev_close != 0 else 0
    else:
        close_v, high_v, low_v, vol_v, oi_v, chg, chg_pct = 0, 0, 0, 0, 0, 0, 0

    na_basis = snap.get("national_avg", 0) if snap else 0
    max_region = snap.get("max_region", "—") if snap else "—"
    max_basis = snap.get("max_basis", 0) if snap else 0
    min_region = snap.get("min_region", "—") if snap else "—"
    min_basis = snap.get("min_basis", 0) if snap else 0

    ha = holdings_analysis or {}
    if ha.get("available"):
        pos_section = f"""- 前20多单合计：**{ha['total_long']:,}**手
- 前20空单合计：**{ha['total_short']:,}**手
- 净持仓：**{ha['net_pos']:+,}**手 → {ha['net_judge']}
- {ha['zhengzhi_summary']}
- {ha['fanzhi_summary']}
- 综合：**{ha['overall_judge']}**"""
    else:
        pos_section = "⚠️ 前20持仓数据缺失（当日大商所数据尚未发布）"

    res_str = "、".join(sr_lines_info.get("resistances", ["暂无"])) if sr_lines_info else "暂无"
    sup_str = "、".join(sr_lines_info.get("supports", ["暂无"])) if sr_lines_info else "暂无"

    md = f"""# 🐷 生猪期货每日分析报告

**报告日期：{cn_date}** ｜ 主力合约：{main_ct}

> 数据来源：涌益咨询现货数据 ｜ 大连商品交易所期货数据

---

## 📊 一、市场概况 — {main_ct}

| 指标 | 数值 |
|------|------|
| 收盘价 | {close_v:.0f} 元/吨 |
| 涨跌 | {chg:+.0f} ({chg_pct:+.2f}%) |
| 最高价 | {high_v:.0f} 元/吨 |
| 最低价 | {low_v:.0f} 元/吨 |
| 成交量 | {vol_v:,} 手 |
| 持仓量 | {oi_v:,} 手 |

📝 {main_ct} 当日收于 **{close_v:.0f}** 元/吨，较{cn_prev} **{chg:+.0f}**（{chg_pct:+.2f}%）。

## 📐 二、基差分析

- 全国均价基差：**{na_basis:+,}元/吨**
- 基差最大区域：**{max_region}**（{max_basis:+,}元/吨）
- 基差最小区域：**{min_region}**（{min_basis:+,}元/吨）

## 💰 三、价差分析

{key_spread_info if key_spread_info else "暂无价差数据"}

## 🏢 四、前20净持仓分析

{pos_section}

## 📉 五、技术分析

- 趋势判断：**{trend_direction or '震荡'}**
- 压力位：{res_str}
- 支撑位：{sup_str}

## 🎯 六、综合结论

> 以上各维度综合判断，当前市场建议关注后续走势。

---

⚠️ **免责声明**：本报告仅供参考，不构成任何投资建议。投资有风险，入市需谨慎。
"""
    return md


def _build_reportlab_pdf(html_content: str, cn_date: str, chart_images: dict = None) -> Optional[bytes]:
    """使用 reportlab 生成 PDF（含嵌入图片），失败返回 None"""
    if not HAS_REPORTLAB:
        return None
    try:
        import re
        from reportlab.platypus import Image as RLImage
        import base64
        import tempfile
        import os

        buf = BytesIO()
        doc = SimpleDocTemplate(buf, pagesize=A4,
            leftMargin=20*mm, rightMargin=20*mm, topMargin=15*mm, bottomMargin=15*mm)
        styles = getSampleStyleSheet()

        title_style = ParagraphStyle('CNTitle', parent=styles['Title'],
            fontName=CN_FONT, fontSize=18, spaceAfter=10, textColor=HexColor('#1a1a2e'))
        h2_style = ParagraphStyle('CNH2', parent=styles['Heading2'],
            fontName=CN_FONT, fontSize=13, spaceBefore=12, spaceAfter=6,
            textColor=HexColor('#34495e'))
        body_style = ParagraphStyle('CNBody', parent=styles['Normal'],
            fontName=CN_FONT, fontSize=10, leading=16, textColor=HexColor('#333333'))
        small_style = ParagraphStyle('CNSmall', parent=styles['Normal'],
            fontName=CN_FONT, fontSize=8, textColor=HexColor('#999999'))

        story = []
        story.append(Paragraph("生猪期货每日分析报告", title_style))
        story.append(Paragraph(f"报告日期：{cn_date}", small_style))
        story.append(Spacer(1, 12))

        # Extract cards from HTML
        cards = re.findall(r'<div class="card">(.*?)</div>', html_content, re.DOTALL)
        tmp_files = []
        for card_html in cards:
            h2_match = re.search(r'<h2>(.*?)</h2>', card_html, re.DOTALL)
            if h2_match:
                title = re.sub(r'<.*?>', '', h2_match.group(1)).strip()
                story.append(Paragraph(title, h2_style))
            content = re.sub(r'<h2>.*?</h2>', '', card_html, flags=re.DOTALL)
            # Check for embedded images
            img_match = re.search(r'<img src="data:image/png;base64,([^"]+)"', content)
            if img_match:
                img_data = base64.b64decode(img_match.group(1))
                tmpf = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
                tmpf.write(img_data)
                tmpf.close()
                tmp_files.append(tmpf.name)
                img = RLImage(tmpf.name, width=450, height=200)
                story.append(img)
                text = re.sub(r'<img[^>]+>', '', content)
            else:
                text = content
            text = re.sub(r'<.*?>', ' ', text).strip()
            text = re.sub(r'\s+', ' ', text)
            for line in text.split('。'):
                line = line.strip()
                if line:
                    story.append(Paragraph(line + '。', body_style))
            story.append(Spacer(1, 6))

        # Conclusion
        concl_match = re.search(r'<div class="conclusion">(.*?)</div>', html_content, re.DOTALL)
        if concl_match:
            concl_text = re.sub(r'<.*?>', '', concl_match.group(1)).strip()
            concl_style = ParagraphStyle('CNConcl', parent=body_style,
                fontName=CN_FONT, fontSize=11, textColor=HexColor('#ffffff'),
                backColor=HexColor('#667eea'), alignment=1)
            story.append(Spacer(1, 10))
            story.append(Paragraph(concl_text, concl_style))

        story.append(Spacer(1, 20))
        story.append(Paragraph("免责声明：本报告仅供参考，不构成任何投资建议。", small_style))

        doc.build(story)

        # Cleanup temp files
        for f in tmp_files:
            try: os.unlink(f)
            except Exception: pass

        return buf.getvalue()
    except Exception:
        return None


# ★ 修改日报逻辑后递增此版本号，使旧缓存自动失效
_DAILY_REPORT_VERSION = 11

@st.cache_data(ttl=120, show_spinner=False)
def _compute_daily_report_cache(main_ct: str, spot_hash: int, _version: int = 0) -> dict:
    """缓存日报计算。期货数据用最新交易日，持仓数据用API实际返回日期。"""
    spot_dict, _ = load_spot(str(SPOT_PATH))
    fut_df, _ = load_futures(main_ct)
    if fut_df is None or fut_df.empty:
        return {"error": "期货数据不可用"}

    # ── 期货最新交易日 ──
    fds = sorted(fut_df["date"].unique())
    ltd_ts = pd.Timestamp(fds[-1])  # 直接用期货数据的最后一天
    td_idx = len(fds) - 1
    prev_td = fds[td_idx - 1] if td_idx > 0 else None

    # ── 基差快照（使用期货数据日期）──
    regions = get_regions(main_ct)
    snap = compute_snapshot(main_ct, spot_dict, fut_df, ltd_ts, regions)

    # ── 持仓数据（如果API数据日期≠报告日期则不显示详情）──
    holdings_df, holdings_actual_date, holdings_source = _get_holdings(
        main_ct, ltd_ts, return_meta=True)
    # 判断持仓数据是否与报告日期一致
    holdings_date_match = False
    try:
        hdt = pd.to_datetime(holdings_actual_date, format="%Y%m%d")
        if hdt.date() == ltd_ts.date():
            holdings_date_match = True
    except Exception:
        pass
    if holdings_date_match:
        holdings_analysis = _analyze_holdings_for_report(
            holdings_df, main_ct, ltd_ts, holdings_actual_date)
    else:
        holdings_analysis = {"available": False, "data_date": str(holdings_actual_date)[:8] if holdings_actual_date else _cn(ltd_ts),
                            "total_long": 0, "total_short": 0, "net_pos": 0,
                            "net_judge": "数据未更新",
                            "zhengzhi_summary": "数据未更新",
                            "fanzhi_summary": "数据未更新",
                            "overall_judge": "数据未更新"}

    # ── 增强基差分析：计算历史分位 ──
    basis_enhanced = _analyze_basis_historical(main_ct, spot_dict, fut_df, ltd_ts, regions, snap)

    # ── 价差（返回(文本, 实际日期)）──
    spread_info, spread_date = _compute_key_spread(main_ct, ltd_ts)

    # ── 成交量/持仓量动能分析 ──
    vol_oi_analysis = _analyze_vol_oi_momentum(fut_df, ltd_ts)

    # ── 技术分析（使用期货数据）──
    trend_dir, sr_info = _quick_technical(fut_df, ltd_ts)

    # ── 生成图表 ──
    chart_images = _generate_report_charts(main_ct, spot_dict, ltd_ts)

    # ── 构建HTML和MD ──
    html = _build_daily_report_html(main_ct, fut_df, spot_dict, ltd_ts, prev_td,
                                     snap, holdings_analysis, spread_info,
                                     trend_dir, sr_info, chart_images, basis_enhanced,
                                     vol_oi_analysis, spread_date)
    md = _build_daily_report_md(main_ct, fut_df, spot_dict, ltd_ts, prev_td,
                                 snap, holdings_analysis, spread_info,
                                 trend_dir, sr_info, basis_enhanced, vol_oi_analysis)

    return {"html": html, "md": md, "error": None,
            "ltd": ltd_ts, "cn_date": _cn(ltd_ts),
            "holdings_date": holdings_actual_date,
            "chart_images": chart_images}


def _analyze_holdings_for_report(holdings_df, main_ct, ltd, holdings_actual_date=None):
    """分析持仓数据。数据必须来自真实API，缺失时标注'数据未更新'。"""
    if holdings_df is None or holdings_df.empty:
        date_str = str(holdings_actual_date)[:8] if holdings_actual_date else _cn(ltd)
        return {
            "available": False,
            "data_date": date_str,
            "total_long": 0, "total_short": 0, "net_pos": 0,
            "net_judge": "数据未更新",
            "zhengzhi_summary": "数据未更新（当日大商所持仓排名尚未发布，通常T+1更新）",
            "fanzhi_summary": "数据未更新",
            "overall_judge": "数据未更新",
        }

    total_long = int(holdings_df["long"].sum())
    total_short = int(holdings_df["short"].sum())
    net_pos = total_long - total_short

    if net_pos > 5000:
        net_judge = "偏多"
    elif net_pos < -5000:
        net_judge = "偏空"
    else:
        net_judge = "中性"

    # 数据日期
    if holdings_actual_date:
        try:
            data_dt = pd.to_datetime(holdings_actual_date, format="%Y%m%d")
            data_date_str = _cn(data_dt)
        except Exception:
            data_date_str = str(holdings_actual_date)
    else:
        data_date_str = _cn(ltd)

    # 正指分析
    zhengzhi_parts = []
    fanzhi_parts = []
    zz_bull, zz_bear = 0, 0
    fz_bull, fz_bear = 0, 0

    for _, row in holdings_df.iterrows():
        co = str(row.get("company", ""))
        lc = int(row.get("long_chg", 0)) if pd.notna(row.get("long_chg")) else 0
        sc = int(row.get("short_chg", 0)) if pd.notna(row.get("short_chg")) else 0

        def _action_text(lc, sc):
            parts = []
            if lc > 0: parts.append(f"加多{lc:,}")
            elif lc < 0: parts.append(f"减多{lc:,}")
            if sc > 0: parts.append(f"加空{sc:,}")
            elif sc < 0: parts.append(f"减空{sc:,}")
            return "、".join(parts) if parts else "持仓不变"

        if co in ZHENGZHI_COMPANIES:
            action = _action_text(lc, sc)
            if lc > 0 and sc < 0:
                intent = "看多"; zz_bull += 1
            elif lc < 0 and sc > 0:
                intent = "看空"; zz_bear += 1
            else:
                intent = "中性"
            zhengzhi_parts.append(f"{co}({action}→{intent})")
        elif co in FANZHI_COMPANIES:
            action = _action_text(lc, sc)
            if lc > 0 and sc < 0:
                intent = "看多"; fz_bull += 1
            elif lc < 0 and sc > 0:
                intent = "看空"; fz_bear += 1
            else:
                intent = "中性"
            fanzhi_parts.append(f"{co}({action}→{intent})")

    # 正指方向
    if zz_bull > zz_bear:
        zz_dir = "看多"; zz_market = "利多"
    elif zz_bear > zz_bull:
        zz_dir = "看空"; zz_market = "利空"
    else:
        zz_dir = "分歧"; zz_market = "中性"

    # 反指方向（反指看多=市场利空）
    if fz_bull > fz_bear:
        fz_dir = "看多"; fz_market = "利空"
    elif fz_bear > fz_bull:
        fz_dir = "看空"; fz_market = "利多"
    else:
        fz_dir = "分歧"; fz_market = "中性"

    bull_score = zz_bull + fz_bear
    bear_score = zz_bear + fz_bull
    if bull_score > bear_score:
        overall = "整体偏多"
    elif bear_score > bull_score:
        overall = "整体偏空"
    else:
        overall = "方向分歧"

    return {
        "available": True,
        "total_long": total_long,
        "total_short": total_short,
        "net_pos": net_pos,
        "net_judge": net_judge,
        "zhengzhi_summary": f"正指{zz_dir}（{zz_market}）：{'；'.join(zhengzhi_parts) if zhengzhi_parts else '无正指席位数据'}",
        "fanzhi_summary": f"反指{fz_dir}（{fz_market}）：{'；'.join(fanzhi_parts) if fanzhi_parts else '无反指席位数据'}",
        "data_date": data_date_str,
        "overall_judge": f"持仓日期：{data_date_str}｜正指{zz_dir}，反指{fz_dir}，{overall}",
    }


def _generate_report_charts(main_ct, spot_dict, ltd) -> dict:
    """生成日报所需图表，返回 {name: base64_png}"""
    import base64
    charts = {}

    # 1. 主力合约基差季节对比图（收盘价基差同比）
    try:
        tmon = ct_month(main_ct)
        same_month = [c for c in ALL_CONTRACTS if ct_month(c) == tmon]
        avail = []
        for c in same_month:
            df_c, _ = load_futures(c)
            if df_c is not None and not df_c.empty:
                avail.append(c)
        if len(avail) >= 2:
            # Build basis series for 全国均价
            series = {}
            md_collector = defaultdict(list)
            for c in avail:
                fdf, _ = load_futures(c)
                na_df = calc_national_basis(spot_dict, fdf)
                if na_df is None or na_df.empty:
                    continue
                na_df["year"] = na_df["date"].dt.year
                na_df["doy"] = na_df["date"].dt.dayofyear
                na_df["plot_date"] = na_df.apply(
                    lambda r: _doy_to_date(int(r["doy"]), int(r["year"])), axis=1)
                for yr, grp in na_df.groupby("year"):
                    grp = grp.sort_values("doy").copy()
                    label = _make_trace_label(c, yr, "全国均价")
                    series[label] = grp
                    for _, row in grp.iterrows():
                        md_collector[(row["date"].month, row["date"].day)].append(row["basis"])

            if series:
                # Add historical avg
                avg_rows = [{"doy": m*100+d, "basis": int(round(np.mean(v))),
                             "plot_date": pd.Timestamp(year=2020, month=m, day=d)}
                            for (m, d), v in sorted(md_collector.items()) if v]
                if avg_rows:
                    series["历史均值"] = pd.DataFrame(avg_rows).sort_values("doy")

                fig_basis = fig_calendar_comparison(series, tmon, "")
                fig_basis.update_layout(height=380, margin=dict(t=40, b=30, l=40, r=20))
                img_bytes = fig_basis.to_image(format="png", scale=1.0, width=900)
                charts["basis_comparison"] = base64.b64encode(img_bytes).decode()
    except Exception:
        pass

    # 2. 价差走势图（主力 vs 次主力，近90天）
    try:
        active = get_active_contracts()
        others = [c for c in active if c != main_ct]
        if others:
            ct_b = others[0]
            dfa, _ = load_futures(main_ct)
            dfb, _ = load_futures(ct_b)
            if dfa is not None and dfb is not None:
                ac = dfa.set_index("date")["close"]
                bc = dfb.set_index("date")["close"]
                cm = sorted(ac.index.intersection(bc.index))
                if len(cm) > 0:
                    recent_cm = cm[-90:]
                    spreads_v = [float(ac[d] - bc[d]) for d in recent_cm]
                    fig_sp = go.Figure()
                    fig_sp.add_trace(go.Scatter(
                        x=recent_cm, y=spreads_v, mode="lines",
                        name=f"{main_ct}-{ct_b}",
                        line=dict(color="#E74C3C", width=2),
                        hovertemplate="%{x|%Y-%m-%d}<br>价差：%{y:+,.0f}<extra></extra>"
                    ))
                    fig_sp.add_hline(y=0, line_dash="dash", line_color="gray", opacity=0.3)
                    fig_sp.update_layout(
                        title=f"{main_ct}-{ct_b} 价差走势（近90日）",
                        xaxis=dict(title="日期", tickformat="%m-%d"),
                        yaxis=dict(title="价差（元/吨）"),
                        template="plotly_white", height=350,
                        margin=dict(t=50, b=40, l=50, r=20),
                    )
                    img_bytes = fig_sp.to_image(format="png", scale=1.0, width=900)
                    charts["spread_trend"] = base64.b64encode(img_bytes).decode()
    except Exception:
        pass

    return charts


def _compute_key_spread(main_ct: str, ltd=None):
    """返回 (html_text, date_str) — date_str为价差实际数据日期"""
    """价差分析 — 使用与 Tab5 完全相同的 spread_collector[(month,day)] 逻辑。
    历史同期均值 = 所有年份同月合约对在同月同日的价差均值。"""
    active = get_active_contracts()
    if len(active) < 2:
        return "暂无足够合约计算价差", ""
    try:
        ct_a = main_ct
        others = [c for c in active if ct_month(c) != ct_month(main_ct)]
        if not others:
            others = [c for c in active if c != ct_a]
        if not others:
            return "暂无其他上市合约", ""
        ct_b = others[0]
        ma, mb = ct_month(ct_a), ct_month(ct_b)

        # ── 当前价差：取最新共同交易日 ──
        dfa, _ = load_futures(ct_a)
        dfb, _ = load_futures(ct_b)
        if dfa is None or dfa.empty or dfb is None or dfb.empty:
            return f"{ct_a}-{ct_b} 价差数据不足", ""
        ac = dfa.set_index("date")["close"]
        bc = dfb.set_index("date")["close"]
        cm_cur = sorted(ac.index.intersection(bc.index))
        if len(cm_cur) == 0:
            return f"{ct_a}-{ct_b} 无共同交易日", ""

        # 用ltd或最新共同日
        if ltd and ltd in cm_cur:
            latest_dt = ltd
        else:
            latest_dt = cm_cur[-1]
        latest_spread = float(ac[latest_dt] - bc[latest_dt])
        target_m, target_d = latest_dt.month, latest_dt.day

        # ── 历史同期：与 Tab5 完全相同的 (month, day) collector ──
        # Tab5: spread_collector[(d.month, d.day)].append(row["spread"])
        spread_by_md = defaultdict(list)

        valid_years = []
        for y in range(21, 28):
            ca, cb = f"LH{y:02d}{ma}", f"LH{y:02d}{mb}"
            if ca not in ALL_CONTRACTS or cb not in ALL_CONTRACTS:
                continue
            dfa_h, _ = load_futures(ca)
            dfb_h, _ = load_futures(cb)
            if dfa_h is None or dfa_h.empty or dfb_h is None or dfb_h.empty:
                continue
            try:
                ach = dfa_h.set_index("date")["close"]
                bch = dfb_h.set_index("date")["close"]
                cm_h = ach.index.intersection(bch.index)
                if len(cm_h) == 0: continue
                for d in cm_h:
                    spread_by_md[(d.month, d.day)].append(float(ach[d] - bch[d]))
                valid_years.append(y)
            except Exception:
                continue

        same_day_vals = spread_by_md.get((target_m, target_d), [])

        if same_day_vals:
            hist_avg = float(np.mean(same_day_vals))
            if hist_avg != 0:
                dev = (latest_spread - hist_avg) / abs(hist_avg) * 100
            else:
                dev = 0
            if dev > 30:
                pos = f"明显高于历史同期均值（偏离+{dev:.0f}%）"
            elif dev > 10:
                pos = f"高于历史同期均值（偏离+{dev:.0f}%）"
            elif dev < -30:
                pos = f"明显低于历史同期均值（偏离{dev:.0f}%）"
            elif dev < -10:
                pos = f"低于历史同期均值（偏离{dev:.0f}%）"
            else:
                pos = f"处于历史同期均值附近（偏离{dev:+.0f}%）"
        else:
            return (f"{ct_a}-{ct_b} 价差 <b>{latest_spread:+,.0f}元/吨</b>"
                    f"（{_cn(pd.Timestamp(latest_dt))}），暂无同月同日历史数据", "")

        return (f"{ct_a}-{ct_b} 价差 <b>{latest_spread:+,.0f}元/吨</b>"
                f"（{_cn(pd.Timestamp(latest_dt))}），"
                f"历史同期均值 <b>{hist_avg:+,.0f}元/吨</b>，{pos}", _cn(pd.Timestamp(latest_dt)))
    except Exception as e:
        return f"价差计算异常: {e}", ""


def _quick_technical(fut_df, ltd) -> Tuple[str, dict]:
    """快速技术分析：趋势 + MACD/RSI/布林带 完整文字结论"""
    if fut_df is None or fut_df.empty:
        return "震荡", {"resistances": [], "supports": []}
    try:
        df = fut_df.sort_values("date").reset_index(drop=True)
        # 计算技术指标
        close = df["close"].astype(float)
        high = df["high"].astype(float) if "high" in df.columns else close
        low = df["low"].astype(float) if "low" in df.columns else close

        n = len(close)
        if n < 20:
            return "数据不足(需≥20日)", {"resistances": [], "supports": []}

        cur = float(close.iloc[-1])

        # ── 均线趋势 ──
        ma5 = close.rolling(5).mean().iloc[-1]
        ma10 = close.rolling(10).mean().iloc[-1]
        ma20 = close.rolling(20).mean().iloc[-1]
        ma5_v = float(ma5) if pd.notna(ma5) else 0
        ma10_v = float(ma10) if pd.notna(ma10) else 0
        ma20_v = float(ma20) if pd.notna(ma20) else 0

        # ── MACD ──
        ema12 = close.ewm(span=12, adjust=False).mean()
        ema26 = close.ewm(span=26, adjust=False).mean()
        dif = ema12 - ema26
        dea = dif.ewm(span=9, adjust=False).mean()
        macd_hist = 2 * (dif - dea)
        dif_v = float(dif.iloc[-1]); dea_v = float(dea.iloc[-1])
        macd_v = float(macd_hist.iloc[-1])
        dif_prev = float(dif.iloc[-2]) if n >= 2 else dif_v
        dea_prev = float(dea.iloc[-2]) if n >= 2 else dea_v

        # ── RSI14 ──
        delta = close.diff()
        gain = delta.clip(lower=0)
        loss = (-delta).clip(lower=0)
        avg_gain = gain.ewm(alpha=1/14, adjust=False).mean()
        avg_loss = loss.ewm(alpha=1/14, adjust=False).mean()
        rs = avg_gain / avg_loss.replace(0, np.nan)
        rsi14_v = float((100 - (100 / (1 + rs))).iloc[-1])

        # ── 布林带 ──
        bb_mid = ma20_v
        std20 = close.rolling(20).std().iloc[-1]
        bb_up = bb_mid + 2 * float(std20) if pd.notna(std20) else cur * 1.05
        bb_low = bb_mid - 2 * float(std20) if pd.notna(std20) else cur * 0.95
        bb_width = (bb_up - bb_low) / bb_mid * 100 if bb_mid > 0 else 0

        # ── 构建文字结论 ──
        tech_lines = []

        # 趋势
        if ma5_v > ma10_v > ma20_v:
            trend = "多头排列，趋势偏强"
        elif ma5_v < ma10_v < ma20_v:
            trend = "空头排列，趋势偏弱"
        elif cur > ma20_v:
            trend = "价格在MA20上方，短期偏多"
        elif cur < ma20_v:
            trend = "价格在MA20下方，短期偏空"
        else:
            trend = "均线交织，震荡格局"

        # MACD
        if dif_prev <= dea_prev and dif_v > dea_v:
            macd_text = "金叉形成（DIF上穿DEA），看涨信号"
        elif dif_prev >= dea_prev and dif_v < dea_v:
            macd_text = "死叉形成（DIF下穿DEA），看跌信号"
        elif dif_v > dea_v:
            macd_text = f"DIF在DEA上方运行，{'红柱' if macd_v > 0 else '绿柱'}，{'动能增强' if abs(macd_v) > 0 else '动能减弱'}"
        else:
            macd_text = f"DIF在DEA下方运行，{'绿柱' if macd_v < 0 else '红柱'}，{'动能增强' if abs(macd_v) > 0 else '动能减弱'}"

        # RSI
        if rsi14_v > 70:
            rsi_text = f"RSI={rsi14_v:.1f}，超买区域（>70），注意回调"
        elif rsi14_v < 30:
            rsi_text = f"RSI={rsi14_v:.1f}，超卖区域（<30），反弹概率增大"
        else:
            rsi_text = f"RSI={rsi14_v:.1f}，中性区间（30-70）"

        # 布林带
        if cur > bb_up:
            bb_text = f"突破布林上轨（{bb_up:.0f}），超强格局，带宽{bb_width:.1f}%"
        elif cur < bb_low:
            bb_text = f"跌破布林下轨（{bb_low:.0f}），超弱格局，带宽{bb_width:.1f}%"
        elif cur > bb_mid:
            pct = (cur - bb_mid) / (bb_up - bb_mid) * 100 if bb_up > bb_mid else 50
            bb_text = f"运行于中轨与上轨之间（{pct:.0f}%），偏强，带宽{bb_width:.1f}%"
        else:
            pct = (cur - bb_low) / (bb_mid - bb_low) * 100 if bb_mid > bb_low else 50
            bb_text = f"运行于中轨与下轨之间（{pct:.0f}%），偏弱，带宽{bb_width:.1f}%"

        # 总结方向
        bull_score = 0; bear_score = 0
        if "多头" in trend or "偏多" in trend: bull_score += 1
        elif "空头" in trend or "偏空" in trend: bear_score += 1
        if "金叉" in macd_text or "上方运行" in macd_text: bull_score += 1
        elif "死叉" in macd_text or "下方运行" in macd_text: bear_score += 1
        if rsi14_v < 30: bull_score += 1
        elif rsi14_v > 70: bear_score += 1
        if cur > bb_mid: bull_score += 1
        elif cur < bb_mid: bear_score += 1

        if bull_score >= 3: direction = "偏多"
        elif bear_score >= 3: direction = "偏空"
        elif bull_score > bear_score: direction = "中性偏多"
        elif bear_score > bull_score: direction = "中性偏空"
        else: direction = "震荡"

        tech_summary = f"{direction} | {trend} | {macd_text} | {rsi_text} | {bb_text}"

        # 支撑/压力
        recent_high = float(close.tail(20).max())
        recent_low = float(close.tail(20).min())
        resistances = [f"前高{recent_high:.0f}"]
        supports = [f"前低{recent_low:.0f}"]
        if bb_up > cur: resistances.append(f"布林上轨{bb_up:.0f}")
        if bb_low < cur: supports.append(f"布林下轨{bb_low:.0f}")
        if ma20_v < cur: supports.append(f"MA20={ma20_v:.0f}")
        else: resistances.append(f"MA20={ma20_v:.0f}")

        return tech_summary, {"resistances": resistances[:2], "supports": supports[:2]}
    except Exception:
        return "震荡(计算异常)", {"resistances": [], "supports": []}


def tab_daily_report():
    """Tab 1: 每日期货分析日报"""
    st.subheader("📋 每日期货分析日报")

    main_ct = get_main_contract()
    spot_dict, _ = load_spot(str(SPOT_PATH))
    spot_hash = _spot_hash(spot_dict)

    # 清除缓存按钮
    col_refresh, _ = st.columns([1, 5])
    with col_refresh:
        if st.button("🔄 刷新日报", key="refresh_daily"):
            _compute_daily_report_cache.clear()
            st.rerun()

    with st.spinner("🔄 正在生成日报…"):
        cache = _compute_daily_report_cache(main_ct, spot_hash, _DAILY_REPORT_VERSION)

    if cache.get("error"):
        st.error(f"❌ {cache['error']}")
        return

    html = cache["html"]
    md = cache["md"]
    cn_date = cache.get("cn_date", "")
    chart_images = cache.get("chart_images", {})

    # ── 下载按钮 ──
    col_title, col_btn1, col_btn2 = st.columns([4, 1, 1])
    with col_title:
        st.caption(f"📅 报告日期：{cn_date} ｜ 主力合约：{main_ct}")
    with col_btn1:
        # PDF with images using weasyprint (more reliable than reportlab for HTML+images)
        pdf_data = _build_weasyprint_pdf(html, cn_date)
        if pdf_data is None:
            pdf_data = _build_reportlab_pdf(html, cn_date, chart_images)
        if pdf_data:
            st.download_button(
                label="📄 下载 PDF",
                data=pdf_data,
                file_name=f"生猪期货日报_{cn_date.replace('年','').replace('月','').replace('日','')}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key="dl_pdf"
            )
        else:
            st.download_button(
                label="📄 下载 HTML",
                data=html.encode("utf-8"),
                file_name=f"生猪期货日报_{cn_date.replace('年','').replace('月','').replace('日','')}.html",
                mime="text/html",
                use_container_width=True,
                key="dl_html",
                help="PDF库不可用，提供HTML格式（浏览器打开后可打印为PDF）"
            )
    with col_btn2:
        png_data = _compose_report_image(chart_images, cn_date, main_ct)
        if png_data:
            st.download_button(
                label="🖼️ 下载图片",
                data=png_data,
                file_name=f"生猪期货日报_{cn_date.replace('年','').replace('月','').replace('日','')}.png",
                mime="image/png",
                use_container_width=True,
                key="dl_img"
            )
        else:
            # 兜底：直接导出HTML
            st.download_button(
                label="🖼️ 下载 HTML",
                data=html.encode("utf-8"),
                file_name=f"生猪期货日报_{cn_date.replace('年','').replace('月','').replace('日','')}.html",
                mime="text/html",
                use_container_width=True,
                key="dl_img_fb",
                help="图片生成失败，提供HTML（浏览器打开后可截图）"
            )

    # ── 渲染 HTML 日报（使用 components.html 避免 sanitizer 破坏样式）──
    st.components.v1.html(html, height=2200, scrolling=True)

    # ── 底部分隔 ──
    st.markdown("---")


def _build_weasyprint_pdf(html: str, cn_date: str) -> Optional[bytes]:
    """使用 weasyprint 生成 PDF，失败返回 None"""
    try:
        from weasyprint import HTML
        buf = BytesIO()
        HTML(string=html).write_pdf(buf)
        return buf.getvalue()
    except ImportError:
        pass
    except Exception:
        pass
    return None


def _compose_report_image(chart_images: dict, cn_date: str, main_ct: str) -> Optional[bytes]:
    """将日报图表生成为一张PNG图片。
    方案1: 直接用plotly subplots组合所有图表（最可靠）
    方案2: PIL拼接已有的chart_images PNG"""
    import base64
    from plotly.subplots import make_subplots

    # ── 方案1: 用plotly直接生成组合图 ──
    try:
        figs_to_combine = []
        titles = []

        # 重建基差季节图
        tmon = ct_month(main_ct)
        same_month = [c for c in ALL_CONTRACTS if ct_month(c) == tmon]
        avail = [c for c in same_month if load_futures(c)[0] is not None and not load_futures(c)[0].empty]
        if len(avail) >= 2:
            spot_dict, _ = load_spot(str(SPOT_PATH))
            series = {}
            md_collector = defaultdict(list)
            for c in avail:
                fdf, _ = load_futures(c)
                na_df = calc_national_basis(spot_dict, fdf)
                if na_df is None or na_df.empty:
                    continue
                na_df["year"] = na_df["date"].dt.year
                na_df["doy"] = na_df["date"].dt.dayofyear
                na_df["plot_date"] = na_df.apply(lambda r: _doy_to_date(int(r["doy"]), int(r["year"])), axis=1)
                for yr, grp in na_df.groupby("year"):
                    grp = grp.sort_values("doy").copy()
                    series[_make_trace_label(c, yr, "全国均价")] = grp
                    for _, row in grp.iterrows():
                        md_collector[(row["date"].month, row["date"].day)].append(row["basis"])
            if series:
                avg_rows = [{"doy": m*100+d, "basis": int(round(np.mean(v))),
                             "plot_date": pd.Timestamp(year=2020, month=m, day=d)}
                            for (m, d), v in sorted(md_collector.items()) if v]
                if avg_rows:
                    series["历史均值"] = pd.DataFrame(avg_rows).sort_values("doy")
                figs_to_combine.append(fig_calendar_comparison(series, tmon, ""))
                titles.append(f"基差季节对比")

        # 重建价差图
        active = get_active_contracts()
        others = [c for c in active if c != main_ct]
        if others:
            ct_b = others[0]
            dfa, _ = load_futures(main_ct)
            dfb, _ = load_futures(ct_b)
            if dfa is not None and dfb is not None:
                ac = dfa.set_index("date")["close"]
                bc = dfb.set_index("date")["close"]
                cm = sorted(ac.index.intersection(bc.index))
                if len(cm) > 0:
                    fig_sp = go.Figure()
                    fig_sp.add_trace(go.Scatter(
                        x=cm[-90:], y=[float(ac[d] - bc[d]) for d in cm[-90:]],
                        mode="lines", name=f"{main_ct}-{ct_b}",
                        line=dict(color="#E74C3C", width=2)))
                    fig_sp.add_hline(y=0, line_dash="dash", line_color="gray", opacity=0.3)
                    fig_sp.update_layout(
                        title=f"{main_ct}-{ct_b} 价差走势（近90日）",
                        xaxis=dict(tickformat="%m-%d"), yaxis=dict(title="价差"),
                        template="plotly_white", height=350, margin=dict(t=50, b=40, l=50, r=20))
                    figs_to_combine.append(fig_sp)
                    titles.append("价差走势")

        if figs_to_combine:
            # 组合为垂直subplots
            n = len(figs_to_combine)
            combined_fig = make_subplots(rows=n, cols=1, subplot_titles=titles,
                                          vertical_spacing=0.08)
            for i, fig in enumerate(figs_to_combine):
                for trace in fig.data:
                    combined_fig.add_trace(trace, row=i+1, col=1)
            combined_fig.update_layout(
                title=dict(text=f"生猪期货每日分析报告  {cn_date}  主力:{main_ct}",
                          font=dict(size=18, color='#1a1a2e')),
                height=400 * n, template="plotly_white",
                margin=dict(t=60, b=30, l=50, r=30))
            img_bytes = combined_fig.to_image(format="png", scale=1.2, width=1100)
            return img_bytes
    except Exception:
        pass

    # ── 方案2: PIL拼接已有chart_images ──
    try:
        from PIL import Image, ImageDraw, ImageFont
        images_to_stack = []
        title_img = Image.new('RGB', (1200, 60), color=(26, 26, 46))
        draw = ImageDraw.Draw(title_img)
        font_title = None
        for fp in ['C:/Windows/Fonts/simhei.ttf', 'C:/Windows/Fonts/msyh.ttf']:
            try:
                font_title = ImageFont.truetype(fp, 24)
                break
            except Exception:
                continue
        if font_title:
            draw.text((15, 15), f"生猪期货每日分析报告  {cn_date}  主力:{main_ct}",
                      fill=(255, 255, 255), font=font_title)
        images_to_stack.append(title_img)
        for key in ["basis_comparison", "spread_trend"]:
            b64 = chart_images.get(key)
            if b64:
                img_bytes = base64.b64decode(b64)
                img = Image.open(BytesIO(img_bytes))
                if img.width > 1200:
                    img = img.resize((1200, int(img.height * 1200 / img.width)), Image.LANCZOS)
                images_to_stack.append(img)
        if len(images_to_stack) > 1:
            total_h = sum(im.height for im in images_to_stack) + 5 * (len(images_to_stack) - 1)
            combined = Image.new('RGB', (1200, total_h + 10), color=(245, 246, 250))
            y = 5
            for im in images_to_stack:
                combined.paste(im, ((1200 - im.width) // 2, y))
                y += im.height + 5
            buf = BytesIO()
            combined.save(buf, format='PNG')
            return buf.getvalue()
    except Exception:
        pass

    return None


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

    # ── 启动时后台下载所有合约信息（仅首次）──
    # Step 1: 下载所有缺失的期货合约数据（并行）
    # Step 2: 下载所有缺失的净持仓聚合数据（并行）
    if "_startup_all_synced" not in st.session_state:
        st.session_state["_startup_all_synced"] = False

    if not st.session_state["_startup_all_synced"]:
        missing_futures = [c for c in ALL_CONTRACTS if not _csv_path(c).exists()]
        all_with_futures_now = [c for c in ALL_CONTRACTS if _csv_path(c).exists()]
        missing_net = [c for c in all_with_futures_now if not _net_agg_path(c).exists()
                       or _load_aggregated_net(c) is None]

        total_tasks = len(missing_futures) + len(missing_net)
        if total_tasks > 0:
            placeholder = st.empty()
            with placeholder.container():
                st.info(
                    f"📡 首次启动：正在后台下载所有历史合约数据…\n\n"
                    f"📦 缺失期货数据：**{len(missing_futures)}** 个合约 ｜ "
                    f"📦 缺失净持仓数据：**{len(missing_net)}** 个合约 ｜ "
                    f"🔢 共 **{total_tasks}** 项任务"
                )
                progress_bar = st.progress(0)
                status_text = st.empty()
                completed = 0

                # ── Step 1: 并行下载所有缺失的期货合约 ──
                if missing_futures:
                    status_text.text(f"⏳ 正在并行下载 {len(missing_futures)} 个合约的期货数据（{_sync_parallel_workers} 线程）…")

                    def _fut_progress(current, total, ct, msg):
                        nonlocal completed
                        completed = current
                        progress_bar.progress(completed / total_tasks,
                            text=f"📈 期货 {completed}/{total} — {ct}: {msg}")
                        status_text.text(f"⏳ 期货下载中… {completed}/{len(missing_futures)} — {ct}: {msg}")

                    sync_all_contracts(max_workers=_sync_parallel_workers,
                                      progress_callback=_fut_progress)
                    completed = len(missing_futures)
                    progress_bar.progress(completed / total_tasks)
                    status_text.text(f"✅ 期货数据下载完成！{len(missing_futures)} 个合约已缓存。")

                # ── Step 2: 并行下载所有缺失的净持仓 ──
                all_with_futures_now = [c for c in ALL_CONTRACTS if _csv_path(c).exists()]
                missing_net = [c for c in all_with_futures_now if not _net_agg_path(c).exists()
                               or _load_aggregated_net(c) is None]
                total_tasks = len(missing_futures) + len(missing_net)
                if total_tasks > 0:
                    progress_bar.progress(completed / total_tasks)
                if missing_net:
                    status_text.text(f"⏳ 正在并行下载 {len(missing_net)} 个合约的净持仓数据（4 线程）…")

                    def _net_progress(current, total, ct, msg):
                        nonlocal completed
                        completed = len(missing_futures) + current
                        progress_bar.progress(min(completed / total_tasks, 1.0),
                            text=f"📊 净持仓 {current}/{total} — {ct}: {msg}")
                        status_text.text(f"⏳ 净持仓下载中… {current}/{len(missing_net)} — {ct}: {msg}")

                    sync_all_net_holdings(max_workers=4, progress_callback=_net_progress)
                    completed = total_tasks
                    progress_bar.progress(1.0)
                    _build_seasonal_net_positions.clear()

                status_text.text(
                    f"✅ 全部完成！期货 {len(missing_futures)} 个 + 净持仓 {len(missing_net)} 个合约数据已就绪。"
                )
            placeholder.empty()
        st.session_state["_startup_all_synced"] = True

    # ── 自动增量同步：每天首次打开时快速检查并拉取最新期货数据 ──
    if "_last_auto_sync_date" not in st.session_state:
        st.session_state["_last_auto_sync_date"] = None

    _today_str = datetime.now().strftime("%Y%m%d")
    if st.session_state["_last_auto_sync_date"] != _today_str:
        # 只增量同步有期货 CSV 的活跃合约（已是最新时 <0.1s 秒过）
        for ct in get_active_contracts():
            cp = _csv_path(ct)
            if cp.exists():
                try:
                    df = pd.read_csv(cp, usecols=["date"])
                    if not df.empty:
                        last_date = pd.to_datetime(df["date"].max()).date()
                        # 数据已经是今天或周末/假日（最近2天内）→ 跳过
                        if last_date >= (datetime.now().date() - timedelta(days=2)):
                            continue
                except Exception:
                    pass
                sync_futures(ct, force_full=False)
        st.session_state["_last_auto_sync_date"] = _today_str

    # ── 八个 Tab ──
    t1, t2, t3, t4, t5, t6, t7, t8 = st.tabs([
        "📋 每日期货分析日报",     # Tab 1 (新增)
        "📊 当日基差分布",         # Tab 2 (原 Tab 1)
        "📈 单合约基差走势",       # Tab 3 (原 Tab 2)
        "🔄 合约基差比较",         # Tab 4 (原 Tab 3)
        "📉 合约价差比较",         # Tab 5 (原 Tab 4)
        "📊 持仓与成交分析",       # Tab 6 (原 Tab 5)
        "📅 季节性持仓对比",       # Tab 7 (原 Tab 6)
        "📉 技术分析",             # Tab 8 (原 Tab 7)
    ])

    with t1: tab_daily_report()
    with t2: tab1()
    with t3: tab2()
    with t4: tab3()
    with t5: tab4()
    with t6: tab5()
    with t7: tab6()
    with t8: tab7()

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
            st.cache_data.clear()
            st.session_state["_startup_all_synced"] = False
            st.session_state["_last_auto_sync_date"] = datetime.now().strftime("%Y%m%d")
            with st.spinner("🔄 正在更新活跃合约最新数据…"):
                # 只增量更新活跃合约（已是最新时秒过，<1秒）
                for ct in get_active_contracts():
                    sync_futures(ct, force_full=False)
                    if _csv_path(ct).exists():
                        sync_net_holdings(ct, force_full=False)
                _build_seasonal_net_positions.clear()
            st.rerun()
    with c2:
        if st.button("🗑️ 清除缓存", use_container_width=True, key="main_clear"):
            st.cache_data.clear()
            st.session_state["_startup_all_synced"] = False
            if FUTURES_DIR.exists(): shutil.rmtree(FUTURES_DIR); FUTURES_DIR.mkdir()
            HOLDINGS_DIR.mkdir(exist_ok=True)
            st.rerun()

    st.caption("⚠️ 免责声明：本平台数据仅供参考，不构成任何投资建议。投资有风险，入市需谨慎。")
    st.markdown("<p style='text-align:right; font-size:0.75rem; color:#adb5bd; margin-top:-8px;'>创作者：chen</p>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
