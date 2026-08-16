#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
涌益咨询生猪周度数据 - 在线可编辑周报系统
=============================================
功能:
1. 自动读取涌益咨询三个数据文件
2. 生成专业图表(季节性同比/双轴图)
3. AI自动分析,按板块输出结论
4. 在线可编辑报告,支持PDF导出
5. 模板化设计,替换数据即可重用

启动: python generate_report.py
访问: http://localhost:8051
"""

import dash
from dash import dcc, html, Input, Output, State, callback_context
import dash_bootstrap_components as dbc
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.express as px
import pandas as pd
import numpy as np
import os
import re
import sys
import json
import warnings
from datetime import datetime, timedelta
from collections import defaultdict

warnings.filterwarnings('ignore')

# ============================================================
# 全局配置
# ============================================================

# 数据目录 - 按优先级扫描：D盘桌面(新数据) → 数据源文件夹(旧数据兜底)
DATA_DIRS = [
    r'D:\CC\Desktop',
    os.path.join(os.path.dirname(os.path.abspath(__file__)), '数据源', '涌益生猪项目数据库'),
]

# 年份颜色 (用户指定)
YR_COLORS = {
    2026: '#E31A1C',  # 红
    2025: '#33A02C',  # 绿
    2024: '#1F1F1F',  # 黑
    2023: '#1F78B4',  # 蓝
    2022: '#FFD700',  # 黄
    2021: '#6A3D9A',  # 紫
}

# 四省 + 广西
TARGET_PROVINCES = ['河南', '四川', '广东', '辽宁', '广西']
PROV_COLORS = {'河南': '#E31A1C', '四川': '#33A02C', '广东': '#1F78B4', '辽宁': '#FF7F00', '广西': '#8E44AD'}

# 报告期
REPORT_PERIOD = ""


# ============================================================
# 数据读取器
# ============================================================

class DataReader:
    """统一数据读取器"""

    def __init__(self, data_dir):
        self.dirs = data_dir if isinstance(data_dir, list) else [data_dir]
        self.file_ridu = None      # 方便拉图表的数据库 (日度数据)
        self.file_zhoudou = None   # 涌益咨询 周度数据
        self.file_tubiao = None    # 涌益咨询 周度图表版
        self._find_files()

    def _find_files(self):
        """自动识别三个文件（多目录扫描，D盘桌面优先）"""
        global REPORT_PERIOD
        REPORT_PERIOD = ""
        for d in self.dirs:
            if not os.path.isdir(d):
                continue
            for f in os.listdir(d):
                if f.startswith('~$') or not f.endswith('.xlsx'):
                    continue
                full = os.path.join(d, f)
                if self.file_ridu is None and ('拉图表' in f or '日度数据' in f):
                    self.file_ridu = full
                elif self.file_zhoudou is None and '周度数据' in f and '图表' not in f:
                    self.file_zhoudou = full
                elif self.file_tubiao is None and '图表版' in f:
                    self.file_tubiao = full
            if not REPORT_PERIOD:
                for f in os.listdir(d):
                    m = re.match(r'(\d{4}\.\d{1,2}\.\d{1,2}-\d{4}\.\d{1,2}\.\d{1,2})', f)
                    if m:
                        REPORT_PERIOD = m.group(1)
                        break
        if not REPORT_PERIOD:
            REPORT_PERIOD = datetime.now().strftime('%Y.%m.%d-%Y.%m.%d')

        print(f"数据文件:")
        print(f"  日度数据: {os.path.basename(self.file_ridu) if self.file_ridu else '未找到'}")
        print(f"  周度数据: {os.path.basename(self.file_zhoudou) if self.file_zhoudou else '未找到'}")
        print(f"  图表版:   {os.path.basename(self.file_tubiao) if self.file_tubiao else '未找到'}")

    # ==================== 通用工具方法 ====================

    def _read_sheet(self, file_path, sheet_idx):
        """读取sheet, 返回header=None的DataFrame"""
        return pd.read_excel(file_path, sheet_name=sheet_idx, header=None)

    def _col_index(self, df, row_idx, keyword):
        """在指定行中查找包含关键词的列索引"""
        for i, v in enumerate(df.iloc[row_idx].values):
            if pd.notna(v) and keyword in str(v):
                return i
        return None

    def _parse_range(self, val_str):
        """解析区间值(如'5.2-5.6')为均值"""
        if pd.isna(val_str):
            return np.nan
        s = str(val_str).strip()
        # 匹配 a-b 或 a - b
        m = re.match(r'([\d.]+)\s*[-~]\s*([\d.]+)', s)
        if m:
            return (float(m.group(1)) + float(m.group(2))) / 2
        try:
            return float(s)
        except:
            return np.nan

    def _sheet_index(self, file_path, keyword):
        """按 sheet 名关键词返回索引，找不到返回 None"""
        try:
            xl = pd.ExcelFile(file_path)
            for i, name in enumerate(xl.sheet_names):
                if keyword in name:
                    return i
        except Exception:
            pass
        return None

    def _header_row(self, df, keywords):
        """找表头行（含关键词的行）"""
        for r in range(min(8, len(df))):
            if any(pd.notna(v) and any(kw in str(v) for kw in keywords) for v in df.iloc[r].values):
                return r
        return 0

    def _col_by_name(self, df, hdr_row, keyword):
        """在表头行找包含关键词的列索引，找不到返回 None"""
        for i, h in enumerate(df.iloc[hdr_row].values):
            if pd.notna(h) and keyword in str(h):
                return i
        return None

    def _wide_timeseries(self, df, prov, val_shift=0):
        """宽表(省份行 × 日期列) → {date, value}。日期列自动检测，值列 = 日期列 + val_shift"""
        hdr = 0
        for r in range(min(3, len(df))):
            if any(pd.notna(v) and isinstance(v, (datetime, pd.Timestamp)) for v in df.iloc[r].values):
                hdr = r
                break
        dates = df.iloc[hdr].values
        prov_row = None
        for r in range(hdr + 1, len(df)):
            if pd.notna(df.iloc[r, 0]) and prov in str(df.iloc[r, 0]):
                prov_row = r
                break
        if prov_row is None:
            return pd.DataFrame(columns=['date', 'value'])
        out = []
        for c in range(1, len(dates)):
            d = dates[c]
            if not (pd.notna(d) and isinstance(d, (datetime, pd.Timestamp))):
                continue
            vc = c + val_shift
            v = df.iloc[prov_row, vc] if vc < len(dates) else np.nan
            try:
                dt = pd.to_datetime(d) if not isinstance(d, datetime) else d
                pv = self._parse_range(v)
                if not np.isnan(pv):
                    out.append({'date': dt, 'value': pv})
            except Exception:
                pass
        if not out:
            return pd.DataFrame(columns=['date', 'value'])
        return pd.DataFrame(out).drop_duplicates('date').sort_values('date')

    def _extract_timeseries(self, df, date_col=0, val_col=3, date_row_start=3):
        """提取通用时间序列 {date: value}"""
        dates, vals = [], []
        for r in range(date_row_start, len(df)):
            d = df.iloc[r, date_col]
            v = df.iloc[r, val_col]
            if pd.notna(d):
                try:
                    dt = pd.to_datetime(d) if not isinstance(d, datetime) else d
                    pv = self._parse_range(v)
                    if not np.isnan(pv):
                        dates.append(dt)
                        vals.append(pv)
                except:
                    pass
        return pd.DataFrame({'date': dates, 'value': vals}).drop_duplicates('date').sort_values('date')

    def _yearly_groups(self, df):
        """按年份分组, 返回 [(year, df), ...]"""
        if df is None or df.empty:
            return []
        df = df.copy()
        df['year'] = df['date'].dt.year
        result = []
        for yr in sorted(df['year'].unique()):
            if 2021 <= yr <= 2026:
                result.append((yr, df[df['year'] == yr].sort_values('date')))
        return result

    def _latest_value(self, df):
        """获取最新值"""
        if df is None or df.empty:
            return np.nan
        return df.sort_values('date')['value'].iloc[-1]

    def _wow_yoy(self, df, weeks_ago=1, weeks_yoy=52):
        """计算环比和同比变化"""
        if df is None or df.empty:
            return (np.nan, np.nan, np.nan, np.nan)
        df = df.sort_values('date')
        cur = df['value'].iloc[-1]
        wow_val = df['value'].iloc[-1-weeks_ago] if len(df) > weeks_ago else np.nan
        yoy_val = df['value'].iloc[-1-weeks_yoy] if len(df) > weeks_yoy else np.nan
        wow_chg = cur - wow_val if not np.isnan(wow_val) else np.nan
        wow_pct = (cur / wow_val - 1) if wow_val and wow_val != 0 else np.nan
        yoy_chg = cur - yoy_val if not np.isnan(yoy_val) else np.nan
        yoy_pct = (cur / yoy_val - 1) if yoy_val and yoy_val != 0 else np.nan
        return wow_chg, wow_pct, yoy_chg, yoy_pct

    # ==================== (1) 日度数据读取 ====================

    def get_province_price(self):
        """各省均价 - 兼容 方便拉图表(日_日价_均价) 与 原始日度(各省份均价)"""
        if not self.file_ridu:
            return {}
        idx = self._sheet_index(self.file_ridu, '省份均价')
        if idx is None:
            idx = self._sheet_index(self.file_ridu, '日价_均价')
        if idx is None:
            return {}
        df = self._read_sheet(self.file_ridu, idx)
        hdr = self._header_row(df, ['河南'])
        headers = df.iloc[hdr].values
        prov_cols = {}
        for i, h in enumerate(headers):
            if pd.notna(h):
                for p in TARGET_PROVINCES:
                    if p in str(h):
                        prov_cols[p] = i
                        break
        # 已知映射兜底: 河南=1, 辽宁=9, 四川=12, 广东=14, 广西=15
        if not prov_cols:
            prov_cols = {'河南': 1, '辽宁': 9, '四川': 12, '广东': 14, '广西': 15}
        ds = hdr + 1
        result = {}
        for p in TARGET_PROVINCES:
            if p in prov_cols:
                col = prov_cols[p]
                result[p] = self._extract_timeseries(df, val_col=col, date_row_start=ds).rename(columns={'value': 'price'})
        return result

    def get_national_price(self):
        """全国均价 - 直接读取 价格+宰量 的 全国均价 列"""
        if not self.file_ridu:
            return None
        idx = self._sheet_index(self.file_ridu, '价格+宰量')
        if idx is None:
            idx = self._sheet_index(self.file_ridu, '价格+宰量')
        if idx is None:
            return None
        df = self._read_sheet(self.file_ridu, idx)
        hdr = self._header_row(df, ['全国均价', '均价'])
        headers = df.iloc[hdr].values
        val_col = 1
        for i, h in enumerate(headers):
            if pd.notna(h) and '均价' in str(h):
                val_col = i
                break
        return self._extract_timeseries(df, val_col=val_col, date_row_start=hdr + 1)

    def get_province_slaughter(self):
        """各省屠宰量 - 兼容 宽表(屠宰企业日度屠宰量) 与 长表(日_屠宰量)"""
        if not self.file_ridu:
            return None
        idx = self._sheet_index(self.file_ridu, '屠宰企业日度屠宰量')
        if idx is not None:
            df = self._read_sheet(self.file_ridu, idx)
            result = {}
            for p in TARGET_PROVINCES:
                ts = self._wide_timeseries(df, p)
                if not ts.empty:
                    result[p] = ts
            return result
        idx = self._sheet_index(self.file_ridu, '日_屠宰量')
        if idx is None:
            return None
        df = self._read_sheet(self.file_ridu, idx)
        hdr = self._header_row(df, ['河南'])
        headers = df.iloc[hdr].values
        prov_cols = {}
        for i, h in enumerate(headers):
            if pd.notna(h):
                for p in TARGET_PROVINCES:
                    if p in str(h):
                        prov_cols[p] = i
                        break
        result = {}
        for p in TARGET_PROVINCES:
            if p in prov_cols:
                result[p] = self._extract_timeseries(df, val_col=prov_cols[p], date_row_start=hdr + 1)
        return result

    def get_fat_standard_spread(self):
        """散户肥标价差 - 兼容 宽表(散户标肥价差) 与 长表(日_散户标肥价差_*)"""
        result = {}
        if not self.file_ridu:
            return result
        idx = self._sheet_index(self.file_ridu, '散户标肥价差')
        if idx is not None:
            df = self._read_sheet(self.file_ridu, idx)
            result['150kg肥标价差'] = self._wide_timeseries(df, '河南', val_shift=1)
            result['175kg肥标价差'] = self._wide_timeseries(df, '河南', val_shift=2)
            return result
        for key, kw in [('150kg肥标价差', '150公斤'), ('175kg肥标价差', '175公斤')]:
            i = self._sheet_index(self.file_ridu, kw)
            if i is not None:
                d = self._read_sheet(self.file_ridu, i)
                result[key] = self._extract_timeseries(d, val_col=3)
        return result

    # ==================== (2) 周度数据读取 ====================

    def get_weight_split(self):
        """体重拆分(集团+散户)"""
        if not self.file_zhoudou:
            return {}
        idx = self._sheet_index(self.file_zhoudou, '体重拆分')
        if idx is None:
            return {}
        df = self._read_sheet(self.file_zhoudou, idx)
        hdr = self._header_row(df, ['全国均重'])
        result = {}
        for kw, label in [('全国均重', '全国平均'), ('集团', '集团'), ('散户', '散户')]:
            col = self._col_by_name(df, hdr, kw)
            if col is None:
                col = {'全国平均': 1, '集团': 2, '散户': 3}[label]
            result[label] = self._extract_timeseries(df, val_col=col, date_col=0, date_row_start=hdr + 1)
        return result

    def _wide_mean_timeseries(self, df):
        """宽表(省份行 × 日期列) → 全国算术平均时间序列（仅日期列，跳过序号/省份/出栏基准列）"""
        hdr = 0
        for r in range(min(3, len(df))):
            if any(pd.notna(v) and isinstance(v, (datetime, pd.Timestamp)) for v in df.iloc[r].values):
                hdr = r
                break
        dates = df.iloc[hdr].values
        date_cols = [c for c in range(len(dates)) if pd.notna(dates[c]) and isinstance(dates[c], (datetime, pd.Timestamp))]
        rows = []
        for c in date_cols:
            vals = []
            for r in range(hdr + 1, len(df)):
                try:
                    v = self._parse_range(df.iloc[r, c])
                    if not np.isnan(v) and v > 0:
                        vals.append(v)
                except Exception:
                    pass
            if vals:
                rows.append({'date': pd.to_datetime(dates[c]), 'value': np.mean(vals)})
        if not rows:
            return None
        return pd.DataFrame(rows).sort_values('date')

    def get_eryu_rate(self):
        """二育栏舍利用率 - 全国各省算术平均"""
        if not self.file_zhoudou:
            return None
        idx = self._sheet_index(self.file_zhoudou, '二育栏舍利用率')
        if idx is None:
            return None
        df = self._read_sheet(self.file_zhoudou, idx)
        return self._wide_mean_timeseries(df)

    def _province_series(self, sheet_keyword, prov='河南'):
        """读周度数据的某省份时间序列（按省份名找列，兼容不同sheet顺序）"""
        if not self.file_zhoudou:
            return None
        idx = self._sheet_index(self.file_zhoudou, sheet_keyword)
        if idx is None:
            return None
        df = self._read_sheet(self.file_zhoudou, idx)
        hdr = self._header_row(df, [prov, '结束日期', '日期'])
        col = self._col_by_name(df, hdr, prov)
        if col is None:
            col = 2
        date_col = 1 if self._col_by_name(df, hdr, '结束日期') is not None else 0
        return self._extract_timeseries(df, date_col=date_col, val_col=col, date_row_start=hdr + 1)

    def get_fresh_sale_rate(self):
        """鲜销率 - 全国均值"""
        if not self.file_zhoudou:
            return None
        idx = self._sheet_index(self.file_zhoudou, '鲜销率')
        if idx is None:
            return None
        df = self._read_sheet(self.file_zhoudou, idx)
        hdr = self._header_row(df, ['全国'])
        col = self._col_by_name(df, hdr, '全国')
        if col is None:
            col = df.shape[1] - 1
        return self._extract_timeseries(df, val_col=col, date_col=1, date_row_start=hdr + 1)

    def get_frozen_stock(self):
        """冻品库存率 - 冻品库存多样本全国均值"""
        if not self.file_zhoudou:
            return None
        idx = self._sheet_index(self.file_zhoudou, '冻品库存多样本')
        if idx is None:
            idx = self._sheet_index(self.file_zhoudou, '冻品库存')
        if idx is None:
            return None
        df = self._read_sheet(self.file_zhoudou, idx)
        hdr = self._header_row(df, ['全国'])
        col = self._col_by_name(df, hdr, '全国')
        if col is None:
            col = df.shape[1] - 1
        return self._extract_timeseries(df, val_col=col, date_col=0, date_row_start=hdr + 1)

    def get_slaughter_profit(self):
        """屠宰利润（白条头均利润）"""
        if not self.file_zhoudou:
            return None
        idx = self._sheet_index(self.file_zhoudou, '河南屠宰白条成本')
        if idx is None:
            return None
        df = self._read_sheet(self.file_zhoudou, idx)
        hdr = self._header_row(df, ['白条头均利润'])
        col = self._col_by_name(df, hdr, '白条头均利润')
        if col is None:
            col = 9
        return self._extract_timeseries(df, val_col=col, date_row_start=hdr + 1)

    def get_cull_sow_price(self):
        """淘汰母猪价格（河南）"""
        return self._province_series('淘汰母猪价格', '河南')

    def get_high_parity_discount(self):
        """高胎母猪折扣（河南）"""
        return self._province_series('高胎淘母折扣', '河南')

    def get_low_parity_discount(self):
        """低胎母猪折扣（河南）"""
        return self._province_series('低胎母猪折扣', '河南')

    def get_binary_sow_price(self):
        """二元母猪价格（河南）"""
        return self._province_series('二元母猪价格', '河南')

    # ==================== (3) 图表版数据读取 ====================

    def get_piglet_15kg(self):
        """15公斤仔猪 - File 3 [9]"""
        if not self.file_tubiao:
            return None
        df = self._read_sheet(self.file_tubiao, 9)
        return self._extract_timeseries(df, date_col=0, val_col=1, date_row_start=2)

    def get_piglet_weaned(self):
        """断奶仔猪 - File 3 [10]"""
        if not self.file_tubiao:
            return None
        df = self._read_sheet(self.file_tubiao, 10)
        return self._extract_timeseries(df, date_col=0, val_col=1, date_row_start=2)

    def get_piglet_sale_profit(self):
        """销售仔猪头均利润 - File 2 [6] 仔猪与商品猪利润对比 col 2"""
        if not self.file_zhoudou:
            return None
        df = self._read_sheet(self.file_zhoudou, 6)
        dates, vals = [], []
        for r in range(3, len(df)):
            d = df.iloc[r, 0]
            v = df.iloc[r, 2]  # Column 2 = 销售仔猪头均利润
            if pd.notna(d) and pd.notna(v):
                try:
                    ds = str(d).strip()
                    # 解析日期范围如 "2022.9.30-2022.10.8", 取结束日期
                    if '-' in ds and '.' in ds:
                        parts = ds.split('-')
                        end_date_str = parts[-1].strip()
                        # Handle "2022.10.8" format
                        dt = pd.to_datetime(end_date_str, format='%Y.%m.%d')
                    else:
                        dt = pd.to_datetime(d)
                    pv = self._parse_range(v)
                    if not np.isnan(pv):
                        dates.append(dt)
                        vals.append(pv)
                except:
                    pass
        return pd.DataFrame({'date': dates, 'value': vals}).sort_values('date')

    def get_maobai_spread(self):
        """毛白价差"""
        if not self.file_zhoudou:
            return None
        idx = self._sheet_index(self.file_zhoudou, '毛白价差')
        if idx is None:
            return None
        df = self._read_sheet(self.file_zhoudou, idx)
        hdr = self._header_row(df, ['价差'])
        # 毛白价差 sheet 为简单表：价差在最后一列（第4列）
        col = self._col_by_name(df, hdr, '价差')
        if col is None or col == 0:
            col = 3
        return self._extract_timeseries(df, val_col=col, date_col=0, date_row_start=hdr + 1)

    def get_breeding_profit(self):
        """养殖利润 - 返回{label: df}"""
        if not self.file_zhoudou:
            return {}
        idx = self._sheet_index(self.file_zhoudou, '养殖利润最新')
        if idx is None:
            idx = self._sheet_index(self.file_zhoudou, '养殖利润')
        if idx is None:
            return {}
        df = self._read_sheet(self.file_zhoudou, idx)
        hdr = self._header_row(df, ['母猪', '利润', '外购'])
        result = {}
        for kw, label in [('母猪50头以下', '母猪50头以下'), ('5000-10000', '5000-10000头'), ('外购', '外购仔猪育肥')]:
            col = self._col_by_name(df, hdr, kw)
            if col is not None:
                result[label] = self._extract_timeseries(df, val_col=col, date_col=1, date_row_start=hdr + 1)
        return result

    # ==================== 批量加载 ====================

    def load_all(self):
        """批量加载全部数据, 返回数据字典"""
        print("\n开始加载数据...")
        data = {}

        # (1) 日度数据
        print("  [1/3] 日度数据...")
        data['province_price'] = self.get_province_price()
        data['national_price'] = self.get_national_price()
        data['province_slaughter'] = self.get_province_slaughter()
        data['fat_std_spread'] = self.get_fat_standard_spread()

        # 从日度数据获取最新日期作为报告截止日期
        global REPORT_PERIOD
        if data.get('national_price') is not None and not data['national_price'].empty:
            latest_dt = data['national_price']['date'].max()
            REPORT_PERIOD = latest_dt.strftime('%Y.%m.%d') if hasattr(latest_dt, 'strftime') else str(latest_dt)[:10]
        elif data.get('province_price'):
            for p, df in data['province_price'].items():
                if df is not None and not df.empty:
                    latest_dt = df['date'].max() if 'date' in df.columns else df.index.max()
                    REPORT_PERIOD = str(latest_dt)[:10]
                    break

        # (2) 周度数据
        print("  [2/3] 周度数据...")
        data['weight_split'] = self.get_weight_split()
        data['eryu_rate'] = self.get_eryu_rate()
        data['fresh_sale_rate'] = self.get_fresh_sale_rate()
        data['frozen_stock'] = self.get_frozen_stock()
        data['slaughter_profit'] = self.get_slaughter_profit()
        data['cull_sow_price'] = self.get_cull_sow_price()
        data['high_parity_disc'] = self.get_high_parity_discount()
        data['low_parity_disc'] = self.get_low_parity_discount()
        data['binary_sow_price'] = self.get_binary_sow_price()

        # (3) 图表版+周度补充
        print("  [3/3] 图表版+补充...")
        data['piglet_15kg'] = self.get_piglet_15kg()
        data['piglet_weaned'] = self.get_piglet_weaned()
        data['piglet_sale_profit'] = self.get_piglet_sale_profit()
        data['maobai_spread'] = self.get_maobai_spread()
        data['breeding_profit'] = self.get_breeding_profit()

        # 打印加载统计
        for k, v in data.items():
            if isinstance(v, dict):
                for k2, v2 in v.items():
                    if v2 is not None and not v2.empty:
                        print(f"    {k}.{k2}: {len(v2)}条")
            elif v is not None and not v.empty:
                print(f"    {k}: {len(v)}条")

        return data


# ============================================================
# 图表生成器
# ============================================================

class ChartBuilder:

    @staticmethod
    def seasonal_yoy(data_dict, title, ylabel, is_pct=False):
        """季节性同比图 (多年度周次对比)"""
        fig = go.Figure()
        for yr, df in data_dict:
            if df is None or df.empty:
                continue
            df = df.sort_values('date').copy()
            df['week'] = range(1, len(df) + 1)
            color = YR_COLORS.get(yr, '#888')
            fig.add_trace(go.Scatter(
                x=df['week'], y=df['value'],
                mode='lines', name=f'{yr}年',
                line=dict(color=color, width=2),
                hovertemplate=f'{yr}年 第%{{x}}周<br>值: %{{y:.2f}}<extra></extra>',
            ))
            # 标注最新值
            if len(df) > 0:
                fig.add_annotation(
                    x=df['week'].iloc[-1], y=df['value'].iloc[-1],
                    text=f"{df['value'].iloc[-1]:.2f}", showarrow=False,
                    font=dict(color=color, size=10), xshift=15,
                )
        fig.update_layout(
            template='plotly_white', title=title,
            xaxis_title='周次', yaxis_title=ylabel,
            hovermode='x unified', height=450,
            legend=dict(orientation='h', y=1.05, x=0),
        )
        if is_pct:
            fig.update_yaxes(tickformat='.1%')
        return fig

    @staticmethod
    def dual_axis(price_data, slaughter_data, province):
        """双轴图: 价格(右轴) + 屠宰量(左轴)"""
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        color_p = PROV_COLORS.get(province, '#E31A1C')

        # 屠宰量 (左轴)
        if province in slaughter_data and not slaughter_data[province].empty:
            s_df = slaughter_data[province].sort_values('date')
            fig.add_trace(go.Bar(
                x=s_df['date'], y=s_df['value'], name=f'{province}屠宰量',
                marker_color=color_p, opacity=0.6, yaxis='y1',
            ), secondary_y=False)

        # 价格 (右轴)
        if province in price_data and not price_data[province].empty:
            p_df = price_data[province].sort_values('date')
            fig.add_trace(go.Scatter(
                x=p_df['date'], y=p_df['price'], name=f'{province}价格',
                line=dict(color='red', width=2.5), yaxis='y2',
            ), secondary_y=True)

        fig.update_layout(
            template='plotly_white',
            title=f'{province} - 屠宰量与价格双轴图',
            hovermode='x unified', height=400,
            legend=dict(orientation='h', y=1.1),
        )
        fig.update_yaxes(title_text='屠宰量 (头/日)', secondary_y=False)
        fig.update_yaxes(title_text='价格 (元/公斤)', secondary_y=True, color='red')
        return fig

    @staticmethod
    def single_timeseries(df, title, ylabel, is_pct=False):
        """单一时间序列图"""
        if df is None or df.empty:
            return go.Figure()
        df = df.sort_values('date')
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=df['date'], y=df['value'], mode='lines+markers',
            line=dict(color='#E31A1C', width=2),
            marker=dict(size=4), name=title,
        ))
        fig.update_layout(
            template='plotly_white', title=title,
            xaxis_title='日期', yaxis_title=ylabel,
            hovermode='x unified', height=400,
        )
        if is_pct:
            fig.update_yaxes(tickformat='.1%')
        return fig

    @staticmethod
    def profit_comparison(profit_dict):
        """养殖利润三合一对比图"""
        fig = go.Figure()
        colors = ['#E31A1C', '#1F78B4', '#33A02C']
        for i, (label, df) in enumerate(profit_dict.items()):
            if df is not None and not df.empty:
                df = df.sort_values('date')
                fig.add_trace(go.Scatter(
                    x=df['date'], y=df['value'], mode='lines',
                    name=label, line=dict(color=colors[i % 3], width=2),
                ))
        fig.add_hline(y=0, line_dash="dash", line_color="gray", opacity=0.5)
        fig.update_layout(
            template='plotly_white', title='养殖利润对比 (元/头)',
            xaxis_title='日期', yaxis_title='利润 (元/头)',
            hovermode='x unified', height=400,
            legend=dict(orientation='h', y=1.05),
        )
        return fig


# ============================================================
# AI 分析引擎
# ============================================================

class Analyzer:
    """自动分析引擎 - 以猪价分析师角度输出结论"""

    @staticmethod
    def _val_col(df):
        """获取DataFrame的值列名"""
        if df is None or df.empty:
            return 'value'
        return 'price' if 'price' in df.columns else 'value'

    @staticmethod
    def _fmt(val, unit='', is_pct=False):
        """格式化数值"""
        if pd.isna(val):
            return 'N/A'
        if is_pct:
            return f'{val*100:+.1f}%'
        return f'{val:+.2f}{unit}'

    @staticmethod
    def _week_period_str(df):
        """返回 (本周自然周日期串, 上周日期串, 去年同周日期串)"""
        if df is None or df.empty:
            return ('—', '—', '—')
        df = df.sort_values('date')
        latest = df['date'].iloc[-1]
        monday = latest - timedelta(days=latest.weekday())
        sunday = monday + timedelta(days=6)
        cur = f"{monday.month}月{monday.day}日-{sunday.month}月{sunday.day}日"
        pm, ps = monday - timedelta(days=7), monday - timedelta(days=1)
        prev = f"{pm.month}月{pm.day}日-{ps.month}月{ps.day}日"
        ly = monday - timedelta(days=364)
        ly2 = ly + timedelta(days=6)
        last = f"{ly.month}月{ly.day}日-{ly2.month}月{ly2.day}日"
        return (cur, prev, last)

    @staticmethod
    def analyze_price(data):
        """价格分析（自然周均价，标注日期，价差只在图中展示）"""
        prov_price = data.get('province_price', {})
        national = data.get('national_price')
        lines = []
        lines.append("【价格分析】")

        # 各省自然周均价（标注自然周日期/上周/去年同期）
        for p in TARGET_PROVINCES:
            df = prov_price.get(p)
            if df is not None and not df.empty:
                cur_s, prev_s, ly_s = Analyzer._week_period_str(df)
                m = Analyzer._natural_week_mean(df, 'price')
                cmp = Analyzer._week_avg_compare(df, 'price')
                if m is not None and cmp:
                    wow_str = Analyzer._fmt(cmp[3], is_pct=True) if cmp[3] is not None else 'N/A'
                    yoy_str = Analyzer._fmt(cmp[4], is_pct=True) if cmp[4] is not None else 'N/A'
                    prev_mean = f"{cmp[1]:.2f}" if cmp[1] is not None else '—'
                    ly_mean = f"{cmp[2]:.2f}" if cmp[2] is not None else '—'
                    lines.append(f"- {p}自然周均价({cur_s}): {m:.2f}元/kg，"
                                 f"上周({prev_s}){prev_mean}元/kg(环比{wow_str})，"
                                 f"去年同期({ly_s}){ly_mean}元/kg(同比{yoy_str})")

        # 全国均价自然周
        if national is not None and not national.empty:
            cur_s, prev_s, ly_s = Analyzer._week_period_str(national)
            m = Analyzer._natural_week_mean(national, 'value')
            cmp = Analyzer._week_avg_compare(national, 'value')
            if m is not None:
                wow_str = Analyzer._fmt(cmp[3], is_pct=True) if cmp and cmp[3] is not None else 'N/A'
                yoy_str = Analyzer._fmt(cmp[4], is_pct=True) if cmp and cmp[4] is not None else 'N/A'
                lines.append(f"- 全国均价自然周({cur_s}): {m:.2f}元/kg，环比{wow_str}，同比{yoy_str}")

        # 肥标价差
        fss = data.get('fat_std_spread', {})
        if isinstance(fss, dict):
            for key in ['150kg肥标价差', '175kg肥标价差']:
                df = fss.get(key)
                if df is not None and not df.empty:
                    latest = df.sort_values('date')['value'].iloc[-1]
                    lines.append(f"- 散户{key}: {latest:.2f}元/kg")

        return '\n'.join(lines)

    @staticmethod
    def analyze_supply(data):
        """供给分析"""
        lines = []
        lines.append("【供给分析】")

        ws = data.get('weight_split', {})
        for key in ['集团', '散户', '全国平均']:
            df = ws.get(key)
            if df is not None and not df.empty:
                wow_chg, wow_pct, yoy_chg, yoy_pct = Analyzer._wow_yoy(df)
                latest = df.sort_values('date')['value'].iloc[-1]
                lines.append(f"- {key}出栏体重: {latest:.1f}kg, "
                           f"环比{Analyzer._fmt(wow_chg, 'kg')}, 同比{Analyzer._fmt(yoy_chg, 'kg')}")

        eryu = data.get('eryu_rate')
        if eryu is not None and not eryu.empty:
            latest = eryu.sort_values('date')['value'].iloc[-1]
            wow_chg, wow_pct, yoy_chg, yoy_pct = Analyzer._wow_yoy(eryu)
            lines.append(f"- 二育栏舍利用率(全国均值): {latest:.1%}, "
                       f"环比{Analyzer._fmt(wow_pct, is_pct=True)}")

        return '\n'.join(lines)

    @staticmethod
    def analyze_slaughter(data):
        """屠宰分析（河南周度屠宰量日均对比 + 其余省份）"""
        lines = []
        lines.append("【屠宰分析】")

        prov_sl = data.get('province_slaughter', {})

        # 河南屠宰量：自然周日均，对比上周日均、去年同周日均
        hn = prov_sl.get('河南')
        if hn is not None and not hn.empty:
            cmp = Analyzer._week_avg_compare(hn, 'value')
            if cmp:
                cur, prev, ly, wow, yoy = cmp
                wow_str = Analyzer._fmt(wow, is_pct=True) if wow is not None else 'N/A'
                yoy_str = Analyzer._fmt(yoy, is_pct=True) if yoy is not None else 'N/A'
                lines.append(f"- 河南周度屠宰量日均: {cur:.0f}头/日, "
                             f"上周日均{prev:.0f}头/日(环比{wow_str}), "
                             f"去年同周日均{ly:.0f}头/日(同比{yoy_str})")

        # 其余省份（自然周日均）
        for p in TARGET_PROVINCES:
            if p == '河南':
                continue
            df = prov_sl.get(p)
            if df is not None and not df.empty:
                m = Analyzer._natural_week_mean(df, 'value')
                if m is not None:
                    lines.append(f"- {p}屠宰量日均: {m:.0f}头/日")

        # 鲜销率
        fresh = data.get('fresh_sale_rate')
        if fresh is not None and not fresh.empty:
            latest = fresh.sort_values('date')['value'].iloc[-1]
            wow_chg, wow_pct, yoy_chg, yoy_pct = Analyzer._wow_yoy(fresh)
            lines.append(f"- 全国屠宰鲜销率: {latest:.1%}, "
                       f"环比{Analyzer._fmt(wow_pct, is_pct=True)}")

        # 冻品库存
        frozen = data.get('frozen_stock')
        if frozen is not None and not frozen.empty:
            latest = frozen.sort_values('date')['value'].iloc[-1]
            wow_chg, wow_pct, yoy_chg, yoy_pct = Analyzer._wow_yoy(frozen)
            lines.append(f"- 全国冻品库存率: {latest:.1%}, "
                       f"环比{Analyzer._fmt(wow_pct, is_pct=True)}")

        # 屠宰利润
        profit = data.get('slaughter_profit')
        if profit is not None and not profit.empty:
            latest = profit.sort_values('date')['value'].iloc[-1]
            lines.append(f"- 河南白条头均利润: {latest:.1f}元/头")

        return '\n'.join(lines)

    @staticmethod
    def analyze_sow(data):
        """母猪分析"""
        lines = []
        lines.append("【母猪分析】")

        cull = data.get('cull_sow_price')
        if cull is not None and not cull.empty:
            latest = cull.sort_values('date')['value'].iloc[-1]
            lines.append(f"- 河南淘汰母猪价格: {latest:.2f}元/kg")

        for key, label in [('high_parity_disc', '高胎母猪折扣'),
                           ('low_parity_disc', '低胎母猪折扣')]:
            df = data.get(key)
            if df is not None and not df.empty:
                latest = df.sort_values('date')['value'].iloc[-1]
                lines.append(f"- 河南{label}: {latest:.1%}")

        binary = data.get('binary_sow_price')
        if binary is not None and not binary.empty:
            latest = binary.sort_values('date')['value'].iloc[-1]
            wow_chg, wow_pct, yoy_chg, yoy_pct = Analyzer._wow_yoy(binary)
            lines.append(f"- 河南二元母猪(50kg)价格: {latest:.0f}元/头, "
                       f"环比{Analyzer._fmt(wow_pct, is_pct=True)}")

        return '\n'.join(lines)

    @staticmethod
    def analyze_piglet(data):
        """仔猪分析"""
        lines = []
        lines.append("【仔猪分析】")

        for key, label in [('piglet_15kg', '15公斤仔猪出厂价'),
                           ('piglet_weaned', '断奶仔猪(7kg)出厂价')]:
            df = data.get(key)
            if df is not None and not df.empty:
                latest = df.sort_values('date')['value'].iloc[-1]
                wow_chg, wow_pct, yoy_chg, yoy_pct = Analyzer._wow_yoy(df)
                lines.append(f"- {label}: {latest:.0f}元/头, "
                           f"环比{Analyzer._fmt(wow_pct, is_pct=True)}, 同比{Analyzer._fmt(yoy_pct, is_pct=True)}")

        return '\n'.join(lines)

    @staticmethod
    def analyze_profit(data):
        """养殖利润分析 (三类分别)"""
        lines = []
        lines.append("【养殖利润分析 (三类分别)】")

        bp = data.get('breeding_profit', {})
        for label, df in bp.items():
            if df is not None and not df.empty:
                latest = df.sort_values('date')['value'].iloc[-1]
                wow_chg, wow_pct, yoy_chg, yoy_pct = Analyzer._wow_yoy(df)
                lines.append(f"- {label}: {latest:.0f}元/头, "
                           f"环比{Analyzer._fmt(wow_pct, is_pct=True)}, 同比{Analyzer._fmt(yoy_pct, is_pct=True)}")
                # 盈亏判断
                if latest < 0:
                    lines.append(f"  → 仍处于亏损状态")
                else:
                    lines.append(f"  → 已实现盈利")

        return '\n'.join(lines)

    @staticmethod
    def overall_conclusion(data):
        """综合结论 - 以猪价分析师角度"""
        lines = []
        lines.append("=" * 50)
        lines.append(f"涌益咨询生猪周度报告 - 核心结论 ({REPORT_PERIOD})")
        lines.append("=" * 50)
        lines.append("")

        # 价格走势判断
        np_df = data.get('national_price')
        if np_df is not None and not np_df.empty:
            latest = np_df.sort_values('date')['value'].iloc[-1]
            wow_chg, wow_pct, yoy_chg, yoy_pct = Analyzer._wow_yoy(np_df)
            direction = '上涨' if wow_chg > 0 else '下跌'
            lines.append(f"本周全国生猪均价{latest:.2f}元/kg, 环比{direction}{abs(wow_chg):.2f}元/kg。")

            if yoy_chg and not np.isnan(yoy_chg):
                yoy_dir = '高于' if yoy_chg > 0 else '低于'
                lines.append(f"同比{yoy_dir}去年同期{abs(yoy_chg):.2f}元/kg。")

        # 供给端判断
        ws = data.get('weight_split', {})
        group_wt = ws.get('集团')
        if group_wt is not None and not group_wt.empty:
            gw_chg, _, gw_yoy, _ = Analyzer._wow_yoy(group_wt)
            if gw_chg > 0:
                lines.append("集团出栏体重增加, 短期供给压力有所上升。")
            else:
                lines.append("集团出栏体重下降, 短期出栏压力缓解。")

        # 屠宰端判断
        fresh = data.get('fresh_sale_rate')
        if fresh is not None and not fresh.empty:
            latest_fresh = fresh.sort_values('date')['value'].iloc[-1]
            if latest_fresh < 0.7:
                lines.append(f"鲜销率偏低({latest_fresh:.1%}), 终端消费疲软。")
            else:
                lines.append(f"鲜销率{latest_fresh:.1%}, 终端走货正常。")

        # 养殖利润判断
        bp = data.get('breeding_profit', {})
        wg = bp.get('外购仔猪育肥')
        if wg is not None and not wg.empty:
            latest_wg = wg.sort_values('date')['value'].iloc[-1]
            if latest_wg > 0:
                lines.append(f"外购仔猪育肥已实现盈利({latest_wg:.0f}元/头), 补栏积极性有望提升。")
            else:
                lines.append(f"外购仔猪育肥仍亏损({latest_wg:.0f}元/头), 补栏意愿低迷。")

        lines.append("")
        lines.append("【期货与后市展望】")
        lines.append("- 当前生猪期货远月升水结构明显, 反映市场对下半年猪价上涨的预期")
        lines.append("- 从供给端看: 能繁母猪存栏持续下降, 下半年生猪供给有望收紧, 支撑猪价回升")
        lines.append("- 从需求端看: 下半年进入猪肉消费旺季, 季节性需求增加将对猪价形成支撑")
        lines.append("- 风险因素: 二育/压栏可能导致短期供给后移; 进口猪肉量变化; 疫病风险; 饲料成本波动")
        lines.append("- 综合判断: 短期猪价底部震荡, 中期看涨但幅度有限, 关注产能去化进度和终端消费恢复情况")
        lines.append("")
        lines.append("免责声明: 以上分析基于涌益咨询周度数据自动生成, 仅供参考, 不构成投资建议。")
        return '\n'.join(lines)

    @staticmethod
    def _wow_yoy(df):
        """工具: 计算环比同比"""
        if df is None or df.empty:
            return np.nan, np.nan, np.nan, np.nan
        df = df.sort_values('date')
        val_col = 'value' if 'value' in df.columns else 'price'
        cur = df[val_col].iloc[-1]
        wow = df[val_col].iloc[-2] if len(df) > 1 else np.nan
        yoy_idx = max(0, len(df) - 53)
        yoy = df[val_col].iloc[yoy_idx] if len(df) > 52 else np.nan
        return (cur - wow, (cur/wow - 1) if wow and wow != 0 else np.nan,
                cur - yoy, (cur/yoy - 1) if yoy and yoy != 0 else np.nan)

    @staticmethod
    def _natural_week_mean(df, val_col):
        """自然周（周一~周日）日均"""
        if df is None or df.empty:
            return None
        df = df.sort_values('date')
        latest = df['date'].iloc[-1]
        monday = latest - timedelta(days=latest.weekday())
        week = df[(df['date'] >= monday) & (df['date'] <= latest)]
        if week.empty:
            return float(df[val_col].iloc[-1])
        return float(week[val_col].mean())

    @staticmethod
    def _week_avg_compare(df, val_col):
        """自然周日均对比：返回 (本周日均, 上周日均, 去年同周日均, 环比%, 同比%)"""
        if df is None or df.empty:
            return None
        df = df.sort_values('date')
        latest = df['date'].iloc[-1]
        monday = latest - timedelta(days=latest.weekday())
        cur = df[(df['date'] >= monday) & (df['date'] <= latest)]
        cur_mean = float(cur[val_col].mean()) if not cur.empty else float(df[val_col].iloc[-1])
        prev_monday = monday - timedelta(days=7)
        prev_sunday = monday - timedelta(days=1)
        prev = df[(df['date'] >= prev_monday) & (df['date'] <= prev_sunday)]
        prev_mean = float(prev[val_col].mean()) if not prev.empty else None
        ly_monday = monday - timedelta(days=364)
        ly = df[(df['date'] >= ly_monday) & (df['date'] <= latest - timedelta(days=364))]
        ly_mean = float(ly[val_col].mean()) if not ly.empty else None
        wow = (cur_mean / prev_mean - 1) if prev_mean else None
        yoy = (cur_mean / ly_mean - 1) if ly_mean else None
        return (cur_mean, prev_mean, ly_mean, wow, yoy)

    @staticmethod
    def full_analysis(data):
        """生成完整分析文本"""
        sections = []
        sections.append(Analyzer.overall_conclusion(data))
        sections.append('')
        sections.append(Analyzer.analyze_price(data))
        sections.append('')
        sections.append(Analyzer.analyze_supply(data))
        sections.append('')
        sections.append(Analyzer.analyze_slaughter(data))
        sections.append('')
        sections.append(Analyzer.analyze_sow(data))
        sections.append('')
        sections.append(Analyzer.analyze_piglet(data))
        sections.append('')
        sections.append(Analyzer.analyze_profit(data))
        return '\n\n'.join(sections)


# ============================================================
# Dash 在线报告应用
# ============================================================

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.FLATLY],
                title=f'涌益生猪周报 - {REPORT_PERIOD}')

# 导航栏
NAVBAR = dbc.Navbar(
    dbc.Container([
        html.Span("涌益咨询 · 生猪周度数据报告", className="navbar-brand",
                  style={'fontWeight': 'bold'}),
        html.Span(REPORT_PERIOD, className="navbar-text", style={'color': '#e94560'}),
        dbc.Nav([
            dbc.NavItem(dbc.NavLink("报告预览", href="#", active=True, id="nav-preview")),
            dbc.NavItem(dbc.NavLink("在线编辑", href="#", id="nav-edit")),
        ], className="ms-auto", navbar=True),
        dbc.Button("下载PDF (浏览器打印→另存PDF)", id="btn-pdf", color="danger", size="sm", className="ms-2"),
        dbc.Button("刷新数据", id="btn-refresh", color="warning", size="sm"),
    ]),
    color="dark", dark=True, className="mb-3",
)


def build_report_layout(analysis_text):
    """构建完整报告页面"""
    return dbc.Container([
        NAVBAR,
        html.Div(id='report-content', children=[
            # 结论摘要
            dbc.Card([
                dbc.CardHeader(html.H4("核心结论", className="mb-0")),
                dbc.CardBody([
                    dcc.Markdown(analysis_text.split('【价格分析】')[0]
                                 if '【价格分析】' in analysis_text
                                 else analysis_text[:2000],
                                 style={'whiteSpace': 'pre-wrap', 'lineHeight': '1.8'}),
                ]),
            ], className='mb-4 border-danger'),

            # 一、价格分析
            build_section('sec-price', '一、价格分析',
                         analysis_extract(analysis_text, '【价格分析】', '【供给分析】'),
                         [
                             dbc.Row([
                                 dbc.Col([dcc.Graph(id='chart-price-yoy-henan')], width=6),
                                 dbc.Col([dcc.Graph(id='chart-price-yoy-sichuan')], width=6),
                             ]),
                             dbc.Row([
                                 dbc.Col([dcc.Graph(id='chart-price-yoy-guangdong')], width=6),
                                 dbc.Col([dcc.Graph(id='chart-price-yoy-liaoning')], width=6),
                             ]),
                             html.H5("价差历史同期（河南 vs 广东/四川/辽宁/全国均价）", className='mt-3'),
                             dbc.Row([
                                 dbc.Col([dcc.Graph(id='chart-spread-hn-gd')], width=6),
                                 dbc.Col([dcc.Graph(id='chart-spread-hn-sc')], width=6),
                             ]),
                             dbc.Row([
                                 dbc.Col([dcc.Graph(id='chart-spread-hn-ln')], width=6),
                                 dbc.Col([dcc.Graph(id='chart-spread-hn-nat')], width=6),
                             ]),
                             html.H5("散户肥标价差季节性同比 & 全国均价", className='mt-3'),
                             dbc.Row([
                                 dbc.Col([dcc.Graph(id='chart-fat-std-150')], width=6),
                                 dbc.Col([dcc.Graph(id='chart-fat-std-175')], width=6),
                             ]),
                             dbc.Row([
                                 dbc.Col([dcc.Graph(id='chart-national-seasonal')], width=6),
                                 dbc.Col([dcc.Graph(id='chart-maobai')], width=6),
                             ]),
                         ]),

            # 二、供给分析
            build_section('sec-supply', '二、供给分析',
                         analysis_extract(analysis_text, '【供给分析】', '【屠宰分析】'),
                         [
                             html.H5("出栏体重季节性同比"),
                             dbc.Row([
                                 dbc.Col([dcc.Graph(id='chart-weight-group')], width=6),
                                 dbc.Col([dcc.Graph(id='chart-weight-retail')], width=6),
                             ]),
                             html.H5("二育栏舍利用率"),
                             dbc.Row([
                                 dbc.Col([dcc.Graph(id='chart-eryu')], width=12),
                             ]),
                         ]),

            # 三、屠宰分析
            build_section('sec-slaughter', '三、屠宰分析',
                         analysis_extract(analysis_text, '【屠宰分析】', '【母猪分析】'),
                         [
                             html.H5("四省屠宰量季节性对比"),
                             dbc.Row([
                                 dbc.Col([dcc.Graph(id='chart-price-sl-henan')], width=6),
                                 dbc.Col([dcc.Graph(id='chart-price-sl-sichuan')], width=6),
                             ]),
                             dbc.Row([
                                 dbc.Col([dcc.Graph(id='chart-price-sl-guangdong')], width=6),
                                 dbc.Col([dcc.Graph(id='chart-price-sl-liaoning')], width=6),
                             ]),
                             html.H5("全国鲜销率 & 冻品库存季节性同比", className='mt-3'),
                             dbc.Row([
                                 dbc.Col([dcc.Graph(id='chart-fresh')], width=6),
                                 dbc.Col([dcc.Graph(id='chart-frozen')], width=6),
                             ]),
                             html.H5("屠宰利润季节性同比", className='mt-3'),
                             dbc.Row([
                                 dbc.Col([dcc.Graph(id='chart-profit')], width=12),
                             ]),
                         ]),

            # 四、母猪分析
            build_section('sec-sow', '四、母猪分析',
                         analysis_extract(analysis_text, '【母猪分析】', '【仔猪分析】'),
                         [
                             dbc.Row([
                                 dbc.Col([dcc.Graph(id='chart-cull-sow')], width=6),
                                 dbc.Col([dcc.Graph(id='chart-low-parity')], width=6),
                                 dbc.Col([dcc.Graph(id='chart-high-parity')], width=6),
                             ]),
                             dbc.Row([
                                 dbc.Col([dcc.Graph(id='chart-binary-sow')], width=6),
                             ]),
                         ]),

            # 五、仔猪分析
            build_section('sec-piglet', '五、仔猪分析',
                         analysis_extract(analysis_text, '【仔猪分析】', '【养殖利润分析 (三类分别)】'),
                         [
                             dbc.Row([
                                 dbc.Col([dcc.Graph(id='chart-piglet-15')], width=6),
                                 dbc.Col([dcc.Graph(id='chart-piglet-weaned')], width=6),
                             ]),
                             dbc.Row([
                                 dbc.Col([dcc.Graph(id='chart-piglet-sale-profit')], width=12),
                             ]),
                         ]),

            # 六、养殖利润
            build_section('sec-profit', '六、养殖利润分析 (三类分别)',
                         analysis_extract(analysis_text, '【养殖利润分析 (三类分别)】', None),
                         [
                             dbc.Row([
                                 dbc.Col([dcc.Graph(id='chart-breed-profit-1')], width=12),
                             dbc.Row([
                                 dbc.Col([dcc.Graph(id='chart-breed-profit-2')], width=12),
                             ]),
                             dbc.Row([
                                 dbc.Col([dcc.Graph(id='chart-breed-profit-3')], width=12),
                             ]),
                             ]),
                         ]),
        ]),

        # 隐藏的编辑区域
        html.Div(id='edit-area', style={'display': 'none'}, children=[
            dbc.Card([
                dbc.CardHeader("在线编辑分析内容"),
                dbc.CardBody([
                    dbc.Textarea(id='editor-textarea',
                                 value=analysis_text,
                                 style={'width': '100%', 'height': '600px',
                                        'fontFamily': 'monospace', 'fontSize': '14px'}),
                    dbc.Button("保存修改", id='btn-save', color='success', className='mt-3'),
                    html.Div(id='save-status'),
                ]),
            ]),
        ]),

        # Store for analysis text
        dcc.Store(id='analysis-store', data=analysis_text),
    ], fluid=True)


def build_section(section_id, title, analysis, charts):
    """构建报告段落"""
    return html.Div([
        html.H3(title, style={'borderBottom': '3px solid #e94560', 'paddingBottom': '10px'}),
        dbc.Card([
            dbc.CardBody([
                dcc.Markdown(analysis if analysis else '暂无分析数据',
                            style={'whiteSpace': 'pre-wrap', 'lineHeight': '1.8',
                                   'background': '#f8f9fa', 'padding': '15px',
                                   'borderRadius': '8px'}),
            ]),
        ], className='mb-3 border-info'),
        *charts,
        html.Hr(className='my-4'),
    ], id=section_id)


def analysis_extract(text, start_marker, end_marker):
    """从分析文本中提取指定段落"""
    if not text:
        return ''
    if start_marker in text:
        start = text.index(start_marker)
        if end_marker and end_marker in text[start:]:
            end = text.index(end_marker, start)
            return text[start:end].strip()
        return text[start:].strip()
    return ''


# ============================================================
# 布局
# ============================================================

# 初始加载
print("初始化数据...")
reader = DataReader(DATA_DIRS)
all_data = reader.load_all()
analysis_text = Analyzer.full_analysis(all_data)
print("\n分析文本生成完毕.")

app.layout = build_report_layout(analysis_text)


# ============================================================
# 图表回调函数
# ============================================================

def make_seasonal_fig(df, title, ylabel, is_pct=False):
    """生成季节性同比图"""
    if df is None or df.empty:
        return go.Figure()
    df = df.copy()
    df['year'] = df['date'].dt.year
    fig = go.Figure()
    for yr in range(2021, 2027):
        yr_df = df[df['year'] == yr].sort_values('date')
        if yr_df.empty:
            continue
        yr_df['week'] = range(1, len(yr_df) + 1)
        color = YR_COLORS.get(yr, '#888')
        fig.add_trace(go.Scatter(
            x=yr_df['week'], y=yr_df['value'], mode='lines',
            name=f'{yr}年', line=dict(color=color, width=2),
        ))
    fig.update_layout(
        template='plotly_white', title=title,
        xaxis_title='周次', yaxis_title=ylabel, height=400,
        legend=dict(orientation='h', y=1.05),
        margin=dict(l=60, r=40, t=60, b=50),
    )
    fig.update_xaxes(range=[0, None])
    if is_pct:
        fig.update_yaxes(tickformat='.1%')
    return fig


def make_dual_fig(price_data, slaughter_data, prov):
    """生成双轴图"""
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    color = PROV_COLORS.get(prov, '#E31A1C')

    if prov in slaughter_data and slaughter_data[prov] is not None and not slaughter_data[prov].empty:
        sd = slaughter_data[prov].sort_values('date')
        fig.add_trace(go.Bar(
            x=sd['date'], y=sd['value'], name=f'{prov}屠宰量',
            marker_color=color, opacity=0.5,
        ), secondary_y=False)

    if prov in price_data and price_data[prov] is not None and not price_data[prov].empty:
        pd_df = price_data[prov].sort_values('date')
        fig.add_trace(go.Scatter(
            x=pd_df['date'], y=pd_df['price'], name=f'{prov}价格',
            line=dict(color='red', width=2.5),
        ), secondary_y=True)

    fig.update_layout(
        template='plotly_white', title=f'{prov}',
        height=350, hovermode='x unified',
        legend=dict(orientation='h', y=1.1),
    )
    fig.update_yaxes(title_text='屠宰量(头)', secondary_y=False)
    fig.update_yaxes(title_text='价格(元/kg)', secondary_y=True, color='red')
    return fig


# 省份英文映射
PROV_EN = {'河南': 'henan', '四川': 'sichuan', '广东': 'guangdong', '辽宁': 'liaoning'}

# 注册所有图表回调
chart_outputs = []
chart_inputs = []

for prov in TARGET_PROVINCES:
    en = PROV_EN.get(prov, prov)
    chart_outputs.append(Output(f'chart-price-sl-{en}', 'figure'))
    chart_outputs.append(Output(f'chart-price-yoy-{en}', 'figure'))

# 价差历史同期（河南 vs 广东/四川/辽宁/全国均价）
for _sp_id in ['gd', 'sc', 'ln', 'nat']:
    chart_outputs.append(Output(f'chart-spread-hn-{_sp_id}', 'figure'))

chart_outputs.extend([
    Output('chart-fat-std-150', 'figure'),
    Output('chart-fat-std-175', 'figure'),
    Output('chart-national-seasonal', 'figure'),
    Output('chart-maobai', 'figure'),
    Output('chart-weight-group', 'figure'),
    Output('chart-weight-retail', 'figure'),
    Output('chart-eryu', 'figure'),
    Output('chart-fresh', 'figure'),
    Output('chart-frozen', 'figure'),
    Output('chart-profit', 'figure'),
    Output('chart-cull-sow', 'figure'),
    Output('chart-low-parity', 'figure'),
    Output('chart-high-parity', 'figure'),
    Output('chart-binary-sow', 'figure'),
    Output('chart-piglet-15', 'figure'),
    Output('chart-piglet-weaned', 'figure'),
    Output('chart-piglet-sale-profit', 'figure'),
    Output('chart-breed-profit-1', 'figure'),
    Output('chart-breed-profit-2', 'figure'),
    Output('chart-breed-profit-3', 'figure'),
])


@app.callback(
    chart_outputs,
    Input('btn-refresh', 'n_clicks'),
)
def update_all_charts(n):
    data = all_data
    results = []

    # 各省屠宰量季节性对比（纯屠宰量）
    for prov in TARGET_PROVINCES:
        sdf = data.get('province_slaughter', {}).get(prov)
        if sdf is not None and not sdf.empty:
            results.append(make_seasonal_fig(sdf, f'{prov}屠宰量季节性对比 (头/日)', '头/日'))
        else:
            results.append(go.Figure())
        if prov in data['province_price']:
            df = data['province_price'][prov].rename(columns={'price': 'value'})
            results.append(make_seasonal_fig(df, f'{prov}价格季节性同比 (元/kg)', '元/kg'))
        else:
            results.append(go.Figure())

    # 价差历史同期（河南 vs 广东/四川/辽宁/全国均价）
    _series = {}
    for _p in TARGET_PROVINCES:
        _pp = data['province_price'].get(_p)
        if _pp is not None and not _pp.empty:
            _series[_p] = _pp.set_index('date')['price']
    if data.get('national_price') is not None and not data['national_price'].empty:
        _series['全国均价'] = data['national_price'].set_index('date')['value']
    for _o in ['广东', '四川', '辽宁', '全国均价']:
        if '河南' in _series and _o in _series:
            _c = _series['河南'].index.intersection(_series[_o].index)
            if len(_c) > 0:
                _sp = (_series['河南'][_c] - _series[_o][_c]).reset_index()
                _sp.columns = ['date', 'value']
                results.append(make_seasonal_fig(_sp, f'河南-{_o}价差历史同期 (元/kg)', '元/kg'))
            else:
                results.append(go.Figure())
        else:
            results.append(go.Figure())

    fss = data.get('fat_std_spread', {})
    for key in ['150kg肥标价差', '175kg肥标价差']:
        df = fss.get(key) if isinstance(fss, dict) else None
        results.append(make_seasonal_fig(df, f'散户{key}季节性同比 (元/kg)', '元/kg'))

    results.append(make_seasonal_fig(data.get('national_price'),
        '全国生猪均价季节性同比 (元/kg)', '元/kg'))
    results.append(make_seasonal_fig(data.get('maobai_spread'),
        '毛白价差季节性同比 (元/kg)', '元/kg'))

    ws = data.get('weight_split', {})
    for key in ['集团', '散户']:
        results.append(make_seasonal_fig(ws.get(key), f'{key}出栏体重季节性同比 (kg/头)', 'kg/头'))

    results.append(make_seasonal_fig(data.get('eryu_rate'),
        '二育栏舍利用率(全国算术均值)季节性同比', '%', is_pct=True))

    results.append(make_seasonal_fig(data.get('fresh_sale_rate'),
        '全国屠宰鲜销率季节性同比', '%', is_pct=True))
    results.append(make_seasonal_fig(data.get('frozen_stock'),
        '全国冻品库存率季节性同比', '%', is_pct=True))
    results.append(make_seasonal_fig(data.get('slaughter_profit'),
        '河南白条头均利润季节性同比 (元/头)', '元/头'))

    results.append(make_seasonal_fig(data.get('cull_sow_price'),
        '河南淘汰母猪价格季节性同比 (元/kg)', '元/kg'))
    results.append(make_seasonal_fig(data.get('low_parity_disc'),
        '河南低胎母猪折扣季节性同比', '折扣率', is_pct=True))
    results.append(make_seasonal_fig(data.get('high_parity_disc'),
        '河南高胎母猪折扣季节性同比', '折扣率', is_pct=True))
    results.append(make_seasonal_fig(data.get('binary_sow_price'),
        '河南二元母猪(50kg)价格季节性同比 (元/头)', '元/头'))

    results.append(make_seasonal_fig(data.get('piglet_15kg'),
        '15公斤仔猪出厂价季节性同比 (元/头)', '元/头'))
    results.append(make_seasonal_fig(data.get('piglet_weaned'),
        '断奶仔猪(7kg)出厂价季节性同比 (元/头)', '元/头'))
    results.append(make_seasonal_fig(data.get('piglet_sale_profit'),
        '销售仔猪头均利润季节性同比 (元/头)', '元/头'))

    bp = data.get('breeding_profit', {})
    for key in ['母猪50头以下', '5000-10000头', '外购仔猪育肥']:
        df = bp.get(key)
        fig = go.Figure()
        if df is not None and not df.empty:
            df2 = df.copy(); df2['year'] = df2['date'].dt.year
            for yr in range(2021, 2027):
                yr_df = df2[df2['year'] == yr].sort_values('date')
                if yr_df.empty: continue
                yr_df['week'] = range(1, len(yr_df)+1)
                c = YR_COLORS.get(yr, '#888')
                fig.add_trace(go.Scatter(x=yr_df['week'], y=yr_df['value'], mode='lines',
                    name=f'{yr}年', line=dict(color=c, width=2 if yr==2026 else 1.2),
                    opacity=0.9 if yr==2026 else 0.5))
        fig.add_hline(y=0, line_dash="dash", line_color="gray", opacity=0.5)
        fig.update_layout(template='plotly_white', title=f'{key}养殖利润季节性同比 (元/头)',
            xaxis_title='周次', yaxis_title='元/头', height=400, margin=dict(l=60, r=40, t=60, b=50))
        results.append(fig)

    return results


# ============================================================
# 编辑和PDF回调
# ============================================================

@app.callback(
    [Output('report-content', 'style'),
     Output('edit-area', 'style'),
     Output('nav-preview', 'active'),
     Output('nav-edit', 'active')],
    [Input('nav-preview', 'n_clicks'),
     Input('nav-edit', 'n_clicks')],
    prevent_initial_call=True,
)
def toggle_view(n_preview, n_edit):
    ctx = callback_context.triggered_id
    if ctx == 'nav-edit':
        return {'display': 'none'}, {'display': 'block'}, False, True
    return {'display': 'block'}, {'display': 'none'}, True, False


@app.callback(
    [Output('analysis-store', 'data'),
     Output('save-status', 'children')],
    Input('btn-save', 'n_clicks'),
    State('editor-textarea', 'value'),
    prevent_initial_call=True,
)
def save_edits(n, text):
    """保存编辑后的分析文本"""
    if text:
        # 重新生成报告内容
        return text, dbc.Alert("保存成功! 请切换到'报告预览'查看", color="success", duration=3000)
    return dash.no_update, dash.no_update


# PDF下载 - 使用浏览器Ctrl+P打印为PDF


# ============================================================
# 启动
# ============================================================

if __name__ == '__main__':
    print("=" * 60)
    print("  涌益咨询生猪周度数据 - 在线可编辑报告系统")
    print(f"  报告期: {REPORT_PERIOD}")
    print("=" * 60)
    print(f"  数据源: {DATA_DIRS}")
    print()
    print("  启动地址: http://localhost:8051")
    print("  - 报告预览: 查看完整图表和分析")
    print("  - 在线编辑: 修改分析文本")
    print("  - 下载PDF: 浏览器打印为PDF")
    print("=" * 60)
    print()
    print("  [提示] 替换数据源文件夹中的Excel文件后重启即可生成新报告")
    print()

    app.run(debug=False, host='0.0.0.0', port=8051)
