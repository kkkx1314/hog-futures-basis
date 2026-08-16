# -*- coding: utf-8 -*-
"""
屠宰量对比 —— 线上版（Streamlit）

功能：
  1. 数据源：涌益咨询日度数据（屠宰量1 + 各省份）+ 自有屠宰量（云南/四川/浙江）
  2. 指标单选；时间段 A/B 先选公历/农历再选日期，自动按农历月日对齐
  3. 对比（4模块）：A时段详情 / B时段详情 / 均值逐年对齐 / 综合汇总
  4. 预测：默认使用 B 时段，输出具体量级（每日预测值）+ 多种方法 + 取值明细

运行：streamlit run 屠宰量对比_web.py
"""

import os
import glob
from datetime import datetime, date, timedelta

import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go

from lunardate import LunarDate
from sklearn.cluster import KMeans
from scipy.stats import pearsonr

st.set_page_config(page_title="屠宰量对比", page_icon="🐷", layout="wide")

_DESKTOP_DIRS = [r"D:\CC\Desktop", os.path.expanduser("~/Desktop")]
_DEFAULT_NONGLI = (r"D:\CC\Documents\WXWork\1688856584623821\WeDrive\神农集团\期货部\数据库\生猪"
                   r"\曹晨现货数据\3.屠宰或批发市场-日度\3.多区域农历屠宰量.xlsx")

_SELF_CONFIG = [
    ("云南屠宰量", "云南屠宰量", "农历日期", "公历日期", "总屠宰量"),
    ("四川屠宰量", "四川屠宰量完整", "农历日期", "公历日期", "屠宰量（头）"),
    ("四川屠宰量", "四川屠宰量", "农历日期", "公历日期", "合计"),
    ("浙江屠宰量", "浙江屠宰量", "农历日期", "公历日期", "屠宰量"),
]

_PLOT_COLORS = ["#E74C3C", "#3498DB", "#27AE60", "#F39C12", "#9B59B6",
                "#1ABC9C", "#E67E22", "#34495E", "#95A5A6", "#D35400"]

# 年份固定配色（2021-2026）
YEAR_COLOR = {
    2026: "#E74C3C",  # 红
    2025: "#27AE60",  # 绿
    2024: "#000000",  # 黑
    2023: "#3498DB",  # 蓝
    2022: "#F1C40F",  # 黄
    2021: "#9B59B6",  # 紫
}


# ══════════════════════════════════════════════════════════════
# 农历工具
# ══════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def _solar2lunar_ts(ts):
    ld = LunarDate.fromSolarDate(ts.year, ts.month, ts.day)
    return (ld.year, ld.month, ld.day, ld.month * 100 + ld.day)


def _solar2lunar(ts):
    return _solar2lunar_ts(ts)


def _parse_date(x):
    if x is None:
        return None
    try:
        if pd.isna(x):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(x, (datetime, pd.Timestamp)):
        return pd.Timestamp(x)
    if isinstance(x, (int, float, np.integer, np.floating)):
        try:
            return pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(x))
        except Exception:
            return None
    s = str(x).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d"):
        try:
            return pd.Timestamp(datetime.strptime(s[:10], fmt))
        except Exception:
            continue
    return None


def solar_range_to_lunar(start, end):
    days = pd.date_range(pd.Timestamp(start), pd.Timestamp(end))
    return [(_s2l[1], _s2l[2], _s2l[3]) for _s2l in (_solar2lunar(d) for d in days)]


def _md_in_range(k, s, e):
    return s <= k <= e if s <= e else (k >= s or k <= e)


def lunar_range_to_keys(start_md, end_md):
    keys = set()
    for year in (2023, 2024, 2025, 2026):
        for d in pd.date_range(f"{year}-01-01", f"{year}-12-31"):
            _, _, _, k = _solar2lunar(d)
            if _md_in_range(k, start_md, end_md):
                keys.add(k)
    return sorted(keys)


def lunar_days_before(start_md, n):
    seq, seen = [], set()
    for year in (2023, 2024, 2025, 2026):
        for d in pd.date_range(f"{year}-01-01", f"{year}-12-31"):
            _, _, _, k = _solar2lunar(d)
            if k not in seen:
                seen.add(k)
                seq.append(k)
    for i, k in enumerate(seq):
        if k == start_md:
            return seq[max(0, i - n):i]
    return []


def _md_label(k):
    return f"{k // 100}月{k % 100}日"


def find_latest_yongyi():
    best, best_mtime = None, 0
    for d in _DESKTOP_DIRS:
        if not os.path.isdir(d):
            continue
        for pat in ("*涌益咨询日度数据*.xlsx", "*涌益咨询*.xlsx"):
            for f in glob.glob(os.path.join(d, pat)):
                if os.path.basename(f).startswith("~$"):
                    continue
                try:
                    mt = os.path.getmtime(f)
                except OSError:
                    continue
                if mt > best_mtime:
                    best_mtime, best = mt, f
    return best


# ══════════════════════════════════════════════════════════════
# 数据加载
# ══════════════════════════════════════════════════════════════
def _finalize(sub):
    if sub is None or sub.empty:
        return None
    sub = sub.dropna(subset=["date", "value"]).copy()
    lunar = sub["date"].map(_solar2lunar)
    sub["lunar_year"] = [l[0] for l in lunar]
    sub["lunar_md"] = [l[3] for l in lunar]
    sub = sub.sort_values("date").reset_index(drop=True)
    return sub if not sub.empty else None


def load_yongyi(path):
    result = {}
    xl = pd.ExcelFile(path)
    if "价格+宰量" in xl.sheet_names:
        df = pd.read_excel(path, sheet_name="价格+宰量")
        dcol = "日期" if "日期" in df.columns else df.columns[0]
        for vcol, label in (("日屠宰量合计1", "屠宰量1"), ("日度屠宰量合计2", "屠宰量合计2")):
            if vcol in df.columns:
                sub = pd.DataFrame({"date": df[dcol].map(_parse_date),
                                    "value": pd.to_numeric(df[vcol], errors="coerce")})
                sub = _finalize(sub)
                if sub is not None:
                    result[label] = sub
    if "屠宰企业日度屠宰量" in xl.sheet_names:
        raw = pd.read_excel(path, sheet_name="屠宰企业日度屠宰量", header=None)
        prov = raw.iloc[1:, 0].astype(str).tolist()
        dates = raw.iloc[0, 1:].map(_parse_date).tolist()
        vals = raw.iloc[1:, 1:].T
        vals.columns = prov
        vals.index = pd.to_datetime(dates, errors="coerce")
        for p in prov:
            if "合计" in p or not p.strip():
                continue
            sub = pd.DataFrame({"date": vals.index, "value": pd.to_numeric(vals[p], errors="coerce")})
            sub = _finalize(sub)
            if sub is not None:
                result[p] = sub
    return result


def load_self(path):
    result = {}
    for prov, sheet, lcol, scol, vcol in _SELF_CONFIG:
        if prov in result:
            continue
        try:
            df = pd.read_excel(path, sheet_name=sheet)
        except Exception:
            continue
        if lcol not in df.columns or vcol not in df.columns:
            continue
        lunar = df[lcol].map(_parse_date)
        solar = df[scol].map(_parse_date) if scol in df.columns else lunar
        sub = pd.DataFrame({
            "date": solar,
            "value": pd.to_numeric(df[vcol], errors="coerce"),
            "lunar_year": [int(l.year) if pd.notna(l) else np.nan for l in lunar],
            "lunar_md": [int(l.month * 100 + l.day) if pd.notna(l) else np.nan for l in lunar],
        })
        sub = sub.dropna(subset=["date", "value", "lunar_year", "lunar_md"]).sort_values("date").reset_index(drop=True)
        if not sub.empty:
            result[prov] = sub
    return result


def extract_by_lunar(df, md_keys):
    key_set = set(md_keys)
    out = {}
    for year, grp in df[df["lunar_md"].isin(key_set)].groupby("lunar_year"):
        out[int(year)] = grp[["date", "lunar_md", "value"]].sort_values("lunar_md").reset_index(drop=True)
    return out


def compute_stats(values):
    v = np.asarray(values, dtype=float)
    return {"均值": round(float(np.mean(v)), 2), "最大值": round(float(np.max(v)), 2),
            "最小值": round(float(np.min(v)), 2),
            "标准差": round(float(np.std(v, ddof=1)) if len(v) > 1 else 0.0, 2), "数据点": int(len(v))}


# ══════════════════════════════════════════════════════════════
# 预测（4方案 + 取值明细）
# ══════════════════════════════════════════════════════════════
def predict_all(df, pred_keys, base_keys):
    n = len(pred_keys)
    years = sorted(df["lunar_year"].dropna().unique())
    if len(years) < 2:
        return None
    this_year = int(years[-1])

    base_today = df[(df["lunar_year"] == this_year) & (df["lunar_md"].isin(set(base_keys)))]
    base_today = base_today.sort_values("lunar_md")
    if len(base_today) < n:
        base_today = df.sort_values("date").tail(n)
    base_daily = base_today["value"].values.astype(float)[:n]
    base_mean = float(base_daily.mean()) if len(base_daily) else np.nan
    if not np.isfinite(base_mean) or base_mean == 0 or len(base_daily) == 0:
        return None

    rates, base_daily_by_year, years_detail = {}, {}, []
    for y in years[:-1]:
        y = int(y)
        b = df[(df["lunar_year"] == y) & (df["lunar_md"].isin(set(base_keys)))].sort_values("lunar_md")
        p = df[(df["lunar_year"] == y) & (df["lunar_md"].isin(set(pred_keys)))].sort_values("lunar_md")
        if b.empty or p.empty:
            continue
        bm, pm = float(b["value"].mean()), float(p["value"].mean())
        if bm > 0:
            rates[y] = pm / bm - 1
            base_daily_by_year[y] = b["value"].values.astype(float)
            years_detail.append({"年份": y, "基准期均值": bm, "预测期均值": pm, "变化率": pm / bm - 1})

    if not rates:
        return None

    def daily(rate):
        return [float(x) * (1 + rate) for x in base_daily]

    avg_rate = float(np.mean(list(rates.values())))
    s1 = {"name": "历史全均值", "rate": avg_rate, "daily": daily(avg_rate),
          "detail": f"全部 {len(rates)} 个历史年份变化率平均"}

    ys = sorted(rates.keys())
    combos = {}
    if len(ys) >= 3:
        combos["近3年均值"] = float(np.mean([rates[y] for y in ys[-3:]]))
    if len(ys) >= 5:
        combos["近5年均值"] = float(np.mean([rates[y] for y in ys[-5:]]))
    if ys:
        abs_close = sorted(ys, key=lambda y: abs(float(df[(df["lunar_year"] == y) & (df["lunar_md"].isin(set(base_keys)))]["value"].mean()) - base_mean))[:3]
        combos["Top3绝对值接近"] = float(np.mean([rates[y] for y in abs_close]))
        corr = []
        for y in ys:
            a, b = base_daily, base_daily_by_year.get(y)
            if b is None or len(a) != len(b) or len(a) < 3:
                continue
            r, _ = pearsonr(a, b)
            if np.isfinite(r):
                corr.append((r, y))
        if corr:
            corr.sort(reverse=True)
            combos["Top3相关系数"] = float(np.mean([rates[y] for _, y in corr[:3]]))
    s2 = {"name": "多年度组合", "combos": {k: {"rate": v, "daily": daily(v)} for k, v in combos.items()}}

    rate_arr = np.array(list(rates.values())).reshape(-1, 1)
    cluster = None
    if len(rate_arr) >= 3:
        k = min(4, max(3, len(rate_arr) // 2))
        km = KMeans(n_clusters=k, n_init=10, random_state=0)
        labels = km.fit_predict(rate_arr)
        centers = km.cluster_centers_.flatten()
        year_labels = {}
        for (y, _), lab in zip(rates.items(), labels):
            year_labels.setdefault(int(lab), []).append(int(y))
        cluster = {"labels": {int(lab): {"center": float(centers[lab]), "years": sorted(year_labels[int(lab)])}
                              for lab in sorted(year_labels)}}
        nearest = int(np.argmin(np.abs(centers - avg_rate)))
        cluster["nearest_center"] = float(centers[nearest])
    s3 = {"name": "聚类分析", "cluster": cluster, "daily": daily(cluster["nearest_center"]) if cluster else None}

    best_year, best_corr = None, -2
    for y, b in base_daily_by_year.items():
        if len(base_daily) != len(b) or len(base_daily) < 3:
            continue
        r, _ = pearsonr(base_daily, b)
        if np.isfinite(r) and r > best_corr:
            best_corr, best_year = r, y
    s4_rate = rates.get(best_year, avg_rate) if best_year else avg_rate
    s4 = {"name": "走势复盘", "rate": s4_rate, "daily": daily(s4_rate),
          "best_year": best_year, "corr": best_corr if best_year else None,
          "detail": f"匹配 {best_year} 年（相关系数 {best_corr:.2f}）" if best_year else "无足够历史年"}

    return {"pred_keys": pred_keys, "base_daily": base_daily, "base_mean": base_mean,
            "pred_mean": float(np.nanmean(daily(avg_rate))),
            "schemes": [s1, s2, s3, s4], "years_detail": years_detail, "环比": float(avg_rate)}


# ══════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════
st.markdown("## 🐷 屠宰量对比分析")
st.caption("公历/农历双历对齐 · 历年同期对比 · 四方案预测")

with st.sidebar:
    st.markdown("### ① 数据源")
    latest = find_latest_yongyi()
    yongyi_f = st.file_uploader("涌益咨询日度数据", type=["xlsx"])
    yongyi_path = None
    if yongyi_f is not None:
        yongyi_path = yongyi_f
    else:
        yongyi_path = latest
        if latest:
            st.caption(f"已自动检索：{os.path.basename(latest)}")

    self_f = st.file_uploader("自有屠宰量（多区域）", type=["xlsx"])
    self_path = self_f if self_f is not None else (_DEFAULT_NONGLI if os.path.exists(_DEFAULT_NONGLI) else None)

    @st.cache_data(show_spinner=False)
    def _load(yongyi_key, self_key):
        yongyi_data = load_yongyi(yongyi_path) if yongyi_path else {}
        self_data = load_self(self_path) if self_path else {}
        return yongyi_data, self_data

    yongyi_data, self_data = _load(str(yongyi_path), str(self_path))

    st.markdown("### ② 指标选择（单选）")
    options = []
    for name in ("屠宰量1", "屠宰量合计2"):
        if name in yongyi_data:
            options.append(("涌益", name))
    for name in sorted(yongyi_data):
        if name not in ("屠宰量1", "屠宰量合计2"):
            options.append(("涌益", name))
    for name in ("云南屠宰量", "四川屠宰量", "浙江屠宰量"):
        if name in self_data:
            options.append(("自有", name))
    opt_labels = [f"涌益 · {n}" if s == "涌益" else f"自有 · {n}" for s, n in options]
    sel_label = st.radio("选择指标", opt_labels, index=0 if opt_labels else None)
    sel_idx = opt_labels.index(sel_label) if sel_label in opt_labels else 0
    sel_src, sel_name = options[sel_idx]
    sel_df = yongyi_data[sel_name] if sel_src == "涌益" else self_data[sel_name]

    st.markdown("### ③ 时间段")
    cal_a = st.radio("A 时段历法", ["公历", "农历"], horizontal=True)
    if cal_a == "公历":
        a_s = st.date_input("A 起始", value=date(2026, 3, 1))
        a_e = st.date_input("A 结束", value=date(2026, 3, 15))
    else:
        c1, c2 = st.columns(2)
        a_m1 = c1.number_input("A 起始月", 1, 12, 7)
        a_d1 = c1.number_input("A 起始日", 1, 30, 3)
        a_m2 = c2.number_input("A 结束月", 1, 12, 7)
        a_d2 = c2.number_input("A 结束日", 1, 30, 9)
    cal_b = st.radio("B 时段历法", ["公历", "农历"], horizontal=True)
    if cal_b == "公历":
        b_s = st.date_input("B 起始", value=date(2026, 6, 1))
        b_e = st.date_input("B 结束", value=date(2026, 6, 15))
    else:
        c1, c2 = st.columns(2)
        b_m1 = c1.number_input("B 起始月", 1, 12, 6)
        b_d1 = c1.number_input("B 起始日", 1, 30, 20)
        b_m2 = c2.number_input("B 结束月", 1, 12, 7)
        b_d2 = c2.number_input("B 结束日", 1, 30, 5)

    run = st.button("🔍 开始分析", type="primary", use_container_width=True)


def _period_keys(cal, s, e):
    if cal == "公历":
        return [k for _, _, k in solar_range_to_lunar(s, e)]
    return lunar_range_to_keys(s[0] * 100 + s[1], e[0] * 100 + e[1])


if run:
    if cal_a == "公历":
        a_keys = _period_keys("公历", a_s, a_e)
    else:
        a_keys = _period_keys("农历", (a_m1, a_d1), (a_m2, a_d2))
    if cal_b == "公历":
        b_keys = _period_keys("公历", b_s, b_e)
    else:
        b_keys = _period_keys("农历", (b_m1, b_d1), (b_m2, b_d2))

    if not a_keys or not b_keys:
        st.error("时间段无有效日期")
        st.stop()

    a_years = extract_by_lunar(sel_df, a_keys)
    b_years = extract_by_lunar(sel_df, b_keys)
    a_all = [v for g in a_years.values() for v in g["value"].tolist()]
    b_all = [v for g in b_years.values() for v in g["value"].tolist()]

    tab1, tab2, tab3, tab4, tab5 = st.tabs(
        ["📊 A时段详情", "📊 B时段详情", "📊 均值逐年对齐", "📊 综合汇总", "🔮 预测结果"])

    with tab1:
        st.markdown(f"#### A 时段各年份每日值（{sel_name}，农历对齐）")
        fig = go.Figure()
        for i, (y, grp) in enumerate(sorted(a_years.items())):
            x_lbl = [_md_label(k) for k in grp["lunar_md"]]
            fig.add_trace(go.Scatter(x=x_lbl, y=grp["value"],
                                     mode="lines+markers", name=str(y),
                                     line=dict(color=YEAR_COLOR.get(y, "#95A5A6"), width=2),
                                     customdata=grp["date"].dt.strftime("%Y-%m-%d").tolist(),
                                     hovertemplate=f"<b>{y}年</b><br>农历%{{x}}（公历%{{customdata}}）<br>值：%{{y:,.2f}}<extra></extra>"))
        fig.update_layout(xaxis_title="农历日期（各年份农历月日一致）", yaxis_title="屠宰量（头）",
                          template="plotly_white", height=420, hovermode="x unified")
        st.plotly_chart(fig, use_container_width=True)
        rows = [{"年份": y, **compute_stats(g["value"].tolist())} for y, g in sorted(a_years.items())]
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    with tab2:
        st.markdown(f"#### B 时段各年份每日值（{sel_name}，农历对齐）")
        fig = go.Figure()
        for i, (y, grp) in enumerate(sorted(b_years.items())):
            x_lbl = [_md_label(k) for k in grp["lunar_md"]]
            fig.add_trace(go.Scatter(x=x_lbl, y=grp["value"],
                                     mode="lines+markers", name=str(y),
                                     line=dict(color=YEAR_COLOR.get(y, "#95A5A6"), width=2),
                                     customdata=grp["date"].dt.strftime("%Y-%m-%d").tolist(),
                                     hovertemplate=f"<b>{y}年</b><br>农历%{{x}}（公历%{{customdata}}）<br>值：%{{y:,.2f}}<extra></extra>"))
        fig.update_layout(xaxis_title="农历日期（各年份农历月日一致）", yaxis_title="屠宰量（头）",
                          template="plotly_white", height=420, hovermode="x unified")
        st.plotly_chart(fig, use_container_width=True)
        rows = [{"年份": y, **compute_stats(g["value"].tolist())} for y, g in sorted(b_years.items())]
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    with tab3:
        st.markdown("#### A vs B 均值逐年对齐")
        a_mean = {y: float(g["value"].mean()) for y, g in a_years.items()}
        b_mean = {y: float(g["value"].mean()) for y, g in b_years.items()}
        years = sorted(set(a_mean) & set(b_mean))
        fig = go.Figure()
        fig.add_trace(go.Bar(x=[str(y) for y in years], y=[a_mean[y] for y in years],
                             name="A时段均值", marker_color="#3498DB"))
        fig.add_trace(go.Bar(x=[str(y) for y in years], y=[b_mean[y] for y in years],
                             name="B时段均值", marker_color="#E67E22"))
        fig.update_layout(barmode="group", xaxis_title="年份", yaxis_title="均值（头）",
                          template="plotly_white", height=400)
        st.plotly_chart(fig, use_container_width=True)
        rows = [{"年份": y, "A均值": round(a_mean[y], 2), "B均值": round(b_mean[y], 2),
                 "差值(B-A)": round(b_mean[y] - a_mean[y], 2),
                 "变化率%": round((b_mean[y] - a_mean[y]) / a_mean[y] * 100, 2) if a_mean[y] else 0} for y in years]
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    with tab4:
        st.markdown("#### 综合汇总")
        if a_all and b_all:
            am, bm = float(np.mean(a_all)), float(np.mean(b_all))
            diff = bm - am
            rate = diff / am * 100 if am else 0
            c1, c2, c3 = st.columns(3)
            c1.metric("A时段总体均值", f"{am:,.2f}")
            c2.metric("B时段总体均值", f"{bm:,.2f}")
            c3.metric("变化率", f"{rate:+.2f}%", f"差值 {diff:+,.2f}")
        else:
            st.info("数据不足")

    with tab5:
        st.markdown(f"#### 预测结果（默认 B 时段，{sel_name}）")
        pred_keys = b_keys
        n = len(pred_keys)
        base_keys = lunar_days_before(pred_keys[0], n)
        pr = predict_all(sel_df, pred_keys, base_keys)
        if pr is None:
            st.warning("数据不足，无法预测")
        else:
            st.markdown(f"**今年基准值：{pr['base_mean']:,.2f} 头**（基准期：预测区间前 {n} 天）")

            # 取值明细
            st.markdown("##### 历史年份取值明细")
            st.dataframe(pd.DataFrame(pr["years_detail"]).sort_values("年份"),
                         use_container_width=True, hide_index=True)

            # 四方案预测（具体量级）
            st.markdown("##### 四方案预测值（具体量级）")
            pred_rows = []
            for s in pr["schemes"]:
                daily = s.get("daily")
                if s["name"] == "多年度组合":
                    for cname, c in s.get("combos", {}).items():
                        pred_rows.append({"方案": f"多年度·{cname}",
                                          "预测均值": float(np.mean(c["daily"])),
                                          "变化率%": c["rate"] * 100})
                elif s["name"] == "聚类分析" and s.get("cluster"):
                    pred_rows.append({"方案": "聚类·最接近类",
                                      "预测均值": float(np.mean(s["daily"])) if s["daily"] else np.nan,
                                      "变化率%": s["cluster"].get("nearest_center", 0) * 100})
                elif daily is not None:
                    extra = s.get("detail", "")
                    pred_rows.append({"方案": s["name"] + (f"（{extra}）" if extra else ""),
                                      "预测均值": float(np.mean(daily)),
                                      "变化率%": s["rate"] * 100})
            pred_df = pd.DataFrame(pred_rows)
            st.dataframe(pred_df.style.format({"预测均值": "{:,.2f}", "变化率%": "{:+.2f}"}),
                         use_container_width=True, hide_index=True)

            # 各方案预测时间段均值（柱状图）
            st.markdown("##### 各方案预测时间段均值")
            fig = go.Figure()
            names, means = [], []
            for s in pr["schemes"]:
                if s["name"] == "多年度组合":
                    for cname, c in s.get("combos", {}).items():
                        names.append(f"多年度·{cname}")
                        means.append(float(np.mean(c["daily"])))
                elif s["name"] == "聚类分析" and s.get("cluster") and s.get("daily"):
                    names.append("聚类·最接近类")
                    means.append(float(np.mean(s["daily"])))
                elif s.get("daily") is not None:
                    names.append(s["name"])
                    means.append(float(np.mean(s["daily"])))
            fig.add_trace(go.Bar(x=names, y=means, marker_color="#2E86AB",
                                 hovertemplate="%{x}<br>预测均值：%{y:,.2f}<extra></extra>"))
            fig.add_hline(y=pr["base_mean"], line_dash="dash", line_color="#E74C3C",
                          annotation_text=f"今年基准均值 {pr['base_mean']:,.2f}")
            fig.update_layout(xaxis_title="方案", yaxis_title="预测时间段均值（头）",
                              template="plotly_white", height=380)
            st.plotly_chart(fig, use_container_width=True)

            # 聚类详情
            cluster = pr["schemes"][2].get("cluster")
            if cluster:
                st.markdown("##### 聚类分析详情")
                cl_rows = [{"类别": lab, "中心变化率%": info["center"] * 100,
                            "样本年份": "、".join(str(y) for y in info["years"])}
                           for lab, info in cluster["labels"].items()]
                st.dataframe(pd.DataFrame(cl_rows), use_container_width=True, hide_index=True)
