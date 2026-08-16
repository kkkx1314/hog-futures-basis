# -*- coding: utf-8 -*-
"""
上传涌益周度数据 → 自动生成生猪周报（可编辑结论）→ 生成 PDF

运行：streamlit run 上传涌益周度数据.py
"""

import os
import glob
from datetime import datetime, timedelta
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.colors import HexColor
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

st.set_page_config(page_title="涌益周度数据 · 周报生成", page_icon="🐷", layout="wide")

_DESKTOP_DIRS = [r"D:\CC\Desktop", os.path.expanduser("~/Desktop")]


def find_latest_weekly():
    best, best_mtime = None, 0
    for d in _DESKTOP_DIRS:
        if not os.path.isdir(d):
            continue
        for pat in ("*涌益*周度*.xlsx", "*周度数据*.xlsx"):
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


def _tail_rows(df, header_rows):
    d = df.iloc[header_rows:].dropna(how="all")
    return d.iloc[-1], (d.iloc[-2] if len(d) > 1 else None)


def _week_rows(df, date_col=1, header_rows=2):
    """返回 (本周, 上周, 去年同周) 三行（按日期列）。"""
    d = df.iloc[header_rows:].copy()
    d[date_col] = pd.to_datetime(d[date_col], errors="coerce")
    d = d.dropna(subset=[date_col]).sort_values(date_col)
    latest = d.iloc[-1]
    prev = d.iloc[-2] if len(d) > 1 else None
    target = latest[date_col] - pd.Timedelta(days=364)
    d2 = d[d[date_col] < latest[date_col]]
    if d2.empty:
        lastyr = latest
    else:
        lastyr = d2.loc[(d2[date_col] - target).abs().idxmin()]
    return latest, prev, lastyr


def extract_weekly(path):
    """从涌益周度数据提取所有指标，返回 dict。"""
    out = {"file": os.path.basename(path)}
    # 1. 商品猪出栏价
    df = pd.read_excel(path, sheet_name="周度-商品猪出栏价", header=None)
    d = df.iloc[2:].dropna(subset=[1]).copy()
    d[1] = pd.to_datetime(d[1], errors="coerce")
    d = d.dropna(subset=[1]).sort_values(1)
    cur, prev = d.iloc[-1], d.iloc[-2]
    _tgt = cur[1] - pd.Timedelta(days=364)
    _d2 = d[d[1] < cur[1]]
    ly = _d2.loc[(_d2[1] - _tgt).abs().idxmin()]
    out["price"] = {
        "全国": float(cur[19]), "上周": float(prev[19]), "去年": float(ly[19]),
        "河南": float(cur[2]), "四川": float(cur[13]), "广东": float(cur[15]),
        "山东": float(cur[8]), "辽宁": float(cur[10]),
        "环比": (float(cur[19]) / float(prev[19]) - 1) * 100,
    }
    # 2. 体重拆分
    w = pd.read_excel(path, sheet_name="周度-体重拆分", header=2)
    w = w.dropna(subset=[w.columns[0]])
    wr = w.iloc[-1]
    out["weight_split"] = {
        "均重": float(wr["全国均重"]), "集团": float(wr["集团"]), "散户": float(wr["散户"]),
        "集团权重": float(wr["集团.1"]), "散户权重": float(wr["散户.1"]),
    }
    # 3. 周度-体重（90kg以下/150kg以上比例）
    wb = pd.read_excel(path, sheet_name="周度-体重", header=None)
    wb = wb.iloc[2:].copy()
    wb[1] = pd.to_datetime(wb[1], errors="coerce")
    wb = wb.dropna(subset=[1])
    latest = wb[1].max()
    ly_target = latest - pd.Timedelta(days=364)
    wb2 = wb[wb[1] < latest]
    ly_date = wb2.loc[(wb2[1] - ly_target).abs().idxmin(), 1]
    out["weight"] = {}
    for kw in ("90kg以下", "150kg以上"):
        r_cur = wb[(wb[1] == latest) & (wb[2].astype(str) == kw)]
        r_ly = wb[(wb[1] == ly_date) & (wb[2].astype(str) == kw)]
        out["weight"][kw] = {
            "本周": float(r_cur.iloc[0][23]) if not r_cur.empty else None,
            "去年": float(r_ly.iloc[0][23]) if not r_ly.empty else None,
        }
    # 4. 鲜销率（全国 c23）
    df = pd.read_excel(path, sheet_name="鲜销率", header=None)
    c, p, ly = _week_rows(df, 1, 2)
    out["鲜销率"] = {"本周": float(c[23]), "上周": float(p[23]), "去年": float(ly[23])}
    # 5. 冻品库存（全国 c21）
    df = pd.read_excel(path, sheet_name="周度-冻品库存", header=None)
    c, p, ly = _week_rows(df, 1, 2)
    out["冻品库存"] = {"本周": float(c[21]), "上周": float(p[21]), "去年": float(ly[21])}
    # 6. 毛白价差
    df = pd.read_excel(path, sheet_name="周度-毛白价差", header=None)
    c, p, ly = _week_rows(df, 0, 1)
    out["毛白价差"] = {"本周": float(c[3]), "上周": float(p[3]), "去年": float(ly[3]),
                      "白条": float(c[1]), "出栏": float(c[2])}
    # 7. 仔猪与商品猪利润
    df = pd.read_excel(path, sheet_name="仔猪与商品猪利润对比", header=None)
    c, p = _tail_rows(df, 3)
    out["利润对比"] = {"仔猪": float(c[1]), "仔猪上周": float(p[1]),
                       "商品猪": float(c[2]), "商品猪上周": float(p[2])}
    # 8. 养殖利润
    pr = pd.read_excel(path, sheet_name="周度-养殖利润最新", header=2)
    pr = pr.dropna(subset=[pr.columns[0]])
    profit = pr[pr.iloc[:, 2].astype(str).str.contains("利润", na=False)]
    out["养殖利润"] = [float(x) for x in profit.iloc[-1, 3:9]]
    # 9. 河南屠宰白条成本（白条头均利润 c9）
    df = pd.read_excel(path, sheet_name="周度-河南屠宰白条成本", header=None)
    c, p = _tail_rows(df, 2)
    out["白条利润"] = {"本周": float(c[9]), "上周": float(p[9])}
    # 10. 淘汰母猪价格（河南 c2）
    df = pd.read_excel(path, sheet_name="周度-淘汰母猪价格", header=None)
    c, p = _tail_rows(df, 2)
    out["淘汰母猪"] = {"本周": str(c[2]), "上周": str(p[2])}
    # 11. 高胎淘母折扣价（河南 c2）
    df = pd.read_excel(path, sheet_name="周度-高胎淘母折扣价", header=None)
    c, p = _tail_rows(df, 2)
    out["高胎折扣"] = {"本周": float(c[2]), "上周": float(p[2])}
    # 12. 低胎母猪折扣价（河南 c2）
    df = pd.read_excel(path, sheet_name="周度-低胎母猪折扣价", header=None)
    c, p = _tail_rows(df, 2)
    out["低胎折扣"] = {"本周": float(c[2]), "上周": float(p[2])}
    # 13. 二育栏舍利用率（全国平均）
    df = pd.read_excel(path, sheet_name="二育栏舍利用率", header=None)
    hdr = df.iloc[1].tolist()
    date_cols = []
    for ci in range(2, len(hdr)):
        try:
            pd.to_datetime(hdr[ci])
            date_cols.append(ci)
        except Exception:
            pass
    vals = pd.to_numeric(df.iloc[2:, date_cols[-1]], errors="coerce")
    pvals = pd.to_numeric(df.iloc[2:, date_cols[-2]], errors="coerce") if len(date_cols) > 1 else vals
    out["二育"] = {"本周": float(vals.mean()), "上周": float(pvals.mean())}
    # 14. 月度出栏完成率
    df = pd.read_excel(path, sheet_name="月度出栏完成率", header=None)
    c, _ = _tail_rows(df, 2)
    out["出栏完成率"] = {"日期": str(c[0])[:10], "全国": float(pd.to_numeric(c[1:], errors="coerce").mean()),
                        "河南": float(c[5]), "山东": float(c[6])}
    return out


# ── 期货数据 ──
_FUTURES_DIR = r"D:\CC\test-claude\sentiment_platform\data\futures"


def list_futures():
    if not os.path.isdir(_FUTURES_DIR):
        return []
    return sorted(os.path.splitext(os.path.basename(f))[0]
                  for f in glob.glob(os.path.join(_FUTURES_DIR, "LH*.csv")))


def get_futures(ct):
    p = os.path.join(_FUTURES_DIR, f"{ct}.csv")
    if not os.path.exists(p):
        return None
    df = pd.read_csv(p)
    df["date"] = pd.to_datetime(df["date"])
    df = df.sort_values("date")
    last, prev = df.iloc[-1], (df.iloc[-2] if len(df) > 1 else df.iloc[-1])
    close = float(last["close"])
    prev_close = float(prev["close"])
    hold = float(last["hold"]) if pd.notna(last["hold"]) else None
    return {
        "ct": ct, "date": last["date"].date(), "close": close, "prev_close": prev_close,
        "chg": close - prev_close, "chg_pct": (close / prev_close - 1) * 100 if prev_close else 0,
        "hold": hold, "volume": float(last["volume"]),
    }


# ── PDF ──
def _font():
    try:
        pdfmetrics.registerFont(TTFont('SimHei', 'C:/Windows/Fonts/simhei.ttf'))
    except Exception:
        pass


def generate_markdown(data, futures):
    """根据提取的数据生成完整周报 Markdown（结论可自行编辑）。"""
    p = data["price"]
    ws = data["weight_split"]
    w = data["weight"]
    L = []
    L.append("# 生猪市场周报")
    L.append("")
    L.append(f"> 数据来源：{data['file']}")
    L.append("")
    L.append("## 一、现货价格")
    L.append("")
    L.append("| 指标 | 本周 | 上周 | 环比 |")
    L.append("|---|---|---|---|")
    L.append(f"| 全国出栏均价 | {p['全国']:.2f} 元/公斤 | {p['上周']:.2f} 元/公斤 | {p['环比']:+.2f}% |")
    L.append(f"| 河南 | {p['河南']:.2f} | — | — |")
    L.append(f"| 四川 | {p['四川']:.2f} | — | — |")
    L.append(f"| 广东 | {p['广东']:.2f} | — | — |")
    L.append(f"| 山东 | {p['山东']:.2f} | — | — |")
    L.append(f"| 辽宁 | {p['辽宁']:.2f} | — | — |")
    L.append("")
    L.append("## 二、出栏体重与结构")
    L.append("")
    L.append("| 指标 | 本周 | 同期 |")
    L.append("|---|---|---|")
    L.append(f"| 全国均重 | {ws['均重']:.1f} 公斤 | 去年同周 {w.get('90kg以下',{}).get('去年') and ''}— |")
    L.append(f"| 集团均重 | {ws['集团']:.1f} 公斤 | — |")
    L.append(f"| 散户均重 | {ws['散户']:.1f} 公斤 | — |")
    L.append(f"| 集团/散户权重 | {ws['集团权重']*100:.1f}% / {ws['散户权重']*100:.1f}% | — |")
    L.append(f"| 90kg以下比例 | {w.get('90kg以下',{}).get('本周',0)*100:.2f}% | 去年 {w.get('90kg以下',{}).get('去年',0)*100:.2f}% |")
    L.append(f"| 150kg以上比例 | {w.get('150kg以上',{}).get('本周',0)*100:.2f}% | 去年 {w.get('150kg以上',{}).get('去年',0)*100:.2f}% |")
    L.append("")
    L.append("## 三、鲜销率 / 冻品库存 / 毛白价差")
    L.append("")
    L.append("| 指标 | 本周 | 上周 | 去年同周 |")
    L.append("|---|---|---|---|")
    L.append(f"| 鲜销率 | {data['鲜销率']['本周']*100:.2f}% | {data['鲜销率']['上周']*100:.2f}% | {data['鲜销率']['去年']*100:.2f}% |")
    L.append(f"| 冻品库存率 | {data['冻品库存']['本周']*100:.2f}% | {data['冻品库存']['上周']*100:.2f}% | {data['冻品库存']['去年']*100:.2f}% |")
    L.append(f"| 毛白价差 | {data['毛白价差']['本周']:.2f} | {data['毛白价差']['上周']:.2f} | {data['毛白价差']['去年']:.2f} |")
    L.append("")
    L.append("## 四、养殖利润")
    L.append("")
    L.append("| 指标 | 本周 | 上周 |")
    L.append("|---|---|---|")
    lc, bai = data["利润对比"], data["白条利润"]
    L.append(f"| 仔猪头均利润 | {lc['仔猪']:.0f} | {lc['仔猪上周']:.0f} |")
    L.append(f"| 商品猪利润 | {lc['商品猪']:.1f} | {lc['商品猪上周']:.1f} |")
    L.append(f"| 河南白条头均利润 | {bai['本周']:.2f} | {bai['上周']:.2f} |")
    L.append("")
    L.append("## 五、淘汰母猪 / 折扣 / 二育 / 出栏完成率")
    L.append("")
    L.append("| 指标 | 本周 | 上周 |")
    L.append("|---|---|---|")
    L.append(f"| 淘汰母猪价(河南) | {data['淘汰母猪']['本周']} | {data['淘汰母猪']['上周']} |")
    L.append(f"| 高胎淘母折扣 | {data['高胎折扣']['本周']*100:.2f}% | {data['高胎折扣']['上周']*100:.2f}% |")
    L.append(f"| 低胎母猪折扣 | {data['低胎折扣']['本周']*100:.2f}% | {data['低胎折扣']['上周']*100:.2f}% |")
    L.append(f"| 二育栏舍利用率 | {data['二育']['本周']*100:.2f}% | {data['二育']['上周']*100:.2f}% |")
    L.append(f"| 月度出栏完成率 | {data['出栏完成率']['全国']*100:.1f}% | — |")
    if futures:
        f = futures
        basis = p["全国"] * 1000 - f["close"]
        L.append("")
        L.append("## 六、期货端")
        L.append("")
        L.append("| 指标 | 数值 |")
        L.append("|---|---|")
        L.append(f"| {f['ct']} 收盘（{f['date']}） | {f['close']:,.0f} 元/吨 |")
        L.append(f"| 较前日 | {f['chg']:+,.0f}（{f['chg_pct']:+.2f}%） |")
        L.append(f"| 持仓量 | {f['hold']:,.0f} 手 |" if f['hold'] else "| 持仓量 | — |")
        L.append(f"| 成交量 | {f['volume']:,.0f} 手 |")
        L.append(f"| 基差（现货全国×1000−期货） | {basis:+,.0f} 元/吨 |")
    L.append("")
    L.append("## 七、结论")
    L.append("")
    L.append("（请在此填写您的分析结论与期货建议）")
    return "\n".join(L)


def _md_inline(text, style):
    """处理 **加粗** 并返回 Paragraph。"""
    html = ""
    for i, seg in enumerate(text.split("**")):
        html += f"<b>{seg}</b>" if i % 2 == 1 else seg
    return Paragraph(html, style)


def _md_table(rows):
    ncol = max(len(r) for r in rows)
    rows = [r + [""] * (ncol - len(r)) for r in rows]
    t = Table(rows, repeatRows=1)
    t.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'SimHei'), ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('BACKGROUND', (0, 0), (-1, 0), HexColor('#2E86AB')), ('TEXTCOLOR', (0, 0), (-1, 0), HexColor('#ffffff')),
        ('ALIGN', (1, 0), (-1, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, HexColor('#cccccc')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [HexColor('#ffffff'), HexColor('#f4f7fa')]),
        ('TOPPADDING', (0, 0), (-1, -1), 4), ('BOTTOMPADDING', (0, 0), (-1, -1), 4)]))
    return t


def markdown_to_pdf(md_text):
    """把 Markdown 文本转成 PDF bytes（支持 # ## 表格 - 列表 > 引用 **加粗**）。"""
    _font()
    h1 = ParagraphStyle('h1', fontName='SimHei', fontSize=18, leading=24, alignment=TA_CENTER,
                        textColor=HexColor('#1a1a2e'), spaceAfter=8)
    h2 = ParagraphStyle('h2', fontName='SimHei', fontSize=12.5, leading=17, textColor=HexColor('#2c3e50'),
                        spaceBefore=10, spaceAfter=5)
    body = ParagraphStyle('body', fontName='SimHei', fontSize=10, leading=16, textColor=HexColor('#333333'))
    bullet = ParagraphStyle('bullet', fontName='SimHei', fontSize=10, leading=16, leftIndent=12, spaceAfter=2)
    quote = ParagraphStyle('quote', fontName='SimHei', fontSize=8.5, leading=13, textColor=HexColor('#888888'), spaceAfter=4)

    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=16 * mm, rightMargin=16 * mm,
                            topMargin=14 * mm, bottomMargin=14 * mm)
    story = []
    lines = md_text.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        if not line.strip():
            i += 1
            continue
        if line.startswith("# "):
            story.append(_md_inline(line[2:], h1))
        elif line.startswith("## "):
            story.append(_md_inline(line[3:], h2))
        elif line.startswith("|"):
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                rows.append([c.strip() for c in lines[i].strip().strip("|").split("|")])
                i += 1
            if len(rows) >= 2 and all(set(c.replace("-", "").replace(":", "").strip()) == set() for c in rows[1]):
                rows.pop(1)
            if rows:
                story.append(_md_table(rows))
            continue
        elif line.startswith("- "):
            story.append(_md_inline(line[2:], bullet))
        elif line.startswith("> "):
            story.append(_md_inline(line[2:], quote))
        else:
            story.append(_md_inline(line, body))
        i += 1
    doc.build(story)
    return buf.getvalue()


# ── UI ──
st.title("🐷 涌益周度数据 → 生猪周报生成器")
st.caption("上传涌益周度数据 Excel，自动提取指标并生成周报，结论可修改后导出 PDF")

col1, col2 = st.columns([2, 1])
with col1:
    up = st.file_uploader("上传涌益周度数据（.xlsx）", type=["xlsx"])
with col2:
    auto = find_latest_weekly()
    if auto:
        st.info(f"自动检测到：\n\n{os.path.basename(auto)}")

path = up if up is not None else auto

if path is None:
    st.warning("请上传涌益周度数据文件，或在 D:\\CC\\Desktop 放置周度数据文件")
    st.stop()

with st.spinner("正在提取指标…"):
    try:
        data = extract_weekly(path)
    except Exception as e:
        st.error(f"解析失败：{e}")
        st.stop()

st.success(f"已解析：{data['file']}")

st.markdown("### 期货合约")
fut_list = list_futures()
if fut_list:
    sel_ct = st.selectbox("选择期货合约", fut_list, index=fut_list.index("LH2609") if "LH2609" in fut_list else 0)
    futures = get_futures(sel_ct)
else:
    futures = None
    st.info("未检测到本地期货数据（可在 D:\\CC\\test-claude\\sentiment_platform\\data\\futures 放置 LH*.csv）")

st.markdown("### 指标预览")
p = data["price"]
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("全国出栏均价", f"{p['全国']:.2f} 元/公斤", f"{p['环比']:+.2f}%")
c2.metric("鲜销率", f"{data['鲜销率']['本周']*100:.2f}%")
c3.metric("冻品库存率", f"{data['冻品库存']['本周']*100:.2f}%", f"去年 {data['冻品库存']['去年']*100:.2f}%")
c4.metric("二育利用率", f"{data['二育']['本周']*100:.2f}%")
if futures:
    basis = p["全国"] * 1000 - futures["close"]
    c5.metric(f"{futures['ct']} 收盘", f"{futures['close']:,.0f}", f"基差 {basis:+,.0f}")

st.markdown("### 编辑周报（全文可自行修改）")
default_md = generate_markdown(data, futures)
md_text = st.text_area("周报内容（Markdown 格式，可改标题/数据/结论，生成 PDF 时按此内容输出）",
                       value=default_md, height=520)

if st.button("📄 生成 PDF", type="primary", use_container_width=True):
    with st.spinner("生成 PDF…"):
        pdf_bytes = markdown_to_pdf(md_text)
    st.download_button("⬇️ 下载 PDF", data=pdf_bytes, file_name="生猪周报.pdf", mime="application/pdf",
                       use_container_width=True)
