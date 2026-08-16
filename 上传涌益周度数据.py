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


# ── PDF ──
def _font():
    try:
        pdfmetrics.registerFont(TTFont('SimHei', 'C:/Windows/Fonts/simhei.ttf'))
    except Exception:
        pass


def build_pdf(data, conclusion):
    """根据提取的数据 + 结论生成 PDF bytes。"""
    _font()
    sts = getSampleStyleSheet()
    title = ParagraphStyle('t', fontName='SimHei', fontSize=19, leading=25, alignment=TA_CENTER,
                           textColor=HexColor('#1a1a2e'), spaceAfter=4)
    sub = ParagraphStyle('s', fontName='SimHei', fontSize=8.5, leading=12, alignment=TA_CENTER,
                         textColor=HexColor('#888888'), spaceAfter=10)
    h = ParagraphStyle('h', fontName='SimHei', fontSize=12.5, leading=17, textColor=HexColor('#2c3e50'),
                       spaceBefore=9, spaceAfter=5)
    body = ParagraphStyle('b', fontName='SimHei', fontSize=9.5, leading=15, textColor=HexColor('#333333'))
    bullet = ParagraphStyle('bl', fontName='SimHei', fontSize=9.5, leading=15, textColor=HexColor('#333333'),
                            leftIndent=12, spaceAfter=3)

    def tbl(data, colw, header_bg='#2E86AB'):
        t = Table(data, colWidths=colw, repeatRows=1)
        t.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), 'SimHei'), ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BACKGROUND', (0, 0), (-1, 0), HexColor(header_bg)), ('TEXTCOLOR', (0, 0), (-1, 0), HexColor('#ffffff')),
            ('ALIGN', (1, 0), (-1, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, HexColor('#cccccc')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [HexColor('#ffffff'), HexColor('#f4f7fa')]),
            ('TOPPADDING', (0, 0), (-1, -1), 4), ('BOTTOMPADDING', (0, 0), (-1, -1), 4)]))
        return t

    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=16 * mm, rightMargin=16 * mm,
                            topMargin=14 * mm, bottomMargin=14 * mm)
    S = []
    p = data.get("price", {})
    S.append(Paragraph("生猪市场周报", title))
    S.append(Paragraph(f"数据来源：{data.get('file','涌益周度数据')}", sub))

    S.append(Paragraph("一、现货价格", h))
    S.append(tbl([['指标', '本周', '上周', '环比'],
                  ['全国出栏均价', f"{p.get('全国'):.2f}", f"{p.get('上周'):.2f}", f"{p.get('环比',0):+.2f}%"],
                  ['河南', f"{p.get('河南'):.2f}", '—', '—'], ['四川', f"{p.get('四川'):.2f}", '—', '—'],
                  ['广东', f"{p.get('广东'):.2f}", '—', '—'], ['山东', f"{p.get('山东'):.2f}", '—', '—'],
                  ['辽宁', f"{p.get('辽宁'):.2f}", '—', '—']], [70 * mm, 38 * mm, 38 * mm, 30 * mm]))

    ws = data.get("weight_split", {})
    w = data.get("weight", {})
    S.append(Paragraph("二、出栏体重与结构", h))
    S.append(tbl([['指标', '本周', '同期'],
                  ['全国均重', f"{ws.get('均重'):.1f} 公斤", '—'],
                  ['集团均重', f"{ws.get('集团'):.1f} 公斤", '—'],
                  ['散户均重', f"{ws.get('散户'):.1f} 公斤", '—'],
                  ['集团/散户权重', f"{ws.get('集团权重')*100:.1f}% / {ws.get('散户权重')*100:.1f}%", '—'],
                  ['90kg以下比例', f"{w.get('90kg以下',{}).get('本周',0)*100:.2f}%", f"去年 {w.get('90kg以下',{}).get('去年',0)*100:.2f}%"],
                  ['150kg以上比例', f"{w.get('150kg以上',{}).get('本周',0)*100:.2f}%", f"去年 {w.get('150kg以上',{}).get('去年',0)*100:.2f}%"]],
                 [60 * mm, 55 * mm, 50 * mm]))

    S.append(Paragraph("三、鲜销率 / 冻品库存 / 毛白价差", h))
    S.append(tbl([['指标', '本周', '上周', '去年同周'],
                  ['鲜销率', f"{data['鲜销率']['本周']*100:.2f}%", f"{data['鲜销率']['上周']*100:.2f}%", f"{data['鲜销率']['去年']*100:.2f}%"],
                  ['冻品库存率', f"{data['冻品库存']['本周']*100:.2f}%", f"{data['冻品库存']['上周']*100:.2f}%", f"{data['冻品库存']['去年']*100:.2f}%"],
                  ['毛白价差', f"{data['毛白价差']['本周']:.2f}", f"{data['毛白价差']['上周']:.2f}", f"{data['毛白价差']['去年']:.2f}"]],
                 [60 * mm, 35 * mm, 35 * mm, 35 * mm]))

    lc = data.get("利润对比", {})
    bai = data.get("白条利润", {})
    S.append(Paragraph("四、养殖利润", h))
    S.append(tbl([['指标', '本周', '上周'],
                  ['仔猪头均利润', f"{lc.get('仔猪'):.0f}", f"{lc.get('仔猪上周'):.0f}"],
                  ['商品猪利润', f"{lc.get('商品猪'):.1f}", f"{lc.get('商品猪上周'):.1f}"],
                  ['河南白条头均利润', f"{bai.get('本周'):.2f}", f"{bai.get('上周'):.2f}"]],
                 [70 * mm, 48 * mm, 48 * mm]))

    S.append(Paragraph("五、淘汰母猪 / 折扣 / 二育 / 出栏完成率", h))
    S.append(tbl([['指标', '本周', '上周'],
                  ['淘汰母猪价(河南)', data['淘汰母猪']['本周'], data['淘汰母猪']['上周']],
                  ['高胎淘母折扣', f"{data['高胎折扣']['本周']*100:.2f}%", f"{data['高胎折扣']['上周']*100:.2f}%"],
                  ['低胎母猪折扣', f"{data['低胎折扣']['本周']*100:.2f}%", f"{data['低胎折扣']['上周']*100:.2f}%"],
                  ['二育栏舍利用率', f"{data['二育']['本周']*100:.2f}%", f"{data['二育']['上周']*100:.2f}%"],
                  ['月度出栏完成率', f"{data['出栏完成率']['全国']*100:.1f}%", '—']],
                 [70 * mm, 48 * mm, 48 * mm]))

    S.append(Paragraph("六、期货操作结论", h))
    for line in conclusion.split("\n"):
        line = line.strip()
        if line:
            S.append(Paragraph(line, bullet))

    doc.build(S)
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

st.markdown("### 指标预览")
p = data["price"]
c1, c2, c3, c4 = st.columns(4)
c1.metric("全国出栏均价", f"{p['全国']:.2f} 元/公斤", f"{p['环比']:+.2f}%")
c2.metric("鲜销率", f"{data['鲜销率']['本周']*100:.2f}%")
c3.metric("冻品库存率", f"{data['冻品库存']['本周']*100:.2f}%", f"去年 {data['冻品库存']['去年']*100:.2f}%")
c4.metric("二育利用率", f"{data['二育']['本周']*100:.2f}%")

st.markdown("### 编辑结论（可自行修改）")
default_conclusion = """1. 短期：现货端关注价格持续性，期货端关注基差与资金动向，建议震荡思路、不追高。
2. 中期：结合旺季消费预期与冻品库存/大猪供应压力，把握基差回归或趋势机会。
3. 产业套保：养殖端深度亏损，建议在期货升水较大时逢高分批卖出保值。

⚠️ 风险提示：旺季消费不及预期、冻品库存高位、出栏体重偏高导致的供应后置。本结论仅供参考，不构成投资建议。"""
conclusion = st.text_area("结论内容（每条一行，生成 PDF 时逐条展示）", value=default_conclusion, height=200)

if st.button("📄 生成 PDF", type="primary", use_container_width=True):
    with st.spinner("生成 PDF…"):
        pdf_bytes = build_pdf(data, conclusion)
    st.download_button("⬇️ 下载 PDF", data=pdf_bytes, file_name="生猪周报.pdf", mime="application/pdf",
                       use_container_width=True)
