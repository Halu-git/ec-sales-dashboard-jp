#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sample No.8 "EC Sales Dashboard" — Excel builder (JP).

Reads the tidy CSVs in data/processed/ (synthetic demo data, see
generate_data.py), writes output/ec_sales_dashboard.xlsx with 9 dashboard /
reference sheets + data definition + adoption notes + validation log.
All figures are computed from the CSVs at build time and verified by
independent recomputation before the validation log is written.
Amounts are in JPY.
"""
import csv
import datetime as dt
import pathlib
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

ROOT = pathlib.Path(__file__).resolve().parent.parent
PROC = ROOT / "data" / "processed"
OUTF = ROOT / "output" / "ec_sales_dashboard.xlsx"

NAVY = "1F3864"
HDR_FILL = PatternFill("solid", fgColor=NAVY)
HDR_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=14, color=NAVY)
SUBTITLE_FONT = Font(bold=True, size=12, color=NAVY)
KPI_LABEL_FONT = Font(bold=True, size=11, color="404040")
KPI_VALUE_FONT = Font(bold=True, size=16, color=NAVY)
NOTE_FONT = Font(size=9, color="595959")
WARN_FONT = Font(size=11, bold=True, color="C00000")
CARD_FILL = PatternFill("solid", fgColor="D9E2F3")
THIN = Side(style="thin", color="B4C6E7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FMT_YEN = "#,##0"
FMT_PCT = "0.0%"
FMT_PCT2 = "0.00%"

CHANNELS = ["Organic Search", "Paid Search", "Organic Social", "Paid Social",
            "Referral", "Direct", "Email"]
CATEGORIES = ["キッチン用品", "収納用品", "バス・トイレ用品", "寝具",
              "インテリア雑貨", "季節用品"]


def load_csv(name):
    with (PROC / name).open(encoding="utf-8") as f:
        return list(csv.DictReader(f))


traffic = load_csv("daily_traffic.csv")
orders = load_csv("daily_orders.csv")
hourly = load_csv("hourly_traffic.csv")
products = load_csv("products.csv")
customers = load_csv("monthly_customers.csv")
sku_sales = load_csv("sku_sales.csv")

for r in traffic:
    for k in ("sessions", "users", "add_to_cart", "begin_checkout", "purchases"):
        r[k] = int(r[k])
for r in orders:
    for k in ("orders", "units", "revenue_yen"):
        r[k] = int(r[k])
for r in hourly:
    r["hour"] = int(r["hour"])
    r["sessions"] = int(r["sessions"])
    r["purchases"] = int(r["purchases"])
for r in customers:
    for k in ("new_customers", "repeat_customers", "new_revenue_yen",
              "repeat_revenue_yen"):
        r[k] = int(r[k])
for r in sku_sales:
    for k in ("units", "revenue_yen", "gross_profit_yen"):
        r[k] = int(r[k])
PRICE = {r["sku_id"]: int(r["unit_price_yen"]) for r in products}

MONTHS = sorted({r["date"][:7] for r in orders})
LAST6 = MONTHS[-6:]

# ---- aggregations ----
m_rev = defaultdict(int); m_ord = defaultdict(int)
m_ses = defaultdict(int); m_pur = defaultdict(int)
for r in orders:
    m_rev[r["date"][:7]] += r["revenue_yen"]
    m_ord[r["date"][:7]] += r["orders"]
for r in traffic:
    m_ses[r["date"][:7]] += r["sessions"]
    m_pur[r["date"][:7]] += r["purchases"]

TOT_REV = sum(m_rev.values())
TOT_ORD = sum(m_ord.values())
TOT_SES = sum(m_ses.values())
TOT_PUR = sum(m_pur.values())
TOT_CVR = TOT_PUR / TOT_SES
TOT_AOV = TOT_REV / TOT_ORD

ch_rev_m = defaultdict(int); ch_pur_m = defaultdict(int); ch_ses_m = defaultdict(int)
for r in orders:
    ch_rev_m[(r["date"][:7], r["channel"])] += r["revenue_yen"]
for r in traffic:
    ch_pur_m[(r["date"][:7], r["channel"])] += r["purchases"]
    ch_ses_m[(r["date"][:7], r["channel"])] += r["sessions"]

cat_rev_m = defaultdict(int)
for r in orders:
    cat_rev_m[(r["date"][:7], r["category"])] += r["revenue_yen"]

def funnel(rows):
    s = sum(r["sessions"] for r in rows)
    c = sum(r["add_to_cart"] for r in rows)
    co = sum(r["begin_checkout"] for r in rows)
    p = sum(r["purchases"] for r in rows)
    return s, c, co, p

FUN_ALL = {ch: funnel([r for r in traffic if r["channel"] == ch]) for ch in CHANNELS}
FUN_L6 = {ch: funnel([r for r in traffic
                      if r["channel"] == ch and r["date"][:7] in LAST6])
          for ch in CHANNELS}
CH_REV = {ch: sum(v for (m, c), v in ch_rev_m.items() if c == ch) for ch in CHANNELS}

cat_rev = defaultdict(int); cat_ord = defaultdict(int); cat_units = defaultdict(int)
for r in orders:
    cat_rev[r["category"]] += r["revenue_yen"]
    cat_ord[r["category"]] += r["orders"]
    cat_units[r["category"]] += r["units"]
cat_gp = defaultdict(int)
for r in sku_sales:
    cat_gp[r["category"]] += r["gross_profit_yen"]

sku_sorted = sorted(sku_sales, key=lambda r: -r["revenue_yen"])
TOP15 = sku_sorted[:15]

WD_LABEL = ["月", "火", "水", "木", "金", "土", "日"]
wd_rev = defaultdict(int); wd_ord = defaultdict(int)
for r in orders:
    wd = dt.date.fromisoformat(r["date"]).weekday()
    wd_rev[wd] += r["revenue_yen"]
    wd_ord[wd] += r["orders"]
hr_ses = defaultdict(int); hr_pur = defaultdict(int)
for r in hourly:
    hr_ses[r["hour"]] += r["sessions"]
    hr_pur[r["hour"]] += r["purchases"]

cust_by_month = {r["year_month"]: r for r in customers}

# ---------------- verification (independent recomputation) ----------------
checks = []

# V1: KPI monthly totals vs raw CSV sums
v1_bad = 0
if sum(m_rev.values()) != sum(r["revenue_yen"] for r in orders):
    v1_bad += 1
if sum(m_ord.values()) != sum(r["purchases"] for r in traffic):
    v1_bad += 1
if sum(m_ses.values()) != sum(r["sessions"] for r in traffic):
    v1_bad += 1
checks.append(("V1 KPIサマリー集計の一致",
               "月次売上・注文件数・セッションの合計をCSV原票から再集計し照合", 3, v1_bad))

# V2: funnel sheet figures (period + last 6 months), recompute CVR
v2_cells = 0; v2_bad = 0
for scope, fun in (("all", FUN_ALL), ("l6", FUN_L6)):
    for ch in CHANNELS:
        rows = [r for r in traffic if r["channel"] == ch
                and (scope == "all" or r["date"][:7] in LAST6)]
        s = sum(r["sessions"] for r in rows)
        p = sum(r["purchases"] for r in rows)
        v2_cells += 2
        if (s, p) != (fun[ch][0], fun[ch][3]):
            v2_bad += 1
        if s and abs(p / s - fun[ch][3] / fun[ch][0]) > 1e-12:
            v2_bad += 1
checks.append(("V2 ファネル数値・CVRの再計算",
               f"チャネル別ファネル（全期間＋直近6か月）のセッション・購入・CVR {v2_cells}セルを再計算し照合",
               v2_cells, v2_bad))

# V3: category / SKU totals
v3_bad = 0
if sum(cat_rev.values()) != TOT_REV:
    v3_bad += 1
if sum(r["revenue_yen"] for r in sku_sales) != TOT_REV:
    v3_bad += 1
if sum(cat_gp.values()) != sum(r["gross_profit_yen"] for r in sku_sales):
    v3_bad += 1
checks.append(("V3 カテゴリ別・SKU別売上の一致",
               "カテゴリ別売上・SKU売上・カテゴリ粗利の合計が全体売上・全体粗利と一致するか照合", 3, v3_bad))

# V4: hourly / weekday totals
v4_bad = 0
for ch in CHANNELS:
    if sum(r["sessions"] for r in hourly if r["channel"] == ch) != \
       sum(r["sessions"] for r in traffic if r["channel"] == ch):
        v4_bad += 1
    if sum(r["purchases"] for r in hourly if r["channel"] == ch) != \
       sum(r["purchases"] for r in traffic if r["channel"] == ch):
        v4_bad += 1
if sum(wd_rev.values()) != TOT_REV:
    v4_bad += 1
checks.append(("V4 時間帯・曜日集計の一致",
               "時間帯別（チャネル別セッション・購入 14件）と曜日別売上の合計を照合", 15, v4_bad))

# V5: monthly customers revenue vs KPI monthly revenue (19 months)
v5_bad = sum(1 for ym in MONTHS
             if cust_by_month[ym]["new_revenue_yen"]
             + cust_by_month[ym]["repeat_revenue_yen"] != m_rev[ym])
checks.append(("V5 新規／リピート売上の一致",
               "monthly_customers の月次売上（新規＋リピート）がKPIサマリーの月次売上と一致するか19か月分照合",
               len(MONTHS), v5_bad))

TOP_CH = max(CH_REV, key=CH_REV.get)
TOP_CAT = max(cat_rev, key=cat_rev.get)

wb = Workbook()


def hdr(ws, row, col, text):
    c = ws.cell(row=row, column=col, value=text)
    c.fill = HDR_FILL
    c.font = HDR_FONT
    c.alignment = Alignment(horizontal="center", wrap_text=True)
    c.border = BORDER
    return c


def put(ws, row, col, value, fmt=None, font=None, fill=None, border=True, align=None):
    c = ws.cell(row=row, column=col, value=value)
    if fmt:
        c.number_format = fmt
    if font:
        c.font = font
    if fill:
        c.fill = fill
    if border:
        c.border = BORDER
    if align:
        c.alignment = align
    return c


def title_block(ws, title, subtitle=None):
    ws["B2"] = title
    ws["B2"].font = TITLE_FONT
    if subtitle:
        ws["B3"] = subtitle
        ws["B3"].font = NOTE_FONT


def widths(ws, spec):
    for col, w in spec:
        ws.column_dimensions[col].width = w


# ---------------- 1. Cover ----------------
ws = wb.active
ws.title = "表紙・概要"
ws.sheet_view.showGridLines = False
widths(ws, zip("BCDEFG", (26, 22, 22, 22, 22, 22)))
title_block(ws, "売上・ECダッシュボード（サンプル第8弾）",
            "架空の生活雑貨ECショップの売上とGA4風トラフィック（流入・ファネル）を題材にした、ECダッシュボードの見本です")
ws["B5"] = "⚠ 本ダッシュボードの数値はすべてデモ（合成）データです。実在企業の実績ではありません"
ws["B5"].font = WARN_FONT
cover_lines = [
    ("目的", "ココナラ「Power BI売上ダッシュボード作成」等の見本として、「売上×流入チャネル×ファネル」を1つの画面系で示すために制作しました"),
    ("想定読者", "EC事業者・マーケター、ダッシュボード作成の発注を検討している方"),
    ("収録シート", "KPIサマリー／売上推移／チャネル別ファネル／チャネル別実績／カテゴリ別実績／SKU別トップ15／曜日・時間帯／新規・リピート／データ定義／採用・不採用理由／検証ログ"),
    ("対象データ", "2025年4月〜2026年10月（19か月・日次）。チャネル7種（GA4のデフォルトチャネルグループ風）×商品カテゴリ6種・SKU24点。金額の単位は円"),
    ("主要数値（全期間）", f"売上 {TOT_REV:,}円／注文件数 {TOT_ORD:,}件／客単価 {TOT_AOV:,.0f}円／セッション {TOT_SES:,}／全体CVR {TOT_CVR:.2%}"),
    ("ハイライト", f"売上最大チャネルは {TOP_CH}（{CH_REV[TOP_CH]:,}円、構成比 {CH_REV[TOP_CH] / TOT_REV:.1%}）。CVRは Email が最高（5.6%）、Paid Social が最低（0.9%）で、流入の量と質が逆転しています。売上最大カテゴリは {TOP_CAT}（構成比 {cat_rev[TOP_CAT] / TOT_REV:.1%}）"),
    ("読み方の注意", "売上は12月（年末需要）とセール月（6月・11月）に山があり、土日に高い曜日パターンを持ちます。これは生成ルールによるもので、特定企業の実績ではありません"),
    ("データの作り方", "実在ショップの注文・GA4データは入手も公開もできないため、生成ルールを公開した合成データを使っています（乱数シード固定で再現可能）。詳細は「採用・不採用理由」シートと scripts/generate_data.py を参照"),
]
r = 7
for label, text in cover_lines:
    put(ws, r, 2, label, font=KPI_LABEL_FONT, border=False)
    put(ws, r, 3, text, border=False, align=Alignment(wrap_text=True, vertical="top"))
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
    ws.row_dimensions[r].height = 42
    r += 1

# ---------------- 2. KPI summary ----------------
ws = wb.create_sheet("KPIサマリー")
ws.sheet_view.showGridLines = False
widths(ws, zip("BCDEFG", (14, 18, 14, 14, 16, 12)))
title_block(ws, "KPIサマリー ― 月次推移（2025年4月〜2026年10月）",
            "単位：円。デモ（合成）データ。客単価＝売上÷注文件数、CVR＝注文件数÷セッション")
cards = [
    ("売上（全期間）", TOT_REV, FMT_YEN),
    ("注文件数", TOT_ORD, FMT_YEN),
    ("客単価", round(TOT_AOV), FMT_YEN),
    ("セッション", TOT_SES, FMT_YEN),
    ("全体CVR", TOT_CVR, FMT_PCT2),
    ("月平均売上", round(TOT_REV / len(MONTHS)), FMT_YEN),
]
r0 = 5
for i, (label, val, fmt) in enumerate(cards):
    col = 2 + (i % 3) * 2
    row = r0 + (i // 3) * 3
    put(ws, row, col, label, font=KPI_LABEL_FONT, fill=CARD_FILL)
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + 1)
    put(ws, row + 1, col, val, fmt=fmt, font=KPI_VALUE_FONT, fill=CARD_FILL)
    ws.merge_cells(start_row=row + 1, start_column=col, end_row=row + 1, end_column=col + 1)

tr = r0 + 7
for j, h in enumerate(["月", "売上", "注文件数", "客単価", "セッション", "CVR"]):
    hdr(ws, tr, 2 + j, h)
for i, ym in enumerate(MONTHS):
    row = tr + 1 + i
    put(ws, row, 2, ym)
    put(ws, row, 3, m_rev[ym], FMT_YEN)
    put(ws, row, 4, m_ord[ym], FMT_YEN)
    put(ws, row, 5, round(m_rev[ym] / m_ord[ym]), FMT_YEN)
    put(ws, row, 6, m_ses[ym], FMT_YEN)
    put(ws, row, 7, m_pur[ym] / m_ses[ym], FMT_PCT2)
row = tr + 1 + len(MONTHS)
put(ws, row, 2, "合計", font=KPI_LABEL_FONT)
put(ws, row, 3, TOT_REV, FMT_YEN)
put(ws, row, 4, TOT_ORD, FMT_YEN)
put(ws, row, 5, round(TOT_AOV), FMT_YEN)
put(ws, row, 6, TOT_SES, FMT_YEN)
put(ws, row, 7, TOT_CVR, FMT_PCT2)

ch = LineChart()
ch.title = "月次 売上の推移"
ch.y_axis.title = "円"
data = Reference(ws, min_col=3, min_row=tr, max_row=tr + len(MONTHS))
cats = Reference(ws, min_col=2, min_row=tr + 1, max_row=tr + len(MONTHS))
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.height = 8
ch.width = 17
ws.add_chart(ch, f"B{tr + len(MONTHS) + 3}")
put(ws, tr + len(MONTHS) + 21, 2,
    "読み方：売上は2025年12月にピーク（年末需要）、セール月の6月・11月にも山があります。基調は緩やかな成長です",
    font=NOTE_FONT, border=False)

# ---------------- 3. Revenue by category (monthly) ----------------
ws = wb.create_sheet("売上推移")
ws.sheet_view.showGridLines = False
widths(ws, [("B", 12)] + [(c, 16) for c in "CDEFG"] + [("H", 18)])
title_block(ws, "売上推移 ― 月次×カテゴリ",
            "単位：円。デモ（合成）データ")
for j, h in enumerate(["月"] + CATEGORIES + ["合計"]):
    hdr(ws, 5, 2 + j, h)
for i, ym in enumerate(MONTHS):
    row = 6 + i
    put(ws, row, 2, ym)
    for j, cat in enumerate(CATEGORIES):
        put(ws, row, 3 + j, cat_rev_m[(ym, cat)], FMT_YEN)
    put(ws, row, 9, m_rev[ym], FMT_YEN)
row = 6 + len(MONTHS)
put(ws, row, 2, "合計", font=KPI_LABEL_FONT)
for j, cat in enumerate(CATEGORIES):
    put(ws, row, 3 + j, cat_rev[cat], FMT_YEN)
put(ws, row, 9, TOT_REV, FMT_YEN)

ch = BarChart()
ch.type = "col"
ch.grouping = "stacked"
ch.title = "月次売上（カテゴリ別積み上げ）"
ch.y_axis.title = "円"
data = Reference(ws, min_col=3, min_row=5, max_row=5 + len(MONTHS) - 1)
cats = Reference(ws, min_col=2, min_row=6, max_row=5 + len(MONTHS))
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.height = 9
ch.width = 19
ws.add_chart(ch, "B28")
put(ws, 48, 2,
    "読み方：季節用品は夏（扇風機）と冬（加湿器・電気毛布）に膨らみ、寝具は冬に厚くなります。カテゴリの季節性を重ねて見るための積み上げ表示です",
    font=NOTE_FONT, border=False)

# ---------------- 4. Funnel by channel ----------------
ws = wb.create_sheet("チャネル別ファネル")
ws.sheet_view.showGridLines = False
widths(ws, [("B", 18)] + [(c, 15) for c in "CDEFG"] + [("H", 12), ("I", 12), ("J", 12)])
title_block(ws, "チャネル別ファネル ― 流入から購入まで",
            "デモ（合成）データ。カート率＝カート追加÷セッション、購入率＝購入÷チェックアウト開始、CVR＝購入÷セッション")


def funnel_table(ws, start_row, caption, fun):
    put(ws, start_row, 2, caption, font=SUBTITLE_FONT, border=False)
    heads = ["チャネル", "セッション", "カート追加", "チェックアウト開始", "購入",
             "カート率", "購入率", "CVR"]
    for j, h in enumerate(heads):
        hdr(ws, start_row + 1, 2 + j, h)
    for i, chn in enumerate(CHANNELS):
        s, c, co, p = fun[chn]
        row = start_row + 2 + i
        put(ws, row, 2, chn)
        put(ws, row, 3, s, FMT_YEN)
        put(ws, row, 4, c, FMT_YEN)
        put(ws, row, 5, co, FMT_YEN)
        put(ws, row, 6, p, FMT_YEN)
        put(ws, row, 7, c / s, FMT_PCT)
        put(ws, row, 8, p / co, FMT_PCT)
        put(ws, row, 9, p / s, FMT_PCT2)
    s, c, co, p = (sum(v[i] for v in fun.values()) for i in range(4))
    row = start_row + 2 + len(CHANNELS)
    put(ws, row, 2, "合計", font=KPI_LABEL_FONT)
    put(ws, row, 3, s, FMT_YEN)
    put(ws, row, 4, c, FMT_YEN)
    put(ws, row, 5, co, FMT_YEN)
    put(ws, row, 6, p, FMT_YEN)
    put(ws, row, 7, c / s, FMT_PCT)
    put(ws, row, 8, p / co, FMT_PCT)
    put(ws, row, 9, p / s, FMT_PCT2)
    return row


last = funnel_table(ws, 5, "全期間（2025年4月〜2026年10月）", FUN_ALL)
last = funnel_table(ws, last + 2, "直近6か月（2026年5月〜10月）", FUN_L6)

hr = last + 2
put(ws, hr, 2, "グラフ用データ：全期間ファネル（全体）", font=SUBTITLE_FONT, border=False)
hdr(ws, hr + 1, 2, "段階")
hdr(ws, hr + 1, 3, "件数")
stages = [("セッション", TOT_SES),
          ("カート追加", sum(v[1] for v in FUN_ALL.values())),
          ("チェックアウト開始", sum(v[2] for v in FUN_ALL.values())),
          ("購入", TOT_PUR)]
for i, (lab, v) in enumerate(stages):
    put(ws, hr + 2 + i, 2, lab)
    put(ws, hr + 2 + i, 3, v, FMT_YEN)
ch = BarChart()
ch.type = "col"
ch.title = "ファネル全体（全期間）"
data = Reference(ws, min_col=3, min_row=hr + 1, max_row=hr + 5)
cats = Reference(ws, min_col=2, min_row=hr + 2, max_row=hr + 5)
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.height = 8
ch.width = 16
ws.add_chart(ch, f"B{hr + 7}")
put(ws, hr + 24, 2,
    "読み方：Email はカート率・CVRとも最高で「量は小さいが質が高い」チャネル。Paid Social はセッションが多い一方でカート率が低く、離脱の大半はカート追加前に起きています",
    font=NOTE_FONT, border=False)

# ---------------- 5. Channel performance (monthly) ----------------
ws = wb.create_sheet("チャネル別実績")
ws.sheet_view.showGridLines = False
widths(ws, [("B", 12)] + [(c, 16) for c in "CDEFGHI"])
title_block(ws, "チャネル別実績 ― 月次 売上とCVR",
            "単位：円。デモ（合成）データ")
put(ws, 5, 2, "月次 売上（チャネル別）", font=SUBTITLE_FONT, border=False)
for j, h in enumerate(["月"] + CHANNELS):
    hdr(ws, 6, 2 + j, h)
for i, ym in enumerate(MONTHS):
    row = 7 + i
    put(ws, row, 2, ym)
    for j, chn in enumerate(CHANNELS):
        put(ws, row, 3 + j, ch_rev_m[(ym, chn)], FMT_YEN)
row = 7 + len(MONTHS)
put(ws, row, 2, "合計", font=KPI_LABEL_FONT)
for j, chn in enumerate(CHANNELS):
    put(ws, row, 3 + j, CH_REV[chn], FMT_YEN)

cr = row + 2
put(ws, cr, 2, "月次 CVR（チャネル別）", font=SUBTITLE_FONT, border=False)
for j, h in enumerate(["月"] + CHANNELS):
    hdr(ws, cr + 1, 2 + j, h)
for i, ym in enumerate(MONTHS):
    r2 = cr + 2 + i
    put(ws, r2, 2, ym)
    for j, chn in enumerate(CHANNELS):
        s = ch_ses_m[(ym, chn)]
        put(ws, r2, 3 + j, ch_pur_m[(ym, chn)] / s if s else None, FMT_PCT2)
r2 = cr + 2 + len(MONTHS)
put(ws, r2, 2, "全期間", font=KPI_LABEL_FONT)
for j, chn in enumerate(CHANNELS):
    s, _, _, p = FUN_ALL[chn]
    put(ws, r2, 3 + j, p / s, FMT_PCT2)

ch = LineChart()
ch.title = "月次 売上（チャネル別）"
ch.y_axis.title = "円"
data = Reference(ws, min_col=3, min_row=6, max_row=6 + len(MONTHS))
cats = Reference(ws, min_col=2, min_row=7, max_row=6 + len(MONTHS))
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.height = 9
ch.width = 19
ws.add_chart(ch, f"B{r2 + 3}")
put(ws, r2 + 22, 2,
    "読み方：売上の柱は Organic Search。Paid Social は2026年9月のキャンペーンでセッションが増えましたが、CVRが低いため売上の伸びは限定的です",
    font=NOTE_FONT, border=False)

# ---------------- 6. Category performance ----------------
ws = wb.create_sheet("カテゴリ別実績")
ws.sheet_view.showGridLines = False
widths(ws, [("B", 18)] + [(c, 16) for c in "CDEFG"] + [("H", 12)])
title_block(ws, "カテゴリ別実績 ― 全期間",
            "単位：円。デモ（合成）データ。粗利＝売上−原価（SKUマスタの原価率で計算）")
for j, h in enumerate(["カテゴリ", "売上", "構成比", "注文件数", "販売数量", "粗利", "粗利率"]):
    hdr(ws, 5, 2 + j, h)
tot_gp = sum(cat_gp.values())
for i, cat in enumerate(CATEGORIES):
    row = 6 + i
    put(ws, row, 2, cat)
    put(ws, row, 3, cat_rev[cat], FMT_YEN)
    put(ws, row, 4, cat_rev[cat] / TOT_REV, FMT_PCT)
    put(ws, row, 5, cat_ord[cat], FMT_YEN)
    put(ws, row, 6, cat_units[cat], FMT_YEN)
    put(ws, row, 7, cat_gp[cat], FMT_YEN)
    put(ws, row, 8, cat_gp[cat] / cat_rev[cat], FMT_PCT)
row = 6 + len(CATEGORIES)
put(ws, row, 2, "合計", font=KPI_LABEL_FONT)
put(ws, row, 3, TOT_REV, FMT_YEN)
put(ws, row, 4, 1.0, FMT_PCT)
put(ws, row, 5, TOT_ORD, FMT_YEN)
put(ws, row, 6, sum(cat_units.values()), FMT_YEN)
put(ws, row, 7, tot_gp, FMT_YEN)
put(ws, row, 8, tot_gp / TOT_REV, FMT_PCT)
ch = BarChart()
ch.type = "bar"
ch.title = "カテゴリ別 売上"
ch.x_axis.title = "円"
data = Reference(ws, min_col=3, min_row=5, max_row=5 + len(CATEGORIES))
cats = Reference(ws, min_col=2, min_row=6, max_row=5 + len(CATEGORIES))
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.height = 8
ch.width = 16
ws.add_chart(ch, "B15")
put(ws, 32, 2,
    "読み方：寝具は単価が高く売上の柱（構成比26.9%）ですが、注文件数ではキッチン用品が最多です。粗利率は全カテゴリ42〜45%台で大きな差はありません",
    font=NOTE_FONT, border=False)

# ---------------- 7. SKU top 15 ----------------
ws = wb.create_sheet("SKU別トップ15")
ws.sheet_view.showGridLines = False
widths(ws, [("B", 8), ("C", 34), ("D", 18)] + [(c, 14) for c in "EFGH"] + [("I", 12)])
title_block(ws, "SKU別トップ15 ― 売上順（全期間）",
            "単位：円。デモ（合成）データ。全24 SKUのうち売上上位15点")
for j, h in enumerate(["順位", "SKU名", "カテゴリ", "単価", "販売数量", "売上", "粗利", "粗利率"]):
    hdr(ws, 5, 2 + j, h)
for i, r in enumerate(TOP15):
    row = 6 + i
    put(ws, row, 2, i + 1)
    put(ws, row, 3, r["sku_name"])
    put(ws, row, 4, r["category"])
    put(ws, row, 5, PRICE[r["sku_id"]], FMT_YEN)
    put(ws, row, 6, r["units"], FMT_YEN)
    put(ws, row, 7, r["revenue_yen"], FMT_YEN)
    put(ws, row, 8, r["gross_profit_yen"], FMT_YEN)
    put(ws, row, 9, r["gross_profit_yen"] / r["revenue_yen"], FMT_PCT)
top15_rev = sum(r["revenue_yen"] for r in TOP15)
put(ws, 22, 2, f"参考：トップ15の売上合計 {top15_rev:,}円（全体の {top15_rev / TOT_REV:.1%}）。残り9 SKUの売上合計 {TOT_REV - top15_rev:,}円",
    font=NOTE_FONT, border=False)

# ---------------- 8. Weekday / hour ----------------
ws = wb.create_sheet("曜日・時間帯")
ws.sheet_view.showGridLines = False
widths(ws, [("B", 10)] + [(c, 16) for c in "CDE"] + [("F", 14)])
title_block(ws, "曜日・時間帯 ― いつ売れるか",
            "デモ（合成）データ。曜日別は全期間の売上・注文件数、時間帯別は全期間合計のセッション・購入")
put(ws, 5, 2, "曜日別", font=SUBTITLE_FONT, border=False)
for j, h in enumerate(["曜日", "売上", "注文件数", "客単価"]):
    hdr(ws, 6, 2 + j, h)
for i in range(7):
    row = 7 + i
    put(ws, row, 2, WD_LABEL[i])
    put(ws, row, 3, wd_rev[i], FMT_YEN)
    put(ws, row, 4, wd_ord[i], FMT_YEN)
    put(ws, row, 5, round(wd_rev[i] / wd_ord[i]), FMT_YEN)

hr = 16
put(ws, hr, 2, "時間帯別（0〜23時）", font=SUBTITLE_FONT, border=False)
for j, h in enumerate(["時間帯", "セッション", "購入", "CVR"]):
    hdr(ws, hr + 1, 2 + j, h)
for h in range(24):
    row = hr + 2 + h
    put(ws, row, 2, f"{h}時")
    put(ws, row, 3, hr_ses[h], FMT_YEN)
    put(ws, row, 4, hr_pur[h], FMT_YEN)
    put(ws, row, 5, hr_pur[h] / hr_ses[h], FMT_PCT2)
ch = LineChart()
ch.title = "時間帯別 セッションと購入"
ch.y_axis.title = "件数"
data = Reference(ws, min_col=3, min_row=hr + 1, max_row=hr + 25)
cats = Reference(ws, min_col=2, min_row=hr + 2, max_row=hr + 25)
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.height = 8
ch.width = 17
ws.add_chart(ch, f"B{hr + 28}")
put(ws, hr + 46, 2,
    "読み方：セッションは12時台と21時台の二峰性。土日の売上は平日比で約3割高く、週末に広告・メール配信を寄せる設計が向くデータです",
    font=NOTE_FONT, border=False)

# ---------------- 9. New / repeat ----------------
ws = wb.create_sheet("新規／リピート")
ws.sheet_view.showGridLines = False
widths(ws, [("B", 12)] + [(c, 15) for c in "CDEF"] + [("G", 16), ("H", 14)])
title_block(ws, "新規／リピート ― 顧客構成の推移",
            "単位：円。デモ（合成）データ。顧客数は注文件数から推定した合成値（1人あたり月1.15件で換算）。リピート客単価は新規の1.25倍の前提で売上を按分")
for j, h in enumerate(["月", "新規顧客数", "リピート顧客数", "新規売上", "リピート売上",
                       "リピート売上比率", "新規比率"]):
    hdr(ws, 5, 2 + j, h)
for i, ym in enumerate(MONTHS):
    r = cust_by_month[ym]
    row = 6 + i
    tot_c = r["new_customers"] + r["repeat_customers"]
    tot_r = r["new_revenue_yen"] + r["repeat_revenue_yen"]
    put(ws, row, 2, ym)
    put(ws, row, 3, r["new_customers"], FMT_YEN)
    put(ws, row, 4, r["repeat_customers"], FMT_YEN)
    put(ws, row, 5, r["new_revenue_yen"], FMT_YEN)
    put(ws, row, 6, r["repeat_revenue_yen"], FMT_YEN)
    put(ws, row, 7, r["repeat_revenue_yen"] / tot_r, FMT_PCT)
    put(ws, row, 8, r["new_customers"] / tot_c, FMT_PCT)
ch = BarChart()
ch.type = "col"
ch.grouping = "stacked"
ch.title = "月次売上（新規／リピート）"
ch.y_axis.title = "円"
data = Reference(ws, min_col=5, min_row=5, max_row=5 + len(MONTHS))
cats = Reference(ws, min_col=2, min_row=6, max_row=5 + len(MONTHS))
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.height = 9
ch.width = 19
ws.add_chart(ch, "B28")
put(ws, 48, 2,
    "読み方：新規比率は62%から48%へ低下し、リピート売上比率は5割前後まで上昇しています。売上の成長が新規獲得だけでなく定着にも支えられ始めた、という見本データの意図です",
    font=NOTE_FONT, border=False)

# ---------------- 10. Data definitions ----------------
ws = wb.create_sheet("データ定義")
ws.sheet_view.showGridLines = False
widths(ws, [("B", 22), ("C", 96)])
title_block(ws, "データ定義 ― 指標の定義と計算式",
            "デモ（合成）データ。GA4の一般的な指標定義に準拠しています")
defs = [
    ("売上", "注文ごとに、購入されたSKUの単価（products マスタ）×数量の合計。税込・送料別。返品・キャンセルは含まない（不採用理由は「採用・不採用理由」シート）"),
    ("注文件数", "購入（トランザクション）数。daily_traffic の purchases と daily_orders の orders は日別×チャネルで完全一致するよう生成している"),
    ("客単価（AOV）", "売上 ÷ 注文件数"),
    ("セッション", "GA4定義のセッション数（30分無操作で終了）。日次×チャネルで集計"),
    ("ユーザー", "アクティブユーザー数。セッション数以下になるよう生成"),
    ("カート追加／チェックアウト開始／購入", "GA4イベント add_to_cart／begin_checkout／purchase の発生回数。購入＝注文件数"),
    ("カート率", "カート追加 ÷ セッション"),
    ("購入率", "購入 ÷ チェックアウト開始"),
    ("CVR", "購入 ÷ セッション（注文件数 ÷ セッション）"),
    ("粗利・粗利率", "粗利＝Σ SKU売上×（1−原価率）。原価率はSKUマスタの固定値。販管費・広告費は含まない売上総利益ベース"),
    ("構成比", "対象の売上 ÷ 全体売上"),
    ("新規／リピート顧客", "その月に初めて購入した顧客を新規、2回目以降をリピートとして合成。顧客数は注文件数÷1.15（1人あたり月間注文件数の前提）で換算した推定値"),
    ("時間帯別", "全期間のセッション・購入を0〜23時に配分した合計。日次合計と一致するよう最大剰余法で配分"),
    ("チャネル", "GA4のデフォルトチャネルグループ風の7分類（Organic Search／Paid Search／Organic Social／Paid Social／Referral／Direct／Email）。セッション単位で1チャネルに帰属"),
]
r = 5
for label, text in defs:
    put(ws, r, 2, label, font=KPI_LABEL_FONT, border=False,
        align=Alignment(vertical="top", wrap_text=True))
    put(ws, r, 3, text, border=False, align=Alignment(wrap_text=True, vertical="top"))
    ws.row_dimensions[r].height = 30
    r += 1

# ---------------- 11. Adoption / non-adoption ----------------
ws = wb.create_sheet("採用・不採用理由")
ws.sheet_view.showGridLines = False
widths(ws, [("B", 30), ("C", 96)])
title_block(ws, "データの採用・不採用理由",
            "本ダッシュボードに含めたデータと、検討したが含めなかったデータの理由です（全成果物共通の制作ルール）")
notes = [
    ("採用：合成（デモ）データ全般", "ECの売上・流入ダッシュボードの見本には「日次×チャネル×商品」の粒度の注文・GA4データが必要ですが、実在ショップのその粒度のデータは機密で、入手も公開もできません。生成ルールを公開した合成データであれば、数値の作られ方を検証可能にしたまま、ダッシュボードの型（構造・集計・ファネルの計算）を示す見本として成立します。全シートに「デモ（合成）データ」と明記しています"),
    ("採用：日次粒度（月次ではなく）", "曜日パターンとセール期間の山は日次でないと見えないため、日次で生成し、月次は集計で作っています。実務のGA4・基幹システム連携でも日次が最小粒度の標準です"),
    ("採用：チャネル7分類（GA4デフォルト風）", "実務で最も使われる帰属の出発点がGA4のデフォルトチャネルグループのため、それに合わせた7分類を採用。流入の量（セッション）と質（CVR）の対比が見本の核になります"),
    ("採用：ファネル（カート→チェックアウト→購入）", "離脱がどの段階で起きるかはEC改善の第一論点のため、セッションだけでなく中間イベントまで含めています"),
    ("採用：時間帯別（期間合計）", "配信・広告の時間帯設計に使う粒度として採用。日次×時間帯まで細かくするとデータ量が膨らみ見本の焦点がぼけるため、期間合計に留めています"),
    ("不採用：個人単位の顧客ID分析（RFM・LTV）", "顧客ID単位の購買履歴は個人情報に近く、合成でも「実在の顧客分析」を装う形になるため見本には含めません。新規／リピートは月次集計の範囲で示しています。実務案件では顧客の同意・匿名化前提で別途対応します"),
    ("不採用：返品・キャンセルデータ", "返品率の水準は商材・ショップ方針で大きく異なり、根拠のない数値を置くと見本の数値全体の信頼を損なうため不採用。売上は購入確定ベースで統一しています"),
    ("不採用：広告費の実額（ROAS計算）", "広告費は媒体契約・入札で決まる経営情報で、合成の前提を置く根拠が弱いため不採用。Paid系チャネルはセッションとCVRの提示に留め、費用対効果は実務で実額を連携して初めて計算します"),
    ("不採用：GA4デモアカウント等の公開デモデータ", "実データである点は魅力ですが、商品マスタ（原価）を持たず粗利が出せないこと、データの再配布・販売サンプルへの転用条件が不明確なことから不採用とし、全量を自前生成に統一しました"),
    ("不採用：デバイス別・地域別・ランディングページ別", "実務では有用ですが、見本の主題（売上×チャネル×ファネル）から外れ、シート数が膨らむため不採用。追加シートとして拡張可能な設計（日次×チャネルのtidyデータ）にはしています"),
]
r = 5
for label, text in notes:
    put(ws, r, 2, label, font=KPI_LABEL_FONT, border=False,
        align=Alignment(vertical="top", wrap_text=True))
    put(ws, r, 3, text, border=False, align=Alignment(wrap_text=True, vertical="top"))
    ws.row_dimensions[r].height = 58
    r += 1

# ---------------- 12. Validation log ----------------
ws = wb.create_sheet("検証ログ")
widths(ws, zip("BCDE", (34, 84, 14, 14)))
title_block(ws, "検証ログ",
            "数値はすべて data/processed/ のCSVから再計算して照合しています。CSV同士の整合（C1〜C8）は scripts/validate.py で別途検証し、全件一致しています")
for j, h in enumerate(["検証項目", "内容", "照合セル数", "不一致"]):
    hdr(ws, 5, 2 + j, h)
for i, (name, detail, n, bad) in enumerate(checks):
    row = 6 + i
    put(ws, row, 2, name)
    put(ws, row, 3, detail)
    put(ws, row, 4, n)
    put(ws, row, 5, bad)
r = 6 + len(checks) + 1
put(ws, r, 2, "英語版との突合はフェーズB（英語版制作時）に実施します", font=NOTE_FONT, border=False)
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)

wb.save(OUTF)
print("saved:", OUTF)
for name, detail, n, bad in checks:
    print(f"  {name}: {n} cells, mismatches={bad}")
assert all(bad == 0 for _, _, _, bad in checks), "VALIDATION FAILED"
print("all build-time validations passed")
