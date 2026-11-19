#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sample No.8 "EC Sales Dashboard" — English edition builder.

Loads the Japanese edition (output/ec_sales_dashboard.xlsx, built by
build_excel.py) and translates every label in place — sheet names, headings,
KPI labels, table headers, chart titles, notes, SKU names — then saves
output/ec_sales_dashboard_en.xlsx. Data values are never touched.
After saving, it cross-checks EVERY numeric cell between the JP and EN
workbooks and writes the result as check V6 into the EN Validation Log.
Run build_excel.py first, then this script.
"""
import pathlib
import re

from openpyxl import load_workbook
from openpyxl.styles import Border, Side

ROOT = pathlib.Path(__file__).resolve().parent.parent
SRC = ROOT / "output" / "ec_sales_dashboard.xlsx"
OUTF = ROOT / "output" / "ec_sales_dashboard_en.xlsx"

SHEET_MAP = {
    "表紙・概要": "Cover & Summary",
    "KPIサマリー": "KPI Summary",
    "売上推移": "Revenue Trend",
    "チャネル別ファネル": "Channel Funnel",
    "チャネル別実績": "Channel Performance",
    "カテゴリ別実績": "Category Performance",
    "SKU別トップ15": "SKU Top 15",
    "曜日・時間帯": "Weekday & Hour",
    "新規／リピート": "New vs Repeat",
    "データ定義": "Data Definitions",
    "採用・不採用理由": "Adoption Notes",
    "検証ログ": "Validation Log",
}

CAT_EN = {
    "キッチン用品": "Kitchenware",
    "収納用品": "Storage",
    "バス・トイレ用品": "Bath & Toilet",
    "寝具": "Bedding",
    "インテリア雑貨": "Home Decor",
    "季節用品": "Seasonal Goods",
}

SKU_EN = {
    "真空断熱タンブラー 600ml": "Vacuum Insulated Tumbler 600ml",
    "シリコン保存容器 6点セット": "Silicone Food Storage Containers, Set of 6",
    "竹製カッティングボード": "Bamboo Cutting Board",
    "計量スプーン・カップセット": "Measuring Spoons & Cups Set",
    "積み重ね収納ボックス M": "Stackable Storage Box M",
    "クローゼット整理ケース 3個組": "Closet Organizer Cases, Set of 3",
    "衣類圧縮収納袋 4枚セット": "Clothing Compression Bags, Set of 4",
    "ケーブル収納ポーチ": "Cable Organizer Pouch",
    "珪藻土バスマット": "Diatomite Bath Mat",
    "マイクロファイバータオル 5枚セット": "Microfiber Towels, Set of 5",
    "詰替ディスペンサー 2本セット": "Refill Dispensers, Set of 2",
    "トイレ収納ラック スリム型": "Slim Toilet Storage Rack",
    "洗える低反発枕": "Washable Memory Foam Pillow",
    "冷感敷きパッド シングル": "Cooling Mattress Pad, Single",
    "綿布団カバー 3点セット": "Cotton Futon Cover Set, 3 Pieces",
    "低反発マットレスパッド": "Memory Foam Mattress Pad",
    "LED間接照明 フロアライト": "LED Indirect Floor Lamp",
    "超音波アロマディフューザー": "Ultrasonic Aroma Diffuser",
    "壁掛けミラー ラウンド型": "Round Wall Mirror",
    "フェイクグリーン ポット付き": "Artificial Plant with Pot",
    "USB卓上扇風機": "USB Desk Fan",
    "超音波加湿器 4L": "Ultrasonic Humidifier 4L",
    "洗える電気毛布 140×80cm": "Washable Electric Blanket 140×80cm",
    "除湿シート 2枚セット": "Dehumidifying Sheets, Set of 2",
}

TERM_EN = {
    # ---- shared atoms ----
    "月": "Month",
    "合計": "Total",
    "全期間": "Full period",
    "売上": "Revenue",
    "注文件数": "Orders",
    "客単価": "AOV",
    "セッション": "Sessions",
    "購入": "Purchases",
    "円": "JPY",
    "件数": "Count",
    "火": "Tue",
    "水": "Wed",
    "木": "Thu",
    "金": "Fri",
    "土": "Sat",
    "日": "Sun",
    # ---- cover ----
    "売上・ECダッシュボード（サンプル第8弾）": "EC Sales Dashboard (Sample No.8)",
    "架空の生活雑貨ECショップの売上とGA4風トラフィック（流入・ファネル）を題材にした、ECダッシュボードの見本です":
        "A sample EC dashboard built on the sales and GA4-style traffic (acquisition & funnel) of a fictional home-goods e-commerce shop",
    "⚠ 本ダッシュボードの数値はすべてデモ（合成）データです。実在企業の実績ではありません":
        "⚠ All figures in this dashboard are demo (synthetic) data. They are not the results of any real company.",
    "目的": "Purpose",
    "ココナラ「Power BI売上ダッシュボード作成」等の見本として、「売上×流入チャネル×ファネル」を1つの画面系で示すために制作しました":
        "Built as a portfolio sample for Power BI dashboard services (e.g. the Coconala 'Power BI Sales Dashboard' service), demonstrating 'revenue × acquisition channel × funnel' in one set of screens",
    "想定読者": "Intended readers",
    "EC事業者・マーケター、ダッシュボード作成の発注を検討している方":
        "E-commerce operators and marketers, and anyone considering commissioning a dashboard",
    "収録シート": "Contents",
    "KPIサマリー／売上推移／チャネル別ファネル／チャネル別実績／カテゴリ別実績／SKU別トップ15／曜日・時間帯／新規・リピート／データ定義／採用・不採用理由／検証ログ":
        "KPI Summary / Revenue Trend / Channel Funnel / Channel Performance / Category Performance / SKU Top 15 / Weekday & Hour / New vs Repeat / Data Definitions / Adoption Notes / Validation Log",
    "対象データ": "Data coverage",
    "2025年4月〜2026年10月（19か月・日次）。チャネル7種（GA4のデフォルトチャネルグループ風）×商品カテゴリ6種・SKU24点。金額の単位は円":
        "Apr 2025 – Oct 2026 (19 months, daily). 7 channels (GA4-style default channel groups) × 6 product categories, 24 SKUs. Amounts in JPY",
    "主要数値（全期間）": "Key figures (full period)",
    "ハイライト": "Highlights",
    "読み方の注意": "How to read",
    "売上は12月（年末需要）とセール月（6月・11月）に山があり、土日に高い曜日パターンを持ちます。これは生成ルールによるもので、特定企業の実績ではありません":
        "Revenue peaks in December (year-end demand) and in the sale months (June and November), and runs higher on weekends. These patterns come from the generation rules and do not describe any real company",
    "データの作り方": "How the data was made",
    "実在ショップの注文・GA4データは入手も公開もできないため、生成ルールを公開した合成データを使っています（乱数シード固定で再現可能）。詳細は「採用・不採用理由」シートと scripts/generate_data.py を参照":
        "Order and GA4 data of real shops can neither be obtained nor published, so this sample uses synthetic data with fully published generation rules (fixed random seed, fully reproducible). See the Adoption Notes sheet and scripts/generate_data.py",
    # ---- KPI summary ----
    "KPIサマリー ― 月次推移（2025年4月〜2026年10月）": "KPI Summary — Monthly Trend (Apr 2025 – Oct 2026)",
    "単位：円。デモ（合成）データ。客単価＝売上÷注文件数、CVR＝注文件数÷セッション":
        "JPY. Demo (synthetic) data. AOV = revenue ÷ orders; CVR = orders ÷ sessions",
    "売上（全期間）": "Revenue (full period)",
    "全体CVR": "Overall CVR",
    "月平均売上": "Average monthly revenue",
    "月次 売上の推移": "Monthly revenue trend",
    "読み方：売上は2025年12月にピーク（年末需要）、セール月の6月・11月にも山があります。基調は緩やかな成長です":
        "How to read: revenue peaks in December 2025 (year-end demand), with secondary peaks in the June and November sale months. The underlying trend is gradual growth",
    # ---- revenue trend ----
    "売上推移 ― 月次×カテゴリ": "Revenue Trend — Month × Category",
    "単位：円。デモ（合成）データ": "JPY. Demo (synthetic) data",
    "月次売上（カテゴリ別積み上げ）": "Monthly revenue by category (stacked)",
    "読み方：季節用品は夏（扇風機）と冬（加湿器・電気毛布）に膨らみ、寝具は冬に厚くなります。カテゴリの季節性を重ねて見るための積み上げ表示です":
        "How to read: Seasonal Goods swell in summer (fans) and winter (humidifiers, electric blankets), and Bedding thickens in winter. The stacked view is for reading category seasonality on top of the total",
    # ---- funnel ----
    "チャネル別ファネル ― 流入から購入まで": "Channel Funnel — From Session to Purchase",
    "デモ（合成）データ。カート率＝カート追加÷セッション、購入率＝購入÷チェックアウト開始、CVR＝購入÷セッション":
        "Demo (synthetic) data. Cart rate = add-to-cart ÷ sessions; completion rate = purchases ÷ begin-checkout; CVR = purchases ÷ sessions",
    "全期間（2025年4月〜2026年10月）": "Full period (Apr 2025 – Oct 2026)",
    "直近6か月（2026年5月〜10月）": "Last 6 months (May – Oct 2026)",
    "チャネル": "Channel",
    "カート追加": "Add to cart",
    "チェックアウト開始": "Begin checkout",
    "カート率": "Cart rate",
    "購入率": "Completion rate",
    "グラフ用データ：全期間ファネル（全体）": "Chart data: full-period funnel (total)",
    "段階": "Stage",
    "ファネル全体（全期間）": "Overall funnel (full period)",
    "読み方：Email はカート率・CVRとも最高で「量は小さいが質が高い」チャネル。Paid Social はセッションが多い一方でカート率が低く、離脱の大半はカート追加前に起きています":
        "How to read: Email has the highest cart rate and CVR — a 'small volume, high quality' channel. Paid Social brings many sessions but a low cart rate: most drop-off happens before add-to-cart",
    # ---- channel performance ----
    "チャネル別実績 ― 月次 売上とCVR": "Channel Performance — Monthly Revenue and CVR",
    "月次 売上（チャネル別）": "Monthly revenue (by channel)",
    "月次 CVR（チャネル別）": "Monthly CVR (by channel)",
    "月次 売上（チャネル別）": "Monthly revenue by channel",
    "読み方：売上の柱は Organic Search。Paid Social は2026年9月のキャンペーンでセッションが増えましたが、CVRが低いため売上の伸びは限定的です":
        "How to read: Organic Search is the revenue pillar. Paid Social sessions rose with the September 2026 campaign, but its low CVR kept the revenue gain limited",
    # ---- category performance ----
    "カテゴリ別実績 ― 全期間": "Category Performance — Full Period",
    "単位：円。デモ（合成）データ。粗利＝売上−原価（SKUマスタの原価率で計算）":
        "JPY. Demo (synthetic) data. Gross profit = revenue − cost (cost from the SKU master's cost ratios)",
    "カテゴリ": "Category",
    "構成比": "Share",
    "販売数量": "Units sold",
    "粗利": "Gross profit",
    "粗利率": "Gross margin",
    "カテゴリ別 売上": "Revenue by category",
    "読み方：寝具は単価が高く売上の柱（構成比26.9%）ですが、注文件数ではキッチン用品が最多です。粗利率は全カテゴリ42〜45%台で大きな差はありません":
        "How to read: Bedding's high price points make it the revenue pillar (26.9% share), while Kitchenware leads in order count. Gross margin sits in the 42–45% range for every category, with no large gaps",
    # ---- SKU top 15 ----
    "SKU別トップ15 ― 売上順（全期間）": "SKU Top 15 — By Revenue (Full Period)",
    "単位：円。デモ（合成）データ。全24 SKUのうち売上上位15点":
        "JPY. Demo (synthetic) data. The top 15 of 24 SKUs by revenue",
    "順位": "Rank",
    "SKU名": "SKU name",
    "単価": "Unit price",
    # ---- weekday & hour ----
    "曜日・時間帯 ― いつ売れるか": "Weekday & Hour — When Sales Happen",
    "デモ（合成）データ。曜日別は全期間の売上・注文件数、時間帯別は全期間合計のセッション・購入":
        "Demo (synthetic) data. Weekday figures are full-period revenue and orders; hourly figures are full-period totals of sessions and purchases",
    "曜日別": "By weekday",
    "曜日": "Weekday",
    "時間帯別（0〜23時）": "By hour (0–23)",
    "時間帯": "Hour",
    "時間帯別 セッションと購入": "Sessions and purchases by hour",
    "読み方：セッションは12時台と21時台の二峰性。土日の売上は平日比で約3割高く、週末に広告・メール配信を寄せる設計が向くデータです":
        "How to read: sessions are bimodal, peaking at 12:00 and 21:00. Weekend revenue runs about 30% above weekdays — a pattern that suits shifting ads and email sends toward weekends",
    # ---- new vs repeat ----
    "新規／リピート ― 顧客構成の推移": "New vs Repeat — Customer Mix Trend",
    "単位：円。デモ（合成）データ。顧客数は注文件数から推定した合成値（1人あたり月1.15件で換算）。リピート客単価は新規の1.25倍の前提で売上を按分":
        "JPY. Demo (synthetic) data. Customer counts are synthetic estimates converted from orders (1.15 orders per customer per month). Revenue is split assuming repeat AOV is 1.25× new-customer AOV",
    "新規顧客数": "New customers",
    "リピート顧客数": "Repeat customers",
    "新規売上": "New revenue",
    "リピート売上": "Repeat revenue",
    "リピート売上比率": "Repeat revenue share",
    "新規比率": "New customer share",
    "月次売上（新規／リピート）": "Monthly revenue (new vs repeat)",
    "読み方：新規比率は62%から48%へ低下し、リピート売上比率は5割前後まで上昇しています。売上の成長が新規獲得だけでなく定着にも支えられ始めた、という見本データの意図です":
        "How to read: the new-customer share falls from 62% to 48% while the repeat revenue share rises to around 50%. By design, the sample shows growth starting to be carried by retention as well as acquisition",
    # ---- data definitions ----
    "データ定義 ― 指標の定義と計算式": "Data Definitions — Metrics and Formulas",
    "デモ（合成）データ。GA4の一般的な指標定義に準拠しています":
        "Demo (synthetic) data. Metric definitions follow the standard GA4 definitions",
    "注文ごとに、購入されたSKUの単価（products マスタ）×数量の合計。税込・送料別。返品・キャンセルは含まない（不採用理由は「採用・不採用理由」シート）":
        "Per order, the sum of unit price (products master) × quantity for the purchased SKUs. Tax included, shipping excluded. Returns and cancellations are not included (see the Adoption Notes sheet)",
    "購入（トランザクション）数。daily_traffic の purchases と daily_orders の orders は日別×チャネルで完全一致するよう生成している":
        "Number of purchases (transactions). daily_traffic.purchases and daily_orders.orders are generated to match exactly by day × channel",
    "客単価（AOV）": "AOV (average order value)",
    "売上 ÷ 注文件数": "Revenue ÷ orders",
    "GA4定義のセッション数（30分無操作で終了）。日次×チャネルで集計":
        "Sessions as defined by GA4 (ending after 30 minutes of inactivity), aggregated daily by channel",
    "ユーザー": "Users",
    "アクティブユーザー数。セッション数以下になるよう生成":
        "Active users, generated to be no greater than the session count",
    "カート追加／チェックアウト開始／購入": "Add to cart / Begin checkout / Purchase",
    "GA4イベント add_to_cart／begin_checkout／purchase の発生回数。購入＝注文件数":
        "Occurrences of the GA4 events add_to_cart / begin_checkout / purchase. Purchases = orders",
    "カート追加 ÷ セッション": "Add to cart ÷ sessions",
    "購入 ÷ チェックアウト開始": "Purchases ÷ begin checkout",
    "購入 ÷ セッション（注文件数 ÷ セッション）": "Purchases ÷ sessions (orders ÷ sessions)",
    "粗利・粗利率": "Gross profit / margin",
    "粗利＝Σ SKU売上×（1−原価率）。原価率はSKUマスタの固定値。販管費・広告費は含まない売上総利益ベース":
        "Gross profit = Σ SKU revenue × (1 − cost ratio). Cost ratios are fixed values in the SKU master. Gross-profit basis only — SG&A and advertising costs are not included",
    "対象の売上 ÷ 全体売上": "Item revenue ÷ total revenue",
    "新規／リピート顧客": "New / repeat customers",
    "その月に初めて購入した顧客を新規、2回目以降をリピートとして合成。顧客数は注文件数÷1.15（1人あたり月間注文件数の前提）で換算した推定値":
        "Customers purchasing for the first time in a month are counted as new; second and later purchases as repeat (synthetic). Customer counts are estimates converted from orders ÷ 1.15 (assumed orders per customer per month)",
    "時間帯別": "By hour",
    "全期間のセッション・購入を0〜23時に配分した合計。日次合計と一致するよう最大剰余法で配分":
        "Full-period sessions and purchases allocated across hours 0–23, apportioned by the largest remainder method so the totals match the daily figures",
    "GA4のデフォルトチャネルグループ風の7分類（Organic Search／Paid Search／Organic Social／Paid Social／Referral／Direct／Email）。セッション単位で1チャネルに帰属":
        "Seven GA4-style default channel groups (Organic Search / Paid Search / Organic Social / Paid Social / Referral / Direct / Email). Each session is attributed to exactly one channel",
    # ---- adoption notes ----
    "データの採用・不採用理由": "Data Adopted / Not Adopted — and Why",
    "本ダッシュボードに含めたデータと、検討したが含めなかったデータの理由です（全成果物共通の制作ルール）":
        "What data this dashboard includes, and what was considered but left out — with reasons (a production rule shared by all samples)",
    "採用：合成（デモ）データ全般": "Adopted: synthetic (demo) data throughout",
    "ECの売上・流入ダッシュボードの見本には「日次×チャネル×商品」の粒度の注文・GA4データが必要ですが、実在ショップのその粒度のデータは機密で、入手も公開もできません。生成ルールを公開した合成データであれば、数値の作られ方を検証可能にしたまま、ダッシュボードの型（構造・集計・ファネルの計算）を示す見本として成立します。全シートに「デモ（合成）データ」と明記しています":
        "A sales/acquisition dashboard sample needs order and GA4 data at daily × channel × product granularity. Data of real shops at that granularity is confidential and can neither be obtained nor published. Synthetic data with published generation rules keeps the numbers auditable while demonstrating the dashboard pattern (structure, aggregation, funnel math). Every sheet is labelled 'Demo (Synthetic) Data'",
    "採用：日次粒度（月次ではなく）": "Adopted: daily granularity (not monthly)",
    "曜日パターンとセール期間の山は日次でないと見えないため、日次で生成し、月次は集計で作っています。実務のGA4・基幹システム連携でも日次が最小粒度の標準です":
        "Weekday patterns and sale-period peaks are invisible at monthly granularity, so the data is generated daily and months are built by aggregation. Daily is also the standard minimum granularity for GA4 and core-system integrations in practice",
    "採用：チャネル7分類（GA4デフォルト風）": "Adopted: 7 GA4-style channel groups",
    "実務で最も使われる帰属の出発点がGA4のデフォルトチャネルグループのため、それに合わせた7分類を採用。流入の量（セッション）と質（CVR）の対比が見本の核になります":
        "GA4's default channel grouping is the most common attribution starting point in practice, so the sample adopts the same 7 groups. The contrast between traffic volume (sessions) and quality (CVR) is the core of the sample",
    "採用：ファネル（カート→チェックアウト→購入）": "Adopted: funnel (cart → checkout → purchase)",
    "離脱がどの段階で起きるかはEC改善の第一論点のため、セッションだけでなく中間イベントまで含めています":
        "Where drop-off occurs is the first question in EC improvement, so the funnel includes the intermediate events, not just sessions",
    "採用：時間帯別（期間合計）": "Adopted: hourly view (period totals)",
    "配信・広告の時間帯設計に使う粒度として採用。日次×時間帯まで細かくするとデータ量が膨らみ見本の焦点がぼけるため、期間合計に留めています":
        "Adopted at the granularity used for send-time and ad-scheduling decisions. Going to daily × hourly would bloat the data and blur the sample's focus, so it stays at period totals",
    "不採用：個人単位の顧客ID分析（RFM・LTV）": "Not adopted: per-customer ID analysis (RFM / LTV)",
    "顧客ID単位の購買履歴は個人情報に近く、合成でも「実在の顧客分析」を装う形になるため見本には含めません。新規／リピートは月次集計の範囲で示しています。実務案件では顧客の同意・匿名化前提で別途対応します":
        "Per-customer purchase histories sit close to personal data; even synthesized, they would look like analysis of real customers, so they are excluded. New vs repeat is shown only as monthly aggregates. In client work this is handled separately, with consent and anonymization as prerequisites",
    "不採用：返品・キャンセルデータ": "Not adopted: returns / cancellations",
    "返品率の水準は商材・ショップ方針で大きく異なり、根拠のない数値を置くと見本の数値全体の信頼を損なうため不採用。売上は購入確定ベースで統一しています":
        "Return rates vary widely by product and shop policy; inserting an ungrounded figure would undermine trust in every other number in the sample. Revenue is kept on a confirmed-purchase basis throughout",
    "不採用：広告費の実額（ROAS計算）": "Not adopted: actual ad spend (ROAS)",
    "広告費は媒体契約・入札で決まる経営情報で、合成の前提を置く根拠が弱いため不採用。Paid系チャネルはセッションとCVRの提示に留め、費用対効果は実務で実額を連携して初めて計算します":
        "Ad spend is management information determined by media contracts and bidding; there is no sound basis for synthesizing it. Paid channels are presented with sessions and CVR only — cost-effectiveness is computed in client work once actual spend is connected",
    "不採用：GA4デモアカウント等の公開デモデータ": "Not adopted: public demo data (GA4 demo account etc.)",
    "実データである点は魅力ですが、商品マスタ（原価）を持たず粗利が出せないこと、データの再配布・販売サンプルへの転用条件が不明確なことから不採用とし、全量を自前生成に統一しました":
        "Real data is attractive, but the GA4 demo has no product master (no costs), so gross profit cannot be computed, and the terms for redistributing the data or reusing it in a sales sample are unclear. All data is therefore self-generated",
    "不採用：デバイス別・地域別・ランディングページ別": "Not adopted: device / region / landing-page breakdowns",
    "実務では有用ですが、見本の主題（売上×チャネル×ファネル）から外れ、シート数が膨らむため不採用。追加シートとして拡張可能な設計（日次×チャネルのtidyデータ）にはしています":
        "Useful in practice, but outside the sample's theme (revenue × channel × funnel) and they would bloat the sheet count. The tidy daily × channel data is designed so they can be added as extra sheets",
    # ---- validation log ----
    "検証ログ": "Validation Log",
    "数値はすべて data/processed/ のCSVから再計算して照合しています。CSV同士の整合（C1〜C8）は scripts/validate.py で別途検証し、全件一致しています":
        "All figures are recomputed from the CSVs in data/processed/ and cross-checked. Consistency among the CSVs themselves (C1–C8) is verified separately by scripts/validate.py — all matched",
    "検証項目": "Check",
    "内容": "Detail",
    "照合セル数": "Cells checked",
    "不一致": "Mismatches",
    "V1 KPIサマリー集計の一致": "V1 KPI summary aggregation",
    "月次売上・注文件数・セッションの合計をCSV原票から再集計し照合":
        "Re-aggregated monthly revenue, orders and sessions from the source CSVs and compared",
    "V2 ファネル数値・CVRの再計算": "V2 Funnel figures & CVR recomputation",
    "V3 カテゴリ別・SKU別売上の一致": "V3 Category / SKU revenue agreement",
    "カテゴリ別売上・SKU売上・カテゴリ粗利の合計が全体売上・全体粗利と一致するか照合":
        "Checked that category revenue, SKU revenue and category gross profit totals agree with total revenue and total gross profit",
    "V4 時間帯・曜日集計の一致": "V4 Hourly / weekday aggregation agreement",
    "時間帯別（チャネル別セッション・購入 14件）と曜日別売上の合計を照合":
        "Compared hourly totals (sessions & purchases by channel, 14 items) and the weekday revenue total",
    "V5 新規／リピート売上の一致": "V5 New / repeat revenue agreement",
    "monthly_customers の月次売上（新規＋リピート）がKPIサマリーの月次売上と一致するか19か月分照合":
        "Checked for all 19 months that monthly revenue in monthly_customers (new + repeat) matches the KPI summary monthly revenue",
    "英語版との突合はフェーズB（英語版制作時）に実施します":
        "The JP↔EN numeric cross-check is recorded as V6 in this sheet (see above)",
}
TERM_EN.update(CAT_EN)
TERM_EN.update(SKU_EN)


def translate_text(s):
    if s in TERM_EN:
        return TERM_EN[s]
    # cover key-figures line (values embedded)
    m = re.fullmatch(
        r"売上 ([\d,]+)円／注文件数 ([\d,]+)件／客単価 ([\d,]+)円／セッション ([\d,]+)／全体CVR ([\d.]+)%", s)
    if m:
        return (f"Revenue ¥{m.group(1)} / Orders {m.group(2)} / AOV ¥{m.group(3)} "
                f"/ Sessions {m.group(4)} / Overall CVR {m.group(5)}%")
    # cover highlight line (top channel / top category embedded)
    m = re.fullmatch(
        r"売上最大チャネルは (.+?)（([\d,]+)円、構成比 ([\d.]+)%）。CVRは Email が最高（([\d.]+)%）、"
        r"Paid Social が最低（([\d.]+)%）で、流入の量と質が逆転しています。売上最大カテゴリは (.+?)（構成比 ([\d.]+)%）", s)
    if m:
        cat = CAT_EN.get(m.group(6), m.group(6))
        return (f"The top channel by revenue is {m.group(1)} (¥{m.group(2)}, {m.group(3)}% share). "
                f"Email has the highest CVR ({m.group(4)}%) and Paid Social the lowest ({m.group(5)}%) — "
                f"traffic volume and quality are inverted. The top category is {cat} ({m.group(7)}% share)")
    # SKU sheet footnote (values embedded)
    m = re.fullmatch(
        r"参考：トップ15の売上合計 ([\d,]+)円（全体の ([\d.]+)%）。残り9 SKUの売上合計 ([\d,]+)円", s)
    if m:
        return (f"Note: the top 15 SKUs total ¥{m.group(1)} ({m.group(2)}% of total revenue). "
                f"The remaining 9 SKUs total ¥{m.group(3)}")
    # validation log V2 detail (cell count embedded)
    m = re.fullmatch(
        r"チャネル別ファネル（全期間＋直近6か月）のセッション・購入・CVR (\d+)セルを再計算し照合", s)
    if m:
        return (f"Recomputed sessions, purchases and CVR for the channel funnel "
                f"(full period + last 6 months) — {m.group(1)} cells — and compared")
    # hour labels on the Weekday & Hour sheet
    m = re.fullmatch(r"(\d+)時", s)
    if m:
        return f"{m.group(1)}:00"
    return s


def _has_jp(s):
    # Japanese = hiragana / katakana / CJK ideographs / fullwidth forms.
    # (Typographic symbols such as — × ÷ • ⚠ in English text are fine.)
    return any(0x3040 <= ord(ch) <= 0x30FF or 0x4E00 <= ord(ch) <= 0x9FFF
               or 0xFF00 <= ord(ch) <= 0xFFEF for ch in s)


def main():
    wb = load_workbook(SRC)
    for old, new in SHEET_MAP.items():
        wb[old].title = new
    leftovers = set()
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str):
                    c.value = translate_text(c.value)
                    if _has_jp(c.value):
                        leftovers.add(c.value)
        for chart in getattr(ws, "_charts", []):
            try:
                for p in chart.title.tx.rich.p:
                    for r in p.r:
                        if r.t:
                            r.t = translate_text(r.t)
                            if _has_jp(r.t):
                                leftovers.add(r.t)
            except Exception:
                pass
            try:
                for ax in (chart.x_axis, chart.y_axis):
                    if ax.title is not None:
                        for p in ax.title.tx.rich.p:
                            for r in p.r:
                                if r.t:
                                    r.t = translate_text(r.t)
            except Exception:
                pass
    # Weekday labels: the generic map turns 月 into "Month" (its header
    # meaning); on the Weekday & Hour sheet cells B7:B13 are weekday labels.
    ws = wb["Weekday & Hour"]
    for i, en in enumerate(["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]):
        ws.cell(row=7 + i, column=2, value=en)
    wb.save(OUTF)
    print("saved:", OUTF)
    if leftovers:
        print("UNTRANSLATED strings remaining:", len(leftovers))
        for s in sorted(leftovers):
            print("  ", repr(s))
    else:
        print("no untranslated strings")

    # ---- V6: JP <-> EN numeric cross-check (every numeric cell) ----
    wb_jp = load_workbook(SRC, data_only=True)
    wb_en = load_workbook(OUTF, data_only=True)
    total = 0
    bad = 0
    for name_jp, name_en in SHEET_MAP.items():
        wsj = wb_jp[name_jp]
        wse = wb_en[name_en]
        max_row = max(wsj.max_row, wse.max_row)
        max_col = max(wsj.max_column, wse.max_column)
        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                vj = wsj.cell(row=r, column=c).value
                ve = wse.cell(row=r, column=c).value
                if isinstance(vj, (int, float)) or isinstance(ve, (int, float)):
                    total += 1
                    if vj != ve:
                        bad += 1
    print(f"V6 JP<->EN numeric cells: {total} compared, mismatches={bad}")

    # write V6 into the EN Validation Log (row 11; the note sits at row 12)
    ws = wb["Validation Log"]
    thin = Side(style="thin", color="B4C6E7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    vals = ["V6 JP↔EN numeric cross-check",
            f"Compared all {total} numeric cells between the Japanese and English editions (same sheet/cell positions); labels only were translated",
            total, bad]
    for j, v in enumerate(vals, start=2):
        c = ws.cell(row=11, column=j, value=v)
        c.border = border
    wb.save(OUTF)
    print("V6 written to EN Validation Log")
    assert bad == 0, "JP/EN numeric mismatch!"
    assert not leftovers, "untranslated strings remain"


if __name__ == "__main__":
    main()
