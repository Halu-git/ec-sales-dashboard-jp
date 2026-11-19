# 売上・ECダッシュボード

架空の生活雑貨ECショップを題材に、「売上×流入チャネル×ファネル」を見るECダッシュボードのサンプルです。Excel完成品（日本語版・英語版）、Power BI設計パック、整形済みCSV、データ生成・検証スクリプトを含みます。

> ⚠ **本リポジトリの数値はすべてデモ（合成）データです。** 題材は架空のショップであり、実在企業の実績ではありません。生成ルールは `scripts/generate_data.py` に全量公開しており、乱数シード固定で完全に再現できます。

📖 Zenn 記事： 「「どの流入が、いくら売っているのか」に即答する ― 売上・ECダッシュボードをExcelとPower BIで作る」（公開後にURLを追記）

## 主な数値（2025年4月〜2026年10月・全期間）

- 売上： **245,675,600円**／注文件数 **52,873件**／客単価 **4,647円**／セッション **1,977,400**／全体CVR **2.67%**
- 売上最大チャネル： **Organic Search**（82,973,520円、構成比33.8%）
- チャネル別CVR： Email 5.57%（最高）〜 Paid Social 0.89%（最低）―「量のPaid Social、質のEmail」
- 売上最大カテゴリ： **寝具**（66,009,040円、構成比26.9%）。注文件数ではキッチン用品が最多
- 月次売上： 初月（2025年4月）11,261,600円 → 直近月（2026年10月）14,393,380円、ピークは2025年12月（18,845,000円）

## 構成

- `excel/ec_sales_dashboard.xlsx` ― Excel完成品・日本語版（12シート、ネイティブグラフ7点）
- `excel/ec_sales_dashboard_en.xlsx` ― 英語版Excel（数値は日本語版と同一・全数値セル突合済み）
- `powerbi/` ― Power BI設計パック（BUILD_NOTES.md / measures.dax / pages.md、4ページ設計、DAXメジャー31本）
- `data/processed/` ― tidy形式CSV 6ファイル（日次トラフィック／日次注文／時間帯別／SKUマスタ／月次顧客／SKU別集計）
- `scripts/` ― データ生成・Excel生成・検証スクリプト（generate_data.py / build_excel.py / build_excel_en.py / validate.py）
- `design.md` ― 設計メモ（目的・データ設計・採用/不採用理由・検証結果）

## データについて

### なぜ合成データなのか

ECダッシュボードの見本には「日次×チャネル×商品」の粒度の注文・GA4データが必要ですが、実在ショップのその粒度のデータは機密で、入手も公開もできません。そのため、生成ルールを公開した合成データを使い、数値の作られ方を検証可能にしています。生成は「トラフィック→注文」の順に行い、ファネルで確定した購入数をカテゴリへ配分して注文を作るため、トラフィック側の購入数と注文側の注文件数が日別×チャネルでも完全一致します。

### 検討したが採用しなかったデータ

- **GA4デモアカウント等の公開デモデータ**： 実データである点は魅力ですが、商品マスタ（原価）を持たず粗利が出せないこと、データの再配布・販売サンプルへの転用条件が不明確なことから不採用
- **返品・キャンセルデータ**： 返品率の水準は商材・ショップ方針で大きく異なり、根拠のない数値を置くと見本全体の信頼を損なうため不採用（売上は購入確定ベース）
- **広告費の実額（ROAS計算）**： 媒体契約・入札で決まる経営情報で、合成の前提を置く根拠が弱いため不採用。Paid系チャネルはセッションとCVRの提示に留める
- **個人単位の顧客ID分析（RFM・LTV）**： 個人情報に近く、合成でも実在の顧客分析を装う形になるため不採用。新規／リピートは月次集計の範囲で提示

### 生成ルールの要約

- 架空の生活雑貨ECショップ、商品カテゴリ6種（キッチン用品／収納用品／バス・トイレ用品／寝具／インテリア雑貨／季節用品）、SKU24点、チャネル7種（GA4のデフォルトチャネルグループ風）
- 季節性（12月ピーク、6月・11月のセール月）、曜日パターン（週末型）、チャネル別のCVR水準（Email 約5.5%〜Paid Social 約0.9%）、カテゴリ別の季節性、セール週・キャンペーンのイベント
- 新規／リピート： 新規顧客比率は62%→48%へ低下、リピート客単価は新規の1.25倍の前提
- 乱数シード： 20261119（再実行で完全再現可能）

### 定義・計算式

- 売上＝注文ごとに購入SKUの単価×数量の合計（税込・送料別、返品・キャンセルを含まない）
- CVR＝購入÷セッション、カート率＝カート追加÷セッション、購入率＝購入÷チェックアウト開始
- 粗利＝Σ SKU売上×（1−原価率）。原価率はSKUマスタの固定値で、販管費・広告費を含まない売上総利益ベース
- 金額の単位は円

## 検証

- CSV同士の整合（C1〜C8、原票のみから独立再計算）： 注文件数＝ファネル購入数（日別×チャネル4,053件）、売上合計の注文／顧客別／SKU別一致、月次売上19か月、時間帯合計＝日次合計、ファネルの単調性、SKU粗利の再計算 ― いずれも不一致0
- Excel側の検証（V1〜V5）： KPI集計、ファネル・CVR再計算（28セル）、カテゴリ・SKU売上、時間帯・曜日、新規／リピート売上（19か月）― いずれも不一致0（各Excelの「検証ログ」シートに記録）
- 英語版との数値突合： 991セル全件一致・不一致0、未翻訳文字列0（英語版の検証ログ V6 に記録）

## 実務での使い方

GA4の集計（日次×チャネル）とEC基幹の注文集計（日次×チャネル×カテゴリ）を、`daily_traffic.csv` / `daily_orders.csv` と同じ列構成に整形して差し替えれば、同じダッシュボードが実データで動きます。差し替え手順は `powerbi/measures.dax` の末尾に記載しています。なお実データではGA4の購入数と基幹の注文件数は一致しないのが普通であり、「KPIの正は基幹、ファネルの正はGA4」と出所を固定して扱います。

---

## English summary

**EC Sales Dashboard** — a sample e-commerce dashboard built on a fictional home-goods EC shop (6 categories, 24 SKUs, Apr 2025 – Oct 2026): revenue × acquisition channel × funnel, with channel CVR, category/SKU performance, weekday/hour patterns, and new-vs-repeat customer mix.

> ⚠ All figures are **demo (synthetic) data** — not the results of any real company. Generation rules are fully documented in `scripts/generate_data.py` (fixed seed, fully reproducible). Real shop data at daily × channel × product granularity is confidential, which is why auditable synthetic data is used; the GA4 demo account, returns/cancellation data, actual ad spend (ROAS), and per-customer ID analysis (RFM/LTV) were considered and rejected for the reasons documented in `design.md` and the workbooks' Adoption Notes sheets.

Key figures (full period): revenue ¥245,675,600, orders 52,873, AOV ¥4,647, sessions 1,977,400, overall CVR 2.67%. Top channel by revenue is Organic Search (33.8% share); channel CVR ranges from Email 5.57% (highest) to Paid Social 0.89% (lowest) — volume and quality are inverted. Top category is Bedding (26.9%). English reports are available — this sample itself is delivered bilingually: the English Excel edition contains identical figures (all 991 numeric cells cross-checked, zero mismatches).
