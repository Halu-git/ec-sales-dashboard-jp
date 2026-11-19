# Power BI ビルドノート ― 売上・ECダッシュボード（モデル設計と注意点）

入力： `data/processed/` の tidy CSV 群（6ファイル）。Excel版と同じCSVをそのまま使う。
注意： 本サンプルのデータは**デモ（合成）データ**。実務案件では、GA4エクスポート（BigQuery／Looker Studio経由の集計でも可）とEC基幹システムの注文データを同じ構造に整形して差し替える（差し替え手順は measures.dax 末尾を参照）。

## テーブル構成

| テーブル | 元CSV | 粒度 | 主な列 |
|---|---|---|---|
| FactOrders | daily_orders.csv | 日付×チャネル×カテゴリ＝18,007行 | date, channel, category, orders, units, revenue_yen |
| FactTraffic | daily_traffic.csv | 日付×チャネル＝4,053行 | date, channel, sessions, users, add_to_cart, begin_checkout, purchases |
| FactHourly | hourly_traffic.csv | 時間帯（0〜23）×チャネル＝168行（**期間合計**） | hour, channel, sessions, purchases |
| FactSkuSales | sku_sales.csv | SKU＝24行（**期間合計**） | sku_id, sku_name, category, units, revenue_yen, gross_profit_yen |
| FactCustomers | monthly_customers.csv | 月＝19行 | year_month, new_customers, repeat_customers, new_revenue_yen, repeat_revenue_yen |
| DimProduct | products.csv | SKU＝24行 | sku_id, sku_name, category, unit_price_yen, cost_ratio |

- 主ファクトは2つ： **売上系は FactOrders、流入・ファネル系は FactTraffic**。どちらを主にするかで数値の出所が決まる（下記の設計ルール参照）。
- FactHourly と FactSkuSales は**期間合計の集計済みテーブル**で、日付列を持たない。ここが本モデルの最大の注意点（落とし穴1・2）。

## ディメンション

| テーブル | 作り方 | 列 |
|---|---|---|
| DimDate | date列から生成（2025-04-01〜2026-10-31の580日） | date, year_month（"2025-04"形式）, year, month, day, weekday（月〜日）, weekday_index（月=0〜日=6）, is_weekend |
| DimChannel | channel列から一意化 | channel, sort_order（GA4の表示順： Organic Search=1, Paid Search=2, Organic Social=3, Paid Social=4, Referral=5, Direct=6, Email=7） |
| DimCategory | category列から一意化 | category, sort_order（キッチン用品=1、収納用品=2、バス・トイレ用品=3、寝具=4、インテリア雑貨=5、季節用品=6） |
| DimHour | 0〜23を生成 | hour, hour_label（"0:00"形式） |

## リレーションシップ

- FactOrders[date] → DimDate[date]（多対一、単方向）
- FactOrders[channel] → DimChannel[channel]（多対一、単方向）
- FactOrders[category] → DimCategory[category]（多対一、単方向）
- FactTraffic[date] → DimDate[date]（多対一、単方向）
- FactTraffic[channel] → DimChannel[channel]（多対一、単方向）
- FactHourly[channel] → DimChannel[channel]（多対一、単方向）、FactHourly[hour] → DimHour[hour]（多対一、単方向）。**DimDate には接続しない**
- FactSkuSales[sku_id] → DimProduct[sku_id]（多対一、単方向）、FactSkuSales[category] → DimCategory[category]（多対一、単方向）。**DimDate・DimChannel には接続しない**
- FactCustomers[year_month] → DimDate[year_month]（多対一、単方向。月粒度の接続）

## 設計ルール（必読）

1. **ファネル指標は FactTraffic 側、売上指標は FactOrders 側から取る。** セッション・カート・チェックアウト・購入はトラフィックにしかなく、売上・販売数量は注文にしかない。CVR は「購入（traffic）÷セッション（traffic）」で、同一テーブル内の率として定義する。
2. **daily_traffic と daily_orders は DimDate × DimChannel という共通ディメンションでのみ接続する。** 2ファクトを直接ジョインしない（粒度が違う： traffic は日×チャネル、orders は日×チャネル×カテゴリ。直接つなぐとカテゴリ数だけ行が膨らむ）。
3. **注文件数（orders）＝購入数（purchases）は生成ルールで構造的に一致する前提**（検証 C1 で日別×チャネル 4,053件全件一致を確認済み）。ただし実務の実データでは GA4 のトランザクションと基幹の注文件数は一致しないのが普通なので、差し替え時は「どちらを正とするか」を先に決める（落とし穴4）。

## 落とし穴

1. **SKU粒度と粗利は FactSkuSales（期間合計）にしかない。** daily_orders には sku_id 列がなく、日次×SKU の売上は存在しない。したがって [粗利] [粗利率] は日付・チャネルスライサーの影響を受けない（カテゴリフィルタだけが DimCategory 経由で効く）。ページ3で期間スライサーを置いても粗利が動かないのは仕様であり、壊れているのではない。日次×SKUが必要なら、生成側（実務では基幹の注文明細）から作り直す。
2. **FactHourly を DimDate に接続しない。** 時間帯テーブルは全期間の合計で日付を持たない。無理に接続すると全行が特定日に寄るか、期間フィルタで全て消える。時間帯ページでは DimDate スライサーを置かない（置く場合は「時間帯は期間合計のため連動しない」旨を注記する）。
3. **FactCustomers は月粒度。** year_month で DimDate に接続するため、日単位の範囲選択（例：月の途中まで）では顧客系メジャーだけ表示が合わなくなる。ページ4は月スライサー前提で設計する。
4. **注文件数と購入数の二重表示に注意。** 両者は同値のため、KPIカードに両方を並べると二重計上に見える。カードは [注文件数]（orders側）に統一し、[購入数] はファネル専用にする。合計行で両者を足すのは厳禁。
5. **CVR を跨ぎ計算しない。** [注文件数] ÷ [セッション] でも本サンプルでは同値になるが、それは構造的一致のおかげ。実データに差し替えた途端に壊れるため、CVR は必ず traffic 側で閉じる（ルール1）。
6. **前年同月比は2026年4月以降しか値が出ない。** データ開始が2025年4月のため、それ以前の DATEADD(-1, YEAR) は空白になる。空白はエラーではない（カードでは「—」表示にする）。
7. **並び順は sort_order で固定する。** チャネルは文字列順だと GA4 の慣例順にならず、カテゴリも五十音順では Excel版と並びがずれる。DimChannel/DimCategory の sort_order を「列で並び替え」に設定すること。曜日は DimDate[weekday_index] で月曜始まりに固定する。
8. **時間帯CVR と日次CVR を同じビジュアルで混ぜない。** 時間帯側は期間合計を最大剰余法で配分した値で、日次側の実績とは別テーブル。出典の違う率を同列に並べると誤読される。
9. **金額の単位は円（千円ではない）。** 第7弾（千円）と単位が違うので、カードには「円」を明記する。桁が大きい場合は ÷1,000,000 で百万円表示に切り替える。

## Excel版との表示一致

- ページ1のカード（売上 245,675,600円／注文件数 52,873件／客単価 4,647円／セッション 1,977,400／全体CVR 2.67%／月平均売上 12,930,295円）は Excel版 KPIサマリーのカードと一致させる。
- ファネルの全期間値（セッション 1,977,400→カート追加→チェックアウト開始→購入 52,873）は Excel版 チャネル別ファネルと一致させる。
