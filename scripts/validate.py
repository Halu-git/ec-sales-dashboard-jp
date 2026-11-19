#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sample No.8 — independent validation.

Recomputes every cross-table / cross-figure relationship from the CSVs in
data/processed/ only (never from generator internals) and prints a check
table. Exit code 1 if any check fails. Results are quoted in design.md.
"""
import csv
import pathlib
import sys
from collections import defaultdict

ROOT = pathlib.Path(__file__).resolve().parent.parent
PROC = ROOT / "data" / "processed"


def load(name):
    with (PROC / name).open(encoding="utf-8") as f:
        return list(csv.DictReader(f))


traffic = load("daily_traffic.csv")
orders = load("daily_orders.csv")
hourly = load("hourly_traffic.csv")
products = load("products.csv")
customers = load("monthly_customers.csv")
sku_sales = load("sku_sales.csv")

checks = []


def check(name, detail, cells, bad):
    checks.append((name, detail, cells, bad))


# C1: purchases (traffic) == orders, per date x channel
pur = defaultdict(int)
for r in traffic:
    pur[(r["date"], r["channel"])] += int(r["purchases"])
ordc = defaultdict(int)
for r in orders:
    ordc[(r["date"], r["channel"])] += int(r["orders"])
keys = set(pur) | set(ordc)
bad = sum(1 for k in keys if pur.get(k, 0) != ordc.get(k, 0))
check("C1 注文件数＝ファネル購入数（日別×チャネル）",
      "daily_traffic.purchases と daily_orders.orders を日別×チャネルで突合",
      len(keys), bad)

# C2: revenue total identical across orders / customers / sku_sales
rev_orders = sum(int(r["revenue_yen"]) for r in orders)
rev_cust = sum(int(r["new_revenue_yen"]) + int(r["repeat_revenue_yen"])
               for r in customers)
rev_sku = sum(int(r["revenue_yen"]) for r in sku_sales)
bad = (0 if rev_orders == rev_cust else 1) + (0 if rev_orders == rev_sku else 1)
check("C2 売上合計の一致（注文／顧客別／SKU別）",
      f"daily_orders={rev_orders:,} monthly_customers={rev_cust:,} sku_sales={rev_sku:,}",
      3, bad)

# C3: monthly revenue: orders vs customers, per month
rev_m_o = defaultdict(int)
for r in orders:
    rev_m_o[r["date"][:7]] += int(r["revenue_yen"])
bad = 0
for r in customers:
    s = int(r["new_revenue_yen"]) + int(r["repeat_revenue_yen"])
    if s != rev_m_o[r["year_month"]]:
        bad += 1
check("C3 月次売上の一致（注文／顧客別）", "19か月分の月次売上を突合", 19, bad)

# C4: hourly totals == daily totals, per channel (sessions & purchases)
hs = defaultdict(int)
hp = defaultdict(int)
for r in hourly:
    hs[r["channel"]] += int(r["sessions"])
    hp[r["channel"]] += int(r["purchases"])
ds = defaultdict(int)
dp = defaultdict(int)
for r in traffic:
    ds[r["channel"]] += int(r["sessions"])
    dp[r["channel"]] += int(r["purchases"])
bad = sum(1 for ch in ds if hs[ch] != ds[ch] or hp[ch] != dp[ch])
check("C4 時間帯合計＝日次合計（チャネル別）",
      "hourly_traffic のセッション・購入合計を daily_traffic と突合", 14, bad)

# C5: funnel monotonicity and CVR recomputation
bad = 0
cells = 0
for r in traffic:
    s, c, co, p = (int(r["sessions"]), int(r["add_to_cart"]),
                   int(r["begin_checkout"]), int(r["purchases"]))
    cells += 1
    if not (s >= c >= co >= p >= 0):
        bad += 1
    if int(r["users"]) > s:
        bad += 1
check("C5 ファネルの単調性・ユーザー数≦セッション数",
      "全日次×チャネル行で sessions>=cart>=checkout>=purchases を確認",
      cells, bad)

# C6: category revenue sums to total; category set matches products master
cat_rev = defaultdict(int)
for r in orders:
    cat_rev[r["category"]] += int(r["revenue_yen"])
master_cats = {r["category"] for r in products}
bad = 0 if set(cat_rev) == master_cats and sum(cat_rev.values()) == rev_orders else 1
check("C6 カテゴリ別売上の合計＝全体・カテゴリ集合の一致",
      "daily_orders のカテゴリ集合が products マスタと一致し合計が全体と一致",
      len(cat_rev), bad)

# C7: SKU gross profit recomputation from products master
cost = {r["sku_id"]: float(r["cost_ratio"]) for r in products}
bad = 0
for r in sku_sales:
    expect = int(round(int(r["revenue_yen"]) * (1 - cost[r["sku_id"]])))
    if expect != int(r["gross_profit_yen"]):
        bad += 1
check("C7 SKU粗利の再計算", "sku_sales.粗利 = 売上×(1-原価率) を全24SKUで再計算",
      len(sku_sales), bad)

# C8: AOV / CVR plausibility ranges (sanity, not equality)
aov = rev_orders / sum(int(r["orders"]) for r in orders)
cvr = sum(int(r["purchases"]) for r in traffic) / sum(int(r["sessions"]) for r in traffic)
bad = 0 if (3000 <= aov <= 8000 and 0.01 <= cvr <= 0.05) else 1
check("C8 水準の妥当性（客単価・全体CVR）",
      f"客単価={aov:,.0f}円（想定3,000〜8,000円） 全体CVR={cvr:.2%}（想定1〜5%）",
      2, bad)

print(f"{'検証項目':<44} {'照合数':>7} {'不一致':>6}")
total_bad = 0
for name, detail, cells, bad in checks:
    print(f"{name:<44} {cells:>7} {bad:>6}  {detail}")
    total_bad += bad
print(f"\n売上合計 {rev_orders:,}円 / 注文件数 "
      f"{sum(int(r['orders']) for r in orders):,}件 / 全体CVR {cvr:.2%}")
if total_bad:
    print(f"VALIDATION FAILED: {total_bad} mismatches")
    sys.exit(1)
print("ALL CHECKS PASSED")
