#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sample No.8 "EC Sales Dashboard" — synthetic data generator.

WHY SYNTHETIC DATA
------------------
This sample is a sales / e-commerce dashboard (orders, revenue and GA4-style
traffic funnels) built as a showcase for Power BI / Excel dashboard services.
Real shop data (orders, GA4 exports) is confidential and cannot be published,
so the dataset is generated for a fictional home-goods e-commerce shop under
the fully documented rules below. No real company or brand names are used.
Every deliverable is labelled "demo (synthetic) data". Fixed seed: the exact
same numbers are reproduced by re-running this script.

GENERATION RULES
----------------
- Shop: fictional home-goods EC shop (kitchen / storage / bath / bedding /
  interior / seasonal goods). 24 SKUs in 6 categories (products.csv).
- Period: 2025-04-01 .. 2026-10-31 (580 days, 19 months), daily grain.
- Channels (GA4 default channel grouping style): Organic Search,
  Paid Search, Organic Social, Paid Social, Referral, Direct, Email.
- Traffic first, orders second (this is what guarantees consistency):
  1. daily sessions per channel = monthly base / 30.4 x growth trend
     x calendar-month seasonality x weekday factor x event factor
     x lognormal noise (sigma 6%).
  2. Funnel per channel with fixed stage rates (binomial draws):
     sessions -> add_to_cart -> begin_checkout -> purchases.
     Channel CVR levels differ by design (Email highest ~5.5%,
     Paid Social lowest ~0.9%).
  3. daily_orders: the day's purchases (transactions) of each channel are
     split across categories with a multinomial draw (category weights have
     their own monthly seasonality and a small channel tilt). Each order
     gets 1 + Poisson(0.45) items; every item is drawn from the category's
     SKUs (popularity weights) and priced from products.csv, so revenue and
     COGS are always consistent with the SKU master.
     => sum(daily_orders.orders) == sum(daily_traffic.purchases) exactly,
        per day and per channel, by construction.
- hourly_traffic: period totals per channel are distributed over 0-23h with
  channel-specific hour profiles (Email peaks 7-8h, Paid Social 21-23h,
  search peaks 12h/20h) using largest-remainder rounding, so hourly totals
  match the daily totals exactly.
- monthly_customers: monthly customers are split into new / repeat with the
  new share declining 62% -> 48% over the period; repeat customers are
  assumed to spend 1.25x the new-customer AOV, and the revenue split is
  rounded so new + repeat revenue == monthly revenue from daily_orders.
- Events baked in:
  * 2025-11-21..30 sale week: paid channels sessions x1.35, others x1.15.
  * 2026-06-19..28 sale week: paid channels sessions x1.30, others x1.12.
  * 2026-09: Paid Social campaign push, sessions x1.25 all month.
  * 2025-12: year-end demand is in the month seasonality (x1.45) and in the
    seasonal-goods category weight.
- Calendar-month seasonality (sessions): Apr 1.00, May 0.97, Jun 1.20
  (sale), Jul 0.95, Aug 0.90, Sep 1.00, Oct 1.02, Nov 1.25 (sale),
  Dec 1.45 (year-end), Jan 0.92, Feb 0.88, Mar 1.05.
- Weekday factor (Mon..Sun): 0.95 0.92 0.93 0.97 1.05 1.25 1.20
  (weekend-heavy, typical for consumer EC).
- Growth trend (monthly, compounding): Organic Search +0.8%, Paid Search
  +1.0%, Organic Social +1.8%, Paid Social +2.5%, Referral +0.3%,
  Direct +0.6%, Email +1.5% (list growth).
"""
import csv
import datetime as dt
import pathlib

import numpy as np

ROOT = pathlib.Path(__file__).resolve().parent.parent
PROC = ROOT / "data" / "processed"
PROC.mkdir(parents=True, exist_ok=True)

SEED = 20261119
rng = np.random.default_rng(SEED)

START = dt.date(2025, 4, 1)
END = dt.date(2026, 10, 31)
DAYS = [START + dt.timedelta(days=i) for i in range((END - START).days + 1)]

CHANNELS = ["Organic Search", "Paid Search", "Organic Social", "Paid Social",
            "Referral", "Direct", "Email"]

# monthly base sessions, monthly growth, funnel stage rates, user ratio
CH_PARAM = {
    #            base    growth  cart   checkout  purchase  user_ratio
    "Organic Search": (26000, 0.008, 0.105, 0.58, 0.525, 0.86),
    "Paid Search":    (15000, 0.010, 0.115, 0.60, 0.55, 0.90),
    "Organic Social": ( 9000, 0.018, 0.055, 0.50, 0.44, 0.80),
    "Paid Social":    (14000, 0.025, 0.042, 0.46, 0.465, 0.82),
    "Referral":       ( 6000, 0.003, 0.080, 0.55, 0.50, 0.84),
    "Direct":         (11000, 0.006, 0.100, 0.58, 0.515, 0.88),
    "Email":          ( 4500, 0.015, 0.180, 0.65, 0.47, 0.92),
}

MONTH_SEASON = {1: 0.92, 2: 0.88, 3: 1.05, 4: 1.00, 5: 0.97, 6: 1.20,
                7: 0.95, 8: 0.90, 9: 1.00, 10: 1.02, 11: 1.25, 12: 1.45}
WEEKDAY_FACTOR = [0.95, 0.92, 0.93, 0.97, 1.05, 1.25, 1.20]  # Mon..Sun

CATEGORIES = ["キッチン用品", "収納用品", "バス・トイレ用品", "寝具",
              "インテリア雑貨", "季節用品"]
CAT_BASE_W = {"キッチン用品": 0.24, "収納用品": 0.20, "バス・トイレ用品": 0.16,
              "寝具": 0.15, "インテリア雑貨": 0.15, "季節用品": 0.10}
# category weight multiplier by calendar month (seasonal goods peak in
# Jul (fans) and Dec (heaters/humidifiers); bedding peaks in winter)
CAT_MONTH_MULT = {
    "キッチン用品":    {m: 1.0 for m in range(1, 13)},
    "収納用品":        {**{m: 1.0 for m in range(1, 13)}, 1: 1.25, 3: 1.2, 4: 1.15},
    "バス・トイレ用品": {m: 1.0 for m in range(1, 13)},
    "寝具":            {**{m: 1.0 for m in range(1, 13)}, 11: 1.3, 12: 1.45, 1: 1.35, 2: 1.2, 6: 0.8, 7: 0.75, 8: 0.8},
    "インテリア雑貨":  {**{m: 1.0 for m in range(1, 13)}, 12: 1.3, 4: 1.15},
    "季節用品":        {**{m: 0.55 for m in range(1, 13)}, 6: 1.5, 7: 1.9, 8: 1.8, 11: 1.4, 12: 1.8, 1: 1.3},
}
# small channel tilt on category mix (rows need not sum to 1; renormalised)
CAT_CHANNEL_TILT = {
    "Organic Search": {"キッチン用品": 1.1, "収納用品": 1.1},
    "Paid Search":    {"キッチン用品": 1.15, "寝具": 1.1},
    "Organic Social": {"インテリア雑貨": 1.5, "季節用品": 1.2},
    "Paid Social":    {"インテリア雑貨": 1.6, "季節用品": 1.25},
    "Referral":       {"バス・トイレ用品": 1.2},
    "Direct":         {"寝具": 1.15, "収納用品": 1.1},
    "Email":          {"寝具": 1.3, "季節用品": 1.3, "収納用品": 1.1},
}

# sku_id, name, category, unit price (JPY), cost ratio, popularity weight
PRODUCTS = [
    ("SKU-001", "真空断熱タンブラー 600ml", "キッチン用品", 2480, 0.55, 1.4),
    ("SKU-002", "シリコン保存容器 6点セット", "キッチン用品", 1980, 0.52, 1.2),
    ("SKU-003", "竹製カッティングボード", "キッチン用品", 2980, 0.58, 0.9),
    ("SKU-004", "計量スプーン・カップセット", "キッチン用品", 1280, 0.50, 1.0),
    ("SKU-005", "積み重ね収納ボックス M", "収納用品", 1780, 0.56, 1.3),
    ("SKU-006", "クローゼット整理ケース 3個組", "収納用品", 3280, 0.57, 1.1),
    ("SKU-007", "衣類圧縮収納袋 4枚セット", "収納用品", 1580, 0.51, 1.2),
    ("SKU-008", "ケーブル収納ポーチ", "収納用品", 1180, 0.49, 0.9),
    ("SKU-009", "珪藻土バスマット", "バス・トイレ用品", 2680, 0.55, 1.3),
    ("SKU-010", "マイクロファイバータオル 5枚セット", "バス・トイレ用品", 1480, 0.50, 1.2),
    ("SKU-011", "詰替ディスペンサー 2本セット", "バス・トイレ用品", 1680, 0.53, 1.0),
    ("SKU-012", "トイレ収納ラック スリム型", "バス・トイレ用品", 3980, 0.60, 0.8),
    ("SKU-013", "洗える低反発枕", "寝具", 4980, 0.58, 1.2),
    ("SKU-014", "冷感敷きパッド シングル", "寝具", 3980, 0.56, 1.1),
    ("SKU-015", "綿布団カバー 3点セット", "寝具", 5980, 0.60, 0.9),
    ("SKU-016", "低反発マットレスパッド", "寝具", 8980, 0.62, 0.8),
    ("SKU-017", "LED間接照明 フロアライト", "インテリア雑貨", 6980, 0.60, 1.0),
    ("SKU-018", "超音波アロマディフューザー", "インテリア雑貨", 3480, 0.54, 1.2),
    ("SKU-019", "壁掛けミラー ラウンド型", "インテリア雑貨", 4480, 0.59, 0.9),
    ("SKU-020", "フェイクグリーン ポット付き", "インテリア雑貨", 2280, 0.52, 1.1),
    ("SKU-021", "USB卓上扇風機", "季節用品", 2980, 0.55, 1.2),
    ("SKU-022", "超音波加湿器 4L", "季節用品", 5480, 0.58, 1.1),
    ("SKU-023", "洗える電気毛布 140×80cm", "季節用品", 6980, 0.61, 0.9),
    ("SKU-024", "除湿シート 2枚セット", "季節用品", 1280, 0.50, 1.0),
]
SKU_BY_CAT = {c: [p for p in PRODUCTS if p[2] == c] for c in CATEGORIES}
PRICE = {p[0]: p[3] for p in PRODUCTS}
COST = {p[0]: p[4] for p in PRODUCTS}


def month_index(d):
    return (d.year - 2025) * 12 + (d.month - 4)  # 2025-04 -> 0


def event_factor(d, channel):
    f = 1.0
    paid = channel in ("Paid Search", "Paid Social")
    if dt.date(2025, 11, 21) <= d <= dt.date(2025, 11, 30):
        f *= 1.35 if paid else 1.15
    if dt.date(2026, 6, 19) <= d <= dt.date(2026, 6, 28):
        f *= 1.30 if paid else 1.12
    if d.year == 2026 and d.month == 9 and channel == "Paid Social":
        f *= 1.25
    return f


def cat_probs(d, channel):
    w = []
    for c in CATEGORIES:
        x = CAT_BASE_W[c] * CAT_MONTH_MULT[c][d.month]
        x *= CAT_CHANNEL_TILT.get(channel, {}).get(c, 1.0)
        w.append(x)
    w = np.array(w, dtype=float)
    return w / w.sum()


# ---------------- 1) daily traffic ----------------
traffic_rows = []
for d in DAYS:
    mi = month_index(d)
    for ch in CHANNELS:
        base, growth, cart_r, co_r, pu_r, ur = CH_PARAM[ch]
        lam = (base / 30.4) * ((1 + growth) ** mi) * MONTH_SEASON[d.month] \
            * WEEKDAY_FACTOR[d.weekday()] * event_factor(d, ch)
        sessions = int(round(lam * rng.lognormal(0, 0.06)))
        users = min(sessions, int(round(sessions * ur * rng.normal(1, 0.02))))
        carts = int(rng.binomial(sessions, min(0.95, cart_r * rng.normal(1, 0.05))))
        checkouts = int(rng.binomial(carts, min(0.95, co_r * rng.normal(1, 0.05))))
        purchases = int(rng.binomial(checkouts, min(0.95, pu_r * rng.normal(1, 0.05))))
        traffic_rows.append({
            "date": d.isoformat(), "channel": ch, "sessions": sessions,
            "users": users, "add_to_cart": carts,
            "begin_checkout": checkouts, "purchases": purchases,
        })

# ---------------- 2) daily orders (from traffic purchases) ----------------
order_rows = []
sku_units = {p[0]: 0 for p in PRODUCTS}
sku_revenue = {p[0]: 0 for p in PRODUCTS}
for tr in traffic_rows:
    d = dt.date.fromisoformat(tr["date"])
    ch = tr["channel"]
    p = tr["purchases"]
    if p == 0:
        continue
    counts = rng.multinomial(p, cat_probs(d, ch))
    for cat, n in zip(CATEGORIES, counts):
        if n == 0:
            continue
        skus = SKU_BY_CAT[cat]
        weights = np.array([s[5] for s in skus], dtype=float)
        weights = weights / weights.sum()
        revenue = 0
        units = 0
        for _ in range(int(n)):
            k = 1 + int(rng.poisson(0.45))
            picks = rng.choice(len(skus), size=k, p=weights)
            for ix in picks:
                sid = skus[ix][0]
                revenue += PRICE[sid]
                units += 1
                sku_units[sid] += 1
                sku_revenue[sid] += PRICE[sid]
        order_rows.append({
            "date": tr["date"], "channel": ch, "category": cat,
            "orders": int(n), "units": units, "revenue_yen": revenue,
        })

# ---------------- 3) hourly traffic (period totals) ----------------
BASE_HOUR = np.array([0.6, 0.3, 0.2, 0.1, 0.1, 0.2, 0.5, 0.9, 1.2, 1.4, 1.6,
                      1.8, 2.2, 1.9, 1.7, 1.6, 1.7, 1.9, 2.1, 2.4, 2.8, 3.0,
                      2.6, 1.6])
HOUR_MULT = {
    "Email":          {7: 2.2, 8: 2.0, 12: 1.3, 21: 1.2},
    "Paid Social":    {21: 1.5, 22: 1.7, 23: 1.6, 12: 0.8},
    "Organic Social": {21: 1.4, 22: 1.5, 23: 1.3},
    "Organic Search": {12: 1.4, 20: 1.3},
    "Paid Search":    {12: 1.3, 20: 1.25},
}
HOUR_CVR = np.array([0.8, 0.7, 0.6, 0.6, 0.6, 0.7, 0.85, 0.95, 1.0, 1.0, 1.0,
                     1.05, 1.1, 1.05, 1.0, 1.0, 1.0, 1.05, 1.1, 1.15, 1.2,
                     1.2, 1.15, 1.0])


def largest_remainder(total, weights):
    w = np.array(weights, dtype=float)
    w = w / w.sum()
    raw = total * w
    base = np.floor(raw).astype(int)
    rem = total - base.sum()
    order = np.argsort(-(raw - base))
    for i in range(int(rem)):
        base[order[i]] += 1
    return base


hourly_rows = []
for ch in CHANNELS:
    tot_s = sum(r["sessions"] for r in traffic_rows if r["channel"] == ch)
    tot_p = sum(r["purchases"] for r in traffic_rows if r["channel"] == ch)
    prof = BASE_HOUR.copy()
    for h, m in HOUR_MULT.get(ch, {}).items():
        prof[h] *= m
    hs = largest_remainder(tot_s, prof)
    hp = largest_remainder(tot_p, prof * HOUR_CVR)
    for h in range(24):
        hourly_rows.append({"hour": h, "channel": ch,
                            "sessions": int(hs[h]), "purchases": int(hp[h])})

# ---------------- 4) monthly customers ----------------
monthly_rev = {}
monthly_orders = {}
for r in order_rows:
    ym = r["date"][:7]
    monthly_rev[ym] = monthly_rev.get(ym, 0) + r["revenue_yen"]
    monthly_orders[ym] = monthly_orders.get(ym, 0) + r["orders"]
months = sorted(monthly_rev)
customer_rows = []
for i, ym in enumerate(months):
    total_orders = monthly_orders[ym]
    total_rev = monthly_rev[ym]
    customers = int(round(total_orders / 1.15))
    new_share = 0.62 - (0.62 - 0.48) * (i / (len(months) - 1))
    new_c = int(round(customers * new_share))
    rep_c = customers - new_c
    # repeat AOV = 1.25 x new AOV
    new_rev = int(round(total_rev * new_c / (new_c + rep_c * 1.25)))
    rep_rev = total_rev - new_rev
    customer_rows.append({
        "year_month": ym, "new_customers": new_c, "repeat_customers": rep_c,
        "new_revenue_yen": new_rev, "repeat_revenue_yen": rep_rev,
    })

# ---------------- write CSVs ----------------
def write_csv(name, rows, fields):
    with (PROC / name).open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)
    print(f"{name}: {len(rows)} rows")


write_csv("daily_traffic.csv", traffic_rows,
          ["date", "channel", "sessions", "users", "add_to_cart",
           "begin_checkout", "purchases"])
write_csv("daily_orders.csv", order_rows,
          ["date", "channel", "category", "orders", "units", "revenue_yen"])
write_csv("hourly_traffic.csv", hourly_rows,
          ["hour", "channel", "sessions", "purchases"])
write_csv("products.csv",
          [{"sku_id": p[0], "sku_name": p[1], "category": p[2],
            "unit_price_yen": p[3], "cost_ratio": p[4]} for p in PRODUCTS],
          ["sku_id", "sku_name", "category", "unit_price_yen", "cost_ratio"])
write_csv("monthly_customers.csv", customer_rows,
          ["year_month", "new_customers", "repeat_customers",
           "new_revenue_yen", "repeat_revenue_yen"])
write_csv("sku_sales.csv",
          [{"sku_id": p[0], "sku_name": p[1], "category": p[2],
            "units": sku_units[p[0]], "revenue_yen": sku_revenue[p[0]],
            "gross_profit_yen": int(round(sku_revenue[p[0]] * (1 - COST[p[0]])))}
           for p in PRODUCTS],
          ["sku_id", "sku_name", "category", "units", "revenue_yen",
           "gross_profit_yen"])

tot_rev = sum(r["revenue_yen"] for r in order_rows)
tot_ord = sum(r["orders"] for r in order_rows)
tot_pur = sum(r["purchases"] for r in traffic_rows)
tot_ses = sum(r["sessions"] for r in traffic_rows)
print(f"TOTAL revenue={tot_rev:,} orders={tot_ord:,} purchases={tot_pur:,} "
      f"sessions={tot_ses:,} CVR={tot_pur / tot_ses:.4f}")
assert tot_ord == tot_pur, "orders != purchases"
print("orders == purchases OK")
