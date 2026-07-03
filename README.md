# Prediction Market Operations Tracker

An operations-analyst tool that pulls **real, live market data from Polymarket**
and produces a self-contained Excel operations workbook where **all analysis is
live Excel formulas** — the kind of monitoring artifact an ops desk at a sports
prediction-market platform actually uses to watch liquidity risk, pricing
integrity, and user trade activity.

> **Snapshot committed to this repo:** `2026-07-03T02:07:21Z` — 100 markets
> (70 sports / 30 non-sports comparison) and 3,000 real trades from the top-10
> sports markets by volume. The exact snapshot metadata lives in
> [`data/raw/snapshot_manifest.json`](data/raw/snapshot_manifest.json) and on the
> workbook's **Dashboard** sheet.

---

## Why this exists

For an operations analyst, the daily questions are: *Where is the platform
exposed? Are our prices internally consistent? Where is the real money moving?*
This tool answers all three from live data and packages them so a non-engineer
can keep the analysis alive in Excel:

- **Liquidity risk** — markets with heavy 24h volume but a shallow order book are
  a real platform-liability signal (hard to hedge, easy to move).
- **Pricing sanity** — YES + NO should sum to ~1.00; deviations flag internal
  pricing discrepancies. Near-certain markets (>0.95 / <0.05) still trading
  heavily are worth a look.
- **Activity** — real user trades broken down by day, category, and order-flow
  side, with whale-trade detection on notional exposure.

Everything recalculates in Excel because it's written as **formulas, not baked-in
numbers** — change a threshold cell and the whole workbook re-evaluates.

---

## Data sources (public, no auth)

| Source | Endpoint | Used for |
|---|---|---|
| Polymarket **Gamma API** | `https://gamma-api.polymarket.com/markets` | market snapshot: question, category, prices, volume, liquidity, end date |
| Polymarket **Data API** | `https://data-api.polymarket.com/trades` | real trades for the top-10 sports markets: timestamp, side, size, price |
| *(optional)* **The Odds API** | `https://api.the-odds-api.com/v4/...` | sportsbook odds for cross-market mispricing (see [Optional](#optional-sportsbook-discrepancy-sheet)) |

Markets are categorized (NBA / NFL / MLB / NHL / Soccer / F1 / Tennis / … vs
Politics / Crypto / Economics / Tech / …) by keyword-matching the market question
plus its Polymarket event ticker/slug, and sports markets are prioritized.

---

## Repo structure

```
fetch_data.py                 API pull → timestamped CSVs + manifest
build_workbook.py             CSVs → output/prediction_market_ops_tracker.xlsx
data/raw/
  markets_latest.csv          committed sample snapshot (markets)
  trades_latest.csv           committed sample snapshot (trades)
  snapshot_manifest.json      snapshot metadata (date, counts, sources)
  markets_YYYYMMDD_*.csv       per-run archives (git-ignored)
  trades_YYYYMMDD_*.csv        per-run archives (git-ignored)
output/
  prediction_market_ops_tracker.xlsx
requirements.txt
README.md
```

Timestamped raw archives are `.gitignore`d (large, reproducible); the
`*_latest.csv` sample snapshot **is committed** so the workbook builds offline
straight from a clone.

---

## How to run / re-run

```bash
pip install -r requirements.txt

# 1) Pull a fresh live snapshot (writes data/raw/*.csv + snapshot_manifest.json)
python3 fetch_data.py                 # optional: --limit 80

# 2) Build the workbook from the snapshot (fully deterministic)
python3 build_workbook.py             # → output/prediction_market_ops_tracker.xlsx
```

- **Build only, no network:** skip step 1 — `build_workbook.py` reads the committed
  sample snapshot and reproduces the exact workbook.
- **Reproducibility:** `build_workbook.py` is deterministic given a CSV snapshot.
  The snapshot date is stamped on the Dashboard and in the manifest, so a build is
  always traceable to a specific pull.
- **Open in Excel 365** to recalculate. The workbook uses `XLOOKUP` (Excel 365 /
  2021+); older Excel will show `#NAME?` for those cells.

### Reliability / rate limits
`fetch_data.py` uses a retrying HTTP session (backoff on 429/5xx), per-request
timeouts, and polite pauses between trade calls. Any API failure is logged and
handled gracefully — a failed trades or odds pull never aborts the market snapshot.

---

## The workbook (4 sheets + 1 optional)

### `Markets` — real market snapshot
One row per market: title, category, YES price (implied probability), NO price,
24h volume, total volume, liquidity, end date, and **`days_to_resolution`** (a
live formula: `end_date − SnapshotDate`).

### `Trades` — real activity data with live formula columns
Trades from the top-10 sports markets (market, timestamp, side, outcome, price,
size). Added **live formula** columns:
- **`notional_usd`** = `price * size`
- **`mkt_category`, `mkt_total_volume`** — pulled from `Markets` via **`XLOOKUP`**
- **`large_trade_flag`** = `IF(notional > 95th-percentile threshold, "WHALE", "")`

### `Analysis` — live `SUMIFS` / `COUNTIFS` / `AVERAGEIFS` / `PERCENTILE`
Editable threshold block at the top drives everything below:
1. **Liquidity risk + pricing sanity** (one row per market):
   - `thin_flag` = `IF(volume_24h ≥ 75th-pct AND liquidity ≤ 25th-pct, "THIN MARKET", "")`
   - `price_flag` = `IF(|YES+NO − 1| > tolerance, "CHECK PRICING", "")`
   - `near_certain_active` = near-certain price **and** heavy volume
2. **Activity trends** — trade volume, count, and average by **day**, by
   **category**, and by **side** (BUY/SELL order-flow imbalance), all via
   `SUMIFS`/`COUNTIFS`/`AVERAGEIFS`.
3. A labeled **`PIVOT TABLES (build manually)`** area.

### `Dashboard` — KPIs, conditional formatting, native charts
- **KPI formulas:** total tracked volume, total & largest 24h volume, markets /
  sports tracked, thin-market flags, pricing-discrepancy flags, whale trades,
  trades tracked.
- **Conditional formatting:** red when thin-market flags exist, yellow for pricing
  discrepancies (also applied on the `Analysis` flag columns and whale trades).
- **2 native Excel charts:** top-10 markets by 24h volume (bar) and daily trade
  volume (line), both driven by live cells.

### Optional: `Discrepancy` sheet
See below — only created when a sportsbook odds snapshot is present.

---

## Optional: sportsbook Discrepancy sheet

If the `THE_ODDS_API_KEY` environment variable is set, `fetch_data.py` also pulls
outright odds from **The Odds API**, de-vigs them into implied probabilities, and
saves `data/raw/odds_latest.csv`. `build_workbook.py` then adds a **`Discrepancy`**
sheet that matches each team to its Polymarket market and computes, as live
formulas, `gap = |Polymarket YES − sportsbook implied|`, flagging **`MISPRICED >5%`**
— literal cross-market mispricing detection.

```bash
export THE_ODDS_API_KEY=your_key_here
python3 fetch_data.py && python3 build_workbook.py
```

> **Status:** the code path is implemented and exercised with a mock odds file,
> but was **not run against the live Odds API** in the committed snapshot (no key
> was available at build time). With a valid key it activates automatically. Team
> name matching (sportsbook ↔ Polymarket) may need light tuning per sport/season.

---

## Skills demonstrated

- **API integration** — two live Polymarket REST endpoints (+ optional Odds API),
  pagination, retries/backoff, rate-limit handling, graceful failure, reproducible
  timestamped snapshots.
- **`XLOOKUP`** — pulling market metadata into the Trades and Discrepancy sheets.
- **`SUMIFS` / `COUNTIFS` / `AVERAGEIFS`** — activity trends by day / category / side.
- **`PERCENTILE` / `MEDIAN` thresholds** — data-driven, editable flag logic.
- **Conditional formatting** — risk-based red/yellow highlighting.
- **Native Excel charts** — bar and line, driven by live cells.
- **Pivot tables** — set up manually (see checklist) with slicers.

---

## Manual checklist (finishing touches in Excel)

The workbook is generated with everything an ops analyst needs to then make it
their own. In Excel:

- [ ] **Build 3 pivot tables** off the `Trades` sheet in the labeled area on
      `Analysis`: (1) volume by market, (2) volume by day × side, (3) whale count
      by market.
- [ ] **Add slicers** on `market` and `side` and wire them to the pivots.
- [ ] **Add screenshots** of the Dashboard (KPIs + both charts) to this README.
- [ ] **Write 3 key findings** from the real data, e.g.:
  - Which markets tripped the **THIN MARKET** flag and why that's a hedging risk.
  - The BUY vs SELL notional imbalance in the top sports markets over the snapshot window.
  - Any **near-certain** markets still absorbing heavy volume (potential
    resolution/settlement risk).

---

*Data © Polymarket, pulled from public endpoints for analysis. This is a
portfolio/operations tool, not trading or investment advice.*