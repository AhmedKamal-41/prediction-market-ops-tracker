# Prediction Market Operations Tracker

I built this to practice the kind of work an operations analyst does at a sports
prediction-market platform: pull real market data, watch for liquidity and
pricing problems, and hand the numbers to non-engineers in a spreadsheet they can
actually poke at.

So that's what it does. Two Python scripts pull live data from Polymarket and turn
it into an Excel workbook where the analysis is all real formulas, not values I
pasted in. Change a threshold and the whole sheet recalculates.

The snapshot committed here is from 2026-07-03 (02:07 UTC): 100 markets and 3,000
trades. Details are in `data/raw/snapshot_manifest.json` and on the Dashboard tab.

## What it tracks

Three questions an ops desk cares about day to day:

1. Liquidity risk. A market with lots of 24h volume but a thin order book is hard
   to hedge and easy to push around. Those get flagged.
2. Pricing sanity. YES + NO should add up to about 1.00. When it drifts too far,
   that's a pricing problem worth a look. Same for near-certain markets (price
   above 0.95 or below 0.05) that are still trading heavily.
3. Activity. Where the real money is moving, broken down by day, category, and
   buy/sell, with the biggest trades flagged as whales.

## Data

Everything comes from Polymarket's public APIs, no key or login needed:

- Markets: `https://gamma-api.polymarket.com/markets` — title, category, prices,
  volume, liquidity, end date.
- Trades: `https://data-api.polymarket.com/trades` — real trades for the 10
  biggest sports markets (time, side, size, price).

Markets get sorted into categories (NBA, NFL, MLB, NHL, soccer, F1, and so on,
plus non-sports buckets like politics and crypto for comparison) by matching
keywords in the question and the Polymarket event slug. Sports markets are pulled
first since that's the focus.

There's also an optional third source, The Odds API, covered further down.

## Files

```
fetch_data.py       pulls the data, writes CSVs to data/raw/
build_workbook.py   reads the CSVs, builds the Excel file
data/raw/           snapshot CSVs + manifest (see note below)
output/             the finished workbook
requirements.txt
```

The timestamped CSVs in `data/raw/` are gitignored because they're big and I can
always re-pull them. The `*_latest.csv` sample snapshot is committed so the
workbook builds from a fresh clone without touching the network.

## Running it

```bash
pip install -r requirements.txt

python3 fetch_data.py        # pull a fresh snapshot (add --limit 80 to keep fewer)
python3 build_workbook.py    # build output/prediction_market_ops_tracker.xlsx
```

If you just want the workbook and don't care about fresh data, skip the first step.
`build_workbook.py` will use the committed sample snapshot and produce the exact
same file every time.

Open the result in Excel 365. It uses XLOOKUP, which older Excel versions don't
have (they'll show `#NAME?` in those cells).

The fetch script retries on failed requests, backs off on rate limits, and won't
crash the whole run if one part of the API is down. It just logs it and moves on.

## The workbook

Four tabs:

- Markets — one row per market with prices, volume, liquidity, and a
  days-to-resolution column that counts from the snapshot date.
- Trades — the real trades, plus formula columns: notional (price × size), market
  category and total volume pulled over with XLOOKUP, and a WHALE flag for trades
  above the 95th percentile.
- Analysis — the flag logic, all live formulas (SUMIFS, COUNTIFS, AVERAGEIFS,
  PERCENTILE). Thresholds sit in editable cells at the top. There's a thin-market
  flag, a pricing-discrepancy flag, a near-certain-but-active flag, and activity
  broken down by day, category, and side. There's also a spot marked off for
  pivot tables to build by hand.
- Dashboard — the KPIs (total volume, thin-market count, discrepancy count, whale
  count, biggest single-market 24h volume), red/yellow conditional formatting on
  the flags, and two charts: top 10 markets by 24h volume, and daily trade volume.

## The Odds API bit (optional)

If you set a `THE_ODDS_API_KEY` environment variable, the fetch script also grabs
sportsbook odds, converts them to implied probabilities (with the vig removed),
and the workbook adds a Discrepancy tab comparing sportsbook probability against
the Polymarket price, flagging any gap over 5%. That's basically cross-market
mispricing detection.

```bash
export THE_ODDS_API_KEY=your_key
python3 fetch_data.py && python3 build_workbook.py
```

Honest note: I didn't have a key when I put this snapshot together, so this path
is written and tested against a mock file but hasn't run against the live Odds
API. It kicks in automatically once a key is set. The team-name matching between
the two sources may need a little tuning depending on the sport and season.

## What's in here, skill-wise

- Pulling and paging real REST APIs, with retries and graceful failure
- XLOOKUP for pulling data across sheets
- SUMIFS / COUNTIFS / AVERAGEIFS for the breakdowns
- Percentile-based thresholds so the flags are data-driven, not hardcoded
- Conditional formatting
- Native Excel charts
- Pivot tables (set up by hand, see below)

## Still to do by hand

The workbook is generated ready to build on. In Excel I'd finish it off by:

- Building the 3 pivot tables in the marked area on the Analysis tab (volume by
  market, volume by day and side, whale count by market)
- Adding slicers on market and side
- Dropping a couple of Dashboard screenshots into this README
- Writing up 3 findings from the actual data — for example which markets tripped
  the thin-market flag, the buy vs sell imbalance in the top markets, and any
  near-certain markets still soaking up volume

---

Data is Polymarket's, pulled from their public endpoints. This is a portfolio
project, not trading advice.
