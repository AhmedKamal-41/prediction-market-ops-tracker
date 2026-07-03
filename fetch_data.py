#!/usr/bin/env python3
"""
fetch_data.py — Prediction Market Operations Tracker (data pull stage)

Pulls REAL live data from Polymarket's public APIs (no auth required):

  * Gamma API   https://gamma-api.polymarket.com/markets
      ~50-100 active markets, prioritizing sports (NBA/NFL/MLB/NHL/soccer/...)
      plus a slice of non-sports markets for comparison.
  * Data API    https://data-api.polymarket.com/trades
      Recent real user trades for the top-10 sports markets by total volume.

Outputs (data/raw/):
  * markets_<UTCSTAMP>.csv   — timestamped archive (git-ignored)
  * trades_<UTCSTAMP>.csv    — timestamped archive (git-ignored)
  * markets_latest.csv       — stable pointer to newest pull (committed sample)
  * trades_latest.csv        — stable pointer to newest pull (committed sample)
  * snapshot_manifest.json   — snapshot metadata read by build_workbook.py

The pull is live, but every run is archived and the manifest records the exact
snapshot timestamp so the downstream workbook build is fully reproducible.

Usage:
    python3 fetch_data.py            # normal live pull
    python3 fetch_data.py --limit 80 # cap number of markets kept
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import time
from datetime import datetime, timezone

import requests

# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #
GAMMA_MARKETS_URL = "https://gamma-api.polymarket.com/markets"
DATA_TRADES_URL = "https://data-api.polymarket.com/trades"
ODDS_API_BASE = "https://api.the-odds-api.com/v4/sports"

# Sport keys tried for The Odds API outright markets when THE_ODDS_API_KEY is set.
# Outright/"winner" markets return one outcome per team, which maps cleanly onto
# Polymarket "Will <team> win ...?" markets. Extend this list for other sports.
ODDS_SPORT_KEYS = [
    "soccer_fifa_world_cup_winner",
    "americanfootball_nfl_super_bowl_winner",
    "basketball_nba_championship_winner",
]

RAW_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "raw")

POOL_PAGES = 6            # pages of 100 markets to scan when building the pool
POOL_PAGE_SIZE = 100
MAX_SPORTS = 70          # cap of sports markets kept
MAX_OTHERS = 30          # cap of non-sports comparison markets kept
HARD_CAP = 100           # never keep more than this many markets total
TOP_N_FOR_TRADES = 10    # pull trades for this many top sports markets
TRADES_PER_MARKET = 300  # recent trades to pull per market
REQUEST_TIMEOUT = 25
RATE_LIMIT_SLEEP = 0.35  # polite pause between trade requests

# Keyword -> (category label, is_sports). Checked in order; first hit wins,
# so more specific leagues are listed before generic buckets.
SPORTS_RULES = [
    (("wnba",), "WNBA"),
    (("nba", "nba-", "basketball"), "NBA"),
    (("nfl", "super-bowl", "superbowl"), "NFL"),
    (("mlb", "world-series", "baseball"), "MLB"),
    (("nhl", "stanley-cup", "hockey"), "NHL"),
    (("fifwc", "fifa", "world-cup", "worldcup", "soccer", "premier-league",
      "premier league", "epl", "uefa", "la-liga", "laliga", "la liga",
      "champions-league", "champions league", "bundesliga", "serie-a",
      "ligue-1", "mls", "ronaldo", "messi", "ballon"), "Soccer"),
    (("2026-f1", "formula-1", "formula 1", "-f1-", "grand-prix", "grand prix"), "F1"),
    (("atp", "wta", "wimbledon", "roland-garros", "us-open-tennis", "tennis"), "Tennis"),
    (("pga", "-golf", "golf-", "masters-golf", "the-open-golf"), "Golf"),
    (("ufc", "-mma", "mma-"), "MMA/UFC"),
    (("boxing", "-boxing"), "Boxing"),
    (("lol-", "league-of-legends", "cs2", "csgo", "cs-go", "dota", "valorant",
      "esports", "e-sports"), "Esports"),
    (("ncaa", "college-football", "college-basketball", "march-madness"), "College"),
    (("cricket", "ipl-"), "Cricket"),
]

OTHER_RULES = [
    (("bitcoin", "ethereum", "crypto", "-btc-", "-eth-", "solana", "dogecoin",
      "xrp", "memecoin"), "Crypto"),
    (("election", "president", "prime-minister", "prime minister", "putin",
      "senate", "congress", "governor", "parliament", "referendum",
      "primary", "cabinet", "impeach"), "Politics"),
    (("fed", "interest-rate", "interest rate", "rate-cut", "rate-hike", "gdp",
      "inflation", "recession", "jobs-report", "cpi"), "Economics"),
    (("largest-company", "nvidia", "openai", "gpt", "-tesla", "apple",
      "-ai-", "chatgpt", "twitter", "tweets", "spacex"), "Tech"),
    (("album", "movie", "box-office", "oscar", "grammy", "emmy", "netflix",
      "spotify", "song", "tour"), "Pop Culture"),
    (("weather", "hurricane", "temperature", "storm"), "Weather"),
]


# --------------------------------------------------------------------------- #
# HTTP helpers
# --------------------------------------------------------------------------- #
def make_session() -> requests.Session:
    """Session with a couple of automatic retries on transient failures."""
    from requests.adapters import HTTPAdapter

    try:
        from urllib3.util.retry import Retry
        retry = Retry(
            total=3,
            backoff_factor=0.8,
            status_forcelist=(429, 500, 502, 503, 504),
            allowed_methods=frozenset(["GET"]),
        )
        adapter = HTTPAdapter(max_retries=retry)
    except Exception:  # pragma: no cover - very old urllib3
        adapter = HTTPAdapter()

    s = requests.Session()
    s.mount("https://", adapter)
    s.headers.update({"User-Agent": "pm-ops-tracker/1.0 (ops-analyst-portfolio)"})
    return s


def get_json(session: requests.Session, url: str, params: dict):
    """GET returning parsed JSON, or None on any failure (logged, non-fatal)."""
    try:
        r = session.get(url, params=params, timeout=REQUEST_TIMEOUT)
        if r.status_code != 200:
            print(f"  ! {url} returned HTTP {r.status_code}", file=sys.stderr)
            return None
        return r.json()
    except requests.RequestException as e:
        print(f"  ! request error for {url}: {e}", file=sys.stderr)
        return None
    except ValueError as e:
        print(f"  ! bad JSON from {url}: {e}", file=sys.stderr)
        return None


# --------------------------------------------------------------------------- #
# Parsing / categorization
# --------------------------------------------------------------------------- #
def _json_list(raw):
    """Polymarket returns some list fields as JSON-encoded strings."""
    if isinstance(raw, list):
        return raw
    if isinstance(raw, str) and raw.strip():
        try:
            v = json.loads(raw)
            return v if isinstance(v, list) else []
        except ValueError:
            return []
    return []


def _to_float(v, default=0.0):
    try:
        if v is None or v == "":
            return default
        return float(v)
    except (TypeError, ValueError):
        return default


def categorize(text: str):
    """Return (category_label, is_sports) from a lowercased combined string."""
    for keywords, label in SPORTS_RULES:
        if any(k in text for k in keywords):
            return label, True
    for keywords, label in OTHER_RULES:
        if any(k in text for k in keywords):
            return label, False
    return "Other", False


def parse_market(m: dict):
    """Flatten one Gamma market record to the row we persist, or None to skip."""
    question = (m.get("question") or "").strip()
    condition_id = m.get("conditionId") or ""
    if not question or not condition_id:
        return None

    outcomes = _json_list(m.get("outcomes"))
    prices = _json_list(m.get("outcomePrices"))

    yes_price = no_price = None
    lower = [str(o).strip().lower() for o in outcomes]
    if "yes" in lower and len(prices) >= 1:
        yes_price = _to_float(prices[lower.index("yes")], None)
    if "no" in lower and len(prices) >= 2:
        no_price = _to_float(prices[lower.index("no")], None)
    # Fallback for markets that aren't literally labelled Yes/No.
    if yes_price is None and len(prices) >= 1:
        yes_price = _to_float(prices[0], None)
    if no_price is None and len(prices) >= 2:
        no_price = _to_float(prices[1], None)

    event = (m.get("events") or [{}])[0] or {}
    combined = " ".join([
        question,
        str(event.get("title") or ""),
        str(event.get("slug") or ""),
        str(event.get("ticker") or ""),
    ]).lower()
    category, is_sports = categorize(combined)

    total_volume = _to_float(m.get("volumeNum"))
    if total_volume == 0.0:
        total_volume = _to_float(m.get("volume"))
    liquidity = _to_float(m.get("liquidityNum"))
    if liquidity == 0.0:
        liquidity = _to_float(m.get("liquidity"))

    return {
        "condition_id": condition_id,
        "question": question,
        "category": category,
        "is_sports": is_sports,
        "yes_price": yes_price if yes_price is not None else "",
        "no_price": no_price if no_price is not None else "",
        "volume_24h": round(_to_float(m.get("volume24hr")), 2),
        "total_volume": round(total_volume, 2),
        "liquidity": round(liquidity, 2),
        "end_date": (m.get("endDate") or m.get("endDateIso") or "")[:10],
        "event_title": str(event.get("title") or "").strip(),
        "event_slug": str(event.get("slug") or "").strip(),
    }


# --------------------------------------------------------------------------- #
# Fetch stages
# --------------------------------------------------------------------------- #
def fetch_pool(session: requests.Session) -> list:
    """Scan several pages of markets ordered by 24h volume; return parsed rows."""
    print("Fetching market pool from Gamma API (ordered by 24h volume)...")
    seen = set()
    rows = []
    for page in range(POOL_PAGES):
        params = {
            "limit": POOL_PAGE_SIZE,
            "offset": page * POOL_PAGE_SIZE,
            "active": "true",
            "closed": "false",
            "order": "volume24hr",
            "ascending": "false",
        }
        data = get_json(session, GAMMA_MARKETS_URL, params)
        if not data:
            print(f"  page {page + 1}: no data (stopping pagination)")
            break
        added = 0
        for m in data:
            parsed = parse_market(m)
            if not parsed:
                continue
            if parsed["condition_id"] in seen:
                continue
            seen.add(parsed["condition_id"])
            rows.append(parsed)
            added += 1
        print(f"  page {page + 1}: kept {added} markets (pool={len(rows)})")
        if len(data) < POOL_PAGE_SIZE:
            break
        time.sleep(0.2)
    return rows


def select_markets(pool: list, limit: int) -> list:
    """Prioritize sports; add a slice of non-sports for comparison."""
    vol24 = lambda m: m["volume_24h"]
    sports = sorted((m for m in pool if m["is_sports"]), key=vol24, reverse=True)
    others = sorted((m for m in pool if not m["is_sports"]), key=vol24, reverse=True)

    selected = sports[:MAX_SPORTS] + others[:MAX_OTHERS]
    # Backfill if we came up short so we always keep a healthy sample.
    if len(selected) < 50:
        extra = sports[MAX_SPORTS:] + others[MAX_OTHERS:]
        selected += extra[: 50 - len(selected)]

    selected = selected[: min(limit, HARD_CAP)]
    n_sports = sum(1 for m in selected if m["is_sports"])
    print(f"Selected {len(selected)} markets "
          f"({n_sports} sports, {len(selected) - n_sports} non-sports).")
    return selected


def fetch_trades(session: requests.Session, markets: list) -> list:
    """Pull recent trades for the top-N sports markets by total volume."""
    sports = [m for m in markets if m["is_sports"] and m["total_volume"] > 0]
    sports.sort(key=lambda m: m["total_volume"], reverse=True)
    targets = sports[:TOP_N_FOR_TRADES]
    print(f"Fetching trades for top {len(targets)} sports markets by volume...")

    all_trades = []
    for i, m in enumerate(targets, 1):
        cid = m["condition_id"]
        collected = []
        offset = 0
        while len(collected) < TRADES_PER_MARKET:
            page = get_json(session, DATA_TRADES_URL, {
                "market": cid,
                "limit": min(500, TRADES_PER_MARKET - len(collected)),
                "offset": offset,
            })
            if not page:
                break
            collected.extend(page)
            if len(page) < 100:  # last page
                break
            offset += len(page)
            time.sleep(RATE_LIMIT_SLEEP)

        for t in collected:
            ts = int(_to_float(t.get("timestamp")))
            iso = datetime.fromtimestamp(ts, tz=timezone.utc).isoformat() if ts else ""
            all_trades.append({
                "condition_id": cid,
                "market": (t.get("title") or m["question"]).strip(),
                "timestamp_unix": ts,
                "timestamp_iso": iso,
                "date": iso[:10],
                "side": t.get("side") or "",
                "outcome": t.get("outcome") or "",
                "price": round(_to_float(t.get("price")), 4),
                "size_shares": round(_to_float(t.get("size")), 4),
                "tx_hash": t.get("transactionHash") or "",
            })
        print(f"  [{i}/{len(targets)}] {m['question'][:50]:50s} -> {len(collected)} trades")
        time.sleep(RATE_LIMIT_SLEEP)

    print(f"Collected {len(all_trades)} trades total.")
    return all_trades


# --------------------------------------------------------------------------- #
# Optional stretch: sportsbook odds (The Odds API) for cross-market mispricing
# --------------------------------------------------------------------------- #
def fetch_odds(session: requests.Session) -> list:
    """
    OPTIONAL. When THE_ODDS_API_KEY is set, pull sportsbook outright odds and
    convert to de-vigged implied probabilities so build_workbook.py can compare
    them against Polymarket prices (literal cross-market mispricing detection).

    Returns a list of rows (possibly empty). Never raises — the odds pull is a
    best-effort add-on and must never break the core Polymarket snapshot.
    """
    api_key = os.environ.get("THE_ODDS_API_KEY")
    if not api_key:
        print("Skipping sportsbook odds (THE_ODDS_API_KEY not set).")
        return []

    print("THE_ODDS_API_KEY detected — pulling sportsbook outright odds...")
    rows = []
    for sport in ODDS_SPORT_KEYS:
        data = get_json(session, f"{ODDS_API_BASE}/{sport}/odds", {
            "apiKey": api_key,
            "regions": "us",
            "markets": "outrights",
            "oddsFormat": "decimal",
        })
        if not data:
            continue
        for event in data:
            # Consensus implied prob = median across books of the de-vigged prob.
            per_team_probs = {}
            for book in event.get("bookmakers", []) or []:
                for market in book.get("markets", []) or []:
                    outs = market.get("outcomes", []) or []
                    inv = [(_to_float(o.get("price")), o.get("name")) for o in outs]
                    inv = [(p, name) for p, name in inv if p and p > 0]
                    overround = sum(1.0 / p for p, _ in inv) or 1.0
                    for p, name in inv:
                        prob = (1.0 / p) / overround  # de-vigged
                        per_team_probs.setdefault(name, []).append(prob)
            for team, probs in per_team_probs.items():
                probs.sort()
                mid = probs[len(probs) // 2]
                rows.append({
                    "sport": sport,
                    "team": (team or "").strip(),
                    "sportsbook_implied_prob": round(mid, 4),
                    "n_books": len(probs),
                })
        time.sleep(RATE_LIMIT_SLEEP)
        if rows:  # got a usable sport; stop to conserve API quota
            print(f"  matched sport: {sport} ({len(rows)} team quotes)")
            break

    if not rows:
        print("  no outright odds returned (sports may be out of season).")
    return rows


# --------------------------------------------------------------------------- #
# Persistence
# --------------------------------------------------------------------------- #
def write_csv(path: str, rows: list, columns: list):
    import csv
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=columns, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)


def main():
    ap = argparse.ArgumentParser(description="Pull live Polymarket data to CSVs.")
    ap.add_argument("--limit", type=int, default=HARD_CAP,
                    help="max markets to keep (default 100)")
    args = ap.parse_args()

    os.makedirs(RAW_DIR, exist_ok=True)
    session = make_session()

    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    snapshot_iso = datetime.now(timezone.utc).replace(microsecond=0).isoformat()

    pool = fetch_pool(session)
    if not pool:
        print("FATAL: could not fetch any markets. Is the API reachable?",
              file=sys.stderr)
        sys.exit(1)

    markets = select_markets(pool, args.limit)
    trades = fetch_trades(session, markets)
    odds = fetch_odds(session)  # optional; empty unless THE_ODDS_API_KEY is set

    market_cols = ["condition_id", "question", "category", "is_sports",
                   "yes_price", "no_price", "volume_24h", "total_volume",
                   "liquidity", "end_date", "event_title", "event_slug"]
    trade_cols = ["condition_id", "market", "timestamp_unix", "timestamp_iso",
                  "date", "side", "outcome", "price", "size_shares", "tx_hash"]

    markets_archive = os.path.join(RAW_DIR, f"markets_{stamp}.csv")
    trades_archive = os.path.join(RAW_DIR, f"trades_{stamp}.csv")
    markets_latest = os.path.join(RAW_DIR, "markets_latest.csv")
    trades_latest = os.path.join(RAW_DIR, "trades_latest.csv")

    for path in (markets_archive, markets_latest):
        write_csv(path, markets, market_cols)
    for path in (trades_archive, trades_latest):
        write_csv(path, trades, trade_cols)

    odds_file = None
    if odds:
        odds_cols = ["sport", "team", "sportsbook_implied_prob", "n_books"]
        odds_latest = os.path.join(RAW_DIR, "odds_latest.csv")
        write_csv(odds_latest, odds, odds_cols)
        write_csv(os.path.join(RAW_DIR, f"odds_{stamp}.csv"), odds, odds_cols)
        odds_file = "odds_latest.csv"

    manifest = {
        "snapshot_utc": snapshot_iso,
        "snapshot_stamp": stamp,
        "markets_file": "markets_latest.csv",
        "trades_file": "trades_latest.csv",
        "markets_archive": os.path.basename(markets_archive),
        "trades_archive": os.path.basename(trades_archive),
        "market_count": len(markets),
        "sports_market_count": sum(1 for m in markets if m["is_sports"]),
        "trade_count": len(trades),
        "odds_file": odds_file,
        "odds_quote_count": len(odds),
        "sources": {
            "markets": GAMMA_MARKETS_URL,
            "trades": DATA_TRADES_URL,
        },
    }
    with open(os.path.join(RAW_DIR, "snapshot_manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2)

    print("\nSnapshot written:")
    print(f"  snapshot_utc : {snapshot_iso}")
    print(f"  markets      : {len(markets)}  -> {os.path.relpath(markets_latest)}")
    print(f"  trades       : {len(trades)}  -> {os.path.relpath(trades_latest)}")
    print(f"  manifest     : data/raw/snapshot_manifest.json")
    print("\nNext: python3 build_workbook.py")


if __name__ == "__main__":
    main()
