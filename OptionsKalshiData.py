"""
Build SQLite databases of:
  - hourly CL=F (WTI crude futures) price data, ~1 year back (Yahoo Finance), and
  - Kalshi's "WTI price at year end" market ladder (KXWTIDIRY), hourly candlesticks.

Requirements:
    pip install yfinance pandas requests openpyxl

Usage:
    python OptionsKalshiData.py                          # CL=F futures (default)
    python OptionsKalshiData.py --db my_oil.db --period 1y
    python OptionsKalshiData.py --source kalshi           # Kalshi WTI year-end ladder
    python OptionsKalshiData.py --source kalshi --print   # inspect existing kalshi db

Notes:
- Yahoo's 1-hour interval is only available for the trailing ~730 days, so 1y
  fits comfortably. If you ask for more than ~2y of hourly data, Yahoo will
  silently truncate the response.
- CL=F is the continuous/generic front-month contract — it splices together
  whichever month is currently front-month, so there are small artificial
  jumps at each monthly rollover. Fine for backtesting volatility/strategy
  behavior; not the same as pulling one specific contract month (e.g. CLZ26).
- Re-running this script is safe: it upserts on datetime, so you can schedule
  it to run daily/weekly to keep the database current without duplicating rows.
- The Kalshi source pulls every strike in the KXWTIDIRY-26DEC31H1430 event (the
  "will WTI be above $X on Dec 31, 2026" ladder), one row per (strike, hour).
  Kalshi's public market-data endpoints (api.elections.kalshi.com) don't require
  authentication.
"""

import argparse
import sqlite3
import sys
import time

import pandas as pd
import requests
import yfinance as yf

SCHEMA = """
CREATE TABLE IF NOT EXISTS cl_hourly (
    datetime_utc TEXT PRIMARY KEY,
    open REAL,
    high REAL,
    low REAL,
    close REAL,
    volume INTEGER
);
"""

KALSHI_API_BASE = "https://api.elections.kalshi.com/trade-api/v2"
KALSHI_SERIES_TICKER = "KXWTIDIRY"
KALSHI_EVENT_TICKER = "KXWTIDIRY-26DEC31H1430"

SCHEMA_KALSHI = """
CREATE TABLE IF NOT EXISTS kalshi_wti_yearend (
    market_ticker TEXT,
    strike REAL,
    datetime_utc TEXT,
    price_close REAL,
    yes_bid_close REAL,
    yes_ask_close REAL,
    volume INTEGER,
    open_interest REAL,
    PRIMARY KEY (market_ticker, datetime_utc)
);
"""


def fetch_hourly(ticker="CL=F", period="1y"):
    print(f"Fetching {ticker} @ 1h interval, period={period} ...")
    df = yf.download(ticker, period=period, interval="1h", auto_adjust=False, progress=False)
    if df.empty:
        print("ERROR: no data returned. Check ticker/network/rate limits.", file=sys.stderr)
        sys.exit(1)

    # yfinance sometimes returns a MultiIndex column header (ticker, field) — flatten it
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [c[0] for c in df.columns]

    df = df.reset_index()
    df = df.rename(columns={
        "Datetime": "datetime_utc",
        "Open": "open", "High": "high", "Low": "low",
        "Close": "close", "Volume": "volume",
    })
    df["datetime_utc"] = pd.to_datetime(df["datetime_utc"], utc=True).dt.strftime("%Y-%m-%d %H:%M:%S")
    return df[["datetime_utc", "open", "high", "low", "close", "volume"]]


def fetch_kalshi_event_markets(event_ticker=KALSHI_EVENT_TICKER):
    resp = requests.get(f"{KALSHI_API_BASE}/markets", params={"event_ticker": event_ticker}, timeout=30)
    resp.raise_for_status()
    markets = resp.json()["markets"]
    if not markets:
        print(f"ERROR: no markets found for event {event_ticker}.", file=sys.stderr)
        sys.exit(1)
    return markets


def fetch_kalshi_candles(market_ticker, series_ticker=KALSHI_SERIES_TICKER,
                          start_ts=None, end_ts=None, period_interval=60):
    end_ts = end_ts or int(time.time())
    max_span = period_interval * 60 * 4500  # stay under the API's per-request candle cap
    rows = []
    window_start = start_ts
    while window_start < end_ts:
        window_end = min(window_start + max_span, end_ts)
        resp = requests.get(
            f"{KALSHI_API_BASE}/series/{series_ticker}/markets/{market_ticker}/candlesticks",
            params={"start_ts": window_start, "end_ts": window_end, "period_interval": period_interval},
            timeout=30,
        )
        resp.raise_for_status()
        for c in resp.json()["candlesticks"]:
            price = c.get("price", {})
            close = price.get("close_dollars") or price.get("previous_dollars")
            rows.append((
                market_ticker,
                pd.to_datetime(c["end_period_ts"], unit="s", utc=True).strftime("%Y-%m-%d %H:%M:%S"),
                float(close) if close is not None else None,
                float(c["yes_bid"]["close_dollars"]) if c.get("yes_bid", {}).get("close_dollars") else None,
                float(c["yes_ask"]["close_dollars"]) if c.get("yes_ask", {}).get("close_dollars") else None,
                int(round(float(c.get("volume_fp", 0) or 0))),
                float(c.get("open_interest_fp", 0) or 0),
            ))
        window_start = window_end
    return pd.DataFrame(rows, columns=[
        "market_ticker", "datetime_utc", "price_close",
        "yes_bid_close", "yes_ask_close", "volume", "open_interest",
    ])


def fetch_kalshi_wti_yearend(event_ticker=KALSHI_EVENT_TICKER, period_interval=60):
    markets = fetch_kalshi_event_markets(event_ticker)
    frames = []
    for m in markets:
        print(f"Fetching {m['ticker']} (strike {m['floor_strike']}) candlesticks ...")
        start_ts = int(pd.Timestamp(m["open_time"]).timestamp())
        end_ts = min(int(time.time()), int(pd.Timestamp(m["close_time"]).timestamp()))
        df = fetch_kalshi_candles(m["ticker"], KALSHI_SERIES_TICKER, start_ts, end_ts, period_interval)
        df["strike"] = m["floor_strike"]
        frames.append(df)
    result = pd.concat(frames, ignore_index=True)
    return result[["market_ticker", "strike", "datetime_utc", "price_close",
                    "yes_bid_close", "yes_ask_close", "volume", "open_interest"]]


def build_kalshi_db(df, db_path):
    conn = sqlite3.connect(db_path)
    conn.execute(SCHEMA_KALSHI)
    rows = list(df.itertuples(index=False, name=None))
    conn.executemany(
        """INSERT INTO kalshi_wti_yearend
               (market_ticker, strike, datetime_utc, price_close, yes_bid_close, yes_ask_close, volume, open_interest)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)
           ON CONFLICT(market_ticker, datetime_utc) DO UPDATE SET
             price_close=excluded.price_close, yes_bid_close=excluded.yes_bid_close,
             yes_ask_close=excluded.yes_ask_close, volume=excluded.volume,
             open_interest=excluded.open_interest""",
        rows,
    )
    conn.commit()

    n = conn.execute("SELECT COUNT(*) FROM kalshi_wti_yearend").fetchone()[0]
    first = conn.execute("SELECT MIN(datetime_utc) FROM kalshi_wti_yearend").fetchone()[0]
    last = conn.execute("SELECT MAX(datetime_utc) FROM kalshi_wti_yearend").fetchone()[0]
    conn.close()
    print(f"Database ready: {db_path}")
    print(f"  {n} rows, {first} -> {last}")


def print_kalshi_db(db_path, limit=10):
    conn = sqlite3.connect(db_path)
    n = conn.execute("SELECT COUNT(*) FROM kalshi_wti_yearend").fetchone()[0]
    first = conn.execute("SELECT MIN(datetime_utc) FROM kalshi_wti_yearend").fetchone()[0]
    last = conn.execute("SELECT MAX(datetime_utc) FROM kalshi_wti_yearend").fetchone()[0]
    print(f"Database: {db_path}")
    print(f"  {n} rows, {first} -> {last}")

    cols = [c[1] for c in conn.execute("PRAGMA table_info(kalshi_wti_yearend)")]
    print("\n" + " | ".join(cols))

    print(f"-- first {limit} --")
    for row in conn.execute(f"SELECT * FROM kalshi_wti_yearend ORDER BY datetime_utc ASC LIMIT {limit}"):
        print(row)

    print(f"-- last {limit} --")
    for row in conn.execute(f"SELECT * FROM kalshi_wti_yearend ORDER BY datetime_utc DESC LIMIT {limit}"):
        print(row)

    conn.close()


def print_db(db_path, limit=10):
    conn = sqlite3.connect(db_path)
    n = conn.execute("SELECT COUNT(*) FROM cl_hourly").fetchone()[0]
    first = conn.execute("SELECT MIN(datetime_utc) FROM cl_hourly").fetchone()[0]
    last = conn.execute("SELECT MAX(datetime_utc) FROM cl_hourly").fetchone()[0]
    print(f"Database: {db_path}")
    print(f"  {n} rows, {first} -> {last}")

    cols = [c[1] for c in conn.execute("PRAGMA table_info(cl_hourly)")]
    print("\n" + " | ".join(cols))

    print(f"-- first {limit} --")
    for row in conn.execute(f"SELECT * FROM cl_hourly ORDER BY datetime_utc ASC LIMIT {limit}"):
        print(row)

    print(f"-- last {limit} --")
    for row in conn.execute(f"SELECT * FROM cl_hourly ORDER BY datetime_utc DESC LIMIT {limit}"):
        print(row)

    conn.close()


def build_db(df, db_path):
    conn = sqlite3.connect(db_path)
    conn.execute(SCHEMA)
    rows = list(df.itertuples(index=False, name=None))
    conn.executemany(
        """INSERT INTO cl_hourly (datetime_utc, open, high, low, close, volume)
           VALUES (?, ?, ?, ?, ?, ?)
           ON CONFLICT(datetime_utc) DO UPDATE SET
             open=excluded.open, high=excluded.high, low=excluded.low,
             close=excluded.close, volume=excluded.volume""",
        rows,
    )
    conn.commit()

    n = conn.execute("SELECT COUNT(*) FROM cl_hourly").fetchone()[0]
    first = conn.execute("SELECT MIN(datetime_utc) FROM cl_hourly").fetchone()[0]
    last = conn.execute("SELECT MAX(datetime_utc) FROM cl_hourly").fetchone()[0]
    conn.close()
    print(f"Database ready: {db_path}")
    print(f"  {n} rows, {first} -> {last}")


def main():
    ap = argparse.ArgumentParser(description="Build SQLite databases of CL=F or Kalshi WTI year-end data")
    ap.add_argument("--source", choices=["cl", "kalshi"], default="cl",
                     help="cl = Yahoo CL=F hourly futures (default), kalshi = Kalshi WTI year-end ladder")
    ap.add_argument("--ticker", default="CL=F", help="[cl] Yahoo ticker (default CL=F, WTI front-month)")
    ap.add_argument("--period", default="1y", help="[cl] How far back (max ~2y for 1h interval)")
    ap.add_argument("--event", default=KALSHI_EVENT_TICKER,
                     help="[kalshi] Event ticker for the year-end ladder")
    ap.add_argument("--interval", type=int, default=60,
                     help="[kalshi] Candlestick period in minutes (1, 60, or 1440)")
    ap.add_argument("--db", default=None, help="Output SQLite file path "
                     "(default: cl_futures.db or kalshi_wti_yearend.db, depending on --source)")
    ap.add_argument("--csv", default=None, help="Optional: also write a CSV copy")
    ap.add_argument("--xlsx", default=None, help="Optional: also write an Excel (.xlsx) copy")
    ap.add_argument("--print", dest="print_only", action="store_true",
                     help="Just print the existing database's contents and exit (no fetch)")
    ap.add_argument("--limit", type=int, default=10,
                     help="Rows to show from each end when using --print (default 10)")
    args = ap.parse_args()

    db_path = args.db or ("kalshi_wti_yearend.db" if args.source == "kalshi" else "cl_futures.db")

    if args.print_only:
        (print_kalshi_db if args.source == "kalshi" else print_db)(db_path, args.limit)
        return

    if args.source == "kalshi":
        df = fetch_kalshi_wti_yearend(args.event, args.interval)
        build_kalshi_db(df, db_path)
    else:
        df = fetch_hourly(args.ticker, args.period)
        build_db(df, db_path)

    if args.csv:
        df.to_csv(args.csv, index=False)
        print(f"Also wrote CSV: {args.csv}")

    if args.xlsx:
        df.to_excel(args.xlsx, index=False)
        print(f"Also wrote Excel: {args.xlsx}")


if __name__ == "__main__":
    main()

