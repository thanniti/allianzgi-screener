# screener.py
import yfinance as yf
import numpy as np
import pandas as pd
import datetime

# ── FUND UNIVERSE ─────────────────────────────────────────────────────────────
# Map your fund names to tickers. Replace with real AllianzGI tickers later.
FUND_UNIVERSE = [
    {"name": "AllianzGI Global MA Blend",    "ticker": "VWRL.L",  "type": "Growth",          "region": "Global",       "esg": 68, "allocation": {"Equity":70,"Bond":20,"Alternatives":10}},
    {"name": "AllianzGI Income & Growth",    "ticker": "IGLN.L",  "type": "Balanced",         "region": "Asia-Pacific", "esg": 78, "allocation": {"Equity":55,"Bond":35,"Cash":10}},
    {"name": "AllianzGI Multi Asset Stab.",  "ticker": "AGBP.L",  "type": "Defensive",        "region": "Global",       "esg": 85, "allocation": {"Equity":30,"Bond":55,"Alternatives":15}},
    {"name": "AllianzGI Absolute Return",    "ticker": "JPST",    "type": "Absolute Return",  "region": "Global",       "esg": 71, "allocation": {"Equity":20,"Bond":40,"Alternatives":40}},
]

# ── DATA PULL ────────────────────────────────────────────────────────────────
def fetch_metrics(ticker: str) -> dict:
    """Pull 1Y price history and compute key metrics."""
    try:
        hist = yf.Ticker(ticker).history(period="1y")["Close"].dropna()
        daily_returns = hist.pct_change().dropna()

        return_1y     = round((hist.iloc[-1] / hist.iloc[0] - 1) * 100, 2)
        volatility    = daily_returns.std() * np.sqrt(252)
        sharpe        = round((daily_returns.mean() * 252) / (volatility), 2) if volatility else 0
        sortino_denom = daily_returns[daily_returns < 0].std() * np.sqrt(252)
        sortino       = round((daily_returns.mean() * 252) / sortino_denom, 2) if sortino_denom else 0
        rolling_max   = hist.cummax()
        drawdown      = ((hist - rolling_max) / rolling_max)
        max_drawdown  = round(drawdown.min() * 100, 2)

        return {
            "return_1y":    return_1y,
            "sharpe":       sharpe,
            "sortino_est":  sortino,
            "max_drawdown": max_drawdown,
            "volatility":   round(volatility * 100, 2),
        }
    except Exception as e:
        print(f"  Could not fetch {ticker}: {e}", flush=True)
        return None

# ── COMPOSITE SCORER ─────────────────────────────────────────────────────────
def composite_score(metrics: dict, esg: int) -> int:
    s  = min(metrics["sharpe"] / 1.5, 1) * 40
    s += (1 - min(abs(metrics["max_drawdown"]) / 25, 1)) * 30
    s += (esg / 100) * 30
    return round(s)

# ── SCREENING FILTERS ────────────────────────────────────────────────────────
FILTERS = {
    "min_sharpe":    0.4,
    "min_esg":       60,
    "max_drawdown": -25.0,   # funds worse than -25% are excluded
}

def run_screener(filters: dict = FILTERS) -> list:
    print("Running fund screener...\n", flush=True)
    passed = []

    for fund in FUND_UNIVERSE:
        print(f"  Fetching {fund['name']}...", flush=True)
        metrics = fetch_metrics(fund["ticker"])
        if not metrics:
            continue

        # Apply filters
        if metrics["sharpe"]      < filters["min_sharpe"]:    continue
        if fund["esg"]            < filters["min_esg"]:       continue
        if metrics["max_drawdown"]< filters["max_drawdown"]:  continue

        score = composite_score(metrics, fund["esg"])

        alloc = fund.get("allocation", {}).copy()
        if "Alternatives" not in alloc and "Cash" in alloc:
            alloc["Alternatives"] = alloc["Cash"]
        # Ensure keys exist for all expected types
        alloc.setdefault("Equity", 0)
        alloc.setdefault("Bond", 0)
        alloc.setdefault("Alternatives", 0)

        passed.append({
            **fund,
            **metrics,
            "allocation": alloc,
            "score": score,
            "date": datetime.date.today().strftime("%B %Y"),
        })
        print(f"  Passed — score {score}/100", flush=True)

    passed.sort(key=lambda x: x["score"], reverse=True)
    print(f"\n{len(passed)} of {len(FUND_UNIVERSE)} funds passed filters.\n", flush=True)
    return passed