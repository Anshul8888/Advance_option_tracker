"""
Data processing layer
  • Aggregate 1-min → 2-min / 5-min candles
  • OBV, VWAP calculation
  • Relative Strength vs ATM
  • Ascending trend detection
  • Sort-score computation
"""

import logging

import pandas as pd
import numpy as np

import config

log = logging.getLogger("tracker.processor")


# ═══════════════════════════════════════════════════════
#  Candle Aggregation
# ═══════════════════════════════════════════════════════
def _prepare_dataframe(minute_candles: list[dict]) -> pd.DataFrame | None:
    """
    Convert raw Kite candle dicts to a clean DataFrame.
    Handles timezone-aware datetimes from Zerodha.
    """
    if not minute_candles or len(minute_candles) < 2:
        return None

    df = pd.DataFrame(minute_candles)
    df["date"] = pd.to_datetime(df["date"])

    # Strip timezone (Zerodha returns IST-aware datetimes)
    if df["date"].dt.tz is not None:
        df["date"] = df["date"].dt.tz_localize(None)

    df.set_index("date", inplace=True)
    df.sort_index(inplace=True)
    return df


def _resample(df: pd.DataFrame, interval: str) -> pd.DataFrame:
    """Resample OHLCV DataFrame to the given interval."""
    market_open = df.index[0].replace(hour=9, minute=15, second=0)

    resampled = (
        df.resample(interval, origin=market_open)
        .agg({
            "open":   "first",
            "high":   "max",
            "low":    "min",
            "close":  "last",
            "volume": "sum",
        })
        .dropna(subset=["close"])
    )
    return resampled


def aggregate_to_2min(
        minute_candles: list[dict],
        num: int = 5,
) -> list[float]:
    """
    1-min candles → 2-min candles → last `num` close prices.
    """
    df = _prepare_dataframe(minute_candles)
    if df is None:
        return []

    resampled = _resample(df, "2min")
    closes = resampled["close"].tolist()
    return closes[-num:] if len(closes) >= num else closes


def aggregate_to_5min(
        minute_candles: list[dict],
        num: int = 3,
) -> list[float]:
    """
    1-min candles → 5-min candles → last `num` close prices.
    """
    df = _prepare_dataframe(minute_candles)
    if df is None:
        return []

    resampled = _resample(df, "5min")
    closes = resampled["close"].tolist()
    return closes[-num:] if len(closes) >= num else closes


# ═══════════════════════════════════════════════════════
#  OBV (On-Balance Volume)
# ═══════════════════════════════════════════════════════
def calculate_obv(candles: list[dict]) -> int:
    """Cumulative OBV from OHLCV candle dicts."""
    if not candles or len(candles) < 2:
        return 0

    obv = 0
    for i in range(1, len(candles)):
        if candles[i]["close"] > candles[i - 1]["close"]:
            obv += candles[i]["volume"]
        elif candles[i]["close"] < candles[i - 1]["close"]:
            obv -= candles[i]["volume"]
    return obv


# ═══════════════════════════════════════════════════════
#  VWAP (Volume-Weighted Average Price)
# ═══════════════════════════════════════════════════════
def calculate_vwap(candles: list[dict]) -> float:
    """
    VWAP = Σ(typical_price × volume) / Σ(volume)
    typical_price = (high + low + close) / 3
    """
    if not candles:
        return 0.0

    cum_tp_vol = 0.0
    cum_vol    = 0

    for c in candles:
        tp = (c["high"] + c["low"] + c["close"]) / 3.0
        vol = c.get("volume", 0)
        cum_tp_vol += tp * vol
        cum_vol    += vol

    return round(cum_tp_vol / cum_vol, 2) if cum_vol > 0 else 0.0


# ═══════════════════════════════════════════════════════
#  Relative Strength vs ATM
# ═══════════════════════════════════════════════════════
def calculate_relative_strength(
        opt_open: float,
        opt_ltp: float,
        atm_open: float,
        atm_ltp: float,
) -> float:
    """
    RS = (option % change from open) / (ATM % change from open)

    RS > 1.0  → outperforming ATM (stronger momentum)
    RS = 1.0  → same as ATM
    RS < 1.0  → underperforming ATM
    RS < 0    → moving opposite to ATM

    Returns 0.0 if data is insufficient.
    """
    if opt_open <= 0 or atm_open <= 0:
        return 0.0

    opt_return = (opt_ltp - opt_open) / opt_open
    atm_return = (atm_ltp - atm_open) / atm_open

    # If ATM hasn't moved, can't compute RS
    if abs(atm_return) < 0.0001:
        # Still show if option itself is up
        if opt_return > 0.01:
            return 1.5    # clearly outperforming
        elif opt_return < -0.01:
            return 0.5    # clearly underperforming
        return 1.0        # both flat

    rs = opt_return / atm_return
    return round(rs, 2)


def get_open_price(candles: list[dict]) -> float:
    """Extract day's opening price from first candle."""
    if candles:
        return candles[0]["open"]
    return 0.0


# ═══════════════════════════════════════════════════════
#  Trend Detection
# ═══════════════════════════════════════════════════════
def is_ascending(closes: list[float]) -> bool:
    """True if every element is strictly greater than the previous."""
    if len(closes) < 2:
        return False
    return all(closes[i] < closes[i + 1] for i in range(len(closes) - 1))


# ═══════════════════════════════════════════════════════
#  Sort Score
# ═══════════════════════════════════════════════════════
def sort_score(
        asc_2m: bool,
        asc_5m: bool,
        rs: float,
        volume: int,
) -> tuple:
    """
    Sort key (higher = shown first).
    Priority:
      1. 5-min ascending
      2. 2-min ascending
      3. Relative Strength (higher = stronger momentum)
      4. Volume (higher = more liquid)
    """
    return (
        1 if asc_5m else 0,
        1 if asc_2m else 0,
        rs if rs else 0,
        volume,
    )


# ═══════════════════════════════════════════════════════
#  Master Row Builder
# ═══════════════════════════════════════════════════════
def build_option_row(
        symbol: str,
        ltp: float,
        volume: int,
        oi: int,
        obv: int,
        vwap: float,
        rs: float,
        candles_2m: list[float],
        candles_5m: list[float],
) -> dict:
    """Build one row of data for the Excel dashboard."""

    asc_2m = is_ascending(candles_2m)
    asc_5m = is_ascending(candles_5m)
    score  = sort_score(asc_2m, asc_5m, rs, volume)

    return {
        "symbol":     symbol,
        "ltp":        ltp,
        "volume":     volume,
        "oi":         oi,
        "obv":        obv,
        "vwap":       vwap,
        "rs":         rs,
        "candles_2m": candles_2m,
        "asc_2m":     asc_2m,
        "candles_5m": candles_5m,
        "asc_5m":     asc_5m,
        "score":      score,
    }