"""
Zerodha KiteConnect wrapper
  • Browser-based login
  • BULK historical data fetch (1 call per token → derive all)
  • Rate limiter for API compliance
  • WebSocket live ticks
"""

import time
import webbrowser
import logging
import threading
from collections import deque
from datetime import datetime, timedelta, date
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Thread

from kiteconnect import KiteConnect, KiteTicker

import config

log = logging.getLogger("tracker.zerodha")


# ═══════════════════════════════════════════════════════
#  Market hours helpers
# ═══════════════════════════════════════════════════════
MARKET_OPEN_HOUR    = 9
MARKET_OPEN_MINUTE  = 15
MARKET_CLOSE_HOUR   = 15
MARKET_CLOSE_MINUTE = 30


def is_market_open() -> bool:
    now = datetime.now()
    mkt_open  = now.replace(hour=MARKET_OPEN_HOUR,  minute=MARKET_OPEN_MINUTE,
                            second=0, microsecond=0)
    mkt_close = now.replace(hour=MARKET_CLOSE_HOUR, minute=MARKET_CLOSE_MINUTE,
                            second=0, microsecond=0)
    return mkt_open <= now <= mkt_close


def _previous_weekday(d: date) -> date:
    d = d - timedelta(days=1)
    while d.weekday() >= 5:
        d -= timedelta(days=1)
    return d


def get_last_trading_session_times() -> tuple[datetime, datetime]:
    """
    Return (start, end) of most recent trading session.
    Before market → previous weekday's session.
    During market → today open to now.
    After market  → today open to close.
    """
    now = datetime.now()
    today_open  = now.replace(hour=MARKET_OPEN_HOUR,  minute=MARKET_OPEN_MINUTE,
                              second=0, microsecond=0)
    today_close = now.replace(hour=MARKET_CLOSE_HOUR, minute=MARKET_CLOSE_MINUTE,
                              second=0, microsecond=0)

    if now >= today_open:
        return today_open, min(now, today_close)
    else:
        prev = _previous_weekday(now.date())
        return (
            datetime.combine(prev, today_open.time()),
            datetime.combine(prev, today_close.time()),
        )


def get_safe_from_to(minutes_back: int = 375) -> tuple[datetime, datetime]:
    """
    Return (from_dt, to_dt) that is ALWAYS valid
    (from_dt < to_dt), handling pre-market hours.
    375 min = 6h15m covers full trading day.
    """
    now = datetime.now()
    today_open = now.replace(hour=MARKET_OPEN_HOUR, minute=MARKET_OPEN_MINUTE,
                             second=0, microsecond=0)

    if now >= today_open:
        to_dt   = now
        from_dt = max(today_open, now - timedelta(minutes=minutes_back))
    else:
        prev = _previous_weekday(now.date())
        prev_open  = datetime.combine(prev, today_open.time())
        prev_close = datetime.combine(
            prev,
            now.replace(hour=MARKET_CLOSE_HOUR, minute=MARKET_CLOSE_MINUTE,
                        second=0, microsecond=0).time(),
        )
        to_dt   = prev_close
        from_dt = max(prev_open, prev_close - timedelta(minutes=minutes_back))

    return from_dt, to_dt


# ═══════════════════════════════════════════════════════
#  Rate Limiter (thread-safe)
# ═══════════════════════════════════════════════════════
class RateLimiter:
    """
    Sliding-window rate limiter.
    Blocks callers beyond `max_calls` per `period` seconds.
    """

    def __init__(self, max_calls: int = 3, period: float = 1.0):
        self.max_calls = max_calls
        self.period    = period
        self._timestamps: deque[float] = deque()
        self._lock = threading.Lock()

    def acquire(self):
        """Block until a request slot is available."""
        while True:
            with self._lock:
                now = time.time()
                # discard timestamps outside the window
                while self._timestamps and now - self._timestamps[0] >= self.period:
                    self._timestamps.popleft()
                if len(self._timestamps) < self.max_calls:
                    self._timestamps.append(now)
                    return
            time.sleep(0.05)


# ═══════════════════════════════════════════════════════
#  Authentication
# ═══════════════════════════════════════════════════════
class ZerodhaAuth:

    def __init__(self):
        self.kite = KiteConnect(api_key=config.API_KEY)

    def login(self) -> KiteConnect:
        if config.ACCESS_TOKEN:
            self.kite.set_access_token(config.ACCESS_TOKEN)
            log.info("Using stored access token.")
            try:
                self.kite.profile()
            except Exception:
                log.warning("Stored token expired – browser login.")
                self._browser_login()
        else:
            self._browser_login()
        return self.kite

    def _browser_login(self):
        url = self.kite.login_url()
        print("\n╔══════════════════════════════════════════════════╗")
        print("║  STEP 1 — Browser opens → Log in to Zerodha     ║")
        print("║  STEP 2 — Copy the FULL redirect URL             ║")
        print("║  STEP 3 — Paste below                            ║")
        print("╚══════════════════════════════════════════════════╝\n")
        webbrowser.open(url)

        redirect_url = input("Paste redirect URL here → ").strip()
        request_token = redirect_url.split("request_token=")[1].split("&")[0]

        data = self.kite.generate_session(request_token, api_secret=config.API_SECRET)
        access_token = data["access_token"]
        self.kite.set_access_token(access_token)

        log.info(f"Login successful.  access_token = {access_token}")
        print(f"\n✅  Logged in!  Save in config.py to skip next time:")
        print(f'    ACCESS_TOKEN = "{access_token}"\n')


# ═══════════════════════════════════════════════════════
#  Instrument Selector
# ═══════════════════════════════════════════════════════
class InstrumentSelector:

    def __init__(self, kite: KiteConnect):
        self.kite = kite
        self.atm_strike: float = 0

    def get_spot_ltp(self) -> float:
        sym_map = {
            "NIFTY":     "NSE:NIFTY 50",
            "BANKNIFTY": "NSE:NIFTY BANK",
            "FINNIFTY":  "NSE:NIFTY FIN SERVICE",
        }
        symbol = sym_map.get(config.UNDERLYING, f"NSE:{config.UNDERLYING}")
        data = self.kite.ltp(symbol)
        return list(data.values())[0]["last_price"]

    def select(self) -> list[dict]:
        spot = self.get_spot_ltp()
        self.atm_strike = round(spot / config.STRIKE_STEP) * config.STRIKE_STEP
        log.info(f"Spot={spot:.2f}  ATM={self.atm_strike}")

        strikes = [
            self.atm_strike + i * config.STRIKE_STEP
            for i in range(-config.NUM_STRIKES, config.NUM_STRIKES + 1)
        ]

        expiry_date = date.fromisoformat(config.EXPIRY)
        all_instruments = self.kite.instruments(config.EXCHANGE)

        selected = []
        for inst in all_instruments:
            if (
                    inst["name"] == config.UNDERLYING
                    and inst["expiry"] == expiry_date
                    and inst["strike"] in strikes
                    and inst["instrument_type"] in config.OPTION_TYPES
            ):
                selected.append(inst)

        selected.sort(key=lambda x: (x["strike"], x["instrument_type"]))
        log.info(f"Tracking {len(selected)} option contracts.")
        return selected


# ═══════════════════════════════════════════════════════
#  Bulk Historical Fetcher  (⚡ 3× faster)
# ═══════════════════════════════════════════════════════
class BulkHistoricalFetcher:
    """
    ONE API call per token (1-min candles for the full day).
    All derivations (2m, 5m, OBV, VWAP) are done locally.

    OLD:  62 tokens × 3 calls × 0.4s gap = ~74 seconds
    NEW:  62 tokens × 1 call  / 3 per sec = ~21 seconds  (3.5× faster)
    """

    def __init__(self, kite: KiteConnect):
        self.kite = kite
        self.rate_limiter = RateLimiter(
            max_calls=config.RATE_LIMIT_PER_SEC,
            period=1.0,
        )

    def fetch_all(
            self,
            tokens: list[int],
            progress_callback=None,
    ) -> dict[int, list[dict]]:
        """
        Fetch 1-min candle data for ALL tokens concurrently.
        Returns: {token: [candle_dict, ...]}
        """
        results: dict[int, list[dict]] = {}
        completed = 0
        total = len(tokens)
        lock = threading.Lock()
        t0 = time.time()

        def _fetch_one(token: int) -> tuple[int, list[dict]]:
            self.rate_limiter.acquire()
            from_dt, to_dt = get_safe_from_to(minutes_back=375)
            try:
                data = self.kite.historical_data(
                    instrument_token=token,
                    from_date=from_dt,
                    to_date=to_dt,
                    interval="minute",
                )
                return token, data
            except Exception as e:
                log.warning(f"Bulk fetch error token={token}: {e}")
                return token, []

        with ThreadPoolExecutor(max_workers=config.BULK_FETCH_WORKERS) as executor:
            futures = {
                executor.submit(_fetch_one, t): t
                for t in tokens
            }

            for future in as_completed(futures):
                token, data = future.result()
                with lock:
                    results[token] = data
                    completed += 1

                    if completed % 10 == 0 or completed == total:
                        elapsed = time.time() - t0
                        log.info(
                            f"Fetched {completed}/{total} tokens "
                            f"({elapsed:.1f}s)"
                        )
                        if progress_callback:
                            progress_callback(completed, total)

        elapsed = time.time() - t0
        log.info(f"Bulk fetch complete: {total} tokens in {elapsed:.1f}s")
        return results


# ═══════════════════════════════════════════════════════
#  WebSocket Tick Streamer
# ═══════════════════════════════════════════════════════
class TickStreamer:

    def __init__(self, kite: KiteConnect, tokens: list[int], on_tick_callback):
        self.kite     = kite
        self.tokens   = tokens
        self.callback = on_tick_callback

        self.ticker = KiteTicker(config.API_KEY, kite.access_token)
        self.ticker.on_ticks   = self._on_ticks
        self.ticker.on_connect = self._on_connect
        self.ticker.on_close   = self._on_close
        self.ticker.on_error   = self._on_error

    def start(self):
        thread = Thread(
            target=self.ticker.connect,
            kwargs={"threaded": True},
            daemon=True,
        )
        thread.start()
        log.info("WebSocket thread started.")

    def _on_connect(self, ws, response):
        log.info(f"WebSocket connected. Subscribing {len(self.tokens)} tokens.")
        ws.subscribe(self.tokens)
        ws.set_mode(ws.MODE_FULL, self.tokens)

    def _on_ticks(self, ws, ticks):
        tick_map = {}
        for t in ticks:
            tick_map[t["instrument_token"]] = {
                "ltp":    t.get("last_price", 0),
                "volume": t.get("volume_traded", t.get("volume", 0)),
                "oi":     t.get("oi", 0),
                "change": t.get("change", 0),
            }
        self.callback(tick_map)

    def _on_close(self, ws, code, reason):
        log.warning(f"WebSocket closed: {code} – {reason}")

    def _on_error(self, ws, code, reason):
        log.error(f"WebSocket error: {code} – {reason}")