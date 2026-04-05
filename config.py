"""
════════════════════════════════════════════════════════
  CONFIGURATION — Edit these values before first run
════════════════════════════════════════════════════════
"""

import os

# ── Zerodha API ──────────────────────────────────────
API_KEY      = os.getenv("API_KEY")    # enter your zerodha api key here
API_SECRET   = os.getenv("API_SECRET")  # enter you zerodha secret key here
ACCESS_TOKEN = ""                 # blank = browser login

# ── Instruments ──────────────────────────────────────
EXCHANGE      = "NFO"
UNDERLYING    = "NIFTY"           # NIFTY | BANKNIFTY | FINNIFTY
EXPIRY        = "2026-03-10"      # nearest expiry YYYY-MM-DD
STRIKE_STEP   = 50                # 50 NIFTY, 100 BANKNIFTY
NUM_STRIKES   = 15                # each side of ATM
OPTION_TYPES  = ["CE", "PE"]

# ── Candle settings ─────────────────────────────────
NUM_2MIN_CANDLES = 5
NUM_5MIN_CANDLES = 3

# ── Refresh rates (seconds) ─────────────────────────
EXCEL_PUSH_INTERVAL   = 3
CANDLE_FETCH_INTERVAL = 30
BULK_FETCH_WORKERS    = 3        # concurrent API threads
RATE_LIMIT_PER_SEC    = 3        # Zerodha allows 3 hist/sec

# ── Excel ────────────────────────────────────────────
WORKBOOK_NAME = "Options_Tracker"

# ── Telegram Alerts (leave blank to disable) ────────
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")          # from @BotFather
TELEGRAM_CHAT_ID   = os.getenv("TELEGRAM_CHAT_ID")          # your chat/group ID
TELEGRAM_ENABLED   = bool(TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID)

# ── Feature toggles ─────────────────────────────────
SHOW_VWAP = True
SHOW_RS   = True                 # Relative Strength vs ATM

# ── Logging ──────────────────────────────────────────
LOG_LEVEL = "INFO"