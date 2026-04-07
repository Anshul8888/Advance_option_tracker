# README.md

# Zerodha Options Tracker

A live Excel-based options tracking dashboard for Zerodha that monitors option contracts around ATM, computes intraday momentum signals, and highlights strong contracts in real time.

It combines:

- Zerodha Kite Connect API
- Live WebSocket ticks
- Bulk historical candle fetching
- Excel dashboard via xlwings
- Optional Telegram alerts

---

## Features

- Track option strikes around ATM
- Supports:
  - `NIFTY`
  - `BANKNIFTY`
  - `FINNIFTY`
- Bulk fetch of 1-minute historical candles
- Local derivation of:
  - 2-minute candles
  - 5-minute candles
  - OBV
  - VWAP
  - Relative Strength vs ATM
- Detect ascending momentum on:
  - 2-minute candles
  - 5-minute candles
- Auto-rank and sort strongest contracts first
- Live Excel dashboard with color coding
- Optional Telegram alerts for newly ascending options
- Handles both:
  - market-open live mode
  - market-closed previous-session mode

---

## How It Works

The tracker follows this pipeline:

1. Authenticate with Zerodha
2. Detect the spot price of the configured underlying
3. Find the ATM strike
4. Select option contracts around ATM for the chosen expiry
5. Start live tick streaming using WebSocket
6. Periodically fetch 1-minute historical candles for all tracked contracts
7. Compute:
   - latest 2m closes
   - latest 5m closes
   - OBV
   - VWAP
   - opening prices
   - Relative Strength vs ATM
8. Rank all contracts by momentum and liquidity
9. Push the results into Excel
10. Send Telegram alerts for new qualifying contracts

---

## Metrics Used

### 1. Ascending Trend Detection

An option is considered ascending if its close prices are strictly increasing.

- **2m ascending** → latest 2-minute closes are rising
- **5m ascending** → latest 5-minute closes are rising
- **Fully ascending** → both 2m and 5m are rising

### 2. OBV (On-Balance Volume)

Calculated from candle close direction and volume to estimate buying/selling pressure.

### 3. VWAP

Volume-weighted average price calculated from intraday candles.

### 4. Relative Strength vs ATM

Compares the option’s percentage move from open against the ATM option of the same type:

- CE compared against ATM CE
- PE compared against ATM PE

This helps identify contracts outperforming the ATM baseline.

---

## Excel Dashboard

The project creates or opens an Excel workbook and updates it continuously.

### Dashboard columns

Typical columns include:

- Option
- LTP
- Volume
- OI
- OBV
- VWAP
- RS vs ATM
- last N 2-minute candle closes
- 2m trend
- last N 5-minute candle closes
- 5m trend

The exact layout depends on your config.

### Color coding

- **Green LTP** → both 2m and 5m are ascending
- **VWAP green** → LTP above VWAP
- **VWAP red** → LTP below VWAP
- **RS green** → strong outperformance
- **RS red** → underperformance
- **Gold option cell** → both 2m and 5m ascending
- **Trend cells**
  - `▲ ASC` = ascending
  - `▼ —` = not ascending

### Workbook behavior

- Workbook name comes from `WORKBOOK_NAME`
- File is saved as `<WORKBOOK_NAME>.xlsx`
- First sheet is named `Tracker`

---

## Telegram Alerts

Telegram alerts are optional.

When enabled, the tracker can send:

- startup notification
- alerts for **new fully ascending options**
- alerts for **new 5-minute ascending options**

Alerts are only sent when an option becomes newly qualified, which helps reduce spam.

### Alert contents may include

- symbol
- LTP
- volume
- OI
- VWAP relation
- RS vs ATM
- OBV
- 2m candle progression
- 5m candle progression

---

## Project Structure

```text
.
├── main.py
├── config.py
├── zerodha_client.py
├── data_processor.py
├── excel_dashboard.py
├── telegram_alert.py
├── README.md
├── CONTRIBUTING.md
└── LICENSE
```

## File overview
main.py → main orchestrator
config.py → user configuration
zerodha_client.py → Zerodha auth, instrument selection, historical fetch, WebSocket
data_processor.py → candle aggregation and signal calculations
excel_dashboard.py → Excel UI via xlwings
telegram_alert.py → Telegram notifier

## Requirements
You need:

Python 3.10+
A Zerodha account with Kite Connect access
Valid Zerodha API credentials
Microsoft Excel installed
Internet connection

Python packages
Install the dependencies:
pip install kiteconnect pandas numpy xlwings requests

Configuration
Before first run, edit config.py.

## Zerodha API
Python

API_KEY      = "your_api_key"
API_SECRET   = "your_api_secret"
ACCESS_TOKEN = ""   # leave blank for browser login
If ACCESS_TOKEN is blank, the app opens a browser login flow
After login, it prints a valid token you can save for later runs

## Instruments
Python

EXCHANGE      = "NFO"
UNDERLYING    = "NIFTY"
EXPIRY        = "2026-04-07"
STRIKE_STEP   = 50
NUM_STRIKES   = 15
OPTION_TYPES  = ["CE", "PE"]
This means the tracker will:

determine ATM strike
track strikes from ATM - NUM_STRIKES * STRIKE_STEP to ATM + NUM_STRIKES * STRIKE_STEP
include both calls and puts

## Candle settings
Python

NUM_2MIN_CANDLES = 5
NUM_5MIN_CANDLES = 3
Controls how many recent resampled closes are shown in the dashboard.

## Refresh rates
Python

EXCEL_PUSH_INTERVAL   = 3
CANDLE_FETCH_INTERVAL = 30
BULK_FETCH_WORKERS    = 3
RATE_LIMIT_PER_SEC    = 3
These values control dashboard responsiveness and API usage.

## Excel
Python

WORKBOOK_NAME = "Options_Tracker"

## Telegram
Python

TELEGRAM_BOT_TOKEN = ""
TELEGRAM_CHAT_ID   = ""
TELEGRAM_ENABLED   = bool(TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID)

## Feature toggles
Python

SHOW_VWAP = True
SHOW_RS   = True

## Logging
Python

LOG_LEVEL = "INFO"

## Zerodha Login Flow
If ACCESS_TOKEN is empty:

Run the app
A browser window opens to Zerodha login
Log in and authorize
Copy the full redirect URL
Paste it into the terminal
The app extracts the request_token
A new access token is generated
Save that token into config.py for future runs

## Running the App
Start the tracker with:

Bash

python main.py
On startup, the app will:

authenticate with Zerodha
select instruments around ATM
create or open the Excel workbook
start live WebSocket tick streaming
fetch initial candle data
begin updating the dashboard continuously

## Market Hours Behavior
The tracker adapts automatically depending on market timing:

During market hours

fetches today’s intraday candles
streams live ticks
Before market opens

uses the previous trading session’s candle data
After market closes

uses today’s completed session data
This allows the dashboard to load useful data even outside live hours.

# Ranking Logic
Rows are sorted using the following priority:

5-minute ascending trend
2-minute ascending trend
Relative Strength vs ATM
Volume
This places stronger and more liquid contracts near the top.

## Performance Notes
Historical data fetching is optimized by:

making one 1-minute historical API call per token
deriving 2m / 5m candles locally
calculating OBV and VWAP locally
using a thread pool for concurrency
using a rate limiter to stay within Zerodha API limits
This is much faster than making multiple interval-specific API calls.

# Example Use Case
This tool is useful for traders who want to quickly identify:

option contracts showing sustained intraday momentum
contracts outperforming ATM
high-volume and liquid contracts
symbols aligned across both short and medium intraday timeframes

# Troubleshooting
## No instruments found
Check:

UNDERLYING
EXPIRY
STRIKE_STEP
EXCHANGE

## Browser login fails
Make sure:

API key and secret are correct
you paste the full redirect URL
your Zerodha session is active

## Excel does not update
Check:

Excel is installed
xlwings is installed
workbook is not blocked by permissions
Excel automation works on your machine

## Telegram alerts not working
Check:

TELEGRAM_BOT_TOKEN
TELEGRAM_CHAT_ID
internet connectivity
whether the bot can message the target chat or group

## Contributing
Contributions are welcome. Please read CONTRIBUTING.md before opening a pull request.

## License
This project is licensed under the MIT License. See the LICENSE file for details.





