"""
Telegram Alert System
  • Tracks which options are fully ascending
  • Sends alert when NEW options become ascending
  • Formats rich messages with option details
"""

import logging
from datetime import datetime
from typing import Optional

import requests as http_requests

import config

log = logging.getLogger("tracker.telegram")


class TelegramNotifier:
    """
    Send alerts via Telegram Bot API when options
    become fully ascending (both 2m and 5m).
    """

    def __init__(self):
        self.bot_token = config.TELEGRAM_BOT_TOKEN
        self.chat_id   = config.TELEGRAM_CHAT_ID
        self.base_url  = f"https://api.telegram.org/bot{self.bot_token}"
        self.enabled   = config.TELEGRAM_ENABLED

        # Track previously ascending options to detect NEW ones
        self._prev_fully_ascending: set[str] = set()
        self._prev_5m_ascending: set[str]    = set()

        if self.enabled:
            self._send_startup_message()

    # ── Public API ───────────────────────────────────

    def check_and_alert(self, sorted_rows: list[dict]):
        """
        Compare current ascending set with previous.
        Send Telegram alert for NEWLY ascending options.
        """
        if not self.enabled:
            return

        # Current sets
        curr_fully_asc = set()     # both 2m AND 5m ascending
        curr_5m_asc    = set()     # only 5m ascending

        for row in sorted_rows:
            sym = row["symbol"]
            if row["asc_5m"] and row["asc_2m"]:
                curr_fully_asc.add(sym)
            elif row["asc_5m"]:
                curr_5m_asc.add(sym)

        # Find newly ascending
        new_fully = curr_fully_asc - self._prev_fully_ascending
        new_5m    = curr_5m_asc - self._prev_5m_ascending

        # Build detail dicts for new ones
        new_fully_rows = [r for r in sorted_rows if r["symbol"] in new_fully]
        new_5m_rows    = [r for r in sorted_rows if r["symbol"] in new_5m]

        # Send alerts
        if new_fully_rows:
            msg = self._format_fully_ascending_alert(new_fully_rows)
            self._send_message(msg)
            log.info(f"Telegram: alerted {len(new_fully_rows)} fully ascending options")

        if new_5m_rows and len(new_5m_rows) <= 5:
            msg = self._format_5m_ascending_alert(new_5m_rows)
            self._send_message(msg)

        # Update state
        self._prev_fully_ascending = curr_fully_asc
        self._prev_5m_ascending    = curr_5m_asc

    # ── Message formatting ───────────────────────────

    def _format_fully_ascending_alert(self, rows: list[dict]) -> str:
        now = datetime.now().strftime("%H:%M:%S")
        lines = [
            "🟢🟢 <b>FULLY ASCENDING OPTIONS</b> 🟢🟢",
            f"<i>Both 2-min and 5-min candles ascending</i>",
            "",
        ]

        for r in rows:
            lines.append(f"📊 <b>{r['symbol']}</b>")
            lines.append(f"   💰 LTP: ₹{r['ltp']:.2f}")
            lines.append(f"   📦 Vol: {r['volume']:,}  |  OI: {r['oi']:,}")

            if r.get("vwap"):
                vwap_rel = "above" if r["ltp"] > r["vwap"] else "below"
                lines.append(f"   📏 VWAP: ₹{r['vwap']:.2f} (LTP {vwap_rel})")

            if r.get("rs") and r["rs"] != 0:
                rs_emoji = "💪" if r["rs"] > 1.0 else "📉"
                lines.append(f"   {rs_emoji} RS vs ATM: {r['rs']:.2f}")

            lines.append(f"   📈 OBV: {r['obv']:,}")

            if r["candles_2m"]:
                c2 = " → ".join(f"{c:.1f}" for c in r["candles_2m"])
                lines.append(f"   2m: {c2} ✅")

            if r["candles_5m"]:
                c5 = " → ".join(f"{c:.1f}" for c in r["candles_5m"])
                lines.append(f"   5m: {c5} ✅")

            lines.append("")

        lines.append(f"⏰ {now}  |  {len(rows)} new fully ascending")
        return "\n".join(lines)

    def _format_5m_ascending_alert(self, rows: list[dict]) -> str:
        now = datetime.now().strftime("%H:%M:%S")
        lines = [
            "🔵 <b>5-MIN ASCENDING OPTIONS</b>",
            "",
        ]

        for r in rows:
            c5 = " → ".join(f"{c:.1f}" for c in r["candles_5m"]) if r["candles_5m"] else "—"
            lines.append(
                f"  • <b>{r['symbol']}</b>  "
                f"₹{r['ltp']:.2f}  "
                f"5m: {c5} ▲"
            )

        lines.append(f"\n⏰ {now}")
        return "\n".join(lines)

    def _send_startup_message(self):
        now = datetime.now().strftime("%H:%M:%S")
        msg = (
            "🚀 <b>Options Tracker Started</b>\n\n"
            f"📋 Underlying: {config.UNDERLYING}\n"
            f"📅 Expiry: {config.EXPIRY}\n"
            f"🎯 Strikes: ±{config.NUM_STRIKES} × {config.STRIKE_STEP}\n"
            f"⏰ Started at: {now}\n\n"
            "<i>You'll receive alerts when options become ascending.</i>"
        )
        self._send_message(msg)

    # ── Low-level send ───────────────────────────────

    def _send_message(self, text: str):
        """Send a message via Telegram Bot API."""
        if not self.enabled:
            return

        url = f"{self.base_url}/sendMessage"
        payload = {
            "chat_id":    self.chat_id,
            "text":       text,
            "parse_mode": "HTML",
            "disable_web_page_preview": True,
        }

        try:
            resp = http_requests.post(url, json=payload, timeout=10)
            if resp.status_code != 200:
                log.warning(
                    f"Telegram API error {resp.status_code}: "
                    f"{resp.text[:200]}"
                )
        except http_requests.exceptions.Timeout:
            log.warning("Telegram send timed out")
        except http_requests.exceptions.ConnectionError:
            log.warning("Telegram connection failed (no internet?)")
        except Exception as e:
            log.warning(f"Telegram send error: {e}")


class DummyNotifier:
    """No-op notifier when Telegram is disabled."""

    def check_and_alert(self, sorted_rows: list[dict]):
        pass


def create_notifier() -> TelegramNotifier | DummyNotifier:
    """Factory: return real or dummy notifier based on config."""
    if config.TELEGRAM_ENABLED:
        log.info("Telegram alerts ENABLED")
        return TelegramNotifier()
    else:
        log.info("Telegram alerts DISABLED (no credentials in config)")
        return DummyNotifier()