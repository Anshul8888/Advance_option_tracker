"""
Excel Dashboard via xlwings
  • Bulk data write (fast)
  • VWAP + RS columns
  • Conditional green/red coloring
  • Auto-sort: ascending options on top
"""

import os
import logging
from typing import Optional

import xlwings as xw

import config

log = logging.getLogger("tracker.excel")

# ── Colour palette ──────────────────────────────────
GREEN_BG       = (198, 239, 206)
GREEN_DARK_FG  = (0, 128, 0)
RED_BG         = (255, 199, 206)
RED_DARK_FG    = (180, 0, 0)
HEADER_BG      = (44, 62, 80)
HEADER_FG      = (255, 255, 255)
NEUTRAL_BG     = (255, 255, 255)
GOLD_BG        = (255, 243, 205)
BLUE_BG        = (219, 234, 254)


# ═══════════════════════════════════════════════════════
#  Column Layout Builder
# ═══════════════════════════════════════════════════════
def _build_static_headers() -> list[str]:
    """First N fixed columns before candle data."""
    headers = ["Option", "LTP", "Volume", "OI", "OBV"]
    if config.SHOW_VWAP:
        headers.append("VWAP")
    if config.SHOW_RS:
        headers.append("RS vs ATM")
    return headers


def _build_full_headers() -> list[str]:
    headers = _build_static_headers()
    for i in range(config.NUM_2MIN_CANDLES):
        headers.append(f"2m-{config.NUM_2MIN_CANDLES - i}")
    headers.append("2m Trend")
    for i in range(config.NUM_5MIN_CANDLES):
        headers.append(f"5m-{config.NUM_5MIN_CANDLES - i}")
    headers.append("5m Trend")
    return headers


class ColumnLayout:
    """Calculate 1-based column indices dynamically."""

    def __init__(self):
        static = _build_static_headers()
        self.num_static = len(static)

        self.COL_OPTION  = 1
        self.COL_LTP     = 2
        self.COL_VOLUME  = 3
        self.COL_OI      = 4
        self.COL_OBV     = 5

        next_col = 6
        if config.SHOW_VWAP:
            self.COL_VWAP = next_col
            next_col += 1
        else:
            self.COL_VWAP = None

        if config.SHOW_RS:
            self.COL_RS = next_col
            next_col += 1
        else:
            self.COL_RS = None

        self.COL_2M_START = next_col
        self.COL_2M_END   = next_col + config.NUM_2MIN_CANDLES - 1
        self.COL_2M_TREND = self.COL_2M_END + 1

        self.COL_5M_START = self.COL_2M_TREND + 1
        self.COL_5M_END   = self.COL_5M_START + config.NUM_5MIN_CANDLES - 1
        self.COL_5M_TREND = self.COL_5M_END + 1

        self.TOTAL_COLS   = self.COL_5M_TREND


# ═══════════════════════════════════════════════════════
#  Excel Dashboard Class
# ═══════════════════════════════════════════════════════
class ExcelDashboard:

    def __init__(self):
        self.wb: Optional[xw.Book]  = None
        self.ws: Optional[xw.Sheet] = None
        self.full_headers = _build_full_headers()
        self.layout       = ColumnLayout()

    # ── Lifecycle ────────────────────────────────────

    def open(self):
        save_path = os.path.join(os.getcwd(), f"{config.WORKBOOK_NAME}.xlsx")

        # Try 1: find already-open workbook
        self.wb = self._find_open_workbook(config.WORKBOOK_NAME)

        # Try 2: open from disk
        if self.wb is None and os.path.exists(save_path):
            try:
                self.wb = xw.Book(save_path)
                log.info(f"Opened existing file: {save_path}")
            except Exception as e:
                log.warning(f"Could not open file: {e}")

        # Try 3: create new
        if self.wb is None:
            self.wb = xw.Book()
            try:
                self.wb.save(save_path)
                log.info(f"Created new workbook: {save_path}")
            except Exception as e:
                log.warning(f"Save failed: {e}. Using unsaved workbook.")

        self.ws = self.wb.sheets[0]
        try:
            self.ws.name = "Tracker"
        except Exception:
            pass

        self._write_headers()
        self._format_headers()
        log.info("Excel dashboard ready ✓")

    def _find_open_workbook(self, name: str) -> Optional[xw.Book]:
        try:
            for app in xw.apps:
                for book in app.books:
                    book_name = book.name.replace(".xlsx", "").replace(".xls", "")
                    if book_name.lower() == name.lower():
                        log.info(f"Found open workbook: {book.name}")
                        return book
        except Exception:
            pass
        return None

    def close(self):
        if self.wb:
            try:
                self.wb.save()
            except Exception:
                pass

    # ── Headers ──────────────────────────────────────

    def _write_headers(self):
        self.ws.range("A1").value = self.full_headers

    def _format_headers(self):
        L = self.layout
        rng = self.ws.range((1, 1), (1, L.TOTAL_COLS))
        rng.color     = HEADER_BG
        rng.font.color = HEADER_FG
        rng.font.bold  = True
        rng.font.size  = 11

        # Column widths
        self.ws.range((1, L.COL_OPTION)).column_width = 22
        self.ws.range((1, L.COL_LTP)).column_width    = 12
        for col in [L.COL_VOLUME, L.COL_OI, L.COL_OBV]:
            self.ws.range((1, col)).column_width = 13
        if L.COL_VWAP:
            self.ws.range((1, L.COL_VWAP)).column_width = 12
        if L.COL_RS:
            self.ws.range((1, L.COL_RS)).column_width = 12
        for c in range(L.COL_2M_START, L.TOTAL_COLS + 1):
            self.ws.range((1, c)).column_width = 11

    # ── Data Write (Optimized: bulk write + format) ──

    def write_all_rows(self, rows: list[dict]):
        """Write sorted rows to Excel with conditional formatting."""
        if not self.ws or not rows:
            return

        app = self.ws.book.app
        app.screen_updating = False
        app.calculation = "manual"
        L = self.layout

        try:
            # ── Step 1: Clear old data ──
            used = self.ws.used_range
            if used.last_cell.row > 1:
                self.ws.range(
                    (2, 1), (used.last_cell.row, L.TOTAL_COLS)
                ).clear()

            # ── Step 2: Build data matrix for bulk write ──
            data_matrix = []
            for row in rows:
                row_list = self._build_row_list(row)
                data_matrix.append(row_list)

            # Bulk write all data at once (much faster than cell-by-cell)
            if data_matrix:
                self.ws.range((2, 1)).value = data_matrix

            # ── Step 3: Apply formatting row by row ──
            for idx, row in enumerate(rows):
                excel_row = idx + 2
                self._format_row(excel_row, row)

            # ── Step 4: Number formatting ──
            end_row = len(rows) + 1
            if end_row > 1:
                self.ws.range((2, L.COL_LTP), (end_row, L.COL_LTP)).number_format = "#,##0.00"
                self.ws.range((2, L.COL_VOLUME), (end_row, L.COL_OBV)).number_format = "#,##0"
                if L.COL_VWAP:
                    self.ws.range((2, L.COL_VWAP), (end_row, L.COL_VWAP)).number_format = "#,##0.00"
                if L.COL_RS:
                    self.ws.range((2, L.COL_RS), (end_row, L.COL_RS)).number_format = "0.00"

        finally:
            app.screen_updating = True
            app.calculation = "automatic"

    def _build_row_list(self, row: dict) -> list:
        """Convert row dict to flat list matching column layout."""
        L = self.layout
        values = [None] * L.TOTAL_COLS

        values[L.COL_OPTION - 1] = row["symbol"]
        values[L.COL_LTP - 1]    = row["ltp"]
        values[L.COL_VOLUME - 1] = row["volume"]
        values[L.COL_OI - 1]     = row["oi"]
        values[L.COL_OBV - 1]    = row["obv"]

        if L.COL_VWAP:
            values[L.COL_VWAP - 1] = row.get("vwap", 0)
        if L.COL_RS:
            values[L.COL_RS - 1] = row.get("rs", 0)

        # 2-min candle closes
        c2 = row["candles_2m"]
        for i in range(config.NUM_2MIN_CANDLES):
            col_idx = L.COL_2M_START - 1 + i
            values[col_idx] = c2[i] if i < len(c2) else ""

        # 2-min trend
        values[L.COL_2M_TREND - 1] = "▲ ASC" if row["asc_2m"] else "▼ —"

        # 5-min candle closes
        c5 = row["candles_5m"]
        for i in range(config.NUM_5MIN_CANDLES):
            col_idx = L.COL_5M_START - 1 + i
            values[col_idx] = c5[i] if i < len(c5) else ""

        # 5-min trend
        values[L.COL_5M_TREND - 1] = "▲ ASC" if row["asc_5m"] else "▼ —"

        return values

    def _format_row(self, excel_row: int, row: dict):
        """Apply conditional colors to a single row."""
        L = self.layout

        # ── LTP cell: green if BOTH timeframes ascending ──
        ltp_cell = self.ws.range((excel_row, L.COL_LTP))
        if row["asc_5m"] and row["asc_2m"]:
            ltp_cell.color     = GREEN_BG
            ltp_cell.font.bold = True
        else:
            ltp_cell.color     = NEUTRAL_BG
            ltp_cell.font.bold = False

        # ── VWAP cell: color based on LTP vs VWAP ──
        if L.COL_VWAP and row.get("vwap") and row["vwap"] > 0:
            vwap_cell = self.ws.range((excel_row, L.COL_VWAP))
            if row["ltp"] > row["vwap"]:
                vwap_cell.color      = GREEN_BG
                vwap_cell.font.color = GREEN_DARK_FG
            elif row["ltp"] < row["vwap"]:
                vwap_cell.color      = RED_BG
                vwap_cell.font.color = RED_DARK_FG
            else:
                vwap_cell.color = NEUTRAL_BG

        # ── RS cell: color based on strength ──
        if L.COL_RS and row.get("rs"):
            rs_cell = self.ws.range((excel_row, L.COL_RS))
            rs = row["rs"]
            if rs > 1.2:
                rs_cell.color      = GREEN_BG
                rs_cell.font.color = GREEN_DARK_FG
                rs_cell.font.bold  = True
            elif rs < 0.8:
                rs_cell.color      = RED_BG
                rs_cell.font.color = RED_DARK_FG
                rs_cell.font.bold  = False
            else:
                rs_cell.color     = NEUTRAL_BG
                rs_cell.font.bold = False

        # ── 2-min candle cells ──
        c2 = row["candles_2m"]
        if c2:
            rng = self.ws.range(
                (excel_row, L.COL_2M_START),
                (excel_row, L.COL_2M_START + len(c2) - 1),
            )
            rng.color = GREEN_BG if row["asc_2m"] else NEUTRAL_BG

        # 2m trend cell
        t2 = self.ws.range((excel_row, L.COL_2M_TREND))
        if row["asc_2m"]:
            t2.color      = GREEN_BG
            t2.font.color = GREEN_DARK_FG
            t2.font.bold  = True
        else:
            t2.color      = RED_BG
            t2.font.color = RED_DARK_FG
            t2.font.bold  = False

        # ── 5-min candle cells ──
        c5 = row["candles_5m"]
        if c5:
            rng = self.ws.range(
                (excel_row, L.COL_5M_START),
                (excel_row, L.COL_5M_START + len(c5) - 1),
            )
            rng.color = GREEN_BG if row["asc_5m"] else NEUTRAL_BG

        # 5m trend cell
        t5 = self.ws.range((excel_row, L.COL_5M_TREND))
        if row["asc_5m"]:
            t5.color      = GREEN_BG
            t5.font.color = GREEN_DARK_FG
            t5.font.bold  = True
        else:
            t5.color      = RED_BG
            t5.font.color = RED_DARK_FG
            t5.font.bold  = False

        # ── Full-row highlight for BOTH ascending ──
        if row["asc_5m"] and row["asc_2m"]:
            opt_cell = self.ws.range((excel_row, L.COL_OPTION))
            opt_cell.color     = GOLD_BG
            opt_cell.font.bold = True

    # ── Status bar ───────────────────────────────────

    def update_status(self, msg: str):
        if not self.ws:
            return
        used = self.ws.used_range
        status_row = max(used.last_cell.row + 2, 3)

        # Clear old status
        self.ws.range((status_row, 1), (status_row + 1, self.layout.TOTAL_COLS)).clear()

        cell = self.ws.range((status_row, 1))
        cell.value      = msg
        cell.font.italic = True
        cell.font.color  = (100, 100, 100)
        cell.font.size   = 10