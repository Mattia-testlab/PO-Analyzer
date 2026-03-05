"""
utils.py — Helper functions for PO_Extract.

Provides date/number parsing, sheet-name sanitization, filename utilities,
and logging setup.
"""

import logging
import re
from datetime import datetime
from pathlib import Path


# ── Date parsing ────────────────────────────────────────────────────────────

def parse_date(date_str: str) -> str:
    """Convert an Italian-format date string to YYYY-MM-DD.

    Supported inputs:
        DD.MM.YY   → 20YY-MM-DD  (e.g. 02.03.26 → 2026-03-02)
        DD.MM.YYYY → YYYY-MM-DD
        DD/MM/YY   → 20YY-MM-DD
        DD/MM/YYYY → YYYY-MM-DD

    Returns the original string if parsing fails.
    """
    if not date_str or not isinstance(date_str, str):
        return date_str

    date_str = date_str.strip()

    for fmt_in, fmt_out in [
        (r"(\d{2})[./](\d{2})[./](\d{2})$",   None),   # DD.MM.YY
        (r"(\d{2})[./](\d{2})[./](\d{4})$",   None),   # DD.MM.YYYY
    ]:
        m = re.match(fmt_in, date_str)
        if m:
            day, month, year = m.groups()
            if len(year) == 2:
                year = f"20{year}"
            try:
                dt = datetime(int(year), int(month), int(day))
                return dt.strftime("%Y-%m-%d")
            except ValueError:
                return date_str
    return date_str


# ── Number parsing ──────────────────────────────────────────────────────────

def parse_number(num_str: str) -> float | None:
    """Convert a European-format number string to a Python float.

    Examples:
        '34,16'      → 34.16
        '1.234,56'   → 1234.56
        '683,20'     → 683.20
        '20'         → 20.0

    Returns None if parsing fails.
    """
    if num_str is None:
        return None

    if isinstance(num_str, (int, float)):
        return float(num_str)

    num_str = str(num_str).strip()
    if not num_str:
        return None

    # Remove thousands separator (dots before the comma)
    # European: 1.234,56 → 1234,56
    num_str = re.sub(r"\.(?=\d{3}(?:[,\s]|$))", "", num_str)
    # Replace decimal comma with dot
    num_str = num_str.replace(",", ".")

    try:
        return float(num_str)
    except ValueError:
        return None


# ── Sheet name utilities ────────────────────────────────────────────────────

_INVALID_SHEET_CHARS = re.compile(r"[\\/*?\[\]:]")


def sanitize_sheet_name(name: str, max_len: int = 31) -> str:
    """Return an Excel-safe sheet name (max *max_len* characters)."""
    name = _INVALID_SHEET_CHARS.sub("_", name)
    # Remove leading/trailing apostrophes (Excel restriction)
    name = name.strip("'")
    # Strip the extension
    name = re.sub(r"\.[Pp][Dd][Ff]$", "", name)
    return name[:max_len] if len(name) > max_len else name


def unique_sheet_name(name: str, existing: set[str]) -> str:
    """Ensure *name* is unique among *existing* sheet names.

    Appends _2, _3, … if needed.
    """
    base = name
    counter = 2
    while name in existing:
        suffix = f"_{counter}"
        name = base[: 31 - len(suffix)] + suffix
        counter += 1
    return name


# ── PO number from filename ────────────────────────────────────────────────

def extract_po_from_filename(filename: str) -> str:
    """Extract the PO number from a filename like 'PO#4559057061_2.PDF'.

    Returns empty string if not found.
    """
    m = re.search(r"(\d{8,})", filename)
    return m.group(1) if m else ""


# ── Logging ─────────────────────────────────────────────────────────────────

class GUILogHandler(logging.Handler):
    """Custom logging handler that appends messages to a GUI callback."""

    def __init__(self, callback):
        super().__init__()
        self.callback = callback

    def emit(self, record):
        msg = self.format(record)
        try:
            self.callback(msg)
        except Exception:
            pass


def setup_logging(
    log_file: str | Path | None = None,
    gui_callback=None,
    level: int = logging.INFO,
) -> logging.Logger:
    """Configure and return the application logger.

    Parameters
    ----------
    log_file : path, optional
        If given, logs are also written to this file.
    gui_callback : callable, optional
        ``callback(message_str)`` — called for every log record so the GUI
        can display it in real time.
    level : int
        Logging level (default ``INFO``).
    """
    logger = logging.getLogger("po_extract")
    logger.setLevel(level)
    logger.handlers.clear()

    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s",
                            datefmt="%H:%M:%S")

    # Console handler
    ch = logging.StreamHandler()
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    # File handler
    if log_file:
        fh = logging.FileHandler(log_file, encoding="utf-8")
        fh.setFormatter(fmt)
        logger.addHandler(fh)

    # GUI handler
    if gui_callback:
        gh = GUILogHandler(gui_callback)
        gh.setFormatter(fmt)
        logger.addHandler(gh)

    return logger
