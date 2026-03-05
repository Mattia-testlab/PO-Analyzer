"""
extractor.py — PDF parsing and data extraction for QVC purchase orders.

Supports three engines:
    - pdfplumber+regex  (default, best for this document type)
    - pdfplumber        (table extraction fallback)
    - camelot           (lattice/stream table detection)

Also provides cleaning / normalisation and MASTER-sheet composition.
"""

import logging
import re
from pathlib import Path

import pandas as pd

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import camelot
except ImportError:
    camelot = None

from utils import extract_po_from_filename, parse_date, parse_number

logger = logging.getLogger("po_extract")

# ── Column schema ───────────────────────────────────────────────────────────

COLUMNS = [
    "nome_file",
    "PO",
    "QVC",
    "MODVR",
    "modello",
    "VR",
    "taglia",
    "descrizione",
    "data_consegna",
    "quantita",
    "prezzo_unitario",
    "valore_netto",
    "paese_origine",
    "note",
]

# ── Regex patterns (derived from real PDF output) ───────────────────────────

# First line of an item block:
#   154957 ONJ W40 Motivi PANTALONE PALAZZO CON CINTUR 02.03.26 20 34,16 683,20
# OR (no space before description, seen on page 3):
#   154958 QPW MMotivi MAGLIA KIMONETTA cacao M 02.03.26 15 15,16 227,40
#
# QVC = 6-digit code + space + alphanumeric variant + space + size (or color-size)
# We capture the QVC part, the description, date, qty, unit price, net value.
# Pattern matches lines like:
#   154957 ONJ W40 Motivi PANTALONE PALAZZO CON CINTUR 02.03.26 20 34,16 683,20
#   154958 QPW MMotivi MAGLIA KIMONETTA cacao M 02.03.26 15 15,16 227,40
# The QVC code is: 6 digits + variant + size.  The size might run into the
# description when the PDF has no space ("QPW MMotivi").
RE_ITEM_LINE = re.compile(
    r"^(\d{6}\s+\S+\s+\S+?)"      # QVC code — non-greedy last token
    r"\s*"                          # optional space (sometimes missing!)
    r"([A-Z][a-z].+?)"             # description starts with a Capital letter
    r"\s+"
    r"(\d{2}\.\d{2}\.\d{2})"       # date DD.MM.YY
    r"\s+"
    r"(\d[\d.]*)"                   # quantity
    r"\s+"
    r"([\d.,]+)"                    # unit price
    r"\s+"
    r"([\d.,]+)"                    # net value
    r"\s*$"
)

RE_COD_ART_FORN = re.compile(r"Cod\.art\.forn\.\s*:\s*(\S+)", re.IGNORECASE)
RE_PAESE = re.compile(r"Paese di origine\s*:\s*(.+)", re.IGNORECASE)
RE_PO_HEADER = re.compile(r"^\s*(\d{10})\s*$")


# ── Engine: pdfplumber + regex ──────────────────────────────────────────────

def extract_with_pdfplumber_regex(pdf_path: str | Path) -> pd.DataFrame:
    """Extract purchase-order line items using pdfplumber text + regex."""
    if pdfplumber is None:
        raise ImportError("pdfplumber is not installed")

    pdf_path = Path(pdf_path)
    filename = pdf_path.name
    po_number = extract_po_from_filename(filename)

    records: list[dict] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            text = page.extract_text()
            if not text:
                continue

            lines = text.split("\n")

            # Try to grab PO from header if not already found
            if not po_number:
                for line in lines:
                    m = RE_PO_HEADER.match(line)
                    if m:
                        po_number = m.group(1)
                        break

            # Parse item blocks
            i = 0
            while i < len(lines):
                line = lines[i].strip()
                m = RE_ITEM_LINE.match(line)
                if m:
                    qvc = m.group(1).strip()
                    descrizione = m.group(2).strip()
                    data_consegna = m.group(3).strip()
                    quantita = m.group(4).strip()
                    prezzo_unitario = m.group(5).strip()
                    valore_netto = m.group(6).strip()

                    # Extract taglia (size) = last token of QVC code
                    qvc_parts = qvc.split()
                    taglia = qvc_parts[-1] if len(qvc_parts) >= 3 else ""

                    # Scan continuation lines for Cod.art.forn and Paese
                    modvr = ""
                    paese = ""
                    note_lines: list[str] = []
                    j = i + 1
                    while j < len(lines):
                        cline = lines[j].strip()
                        # Stop if we hit the next item line
                        if RE_ITEM_LINE.match(cline):
                            break

                        # For lines with underscores, strip them to check
                        # for embedded text (pdfplumber rendering artifact).
                        # e.g. "__P_a_e_s_e_ _d_i _o_ri_g_in_e_ : Myanmar__"
                        clean_line = cline.replace("_", "")

                        # Check for article code (may have empty value)
                        cm = RE_COD_ART_FORN.search(cline)
                        if cm:
                            modvr = cm.group(1).strip()
                            j += 1
                            continue
                        # Handle "Cod.art.forn.:" with NO value after it
                        if re.search(r"Cod\.art\.forn\.\s*:\s*$", cline, re.IGNORECASE):
                            j += 1
                            continue

                        # Check for country of origin (also in underscored lines)
                        pm = RE_PAESE.search(cline) or RE_PAESE.search(clean_line)
                        if pm:
                            raw_paese = pm.group(1).strip()
                            # Clean underscores out of the value too
                            paese = raw_paese.replace("_", "").strip()
                            j += 1
                            continue

                        # Skip lines that are mostly underscores (separators,
                        # totale-netto line, etc.)  — if >40% of chars are '_'
                        if len(cline) > 4 and cline.count("_") / len(cline) > 0.4:
                            j += 1
                            continue

                        # Stop scanning if we hit the Terms & Conditions section
                        if cline.startswith(".Assicurarsi") or cline.startswith("Si prega il fornitore"):
                            break

                        # Skip known boilerplate
                        boilerplate_hit = any(kw in cline.lower() for kw in [
                            "the article must", "be delivered according",
                            "requirements of the", "logistics of manual",
                            "of manual accordingly", "according to offer",
                            "qvc italia", "sede legale",
                            "cf/pi/registro",
                            # T&C section keywords
                            "assicurarsi che", "codice a barre",
                            "ordine di acquisto", "contratto approvato",
                            "manuale di controllo", "manuale logistico",
                            "data di consegna dell", "avviso della consegna",
                            "notifica di consegna", "garanzie di qualit",
                            "requisiti indicati", "omessa adesione",
                            "controllo qualita", "imballo sia perfettamente",
                            "si prega il fornitore", "si prega di controllare",
                            "qvc.inbound@geodis",
                            "corrispondenza", "codici articolo",
                        ])
                        # Also skip lines like "154958: The article must ..."
                        if not boilerplate_hit and re.match(r'^\d{6}:', cline):
                            boilerplate_hit = True
                        # Skip lines starting with . or - (T&C bullet points)
                        if not boilerplate_hit and re.match(r'^[.\-]', cline):
                            boilerplate_hit = True
                        if boilerplate_hit:
                            j += 1
                            continue
                        # Anything else → note
                        if cline and not cline.startswith("_"):
                            note_lines.append(cline)
                        j += 1

                    # Build modello / VR from MODVR
                    modello = modvr[:6] if len(modvr) >= 6 else modvr
                    vr = modvr[6:] if len(modvr) > 6 else ""

                    records.append({
                        "nome_file": filename,
                        "PO": po_number,
                        "QVC": qvc,
                        "MODVR": modvr,
                        "modello": modello,
                        "VR": vr,
                        "taglia": taglia,
                        "descrizione": descrizione,
                        "data_consegna": parse_date(data_consegna),
                        "quantita": int(parse_number(quantita) or 0),
                        "prezzo_unitario": parse_number(prezzo_unitario),
                        "valore_netto": parse_number(valore_netto),
                        "paese_origine": paese,
                        "note": "; ".join(note_lines) if note_lines else "",
                    })

                    i = j  # skip to the line after the block
                else:
                    i += 1

    logger.info("pdfplumber+regex: %d record(s) from '%s'", len(records), filename)
    df = pd.DataFrame(records, columns=COLUMNS)
    return df


# ── Engine: pdfplumber table extraction ─────────────────────────────────────

def extract_with_pdfplumber(pdf_path: str | Path) -> pd.DataFrame:
    """Try pdfplumber's built-in table extraction (often empty for these PDFs)."""
    if pdfplumber is None:
        raise ImportError("pdfplumber is not installed")

    pdf_path = Path(pdf_path)
    filename = pdf_path.name
    po_number = extract_po_from_filename(filename)
    all_rows: list[list] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    all_rows.append(row)

    if not all_rows:
        logger.warning("pdfplumber table extraction found 0 tables in '%s'", filename)
        return pd.DataFrame(columns=COLUMNS)

    # Best-effort: try to map columns
    records = []
    for row in all_rows:
        if row and len(row) >= 6:
            records.append({
                "nome_file": filename,
                "PO": po_number,
                "QVC": str(row[0] or "").strip(),
                "MODVR": "",
                "modello": "",
                "VR": "",
                "taglia": "",
                "descrizione": str(row[1] or "").strip(),
                "data_consegna": parse_date(str(row[2] or "")),
                "quantita": int(parse_number(str(row[3] or "")) or 0),
                "prezzo_unitario": parse_number(str(row[4] or "")),
                "valore_netto": parse_number(str(row[5] or "")),
                "paese_origine": "",
                "note": "",
            })

    logger.info("pdfplumber tables: %d record(s) from '%s'", len(records), filename)
    return pd.DataFrame(records, columns=COLUMNS)


# ── Engine: Camelot ─────────────────────────────────────────────────────────

def extract_with_camelot(pdf_path: str | Path) -> pd.DataFrame:
    """Extract tables using Camelot (lattice + stream)."""
    if camelot is None:
        raise ImportError("camelot-py is not installed")

    pdf_path = Path(pdf_path)
    filename = pdf_path.name
    po_number = extract_po_from_filename(filename)

    records = []
    for flavor in ("lattice", "stream"):
        try:
            tables = camelot.read_pdf(str(pdf_path), pages="all", flavor=flavor)
            for tbl in tables:
                df_t = tbl.df
                for _, row in df_t.iterrows():
                    vals = row.tolist()
                    if len(vals) >= 6:
                        records.append({
                            "nome_file": filename,
                            "PO": po_number,
                            "QVC": str(vals[0]).strip(),
                            "MODVR": "",
                            "modello": "",
                            "VR": "",
                            "taglia": "",
                            "descrizione": str(vals[1]).strip(),
                            "data_consegna": parse_date(str(vals[2])),
                            "quantita": int(parse_number(str(vals[3])) or 0),
                            "prezzo_unitario": parse_number(str(vals[4])),
                            "valore_netto": parse_number(str(vals[5])),
                            "paese_origine": "",
                            "note": "",
                        })
            if records:
                break
        except Exception as exc:
            logger.debug("camelot %s failed: %s", flavor, exc)

    logger.info("camelot: %d record(s) from '%s'", len(records), filename)
    return pd.DataFrame(records, columns=COLUMNS)


# ── Engine detection ────────────────────────────────────────────────────────

def detect_pdf_type(pdf_path: str | Path) -> str:
    """Return the recommended engine name for a given PDF.

    Currently always returns 'pdfplumber+regex' as QVC POs are text-based.
    """
    return "pdfplumber+regex"


# ── Dispatch ────────────────────────────────────────────────────────────────

ENGINE_MAP = {
    "pdfplumber+regex": extract_with_pdfplumber_regex,
    "pdfplumber": extract_with_pdfplumber,
    "camelot": extract_with_camelot,
}


def extract(pdf_path: str | Path, engine: str = "pdfplumber+regex") -> pd.DataFrame:
    """High-level extraction dispatcher."""
    func = ENGINE_MAP.get(engine)
    if func is None:
        raise ValueError(f"Unknown engine: {engine!r}")
    return func(pdf_path)


# ── Cleaning / normalisation ───────────────────────────────────────────────

def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Apply normalisation rules to an extracted DataFrame.

    - dates  → YYYY-MM-DD
    - prices → float
    - qty    → int
    - MODVR  → modello (first 6) + VR (rest)
    """
    df = df.copy()

    # Date normalisation (in case raw strings slipped through)
    if "data_consegna" in df.columns:
        df["data_consegna"] = df["data_consegna"].apply(
            lambda v: parse_date(str(v)) if pd.notna(v) else v
        )

    # Numeric normalisation
    for col in ("prezzo_unitario", "valore_netto"):
        if col in df.columns:
            df[col] = df[col].apply(
                lambda v: parse_number(v) if isinstance(v, str) else v
            )

    if "quantita" in df.columns:
        df["quantita"] = df["quantita"].apply(
            lambda v: int(parse_number(v) or 0) if isinstance(v, str) else int(v or 0)
        )

    # MODVR split
    if "MODVR" in df.columns:
        df["modello"] = df["MODVR"].apply(lambda v: str(v)[:6] if v else "")
        df["VR"] = df["MODVR"].apply(lambda v: str(v)[6:] if v and len(str(v)) > 6 else "")

    return df


# ── MASTER sheet composition ───────────────────────────────────────────────

def compose_master(dfs: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Concatenate all per-file DataFrames into a single MASTER DataFrame.

    The *nome_file* column is already present in each DataFrame.
    """
    if not dfs:
        return pd.DataFrame(columns=COLUMNS)
    master = pd.concat(dfs.values(), ignore_index=True)
    return master[COLUMNS]
