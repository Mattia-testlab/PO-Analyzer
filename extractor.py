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
    "tipo_ordine",
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

# Two-step approach for item lines:
# Step 1: Match the overall structure:  6-digit + middle + date + qty + price + value
# Step 2: Post-process 'middle' to separate QVC code from description.
#
# Examples of item lines:
#   154957 ONJ W40 Motivi PANTALONE PALAZZO CON CINTUR 02.03.26 20 34,16 683,20
#   154890 AJG W42 oltre TRENCH CORTO IN FINTA PELLE 02.03.26 30 52,00 1.560,00
#   154951 QPW W40Motivi GIUBBINO IN ECOPELLE 02.03.26 10 35,00 350,00
#   154887 AAA XLoltre T-shirt bimaterica 02.03.26 29 12,50 362,50
#   154892 AAP XLOltre Cardigan lungo 02.03.26 9 37,96 341,64

RE_ITEM_LINE = re.compile(
    r"^(\d{6})"                      # 6-digit article code
    r"\s+"
    r"(.+?)"                          # middle: variant + size + description
    r"\s+"
    r"(\d{2}\.\d{2}\.\d{2})"         # date DD.MM.YY
    r"\s+"
    r"(\d[\d.]*)"                     # quantity
    r"\s+"
    r"([\d.,]+)"                      # unit price
    r"\s+"
    r"([\d.,]+)"                      # net value
    r"\s*$"
)

# Known size tokens (order matters: longer patterns first)
_SIZE_TOKENS = [
    "NOSIZE", "OVR", "OVQ", "XL", "XS", "XXL", "XXS",
    "W36", "W38", "W40", "W42", "W44", "W46", "W48", "W50", "W52",
    "S", "M", "L", "X",
]
# Known brand prefixes that mark the start of the description when glued
_BRAND_PREFIXES = ["Motivi", "Oltre", "oltre", "olte", "Fiorella", "fiorella"]

# Regex to split: variant (2-3 alpha) + size (possibly glued to description)
_RE_VARIANT_SIZE = re.compile(
    r"^([A-Z]{2,4})"                # variant code (e.g., ONJ, AJG, QPW, AAA)
    r"\s+"
    r"(\S+)"                         # size token (may be glued to description)
    r"(?:\s+(.*))?$"                 # rest = description (may be absent if glued)
)


def _split_qvc_description(article_code: str, middle: str) -> tuple[str, str, str]:
    """Split the middle part of an item line into (qvc, taglia, descrizione).

    Returns (qvc_code, taglia, descrizione).
    The qvc_code includes the 6-digit article code, variant, and size.
    """
    m = _RE_VARIANT_SIZE.match(middle)
    if not m:
        # Fallback: return article_code as QVC, no taglia, rest as description
        return article_code, "", middle.strip()

    variant = m.group(1)          # e.g., "ONJ", "AJG"
    size_raw = m.group(2)         # e.g., "W40", "W40Motivi", "XLoltre", "QKU"
    rest = (m.group(3) or "").strip()  # description after space (may be empty)

    # Check if size_raw contains a glued brand/description
    taglia = size_raw
    glued_desc = ""

    # First check: is size_raw an exact known token? (e.g., "XS", "W40", "M")
    if size_raw in _SIZE_TOKENS or re.match(r'^[A-Z]\d+$', size_raw):
        taglia = size_raw
    else:
        # Try splitting known size tokens from the start of size_raw
        for st in _SIZE_TOKENS:
            if size_raw.startswith(st) and len(size_raw) > len(st):
                taglia = st
                glued_desc = size_raw[len(st):]
                break
        else:
            # No known size prefix found
            if re.match(r'^(W\d+|QKU)', size_raw):
                # Size pattern like W40, QKU followed by text
                sm = re.match(r'^(W\d+|QKU)(.*)', size_raw)
                if sm:
                    taglia = sm.group(1)
                    glued_desc = sm.group(2)

    # Build description from glued part + rest
    if glued_desc and rest:
        descrizione = glued_desc + " " + rest
    elif glued_desc:
        descrizione = glued_desc
    else:
        descrizione = rest

    descrizione = descrizione.strip()
    
    # Clean up trailing boilerplate. Some PDF layouts glue the next item's article_code prefix 
    # to the end of the current description. e.g. "TOP LINGERIE naturale XL 154908:"
    trailing_boiler = f"{article_code}:"
    if descrizione.endswith(trailing_boiler):
        descrizione = descrizione[:-len(trailing_boiler)].strip()

    qvc = f"{article_code} {variant} {taglia}"
    return qvc, taglia, descrizione


RE_COD_ART_FORN = re.compile(r"Cod\.art\.forn\.\s*:\s*(\S+)", re.IGNORECASE)
RE_PAESE = re.compile(r"Paese di origine\s*:\s*(.+)", re.IGNORECASE)
RE_PO_HEADER = re.compile(r"^\s*(\d{10})\s*$")
RE_TIPO_ORDINE = re.compile(r"Tipo Ordine:\s*(.+?)\s+Termini", re.IGNORECASE)


# ── Engine: pdfplumber + regex ──────────────────────────────────────────────

def extract_with_pdfplumber_regex(pdf_path: str | Path) -> pd.DataFrame:
    """Extract purchase-order line items using pdfplumber text + regex."""
    if pdfplumber is None:
        raise ImportError("pdfplumber is not installed")

    pdf_path = Path(pdf_path)
    filename = pdf_path.name
    po_number = extract_po_from_filename(filename)
    tipo_ordine = ""

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

            # Extract Tipo Ordine from first page only
            if page_num == 1 and not tipo_ordine:
                for line in lines:
                    tm = RE_TIPO_ORDINE.search(line)
                    if tm:
                        tipo_ordine = tm.group(1).strip()
                        break

            # Parse item blocks
            i = 0
            while i < len(lines):
                line = lines[i].strip()
                m = RE_ITEM_LINE.match(line)
                if m:
                    article_code = m.group(1).strip()
                    middle = m.group(2).strip()
                    data_consegna = m.group(3).strip()
                    quantita = m.group(4).strip()
                    prezzo_unitario = m.group(5).strip()
                    valore_netto = m.group(6).strip()

                    # Split middle into QVC code, taglia, and description
                    qvc, taglia, descrizione = _split_qvc_description(
                        article_code, middle
                    )

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
                        "tipo_ordine": tipo_ordine,
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
                "tipo_ordine": "",
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
                            "tipo_ordine": "",
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
