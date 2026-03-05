"""
app.py — Streamlit web application for PO_Extract.

Provides a modern graphical interface to:
  - Load one or more purchase-order PDF files (max 50 per batch)
  - Choose an extraction engine
  - Extract data and download an Excel workbook
  - Preview extracted rows
  - View real-time logs and progress
"""

import io
import logging
from pathlib import Path
import time

import pandas as pd
import streamlit as st

from extractor import COLUMNS, clean_dataframe, compose_master, extract
from utils import sanitize_sheet_name, unique_sheet_name

# ── App Configuration ───────────────────────────────────────────────────────
st.set_page_config(
    page_title="PO Extract - Purchase Order Data Extractor",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Constants
ENGINES = ["pdfplumber+regex", "pdfplumber", "camelot"]
MAX_PREVIEW_ROWS = 20
MAX_UPLOAD_FILES = 50

# Custom CSS for modern UI
st.markdown("""
<style>
    .stApp {
        background-color: #0e1117;
    }
    .main-header {
        font-size: 2.5rem;
        font-weight: 700;
        color: #4da6ff;
        margin-bottom: 0px;
    }
    .sub-header {
        font-size: 1.1rem;
        color: #90a4ae;
        margin-bottom: 2rem;
    }
    .metric-value {
        font-size: 1.8rem;
        font-weight: bold;
        color: #e0e0e0;
    }
    .metric-label {
        font-size: 0.9rem;
        color: #90a4ae;
    }
    .log-container {
        font-family: 'Consolas', 'Courier New', monospace;
        background-color: #1e1e1e;
        color: #d4d4d4;
        padding: 1rem;
        border-radius: 5px;
        height: 300px;
        overflow-y: auto;
        font-size: 0.85rem;
    }
    .stButton>button {
        width: 100%;
        border-radius: 5px;
        font-weight: bold;
        transition: all 0.3s;
    }
    div[data-testid="stFileUploader"] {
        border: 2px dashed #4da6ff;
        border-radius: 10px;
        padding: 2rem;
        text-align: center;
        transition: all 0.3s ease;
    }
    div[data-testid="stFileUploader"]:hover {
        border-color: #82b1ff;
        background-color: rgba(77, 166, 255, 0.05);
    }
</style>
""", unsafe_allow_html=True)


# ── Custom Logging Handler for Streamlit ────────────────────────────────────

class StreamlitLogHandler(logging.Handler):
    """Log handler that writes to a Streamlit list in session state."""
    def __init__(self):
        super().__init__()
        if "log_messages" not in st.session_state:
            st.session_state.log_messages = []
            
    def emit(self, record):
        msg = self.format(record)
        st.session_state.log_messages.append(msg)

def setup_st_logging():
    """Configure logger for Streamlit."""
    logger = logging.getLogger("po_extract")
    logger.setLevel(logging.INFO)
    
    # Only add handler if not already present
    if not any(isinstance(h, StreamlitLogHandler) for h in logger.handlers):
        logger.handlers.clear()
        
        fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S")
        
        # Console handler
        ch = logging.StreamHandler()
        ch.setFormatter(fmt)
        logger.addHandler(ch)
        
        # Streamlit handler
        sh = StreamlitLogHandler()
        sh.setFormatter(fmt)
        logger.addHandler(sh)
        
    return logger


# ── Helpers ─────────────────────────────────────────────────────────────────

def human_size(nbytes: int) -> str:
    """Return a human-readable file size."""
    for unit in ("B", "KB", "MB"):
        if nbytes < 1024:
            return f"{nbytes:.0f} {unit}"
        nbytes /= 1024
    return f"{nbytes:.1f} GB"

def reset_state():
    """Reset session state variables between runs."""
    st.session_state.log_messages = []
    st.session_state.extraction_done = False
    st.session_state.excel_buffer = None
    st.session_state.master_dfs = {}
    st.session_state.edited_master = None


# ── Main App ────────────────────────────────────────────────────────────────

def main():
    logger = setup_st_logging()
    
    # Initialize session state
    if "extraction_done" not in st.session_state:
        st.session_state.extraction_done = False
    if "excel_buffer" not in st.session_state:
        st.session_state.excel_buffer = None
    if "preview_df" not in st.session_state:
        st.session_state.preview_df = None
    if "master_dfs" not in st.session_state:
        st.session_state.master_dfs = {}
    if "edited_master" not in st.session_state:
        st.session_state.edited_master = None
        
    # Header
    st.markdown('<p class="main-header">PO Extract</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Estrai dati da ordini di acquisto PDF → Excel in modo intelligente</p>', unsafe_allow_html=True)
    
    # Sidebar: Options & Config
    with st.sidebar:
        st.header("⚙️ Opzioni")
        st.markdown("Configura le preferenze di estrazione.")
        
        engine = st.selectbox(
            "Motore di Estrazione",
            options=ENGINES,
            index=0,
            help="Scegli l'algoritmo per estrarre il testo. 'pdfplumber+regex' è ottimizzato per ordini QVC."
        )
        
        force_ocr = st.checkbox("Forza OCR (sperimentale)", value=False, help="Forza la lettura ottica se il PDF è un'immagine scansionata.")
        
        st.divider()
        st.markdown("### ℹ️ Informazioni")
        st.info(f"**Limite batch:** Max {MAX_UPLOAD_FILES} file per volta per salvaguardare la memoria.")
        st.caption("v2.0.0 - Powered by Streamlit")

    # Layout: Two columns for upload and status
    col1, col2 = st.columns([6, 4])
    
    with col1:
        st.subheader("📂 Carica file PDF")
        uploaded_files = st.file_uploader(
            "Trascina qui i tuoi PDF o clicca per sfogliare",
            type="pdf",
            accept_multiple_files=True,
            on_change=reset_state
        )
        
        if uploaded_files:
            file_count = len(uploaded_files)
            total_size = sum(f.size for f in uploaded_files)
            
            # Display metrics
            mc1, mc2, mc3 = st.columns(3)
            mc1.metric("File caricati", file_count)
            mc2.metric("Dimensione totale", human_size(total_size))
            mc3.metric("Motore", engine.split('+')[0])
            
            # Enforce limits
            if file_count > MAX_UPLOAD_FILES:
                st.error(f"⚠️ Hai caricato {file_count} file. Il limite massimo è di {MAX_UPLOAD_FILES} file per elaborazione. Per favore rimuovi dei file.")
                st.stop()
                
            # Quick preview table of uploaded files
            with st.expander("Mostra dettagli file", expanded=False):
                file_details = [{"Nome File": f.name, "Dimensione": human_size(f.size)} for f in uploaded_files]
                st.dataframe(pd.DataFrame(file_details), use_container_width=True, hide_index=True)

    with col2:
        st.subheader("▶️ Azioni")
        
        # Preview block
        if uploaded_files:
            if st.button("🔍 Genera Anteprima Rapida", help="Estrai le prime righe dal primo file per verificare i dati"):
                with st.spinner("Estrazione anteprima in corso..."):
                    try:
                        first_file = uploaded_files[0]
                        # Temporary save since extract needs a path or file-like object
                        with open(f"/tmp/{first_file.name}", "wb") as f:
                            f.write(first_file.getbuffer())
                            
                        df = extract(f"/tmp/{first_file.name}", engine=engine)
                        df = clean_dataframe(df)
                        st.session_state.preview_df = df.head(MAX_PREVIEW_ROWS)
                        st.toast(f"Anteprima generata con successo per {first_file.name}", icon="✅")
                    except Exception as e:
                        st.error(f"Errore durante l'anteprima: {str(e)}")
                        
        # Extract block
        extract_button = st.button(
            "🚀 Avvia Estrazione Completa", 
            type="primary", 
            disabled=not uploaded_files,
            use_container_width=True
        )

    # Show Preview if available
    if st.session_state.preview_df is not None and not st.session_state.extraction_done:
        st.subheader("👀 Anteprima Dati")
        st.dataframe(st.session_state.preview_df, use_container_width=True)

    # Execution logic
    if extract_button and uploaded_files:
        st.session_state.log_messages = []
        st.session_state.extraction_done = False
        
        st.subheader("⏳ Stato di Elaborazione")
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        log_container = st.empty()
        
        dfs: dict[str, pd.DataFrame] = {}
        total = len(uploaded_files)
        
        for i, up_file in enumerate(uploaded_files, 1):
            pct = int((i - 1) / total * 90)
            progress_bar.progress(pct)
            status_text.text(f"Elaborazione ({i}/{total}): {up_file.name}...")
            
            logger.info("Elaborazione file: %s", up_file.name)
            
            # Fallback to display logs in UI quickly
            log_text = "\\n".join(st.session_state.log_messages[-15:])
            log_container.markdown(f'<div class="log-container">{log_text}</div>', unsafe_allow_html=True)
            
            try:
                # Save uploaded file temporarily for extraction
                temp_path = f"/tmp/{up_file.name}"
                with open(temp_path, "wb") as f:
                    f.write(up_file.getbuffer())
                    
                df = extract(temp_path, engine=engine)
                df = clean_dataframe(df)
                dfs[up_file.name] = df
                
                logger.info("  → %d righe estratte con successo", len(df))
            except Exception as exc:
                logger.error("  ✗ Errore su %s: %s", up_file.name, exc)
                
            # Update log view
            log_text = "\\n".join(st.session_state.log_messages[-20:])
            log_container.markdown(f'<div class="log-container">{log_text}</div>', unsafe_allow_html=True)

        if not dfs:
            status_text.text("Nessun dato estratto dai file.")
            st.error("Nessun dato valido è stato estratto dai file PDF caricati. Controlla i log per i dettagli.")
            progress_bar.progress(100)
            st.stop()

        st.session_state.master_dfs = dfs
        master = compose_master(dfs)
        st.session_state.edited_master = master
        st.session_state.extraction_done = True
        
        progress_bar.progress(100)
        status_text.text("Completato! 🎉")
        logger.info("✅ Estrazione completata con successo!")
        
        # Final log update
        log_text = "\\n".join(st.session_state.log_messages[-20:])
        log_container.markdown(f'<div class="log-container">{log_text}</div>', unsafe_allow_html=True)
        st.success("Tutti i PDF sono stati elaborati correttamente.")

    # After extraction is done, show Editor and Download options
    if st.session_state.extraction_done and st.session_state.edited_master is not None:
        st.divider()
        
        st.subheader("📊 Riepilogo e Controllo Dati")
        st.markdown("Puoi **modificare o correggere** i dati direttamente nella tabella qui sotto prima di scaricare l'Excel. Eventuali errori di OCR possono essere corretti con un doppio clic sulla cella.")
        
        # Calculate summary metrics
        master = st.session_state.edited_master
        total_items = len(master)
        
        # Safe numeric parsing for sum
        quantita_sum = pd.to_numeric(master.get("quantita", pd.Series([0])), errors='coerce').sum()
        valore_sum = pd.to_numeric(master.get("valore_netto", pd.Series([0])), errors='coerce').sum()
        
        met1, met2, met3 = st.columns(3)
        met1.metric("Totale Righe Ordinate", f"{total_items}")
        met2.metric("Quantità Totale", f"{int(quantita_sum)}")
        met3.metric("Valore Netto Totale", f"€ {valore_sum:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        
        # Interactive Editor
        edited_df = st.data_editor(
            master,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            key="master_editor"
        )
        
        # Excel Generation (on the fly to capture edits)
        output = io.BytesIO()
        existing_sheets: set[str] = set()
        
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            # Save the edited MASTER
            edited_df.to_excel(writer, sheet_name="MASTER", index=False)
            existing_sheets.add("MASTER")
            
            # Save individual files (grouped from edited master)
            if "nome_file" in edited_df.columns:
                for fname, group_df in edited_df.groupby("nome_file"):
                    sheet = sanitize_sheet_name(str(fname))
                    sheet = unique_sheet_name(sheet, existing_sheets)
                    existing_sheets.add(sheet)
                    group_df.to_excel(writer, sheet_name=sheet, index=False)
                    
        excel_data = output.getvalue()
        
        st.divider()
        st.subheader("📥 Download Risultati")
        
        dl_col1, dl_col2, dl_col3 = st.columns([1, 2, 1])
        with dl_col2:
            st.download_button(
                label="⬇️ Scarica File Excel Aggiornato (PO_Extract.xlsx)",
                data=excel_data,
                file_name="PO_Extract_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )

if __name__ == "__main__":
    main()
