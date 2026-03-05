---
title: PO Extract
emoji: 📄
colorFrom: blue
colorTo: indigo
sdk: streamlit
sdk_version: 1.32.2
app_file: app.py
pinned: false
---

# PO Extract

Un'applicazione per l'estrazione intelligente dei dati dai file PDF (ordini di acquisto) verso formato Excel.

## Come installare su Hugging Face Spaces

1. Crea un nuovo **Space** su Hugging Face: [https://huggingface.co/spaces](https://huggingface.co/spaces)
2. Seleziona **Streamlit** come SDK.
3. Carica i file di questo progetto nello Space:
    - `app.py`
    - `extractor.py`
    - `utils.py`
    - `requirements.txt`
4. L'applicazione verrà automaticamente configurata e avviata!

## Funzionalità principali

- **Interfaccia Web Moderna**: UI interamente rinnovata con Streamlit.
- **Supporto multi-motore**: Estrae dati usando `pdfplumber+regex`, `pdfplumber` base o `camelot`.
- **Limiti di elaborazione**: Protegge dalle limitazioni di memoria (50 PDF massimi alla volta), ideale per l'hosting Cloud in contenitori di base.
- **Privacy First**: I dati vengono processati in memoria e non salvati su disco per periodi prolungati. Il file Excel finale viene reso disponibile tramite download-stream.
