# QC Certificate Extraction Tool

A Streamlit web app that extracts data from Mozambique LNG / TOTAL E&P / Yanda-CCSJV QC certificate PDFs and exports it to a formatted Excel register.

---

## Supported Form Types

| Prefix | Form | Description |
|--------|------|-------------|
| **GAL** | FORM-GAL-203 | Surface Preparation & Galvanizing |
| **PIP** | FORM-PIP-101 | Prefabricated Piping Spool Release |
| **PAI** | FORM-PAI-110 | DFT Check (Painting Inspection) |

---

## Features

- Upload multiple PDFs at once
- **Dual extraction engine**:
  - **Primary** — Gemini 2.0 Flash Vision (AI, requires API key)
  - **Fallback** — Tesseract OCR pipeline (local, no API key needed)
- Extracts per-row item tags plus certificate metadata (Report No., NOI No., dates, inspector names)
- Preview & edit results in an interactive table before downloading
- Colour-coded rows: green (complete) / amber (partial) / red (failed)
- One-click Excel download with timestamped filename

---

## Prerequisites

### Python
Python 3.10 or higher.

### Tesseract OCR (required for the OCR fallback)
Download and install from the [Tesseract releases page](https://github.com/UB-Mannheim/tesseract/wiki).

Default expected path (edit `extractor.py` if different):
```
C:\Program Files\Tesseract-OCR\tesseract.exe
```

### Poppler (required for PDF-to-image conversion)
Download from the [poppler-windows releases](https://github.com/oschwartz10612/poppler-windows/releases).

Default expected path (edit `extractor.py` if different):
```
C:\Users\<you>\Downloads\Release-26.02.0-0\poppler-26.02.0\Library\bin
```

---

## Installation

```bash
git clone https://github.com/<your-username>/qc-certificate-extractor.git
cd qc-certificate-extractor
pip install -r requirements.txt
```

---

## Configuration

Copy `.env.example` to `.env` and optionally add your Gemini API key:

```bash
copy .env.example .env
```

```env
# .env
GEMINI_API_KEY=your_key_here   # leave blank to use OCR-only mode
```

Get a free key at [aistudio.google.com](https://aistudio.google.com/app/apikey).  
If the key is not set, the app runs in OCR-only mode — all features still work.

---

## Running the App

```bash
streamlit run app.py
```

The app opens at `http://localhost:8501`.

---

## Usage

1. **Upload** — drag and drop one or more GAL / PIP / PAI certificate PDFs
2. **Process PDFs** — click the button; a progress bar shows per-file status
3. **Preview & Edit** — review the extracted table, edit any cell, delete wrong rows
4. **Download** — export as `QC_Register_YYYYMMDD_HHMMSS.xlsx`

---

## Output Columns

| Column | Description |
|--------|-------------|
| Item | Spool / tag identifier (one row per item) |
| RFI | Inspection Notification Reference (NOI No.) |
| Certificate_Code | Report number |
| Issue_date | Certificate issue date (DD-MMM-YY) |
| Notes | Free-text notes (editable) |
| Quality Signature Name | CCSJV inspector name |
| Quality Signature Date | CCSJV sign-off date |
| Client Signature Name | Client / TPI inspector name |
| Client Signature Date | Client sign-off date |

---

## Project Structure

```
├── app.py            # Streamlit UI
├── extractor.py      # PDF extraction logic (Gemini + OCR pipeline)
├── excel_writer.py   # Excel output formatting
├── requirements.txt
├── .env.example
└── PDF_Folder/       # Temporary storage for uploaded PDFs
```

---

## Extraction Accuracy

| Mode | Typical accuracy |
|------|-----------------|
| Gemini Vision (API key set) | ~99% — recommended for production |
| Tesseract OCR fallback | ~85–95% depending on scan quality |

For best results on scanned PDFs, ensure the scan is ≥ 300 DPI and that the form type prefix (GAL / PIP / PAI) is in the filename.

---

## License

MIT
