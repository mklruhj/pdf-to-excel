"""
app.py — QC Certificate Extraction Tool (Main Streamlit Application)
Mozambique LNG · TOTAL E&P · Yanda / CCSJV · Version 1.0
"""
import os
from datetime import datetime

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from extractor import process_pdf
from excel_writer import build_excel, COLUMNS

load_dotenv(override=True)

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="QC Certificate Extraction Tool",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Session state initialisation ───────────────────────────────────────────────
for key, default in [
    ("rows",            []),
    ("file_statuses",   {}),
    ("done",            False),
]:
    if key not in st.session_state:
        st.session_state[key] = default

api_key = os.getenv("GEMINI_API_KEY", "")

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.title("Info")
    st.divider()
    st.caption("**Supported form types**")
    st.caption("• GAL — Surface Prep / Galvanizing")
    st.caption("• PIP — Prefab Piping Spool Release")
    st.caption("• PAI — DFT Check (Painting Inspection)")
    st.divider()
    if api_key:
        st.success("Gemini Vision active", icon="🤖")
        st.caption("Primary: Gemini 2.0 Flash Vision")
        st.caption("Fallback: Tesseract OCR")
    else:
        st.warning("Gemini key not set", icon="⚠️")
        st.caption("Set GEMINI_API_KEY in .env for")
        st.caption("maximum extraction accuracy.")
        st.caption("Fallback: Tesseract OCR only")
    st.divider()
    st.caption("v1.1 · June 2026")

# ══════════════════════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════════════════════
st.title("QC Certificate Extraction Tool")
st.caption("Mozambique LNG · TOTAL E&P · Yanda / CCSJV")

st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# SCREEN 1 — UPLOAD
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("1 — Upload PDF Certificates")

uploaded_files = st.file_uploader(
    "Drag and drop GAL / PIP / PAI certificates here (multiple files accepted)",
    type=["pdf"],
    accept_multiple_files=True,
)

if uploaded_files:
    st.info(f"{len(uploaded_files)} file(s) loaded and ready to process.")

process_btn = st.button(
    "Process PDFs",
    type="primary",
    disabled=not uploaded_files,
)

# ══════════════════════════════════════════════════════════════════════════════
# SCREEN 2 — PROCESSING
# ══════════════════════════════════════════════════════════════════════════════
if process_btn and uploaded_files:
    st.divider()
    st.subheader("2 — Processing")

    os.makedirs("PDF_Folder", exist_ok=True)
    all_rows, file_statuses = [], {}
    n = len(uploaded_files)

    progress_bar = st.progress(0, text="Starting…")
    log = st.empty()
    log_lines = []

    for idx, uf in enumerate(uploaded_files):
        progress_bar.progress(idx / n, text=f"Processing {uf.name} ({idx + 1}/{n})…")

        pdf_path = os.path.join("PDF_Folder", uf.name)
        with open(pdf_path, "wb") as f:
            f.write(uf.getbuffer())

        rows, status, error = process_pdf(pdf_path, api_key)
        all_rows.extend(rows)
        file_statuses[uf.name] = (status, error)

        icon = "✅" if status == "complete" else ("⚠️" if status == "partial" else "❌")
        msg = (
            f"{icon} **{uf.name}** — {len(rows)} row(s) · {status}"
            if not error
            else f"{icon} **{uf.name}** — {error}"
        )
        log_lines.append(msg)
        log.markdown("\n\n".join(log_lines))

    progress_bar.progress(1.0, text="Done!")

    st.session_state.rows          = all_rows
    st.session_state.file_statuses = file_statuses
    st.session_state.done          = True

# ══════════════════════════════════════════════════════════════════════════════
# SCREEN 3 — PREVIEW TABLE
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.done and st.session_state.rows:
    st.divider()
    st.subheader("3 — Preview & Edit")

    rows = st.session_state.rows
    total    = len(rows)
    complete = sum(1 for r in rows if r.get("_status") == "complete")
    partial  = sum(1 for r in rows if r.get("_status") == "partial")
    failed   = sum(1 for r in rows if r.get("_status") == "failed")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Rows",   total)
    m2.metric("Complete ✅",  complete)
    m3.metric("Partial ⚠️",   partial)
    m4.metric("Failed ❌",    failed)

    st.caption(
        "🟢 Green = all fields complete  "
        "🟡 Amber = missing optional fields  "
        "🔴 Red = missing required fields"
    )
    st.caption("You can edit any cell or delete rows using the table below before downloading.")

    # Build display dataframe (hide internal _status column)
    df = pd.DataFrame([{c: r.get(c, "") for c in COLUMNS} for r in rows])

    # Colour-code rows via pandas Styler
    status_list = [r.get("_status", "complete") for r in rows]

    def _colour(row):
        s = status_list[row.name] if row.name < len(status_list) else "complete"
        bg = (
            "background-color: #C6EFCE" if s == "complete" else
            "background-color: #FFEB9C" if s == "partial"  else
            "background-color: #FFC7CE"
        )
        return [bg] * len(row)

    styled_df = df.style.apply(_colour, axis=1)

    # Editable table (allows row deletion and cell edits)
    edited_df = st.data_editor(
        df,
        use_container_width=True,
        num_rows="dynamic",
        key="editor",
        hide_index=False,
    )

    # Persist edits back to session state
    if edited_df is not None:
        updated_rows = []
        for i, (_, row_data) in enumerate(edited_df.iterrows()):
            r = dict(row_data)
            r["_status"] = status_list[i] if i < len(status_list) else "complete"
            updated_rows.append(r)
        st.session_state.rows = updated_rows

# ══════════════════════════════════════════════════════════════════════════════
# SCREEN 4 — DOWNLOAD
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.done and st.session_state.rows:
    st.divider()
    st.subheader("4 — Download")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename  = f"QC_Register_{timestamp}.xlsx"

    excel_buf = build_excel(st.session_state.rows)

    col_dl, col_reset = st.columns([2, 1])
    with col_dl:
        st.download_button(
            label="⬇ Download Excel",
            data=excel_buf,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        st.caption(f"Filename: `{filename}`")

    with col_reset:
        if st.button("Process New Files", use_container_width=True):
            st.session_state.rows          = []
            st.session_state.file_statuses = {}
            st.session_state.done          = False
            st.rerun()
