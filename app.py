#!/usr/bin/env python3
"""
KPI Extraction Platform -- local web app
==========================================
A browser front-end for the extract_kpis.py / calculate_derived_kpis.py
pipeline. No terminal commands needed after the initial launch: upload a
report (PDF, Excel, or CSV), pick a country, click a button.

This app imports extract_kpis.py and calculate_derived_kpis.py directly and
calls their functions -- it does NOT re-implement the extraction or
calculation logic. There is exactly one copy of that logic (the two .py
files); this file is purely the UI on top of it. If you edit the KPI logic,
edit those two files and both the CLI and this app pick up the change.

Setup
-----
    pip install streamlit google-genai pydantic openpyxl pandas

Run
---
    streamlit run app.py

This opens a browser tab at http://localhost:8501. Leave the terminal
window open while you use the app -- closing it stops the server. To stop
manually, click back into that terminal and press Ctrl+C.

Folder layout
-------------
Must sit in the same folder as: extract_kpis.py, calculate_derived_kpis.py,
kpi_reference.json, country_config.csv -- exactly the same requirement as
running the CLI scripts directly.
"""

import os
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

import extract_kpis
import calculate_derived_kpis

st.set_page_config(page_title="KPI Extraction Platform", page_icon="📊", layout="wide")

DEFAULT_DB_PATH = SCRIPT_DIR / "KPI_MASTER_DB.xlsx"


# ---------------------------------------------------------------------------
# Sidebar -- settings shared across tabs
# ---------------------------------------------------------------------------
st.sidebar.header("Settings")

api_key_input = st.sidebar.text_input(
    "Gemini API key",
    value=os.environ.get("GEMINI_API_KEY", ""),
    type="password",
    help="Free key: https://aistudio.google.com/apikey -- kept only in this "
         "browser session's memory, never written to disk.",
)
if api_key_input:
    os.environ["GEMINI_API_KEY"] = api_key_input

db_path_input = st.sidebar.text_input(
    "Master database file",
    value=str(DEFAULT_DB_PATH),
    help="All extractions and derived-KPI runs read/write this one .xlsx file.",
)
db_path = Path(db_path_input)

st.sidebar.caption(f"Model: `{extract_kpis.MODEL}`")
if db_path.exists():
    st.sidebar.caption(f"✅ Database exists ({db_path.stat().st_size // 1024} KB)")
else:
    st.sidebar.caption("⚠️ Database doesn't exist yet -- extract a report to create it.")


st.title("📊 Telecom KPI Extraction Platform")

tab_extract, tab_derive, tab_view = st.tabs(
    ["📄 Extract from a report", "🧮 Compute Derived KPIs", "📋 View Database"]
)

# ---------------------------------------------------------------------------
# Tab 1: Extract raw KPIs from a report (PDF, Excel, or CSV)
# ---------------------------------------------------------------------------
with tab_extract:
    st.subheader("Extract KPIs from a report")
    st.caption(
        "Sends the report to Gemini, which reports only values explicitly "
        "stated in it -- nothing is calculated at this stage. Accepts a PDF, "
        "an Excel workbook, or a CSV; Excel/CSV are converted to a plain "
        "text table before being sent, since Gemini reads PDFs but not "
        "spreadsheets directly."
    )

    try:
        countries = extract_kpis.list_countries()
    except FileNotFoundError:
        countries = []
        st.error("country_config.csv not found next to app.py.")

    uploaded_file = st.file_uploader("Report file", type=["pdf", "xlsx", "xls", "csv"])

    sheet_name = None
    if uploaded_file and Path(uploaded_file.name).suffix.lower() in (".xlsx", ".xls"):
        try:
            sheets = extract_kpis.list_excel_sheets(uploaded_file.getvalue())
        except Exception as e:
            sheets = []
            st.error(f"Could not read sheet names from this workbook: {e}")

        if len(sheets) == 1:
            sheet_name = sheets[0]
        elif sheets:
            sheet_name = st.selectbox(
                "Sheet to extract from",
                sheets,
                index=None,
                placeholder="Pick the sheet that actually holds the KPI data...",
                help="This workbook has multiple tabs -- some may be notes or "
                     "unrelated data rather than KPI tables. Sending the whole "
                     "workbook would dilute what the model is looking for, so "
                     "pick the one sheet for this extraction.",
            )

    col1, col2 = st.columns(2)
    with col1:
        country = st.selectbox("Country", countries) if countries else st.text_input("Country")
    with col2:
        default_label = Path(uploaded_file.name).stem.replace("_", " ") if uploaded_file else ""
        source_label = st.text_input(
            "Source label",
            value=default_label,
            help="Short human-readable name for this report, e.g. "
                 "'ARTP Tableau de bord Mobile T2 2025'. Goes into the "
                 "source_id column so you always know where a number came from.",
        )

    needs_sheet = uploaded_file and Path(uploaded_file.name).suffix.lower() in (".xlsx", ".xls")
    extract_disabled = not uploaded_file or (needs_sheet and not sheet_name)
    extract_clicked = st.button("Extract KPIs", type="primary", disabled=extract_disabled)

    if extract_clicked:
        if not os.environ.get("GEMINI_API_KEY"):
            st.error("Enter your Gemini API key in the sidebar first.")
        elif not country:
            st.error("Enter or select a country.")
        else:
            with st.spinner(f"Sending {uploaded_file.name} to {extract_kpis.MODEL}..."):
                try:
                    file_bytes = uploaded_file.getvalue()
                    result = extract_kpis.extract_from_source(
                        file_bytes, uploaded_file.name, country, source_label, db_path,
                        sheet_name=sheet_name,
                    )
                except Exception as e:
                    st.error(f"Extraction failed: {e}")
                    result = None

            if result is not None:
                st.success(f"Extracted {len(result['rows'])} KPI values → {db_path.name}")

                if result["rows"]:
                    st.dataframe(pd.DataFrame(result["rows"]), use_container_width=True, hide_index=True)

                    review_rows = [r for r in result["rows"] if r["ob_confidence"] == "differs from expected — review"]
                    if review_rows:
                        st.warning(
                            f"{len(review_rows)} row(s) have an Observation Basis that "
                            f"contradicts the usual assumption for that KPI, based on this "
                            f"report's own wording -- worth a look: "
                            f"{[r['attribute_id'] for r in review_rows]}"
                        )
                else:
                    st.info("No KPIs from this catalogue were found explicitly stated in this report.")

                if result["skipped_ids"]:
                    st.warning(
                        f"{len(result['skipped_ids'])} returned attribute_id(s) not in "
                        f"kpi_reference.json (dropped): {result['skipped_ids']}"
                    )
                if result["not_found_notes"]:
                    st.info(f"Model notes: {result['not_found_notes']}")

                st.caption(f"{result['found_count']}/{result['total_count']} KPIs found in this report.")
                with st.expander("KPIs not found in this report"):
                    st.write(", ".join(result["missing_ids"]) or "None -- everything was found.")


# ---------------------------------------------------------------------------
# Tab 2: Compute derived KPIs (pure Python arithmetic, no model)
# ---------------------------------------------------------------------------
with tab_derive:
    st.subheader("Compute derived KPIs")
    st.caption(
        "Reads the raw values already in the master database and computes "
        "everything in kpi_reference.json marked data_origin: calculated "
        "(ARPU, market share, growth rates, HHI, ...) -- plain Python "
        "arithmetic, no model involved, so every number traces back to a "
        "formula you can check."
    )

    derive_clicked = st.button("Compute Derived KPIs", type="primary")

    if derive_clicked:
        if not db_path.exists():
            st.error(f"{db_path} doesn't exist yet -- extract at least one report first.")
        else:
            with st.spinner("Computing..."):
                try:
                    new_rows, skipped_dupes, unresolved = calculate_derived_kpis.run_derived_calculations(db_path)
                except Exception as e:
                    st.error(f"Calculation failed: {e}")
                    new_rows = None

            if new_rows is not None:
                st.success(f"Computed {len(new_rows)} new derived KPI values → {db_path.name}")
                if skipped_dupes:
                    st.caption(
                        f"{skipped_dupes} matched an already-present explicit value "
                        f"and were not duplicated."
                    )
                if new_rows:
                    st.dataframe(pd.DataFrame(new_rows), use_container_width=True, hide_index=True)

                seen = set()
                unresolved_rows = []
                for metric, formula, reason in unresolved:
                    if metric in seen:
                        continue
                    seen.add(metric)
                    unresolved_rows.append({"metric": metric, "formula": formula, "reason": reason})

                if unresolved_rows:
                    st.write(f"**Could not auto-compute ({len(unresolved_rows)} formulas):**")
                    st.dataframe(pd.DataFrame(unresolved_rows), use_container_width=True, hide_index=True)
                    st.caption(
                        "These need either a fix to the ALIASES table in "
                        "calculate_derived_kpis.py, a source that states the value "
                        "directly (so extraction can pick it up), or a decision that "
                        "they stay human-calculated for now."
                    )


# ---------------------------------------------------------------------------
# Tab 3: View / filter / download the master database
# ---------------------------------------------------------------------------
with tab_view:
    st.subheader("Master KPI database")

    if not db_path.exists():
        st.info("No database yet -- extract at least one report first.")
    else:
        df = pd.read_excel(db_path)

        c1, c2, c3 = st.columns(3)
        with c1:
            country_filter = st.multiselect(
                "Filter by country", sorted(df["Country_name"].dropna().unique())
            )
        with c2:
            module_filter = st.multiselect(
                "Filter by module", sorted(df["Module"].dropna().unique())
            )
        with c3:
            confidence_filter = st.multiselect(
                "Filter by Observation Basis confidence",
                sorted(df["ob_confidence"].dropna().unique()) if "ob_confidence" in df.columns else [],
                help="'differs from expected — review' and 'assumed' are the ones "
                     "worth a human look -- 'confirmed', 'formula-defined', and "
                     "'derived (calculated)' don't need one.",
            )

        filtered = df
        if country_filter:
            filtered = filtered[filtered["Country_name"].isin(country_filter)]
        if module_filter:
            filtered = filtered[filtered["Module"].isin(module_filter)]
        if confidence_filter:
            filtered = filtered[filtered["ob_confidence"].isin(confidence_filter)]

        st.dataframe(filtered, use_container_width=True, hide_index=True)
        st.caption(f"{len(filtered)} of {len(df)} rows shown.")

        with open(db_path, "rb") as f:
            st.download_button(
                f"Download {db_path.name}",
                data=f.read(),
                file_name=db_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )