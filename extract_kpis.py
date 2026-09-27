"""
KPI Extraction Agent
=====================
Feeds a telecom regulator/operator report to Gemini and extracts the KPI
values defined in kpi_reference.json, writing them into an Excel file with
the same column structure as the "KPI MASTER DB" tab.

Accepted source files: PDF, Excel (.xlsx), or CSV. A PDF is sent to Gemini
as a document, exactly as before. Excel/CSV sources don't have a Gemini
document type, so they're converted to a plain text table first and sent as
text -- the model still does all the interpretive work (messy headers,
inconsistent layouts across countries), it just reads a table instead of a
scanned page. An Excel workbook can contain several sheets, so callers pick
one sheet per extraction (see list_excel_sheets / app.py's sheet picker) --
sending an entire multi-sheet workbook as one blob would dilute the model's
attention the same way an oversized PDF does.

Setup
-----
    pip install google-genai pydantic openpyxl pandas

    export GEMINI_API_KEY="your-key-here"

Get a free key at https://aistudio.google.com/apikey -- no credit card
required. The free tier has rate limits (requests per minute/day), so it's
fine for processing a handful of reports at a time, not thousands per day.

Usage
-----
    python extract_kpis.py \
        --file "ARTP_Tableau_Bord_Mobile_T2_2025.pdf" \
        --country "Senegal" \
        --source-label "ARTP Tableau de bord Mobile T2 2025" \
        --output "KPI_MASTER_DB_new.xlsx"

    python extract_kpis.py \
        --file "regulator_data.xlsx" --sheet "Ivory coast" \
        --country "Ivory Coast" --source-label "Regulator workbook 2025" \
        --output "KPI_MASTER_DB_new.xlsx"

Run it once per report. Extracted rows are appended to --output, so you can
point every report at the same file and build up your database.

Design notes
------------
- The model is ONLY asked to find values that are explicitly stated in the
  document for a specific period (year/quarter/month) and, where relevant,
  a specific operator. It is explicitly told NOT to calculate derived KPIs
  itself (e.g. ARPU, penetration rates) -- those stay computed by formula
  in Excel afterwards, exactly like Phase 1's manual process. This avoids
  the model silently inventing numbers.
- Deterministic fields (attribute_name, Module, frequency, data_type,
  Country_name, Region_name) are filled in by this script from
  kpi_reference.json / country_config.csv, NOT by the model -- the model
  only has to report id, value, unit, period, operator, and where it found
  it. This keeps the model's job narrow and makes mistakes easy to spot.
- "Observation Basis" (what kind of number this is -- a snapshot, a period
  total, or a period average) works the same way for most KPIs: it's a
  fixed, formula-defined property of the KPI (kpi_reference.json's
  observation_basis_category, via kpi_shared.py), not something the model
  decides. For the subset of KPIs where the real-world basis is known to
  vary by regulator/report (observation_basis_check_source == true --
  averages, stock/stock ratios like penetration, and regulatory rate
  floors), the model is additionally asked to report what the document's
  own wording implies, and that's compared against the default -- see
  kpi_shared.resolve_observation_basis for the exact rule.
- Extraction is forced through a JSON schema (structured output), not free
  text, so the result is reliably parseable every time.
"""

import argparse
import base64
import csv
import io
import json
import os
import sys
from pathlib import Path
from typing import List, Optional

try:
    from google import genai
except ImportError:
    sys.exit("Missing dependency. Run: pip install google-genai --break-system-packages")

try:
    from pydantic import BaseModel, Field
except ImportError:
    sys.exit("Missing dependency. Run: pip install pydantic --break-system-packages")

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font
except ImportError:
    sys.exit("Missing dependency. Run: pip install openpyxl --break-system-packages")

try:
    import pandas as pd
except ImportError:
    sys.exit("Missing dependency. Run: pip install pandas --break-system-packages")

from kpi_shared import OUTPUT_COLUMNS, resolve_observation_basis, ensure_headers

SCRIPT_DIR = Path(__file__).resolve().parent
KPI_REFERENCE_PATH = SCRIPT_DIR / "kpi_reference.json"
COUNTRY_CONFIG_PATH = SCRIPT_DIR / "country_config.csv"

# Google retires/renames Gemini models fairly often -- if this 404s with a
# "model no longer available" error, check https://ai.google.dev/gemini-api/docs/models
# for the current Flash model name and swap it in here.
MODEL = "gemini-3.6-flash"

SUPPORTED_EXTENSIONS = {".pdf", ".xlsx", ".xls", ".csv"}


# ---------------------------------------------------------------------------
# Structured output schema
# ---------------------------------------------------------------------------
# Gemini is asked to return JSON matching this schema directly (via
# response_format below). Same idea as a forced tool-call: a reliably
# parseable, narrowly-scoped result instead of free text.

class ExtractedKPI(BaseModel):
    attribute_id: str = Field(description="The attribute_id from the KPI catalogue, e.g. 'A005'.")
    value: float
    unit: str = Field(description="The unit as stated in the report.")
    year: int
    quarter: Optional[str] = Field(default=None, description="e.g. 'Q1', or null if not quarterly.")
    month: Optional[int] = Field(default=None, description="1-12, or null if not monthly.")
    operator: str = Field(description="'ALL' for country-level KPIs or market totals, otherwise the specific operator name.")
    source_snippet: str = Field(description="A few words of surrounding text or the table/section name, for verification.")
    observation_basis_hint: Optional[str] = Field(
        default=None,
        description=(
            "ONLY fill this in for an attribute_id marked 'verify_basis: yes' in the "
            "catalogue below, and only if the document's own wording makes it clear. "
            "One of: 'point_in_time' (a snapshot as of a specific date, e.g. "
            "\"subscribers at end of March\"), 'period_total' (summed over the period, "
            "e.g. \"total Q1 revenue\"), or 'period_average' (an average over the "
            "period, e.g. \"average monthly ARPU\"). Leave null if the wording doesn't "
            "make this clear, or if the attribute_id isn't marked 'verify_basis: yes'."
        ),
    )


class ExtractionResult(BaseModel):
    extracted: List[ExtractedKPI] = Field(default_factory=list)
    not_found_notes: Optional[str] = Field(
        default=None,
        description="Optional: brief note on KPIs you expected to find but couldn't, or ambiguities.",
    )


def load_kpi_reference():
    with open(KPI_REFERENCE_PATH, encoding="utf-8") as f:
        return json.load(f)


def load_country_config(country):
    with open(COUNTRY_CONFIG_PATH, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            if row["country"].strip().lower() == country.strip().lower():
                return row
    available = [row["country"] for row in csv.DictReader(open(COUNTRY_CONFIG_PATH, encoding="utf-8"))]
    raise ValueError(
        f"Country '{country}' not found in country_config.csv.\n"
        f"Available: {', '.join(available)}\n"
        f"Add a new row to country_config.csv to support a new country -- "
        f"that's the whole point of the config table."
    )


def list_countries():
    """Country names from country_config.csv, for populating a dropdown."""
    with open(COUNTRY_CONFIG_PATH, encoding="utf-8") as f:
        return [row["country"] for row in csv.DictReader(f)]


def list_excel_sheets(file_bytes):
    """Sheet names in an uploaded .xlsx, for app.py's sheet picker. A
    workbook can mix real data tables with notes/todo tabs (as yours does),
    so the caller -- a human -- picks which sheet(s) are actually the data
    source for this extraction."""
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True)
    try:
        return wb.sheetnames
    finally:
        wb.close()


def _read_excel_as_text(file_bytes, sheet_name):
    if not sheet_name:
        raise ValueError(
            "This Excel file has more than one sheet -- pick which sheet to extract "
            "from (see list_excel_sheets)."
        )
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)
    return df.to_csv(index=False, header=False)


def _read_csv_as_text(file_bytes):
    # Regulator CSVs are often exported from Excel with French text (accents
    # like é/à/ç), which sometimes isn't UTF-8. Try UTF-8 first, then the two
    # encodings "Excel CSV" commonly uses, rather than silently mangling
    # characters.
    for encoding in ("utf-8", "cp1252", "latin-1"):
        try:
            return file_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError("Could not decode this CSV with utf-8, cp1252, or latin-1.")


def source_kind_from_filename(filename):
    ext = Path(filename).suffix.lower()
    if ext == ".pdf":
        return "pdf"
    if ext in (".xlsx", ".xls"):
        return "excel"
    if ext == ".csv":
        return "csv"
    raise ValueError(
        f"Unsupported file type '{ext}'. Accepted: PDF (.pdf), Excel (.xlsx), CSV (.csv)."
    )


def build_kpi_list_text(kpi_ref, operators):
    """Human-readable KPI catalogue for the prompt."""
    lines = []
    for k in kpi_ref:
        scope = []
        if k["country_level"]:
            scope.append("country-level")
        if k["operator_level"]:
            scope.append("per-operator")
        scope_str = " / ".join(scope) if scope else "n/a"
        verify_flag = "yes" if k.get("observation_basis_check_source") else "no"
        lines.append(
            f"- {k['attribute_id']} | {k['metric_name']} | {k['name_en']} "
            f"| unit: {k['unit']} | frequency: {k['frequency']} "
            f"| origin: {k['data_origin']} | scope: {scope_str} "
            f"| verify_basis: {verify_flag}"
        )
    return "\n".join(lines)


def build_prompt(country, regulator, operators, source_label, kpi_ref, source_kind):
    kpi_list_text = build_kpi_list_text(kpi_ref, operators)

    if source_kind == "pdf":
        source_note = "The attached document is the report to extract from."
    else:
        source_note = (
            "The attached text is a table converted from a spreadsheet (Excel or CSV), "
            "not the original file -- ignore artifacts like repeated values from merged "
            "cells or generic column headers (e.g. 'Column1'), and use the surrounding "
            "row/column context to figure out what each number represents."
        )

    return f"""You are extracting telecom KPI values from an attached report for {country}.

Regulator for this country: {regulator}
Operators active in this market: {operators}
Report source label: {source_label}

{source_note}

Below is the catalogue of KPIs we track. For each one, "origin: raw" means the
value should appear directly in the report as a stated number. "origin:
calculated" means it is normally computed afterward from other KPIs using a
formula -- DO NOT calculate these yourself. Only report a "calculated" KPI if
the report states that exact computed number explicitly (e.g. the document
itself prints an "EBITDA margin: 34%" line) -- never derive it from other
numbers you found.

"verify_basis: yes" marks KPIs where you should also try to fill in
observation_basis_hint (see the field description) from the document's own
wording. Leave observation_basis_hint null for every other KPI, and null even
for a "verify_basis: yes" KPI if the wording isn't clear.

KPI catalogue:
{kpi_list_text}

Instructions:
1. Search the attached content for any of the KPIs above.
2. For each value you find, report: the attribute_id, the numeric value, the
   unit as stated in the report (convert only trivial unit prefixes, e.g.
   "000s" to actual count -- do not do unit conversions you're unsure of),
   the year, quarter (e.g. "Q1", or null if not quarterly), month (1-12, or
   null if not monthly), which operator it applies to (use "ALL" for
   country-level KPIs or when the report gives a market total, otherwise the
   specific operator name), a short source_snippet (a few words of
   surrounding text or the table/section name) showing exactly where you
   found it, for verification, and observation_basis_hint where applicable
   (see above).
3. Do not report a KPI you cannot find a specific stated number for. It is
   fine, and expected, to leave most of the catalogue unmatched -- one
   report typically only covers a subset of KPIs.
4. Do not guess, estimate, or calculate values that are not explicitly
   stated.
5. Respond with the structured JSON output containing everything you found.
"""


def call_gemini(source_kind, source_payload, prompt):
    """source_payload is raw PDF bytes when source_kind == 'pdf', or a plain
    text string (the converted table) when source_kind is 'excel'/'csv'."""
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError(
            "GEMINI_API_KEY is not set. Get a free key at https://aistudio.google.com/apikey."
        )
    client = genai.Client(api_key=api_key)

    if source_kind == "pdf":
        pdf_b64 = base64.b64encode(source_payload).decode("utf-8")
        input_blocks = [
            {"type": "document", "data": pdf_b64, "mime_type": "application/pdf"},
            {"type": "text", "text": prompt},
        ]
    else:
        # No Gemini document type for spreadsheets -- send the converted
        # table as plain text alongside the prompt instead.
        input_blocks = [
            {"type": "text", "text": f"Source table:\n\n{source_payload}"},
            {"type": "text", "text": prompt},
        ]

    interaction = client.interactions.create(
        model=MODEL,
        input=input_blocks,
        response_format={
            "type": "text",
            "mime_type": "application/json",
            "schema": ExtractionResult.model_json_schema(),
        },
    )

    try:
        result = ExtractionResult.model_validate_json(interaction.output_text)
    except Exception as e:
        raise RuntimeError(f"Gemini did not return valid structured JSON: {e}\n\nRaw response:\n{interaction.output_text}")
    return result


def enrich_rows(extracted, kpi_ref_by_id, country, region, source_label):
    rows = []
    skipped = []
    for item in extracted:
        ref = kpi_ref_by_id.get(item.attribute_id)
        if ref is None:
            skipped.append(item.attribute_id)
            continue

        observation_basis, ob_confidence = resolve_observation_basis(
            category=ref["observation_basis_category"],
            check_source=ref["observation_basis_check_source"],
            year=item.year,
            quarter=item.quarter,
            month=item.month,
            model_hint=item.observation_basis_hint,
        )

        rows.append({
            "attribute_id": item.attribute_id,
            "attribute_name": ref["name_en"],
            "value": item.value,
            "unit": item.unit or ref["unit"],
            "Country_name": country,
            "Operator_name": item.operator or "ALL",
            "Region_name": region,
            "year": item.year,
            "quarter": item.quarter,
            "month": item.month,
            "Module": ref["module"],
            "source_id": source_label,
            "data_type": ref["data_origin"],
            "frequency_name": ref["frequency"],
            "source_snippet": item.source_snippet,
            "observation_basis": observation_basis,
            "ob_confidence": ob_confidence,
        })
    return rows, skipped


def write_output(rows, output_path):
    output_path = Path(output_path)
    if output_path.exists():
        wb = load_workbook(output_path)
        ws = wb.active
        ensure_headers(ws)  # upgrade an older database file's header row if needed
        start_row = ws.max_row + 1
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "KPI MASTER DB"
        for col_idx, header in enumerate(OUTPUT_COLUMNS, start=1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = Font(name="Arial", bold=True)
        start_row = 2

    for r_offset, row in enumerate(rows):
        for col_idx, col_name in enumerate(OUTPUT_COLUMNS, start=1):
            cell = ws.cell(row=start_row + r_offset, column=col_idx, value=row.get(col_name))
            cell.font = Font(name="Arial")

    wb.save(output_path)


def extract_from_source(file_bytes, filename, country, source_label, output_path, sheet_name=None):
    """End-to-end: source file bytes (PDF, Excel, or CSV) -> Gemini
    extraction -> enriched rows -> appended to output_path. Used by both the
    CLI (main, below) and app.py. Raises ValueError/RuntimeError on failure
    instead of exiting, so callers running inside a long-lived process (like
    a Streamlit app) can catch and display the error instead of the whole
    process dying.

    sheet_name is required when filename is a multi-sheet .xlsx (see
    list_excel_sheets); ignored for PDF/CSV.
    """
    source_kind = source_kind_from_filename(filename)

    kpi_ref = load_kpi_reference()
    kpi_ref_by_id = {k["attribute_id"]: k for k in kpi_ref}
    config = load_country_config(country)  # raises ValueError if unknown

    if source_kind == "pdf":
        source_payload = file_bytes
    elif source_kind == "excel":
        source_payload = _read_excel_as_text(file_bytes, sheet_name)
    else:  # csv
        source_payload = _read_csv_as_text(file_bytes)

    prompt = build_prompt(
        country=country,
        regulator=config["regulator"],
        operators=config["operators"],
        source_label=source_label,
        kpi_ref=kpi_ref,
        source_kind=source_kind,
    )

    result = call_gemini(source_kind, source_payload, prompt)  # raises RuntimeError on failure

    rows, skipped_ids = enrich_rows(
        result.extracted, kpi_ref_by_id, country, config["region"], source_label
    )
    write_output(rows, output_path)

    found_ids = {r["attribute_id"] for r in rows}
    all_ids = {k["attribute_id"] for k in kpi_ref}

    return {
        "rows": rows,
        "skipped_ids": skipped_ids,
        "not_found_notes": result.not_found_notes,
        "found_count": len(found_ids),
        "total_count": len(all_ids),
        "missing_ids": sorted(all_ids - found_ids),
    }


def extract_from_pdf(pdf_bytes, country, source_label, output_path):
    """Backward-compatible wrapper kept for existing callers -- PDF-only,
    same as extract_from_source with filename='report.pdf'."""
    return extract_from_source(pdf_bytes, "report.pdf", country, source_label, output_path)


def main():
    parser = argparse.ArgumentParser(description="Extract telecom KPIs from a report file using Gemini.")
    parser.add_argument("--file", help="Path to the report file: PDF, Excel (.xlsx), or CSV.")
    parser.add_argument("--pdf", help=argparse.SUPPRESS)  # old flag, kept working as an alias for --file
    parser.add_argument("--sheet", default=None, help="Sheet name to extract from, if --file is a multi-sheet .xlsx.")
    parser.add_argument("--country", required=True, help="Country name, must match country_config.csv.")
    parser.add_argument("--source-label", required=True, help="Human-readable label for this report, e.g. 'ARTP Tableau de bord Mobile T2 2025'.")
    parser.add_argument("--output", default="KPI_MASTER_DB_new.xlsx", help="Output .xlsx file (created or appended to).")
    args = parser.parse_args()

    file_arg = args.file or args.pdf
    if not file_arg:
        sys.exit("Provide a source file with --file (PDF, .xlsx, or .csv).")

    file_path = Path(file_arg)
    if not file_path.exists():
        sys.exit(f"File not found: {file_path}")

    try:
        source_kind_from_filename(file_path.name)
    except ValueError as e:
        sys.exit(str(e))

    print(f"Sending {file_arg} to {MODEL} for extraction...")
    try:
        result = extract_from_source(
            file_path.read_bytes(), file_path.name, args.country, args.source_label,
            args.output, sheet_name=args.sheet,
        )
    except (ValueError, RuntimeError) as e:
        sys.exit(str(e))

    print(f"\nExtracted {len(result['rows'])} KPI values -> {args.output}")
    if result["skipped_ids"]:
        print(f"Warning: {len(result['skipped_ids'])} returned attribute_id(s) not in kpi_reference.json (dropped): {result['skipped_ids']}")
    if result["not_found_notes"]:
        print(f"\nModel notes: {result['not_found_notes']}")

    print(f"\n{result['found_count']}/{result['total_count']} KPIs found in this report.")
    print("Not found in this report (may be in a different source document):")
    print(", ".join(result["missing_ids"]))


if __name__ == "__main__":
    main()