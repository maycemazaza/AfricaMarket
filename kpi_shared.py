#!/usr/bin/env python3
"""
Shared constants and helpers for the KPI pipeline
===================================================
Both extract_kpis.py (Stage 1) and calculate_derived_kpis.py (Stage 2) write
rows with the same column structure and the same "Observation Basis" wording
rules. This module is the one place that structure/wording lives, so the two
stages can't quietly drift apart -- same "one copy of the logic" principle
the rest of this project already follows for the extraction/calculation code
itself.
"""

# Columns of the master database. Both scripts append rows shaped like this.
OUTPUT_COLUMNS = [
    "attribute_id", "attribute_name", "value", "unit", "Country_name",
    "Operator_name", "Region_name", "year", "quarter", "month", "Module",
    "source_id", "data_type", "frequency_name", "source_snippet",
    "observation_basis", "ob_confidence",
]

# Human-readable phrasing per (category, period granularity), matching the
# vocabulary already used in the manually-reviewed Consolidated Data sheet
# (its "Points d'amelioration" tab: Annual Average / End of Year / Annual
# Total / Annual Rate / Annual Index / Quarterly Total / End of Month /
# Monthly average / Monthly Sum). Quarter- and month-level "average"/"rate"/
# "index"/"ratio" wording is extended from that same pattern for KPIs that
# are reported more often than annually.
OBSERVATION_BASIS_TEMPLATES = {
    "stock":           {"Year": "End of Year",       "Quarter": "End of Quarter",       "Month": "End of Month"},
    "flow":            {"Year": "Annual Total",       "Quarter": "Quarterly Total",       "Month": "Monthly Sum"},
    "average":         {"Year": "Annual Average",     "Quarter": "Quarterly Average",     "Month": "Monthly Average"},
    "rate":            {"Year": "Annual Rate",        "Quarter": "Quarterly Rate",        "Month": "Monthly Rate"},
    "index":           {"Year": "Annual Index",       "Quarter": "Quarterly Index",       "Month": "Monthly Index"},
    "ratio":           {"Year": "Annual Ratio",       "Quarter": "Quarterly Ratio",       "Month": "Monthly Ratio"},
    # snapshot_ratio and regulatory_rate behave like "stock" for wording
    # purposes -- both describe a value that holds at a point in time rather
    # than something summed or averaged over the period.
    "snapshot_ratio":  {"Year": "End of Year",        "Quarter": "End of Quarter",        "Month": "End of Month"},
    "regulatory_rate": {"Year": "End of Year",        "Quarter": "End of Quarter",        "Month": "End of Month"},
}

# Maps the model's optional, report-wording-based hint (see extract_kpis.py's
# ExtractedKPI.observation_basis_hint) onto the same phrasing scheme above, so
# a hint can be compared against -- or substituted for -- the category default.
HINT_TO_CATEGORY = {
    "point_in_time": "stock",
    "period_total": "flow",
    "period_average": "average",
}

CONFIDENCE_FORMULA_DEFINED = "formula-defined"   # not source-checked; category is certain by definition
CONFIDENCE_CONFIRMED = "confirmed"                # report wording matched the expected category
CONFIDENCE_DIFFERS = "differs from expected — review"  # report wording contradicted the expected category
CONFIDENCE_ASSUMED = "assumed"                    # source-checked category, but report wording wasn't clear
CONFIDENCE_DERIVED = "derived (calculated)"       # Stage 2 computed the value itself; category isn't a guess


def period_label(year, quarter, month):
    """Return 'Month', 'Quarter', or 'Year' -- the finest granularity implied
    by which period fields are actually filled in on this row."""
    if month:
        return "Month"
    if quarter:
        return "Quarter"
    return "Year"


def category_default_text(category, year, quarter, month):
    """The formula-defined Observation Basis phrase for this category at this
    row's period granularity, e.g. category_default_text('flow', 2026, 'Q1', None)
    -> 'Quarterly Total'."""
    templates = OBSERVATION_BASIS_TEMPLATES.get(category)
    if templates is None:
        return None
    return templates[period_label(year, quarter, month)]


def hint_text(hint, year, quarter, month):
    """Render the model's observation_basis_hint (point_in_time / period_total
    / period_average) using the same wording scheme as the category default,
    so the two are directly comparable."""
    mapped_category = HINT_TO_CATEGORY.get(hint)
    if mapped_category is None:
        return None
    return category_default_text(mapped_category, year, quarter, month)


def ensure_headers(ws):
    """Backward compatibility for existing database files: if the sheet's
    header row (row 1) predates the Observation Basis columns (or any other
    future OUTPUT_COLUMNS addition), append the missing header(s) at the end
    rather than silently writing data into unlabeled columns. Safe to call
    every time a file is opened for appending -- a no-op once headers match."""
    from openpyxl.styles import Font

    existing = [c.value for c in ws[1]]
    for col_name in OUTPUT_COLUMNS:
        if col_name not in existing:
            new_col_idx = ws.max_column + 1
            cell = ws.cell(row=1, column=new_col_idx, value=col_name)
            cell.font = Font(name="Arial", bold=True)
            existing.append(col_name)


def resolve_observation_basis(category, check_source, year, quarter, month, model_hint=None):
    """Core decision used by extract_kpis.py for every extracted row.
    Returns (observation_basis_text, confidence_label).

    - Not flagged for source-checking -> always the category default,
      confidence 'formula-defined'.
    - Flagged, model gave a clear hint -> compare hint to the default:
      match -> 'confirmed' (keep the default text); mismatch -> use the
      model's hint text instead, confidence 'differs from expected'.
    - Flagged, model gave no hint -> category default, confidence 'assumed'.
    """
    default_text = category_default_text(category, year, quarter, month)

    if not check_source:
        return default_text, CONFIDENCE_FORMULA_DEFINED

    if model_hint:
        candidate = hint_text(model_hint, year, quarter, month)
        if candidate is not None:
            if candidate == default_text:
                return default_text, CONFIDENCE_CONFIRMED
            return candidate, CONFIDENCE_DIFFERS

    return default_text, CONFIDENCE_ASSUMED
