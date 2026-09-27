"""
Derived KPI Calculator
=======================
Second stage of the pipeline. Reads the raw + explicitly-stated values
produced by extract_kpis.py and computes the "calculated" KPIs (data_origin
== "calculated" in kpi_reference.json) using the formulas already defined
there -- no LLM involved, pure deterministic arithmetic, so results are
100% traceable to a formula you can check.

Usage
-----
    python calculate_derived_kpis.py --file KPI_MASTER_DB_new.xlsx

Rows are appended to the same file with data_type = "calculated (auto)".

Important: not every "calculated" KPI in kpi_reference.json can actually be
computed. Several formulas reference variables that were never given a
tracked metric anywhere in your KPI list (e.g. NB_SITES, SPECTRUM, regional
benchmarks) or are undefined placeholder functions ("f(...)"). Run this
script and read the final report -- the list of what it could NOT compute,
and why, is directly useful for your project's "which KPIs need human
intervention" reflection.
"""

import argparse
import json
import re
import sys
from pathlib import Path
from collections import defaultdict

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Font
except ImportError:
    sys.exit("Missing dependency. Run: pip install openpyxl --break-system-packages")

from kpi_shared import OUTPUT_COLUMNS, category_default_text, CONFIDENCE_DERIVED, ensure_headers


SCRIPT_DIR = Path(__file__).resolve().parent
KPI_REFERENCE_PATH = SCRIPT_DIR / "kpi_reference.json"

# Fixes for known inconsistent variable naming between the formula text in
# the KPIs tab and the actual tracked metric_name. Edit this if you decide
# a mapping is wrong, or add more as you find them -- that's the point of
# keeping it as an explicit, visible table rather than guessing silently.
ALIASES = {
    "REVENUE": "CA_TOTAL_MOBILE",
    "SUBS": "SUBS_MOBILE",
    "TRAFFIC": "TRAFFIC_DATA",
    "ARPU": "ARPU_MOBILE",
    "MARKET_SHARE": "MARKET_SHARE_MOBILE",
    "INBOUND": "TRAFFIC_INBOUND",
    "OUTBOUND": "TRAFFIC_OUTBOUND",
    "TRAFFIC_VOICE": "TRAFFIC_VOICE_MOBILE",
    "SUBS_DATA": "SUBS_DATA_MOBILE",
    "GROWTH_TRAFFIC": "GROWTH_TRAFFIC_DATA",
    "GROWTH_REV": "GROWTH_REVENUE",
}

GROWTH_RE = re.compile(r"^\(([A-Za-z_]+)_t\s*-\s*\1_t-1\)/\1_t-1$")
SUMSQ_RE = re.compile(r"^SUM\(([A-Za-z_]+)\u00b2\)$")
BINOP_RE = re.compile(r"^([A-Za-z_]+)\s*([/\-\+])\s*([A-Za-z_]+)$")


def load_kpi_reference():
    with open(KPI_REFERENCE_PATH, encoding="utf-8") as f:
        return json.load(f)


def read_existing_rows(path):
    wb = load_workbook(path)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] is None:
            continue
        rows.append(dict(zip(headers, r)))
    return rows


def resolve_name(name, tracked):
    """Return the tracked metric_name this formula variable refers to, or None."""
    if name in tracked:
        return name
    if name in ALIASES and ALIASES[name] in tracked:
        return ALIASES[name]
    if name.startswith("TOTAL_"):
        base = name[len("TOTAL_"):]
        base = ALIASES.get(base, base)
        if base in tracked:
            return "TOTAL_" + base
    return None


def period_key(row):
    return (row.get("year"), row.get("quarter"), row.get("month"))


def sort_key(period):
    year, quarter, month = period
    q = int(quarter[1]) if quarter and quarter[0] in "Qq" else 0
    m = month or 0
    return (year or 0, q, m)


class ValueStore:
    """Maps (metric_name, country, operator, period) -> value, with an
    operator fallback to 'ALL' for country-only metrics."""

    def __init__(self):
        self.data = {}
        # index of (metric_name, country, operator) -> list of (period, value), for growth/lookback
        self.series = defaultdict(list)

    def set(self, metric, country, operator, period, value):
        self.data[(metric, country, operator, period)] = value
        self.series[(metric, country, operator)].append((period, value))

    def get(self, metric, country, operator, period):
        v = self.data.get((metric, country, operator, period))
        if v is not None:
            return v
        if operator != "ALL":
            return self.data.get((metric, country, "ALL", period))
        return None

    def get_total_across_operators(self, metric, country, period):
        total, found = 0.0, False
        for (m, c, op), series in self.series.items():
            if m == metric and c == country and op != "ALL":
                for p, v in series:
                    if p == period:
                        total += v
                        found = True
        return total if found else None

    def get_all_values_across_operators(self, metric, country, period):
        vals = []
        for (m, c, op), series in self.series.items():
            if m == metric and c == country and op != "ALL":
                for p, v in series:
                    if p == period:
                        vals.append(v)
        return vals

    def get_previous(self, metric, country, operator, period):
        series = sorted(self.series.get((metric, country, operator), []), key=lambda pv: sort_key(pv[0]))
        periods = [p for p, _ in series]
        if period not in periods:
            return None
        idx = periods.index(period)
        if idx == 0:
            return None
        return series[idx - 1][1]


def compute(kpi_ref, store, existing_periods_by_country):
    tracked = {k["metric_name"] for k in kpi_ref}
    id_by_metric = {k["metric_name"]: k["attribute_id"] for k in kpi_ref}
    calc_kpis = [k for k in kpi_ref if k["data_origin"] == "calculated"]

    resolved_rows = []
    unresolved = []
    already_present = []

    for k in calc_kpis:
        formula = k["formula"]
        metric = k["metric_name"]

        if formula.startswith("f("):
            unresolved.append((metric, formula, "undefined function placeholder -- needs a human-defined rule"))
            continue

        m = GROWTH_RE.match(formula)
        sq = SUMSQ_RE.match(formula)
        bo = BINOP_RE.match(formula)

        if m:
            base = resolve_name(m.group(1), tracked)
            if base is None:
                unresolved.append((metric, formula, f"base variable '{m.group(1)}' not tracked and no alias defined"))
                continue
            kind = ("growth", base)
        elif sq:
            base = resolve_name(sq.group(1), tracked)
            if base is None:
                unresolved.append((metric, formula, f"base variable '{sq.group(1)}' not tracked and no alias defined"))
                continue
            kind = ("sumsq", base)
        elif bo:
            a_raw, op, b_raw = bo.groups()
            a, b = resolve_name(a_raw, tracked), resolve_name(b_raw, tracked)
            missing = [v for v, r in [(a_raw, a), (b_raw, b)] if r is None]
            if missing:
                unresolved.append((metric, formula, f"variable(s) not tracked and no alias defined: {missing}"))
                continue
            kind = ("binop", a, op, b)
        else:
            unresolved.append((metric, formula, "formula pattern not recognized by the calculator"))
            continue

        # Try to compute a value for every (country, operator, period) combo we've seen data for.
        computed_any = False
        computed_keys = set()  # dedupe: some kinds (sumsq) yield the same key across operator iterations
        for country, periods in existing_periods_by_country.items():
            operators = {op for (m2, c2, op, p2) in store.data if c2 == country}
            for operator in operators:
                for period in periods:
                    value = None
                    if kind[0] == "growth":
                        base = kind[1]
                        cur = store.get(base, country, operator, period)
                        prev = store.get_previous(base, country, operator, period)
                        if cur is not None and prev not in (None, 0):
                            value = (cur - prev) / prev
                    elif kind[0] == "sumsq":
                        base = kind[1]
                        vals = store.get_all_values_across_operators(base, country, period)
                        if vals:
                            value = sum(v * v for v in vals)
                            operator = "ALL"
                    elif kind[0] == "binop":
                        _, a, op, b = kind
                        av = _lookup(store, a, country, operator, period)
                        bv = _lookup(store, b, country, operator, period)
                        if av is not None and bv is not None:
                            if op == "/":
                                value = av / bv if bv != 0 else None
                            elif op == "+":
                                value = av + bv
                            elif op == "-":
                                value = av - bv

                    if value is not None:
                        key = (metric, country, operator, period)
                        if key in computed_keys:
                            continue
                        computed_keys.add(key)
                        store.set(metric, country, operator, period, value)
                        # Stage 2 has no report to check against -- the category
                        # default isn't a guess here, it's guaranteed by the
                        # arithmetic we just performed (e.g. a period total
                        # divided by a stock count IS a period average), so it
                        # always gets the 'derived (calculated)' confidence
                        # rather than 'formula-defined'/'assumed'.
                        observation_basis = category_default_text(
                            k.get("observation_basis_category"), period[0], period[1], period[2]
                        )
                        resolved_rows.append({
                            "attribute_id": k["attribute_id"],
                            "attribute_name": k["name_en"],
                            "value": value,
                            "unit": k["unit"],
                            "Country_name": country,
                            "Operator_name": operator,
                            "Region_name": None,
                            "year": period[0],
                            "quarter": period[1],
                            "month": period[2],
                            "Module": k["module"],
                            "source_id": "calculated (auto)",
                            "data_type": "calculated (auto)",
                            "frequency_name": k["frequency"],
                            "source_snippet": f"formula: {formula}",
                            "observation_basis": observation_basis,
                            "ob_confidence": CONFIDENCE_DERIVED,
                        })
                        computed_any = True
        if not computed_any:
            unresolved.append((metric, formula, "no matching raw data found for any period/operator combo"))

    return resolved_rows, unresolved


def _lookup(store, metric, country, operator, period):
    if metric.startswith("TOTAL_"):
        base = metric[len("TOTAL_"):]
        return store.get_total_across_operators(base, country, period)
    return store.get(metric, country, operator, period)


def run_derived_calculations(path):
    """Core logic: read existing rows from path, compute every derived KPI it
    can, append the new rows to the workbook, and return
    (new_rows, skipped_dupes, unresolved) for the caller to report on.
    Raises FileNotFoundError if path doesn't exist. Used by both the CLI
    (main, below) and app.py."""
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    kpi_ref = load_kpi_reference()
    id_to_metric = {k["attribute_id"]: k["metric_name"] for k in kpi_ref}
    existing_rows = read_existing_rows(path)

    store = ValueStore()
    periods_by_country = defaultdict(set)

    for row in existing_rows:
        metric = id_to_metric.get(row["attribute_id"])
        if metric is None:
            continue
        country = row["Country_name"]
        operator = row["Operator_name"] or "ALL"
        period = period_key(row)
        store.set(metric, country, operator, period, row["value"])
        periods_by_country[country].add(period)
        # If this row is itself an already-computed/explicit calculated value, don't recompute it.

    resolved_rows, unresolved = compute(kpi_ref, store, periods_by_country)

    # Don't duplicate rows for (attribute_id, country, operator, period) combos
    # that already exist in the file (e.g. the report explicitly stated a
    # calculated value like EBITDA margin -- prefer the reported number).
    existing_keys = {
        (r["attribute_id"], r["Country_name"], r["Operator_name"] or "ALL", period_key(r))
        for r in existing_rows
    }
    new_rows = [
        r for r in resolved_rows
        if (r["attribute_id"], r["Country_name"], r["Operator_name"], (r["year"], r["quarter"], r["month"])) not in existing_keys
    ]
    skipped_dupes = len(resolved_rows) - len(new_rows)

    # Fill Region_name from any existing row for that country.
    region_by_country = {r["Country_name"]: r["Region_name"] for r in existing_rows if r.get("Region_name")}
    for r in new_rows:
        r["Region_name"] = region_by_country.get(r["Country_name"])

    # Append to the workbook.
    wb = load_workbook(path)
    ws = wb.active
    ensure_headers(ws)  # upgrade an older database file's header row if needed
    start_row = ws.max_row + 1
    for r_offset, row in enumerate(new_rows):
        for col_idx, col_name in enumerate(OUTPUT_COLUMNS, start=1):
            cell = ws.cell(row=start_row + r_offset, column=col_idx, value=row.get(col_name))
            cell.font = Font(name="Arial")
    wb.save(path)

    return new_rows, skipped_dupes, unresolved


def format_unresolved_report(unresolved):
    """De-duped, human-readable text block for the CLI. app.py renders the
    same (metric, formula, reason) tuples as a table instead."""
    seen = set()
    lines = [f"\n{'='*60}\nCOULD NOT AUTO-COMPUTE ({len(unresolved)} formulas):\n{'='*60}"]
    for metric, formula, reason in unresolved:
        if metric in seen:
            continue
        seen.add(metric)
        lines.append(f"  {metric:30} = {formula:40} -> {reason}")
    lines.append(
        "\nThese need either: a fix to the ALIASES table in this script, a source that\n"
        "states the value directly (extract_kpis.py can then pick it up), or a decision\n"
        "that they stay human-calculated for now -- this list is your 'Reflexion finale' answer."
    )
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description="Compute derived KPIs from already-extracted raw values.")
    parser.add_argument("--file", required=True, help="The .xlsx produced by extract_kpis.py (rows are appended to it).")
    args = parser.parse_args()

    try:
        new_rows, skipped_dupes, unresolved = run_derived_calculations(args.file)
    except FileNotFoundError as e:
        sys.exit(str(e))

    print(f"Computed {len(new_rows)} new derived KPI values -> {args.file}")
    if skipped_dupes:
        print(f"({skipped_dupes} matched an already-present explicit value and were not duplicated)")

    print(format_unresolved_report(unresolved))


if __name__ == "__main__":
    main()