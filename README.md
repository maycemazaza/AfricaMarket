# KPI Extraction Agent

Two-stage pipeline:

1. **`extract_kpis.py`** — feed it a regulator/operator report (PDF, Excel,
   or CSV), get back the KPI values *explicitly stated* in that report, in
   the same column structure as your `KPI MASTER DB` tab. This is the
   "universal prompt," wired into the API instead of typed into a chat
   window each time.
2. **`calculate_derived_kpis.py`** — computes the KPIs marked
   `data_origin: calculated` in your `KPIs` tab (ARPU, market share, HHI,
   growth rates, etc.) from whatever raw values stage 1 has collected so
   far. This is plain Python arithmetic, not the model -- every number is
   traceable to the exact formula that produced it.

Run stage 1 once per report you have. Run stage 2 whenever you want the
derived KPIs refreshed from the latest raw data.

## What's in this folder

| File | What it is |
|---|---|
| `app.py` | Local web app (Streamlit) -- upload a report and click buttons instead of running commands. See "Running as a local web app" below. |
| `extract_kpis.py` | Stage 1: report → raw KPI rows. Also the CLI entry point. |
| `calculate_derived_kpis.py` | Stage 2: raw KPI rows → derived/calculated KPI rows, using the formulas in `kpi_reference.json`. Also the CLI entry point. |
| `kpi_shared.py` | Column list and Observation Basis wording rules shared by both stages, so they can't drift apart. Neither script duplicates this logic anymore. |
| `classify_observation_basis.py` | One-off script that assigned each KPI's `observation_basis_category`/`observation_basis_check_source` in `kpi_reference.json`. Not part of the running pipeline -- kept as a record of how/why each KPI was classified, in case one needs reclassifying later. |
| `kpi_reference.json` | Your 86 KPIs, pulled straight from the `KPIs` tab (id, formula, unit, module, whether it's country- or operator-level, raw vs calculated, Observation Basis category). Both scripts read this. |
| `country_config.csv` | Country → regulator → operators. Add a row here to support a new country -- that's the "fill in one row instead of 2 weeks of manual work" step from Phase 2. |
| `README.md` | This file. |

`app.py` doesn't reimplement any KPI logic -- it imports `extract_kpis.py`
and `calculate_derived_kpis.py` directly and calls their functions. There is
exactly one copy of the extraction/calculation logic; the CLI and the web
app are just two different front ends onto it.

## Source file types

`extract_kpis.py` accepts three kinds of source file:

- **PDF** — sent to Gemini as a document, same as before.
- **Excel (.xlsx)** — converted to a plain text table and sent as text
  (Gemini has no native spreadsheet input type). If the workbook has more
  than one sheet, you pick which one to extract from -- sending an entire
  multi-sheet workbook as one blob would dilute the model's attention the
  same way an oversized PDF does. The CLI's `--sheet` flag / the app's sheet
  picker is where you choose.
- **CSV** — read directly as text. Encoding is tried as UTF-8 first, then
  `cp1252`/`latin-1` as a fallback, since regulator CSVs exported from Excel
  with French text (é, à, ç) aren't always UTF-8.

Either way, the model still does all the interpretive work (messy headers,
inconsistent layouts) -- converting to text doesn't mean parsing it into
structure first, just making it something Gemini can read.

## Observation Basis

Every row also gets an **Observation Basis** column: what *kind* of number
the value is for its period -- a point-in-time snapshot ("End of Month"), a
period total ("Quarterly Total"), a period average ("Annual Average"), a
growth rate, an index, or a ratio. This mirrors a column already used in
your manually-reviewed Consolidated Data sheet.

Each KPI in `kpi_reference.json` has an `observation_basis_category`
(`stock`, `flow`, `average`, `rate`, `index`, `snapshot_ratio`, `ratio`, or
`regulatory_rate`) and an `observation_basis_check_source` flag:

- **`false`** (~50 KPIs — subscriber/traffic/revenue counts, growth rates,
  indices, ratio-of-ratios): the category is guaranteed by definition or by
  the arithmetic that computed it. No model involvement; the column is
  filled in from `kpi_reference.json` exactly like `Module` or
  `frequency_name` are today. Confidence label: `formula-defined`.
- **`true`** (~36 KPIs — averages like ARPU, stock/stock ratios like
  penetration and market share, and regulatory rate floors): this is where
  real-world methodology is known to vary by regulator. Proven, not assumed
  -- two KPIs with an identical formula shape (`stock / stock`, both %)
  came back with different real Observation Basis values in the
  already-collected data: mobile penetration was "End of Month", smartphone
  penetration was "Average Quarter". For these, `extract_kpis.py` also asks
  the model to report what the report's own wording implies
  (`observation_basis_hint`), then:
  - hint matches the category default → keep the default, confidence
    `confirmed`
  - hint contradicts the default → use the hint instead (trust the report's
    own wording over a generic assumption), confidence `differs from
    expected — review`
  - report doesn't state it clearly → category default, confidence
    `assumed`

`calculate_derived_kpis.py` (Stage 2) always uses the category default,
confidence `derived (calculated)` -- there's no report to check against at
that stage, but the category isn't a guess either: it's guaranteed by the
arithmetic Stage 2 just performed (a period total divided by a stock count
*is* a period average).

The View Database tab in the app has a filter on this confidence column, so
you can pull up every row that's `assumed` or flagged `differs from
expected — review` without scrolling the whole database.

If you disagree with a KPI's category, or want to move one in/out of the
source-check group, edit `observation_basis_category` /
`observation_basis_check_source` directly in `kpi_reference.json` -- same
philosophy as the `ALIASES` table in `calculate_derived_kpis.py`: an
explicit, visible, editable mapping rather than a hidden rule.

## Setup (one time)

```bash
pip install google-genai pydantic openpyxl pandas
$env:GEMINI_API_KEY="YOUR_API_KEY"
```

(`pandas` is needed now too -- it's how Excel/CSV sources get converted to
the text table that's sent to Gemini.)

Get a free key at https://aistudio.google.com/apikey if you don't have one --
no credit card required. The free tier has rate limits (requests per
minute/day), which is fine for processing a handful of reports at a time.

## Running as a local web app (recommended)

```bash
python -m pip install streamlit pandas
python -m streamlit run app.py
```

This opens a browser tab at `http://localhost:8501`. Three tabs:

1. **Extract from a report** -- paste your Gemini API key into the sidebar
   once per session, upload a report (PDF, Excel, or CSV), pick the country,
   click **Extract KPIs**. If you upload a multi-sheet Excel workbook, a
   sheet picker appears first -- choose the tab that actually holds the KPI
   data. Results show up as a table right there, with a warning if any row's
   Observation Basis contradicted the usual assumption for that KPI.
2. **Compute Derived KPIs** -- click a button to run stage 2 against the
   current database. Shows what it computed and what it couldn't, same as
   the CLI's printed report but as tables.
3. **View Database** -- browse/filter the whole master database and
   download the `.xlsx`.

Leave the terminal window open while you use the app -- it's running a
local server; closing the terminal (or pressing Ctrl+C in it) stops the
app. The API key you type into the sidebar only lives in that browser
session's memory, not written to disk.

If you'd rather not paste the key in every time you restart the app, set it
as an environment variable first (see the CLI section below) -- the app
picks it up automatically and pre-fills the sidebar field.

## Running from the command line instead

Both scripts still work as standalone CLI tools, if you prefer that or want
to script/automate them (e.g. a scheduled batch job):

```bash
python extract_kpis.py \
    --file "ARTP_Tableau_Bord_Mobile_T2_2025.pdf" \
    --country "Senegal" \
    --source-label "ARTP Tableau de bord Mobile T2 2025" \
    --output "KPI_MASTER_DB_new.xlsx"

# Excel source, picking one sheet out of a multi-tab workbook:
python extract_kpis.py \
    --file "regulator_data.xlsx" --sheet "Ivory coast" \
    --country "Ivory Coast" \
    --source-label "Regulator workbook 2025" \
    --output "KPI_MASTER_DB_new.xlsx"
```

- `--file`: the report you're extracting from -- PDF, `.xlsx`, or `.csv`.
  (`--pdf` still works as an alias, for existing scripts/muscle memory.)
- `--sheet`: which sheet to read, if `--file` is a multi-sheet `.xlsx`. Not
  needed for a single-sheet workbook, PDF, or CSV.
- `--country`: must match a row in `country_config.csv` (case-insensitive).
- `--source-label`: a short human-readable name for this report -- goes into
  the `source_id` column so you always know which document a number came
  from.
- `--output`: rows get **appended** to this file, so point every report at
  the same output file to build up your database. It's created
  automatically on the first run. If it's an older database file from
  before the Observation Basis columns existed, the header row is upgraded
  in place the first time you extract into it.

Run it once per report. A country usually needs several runs (one per
quarterly bulletin, one for the annual report, etc.) -- exactly like your
manual process in Phase 1.

## Running stage 2 (derived KPIs)

```bash
python calculate_derived_kpis.py --file KPI_MASTER_DB_new.xlsx
```

It reads whatever raw rows are already in that file, computes every
`calculated` KPI it can, appends the results, and prints two things:

- how many derived rows it added
- a list of formulas it **could not** compute automatically, and why

That second list is not a bug report -- it's real signal. Of your 51
calculated KPIs, 27 resolve cleanly today (ratios, growth rates, share-of-
total). The other 24 fail for one of two reasons:

- **Missing underlying data**: the formula references something no KPI in
  your schema tracks at all -- `NB_SITES`, `SPECTRUM`, `CA_VOICE` (mobile
  voice revenue isn't tracked separately from total mobile revenue),
  `LOCAL`/`REGIONAL` (no regional benchmark dataset exists anywhere in this
  project). No script can compute these without that data being collected
  first, from a report or another source.
- **Undefined formula**: two KPIs (`PRESSURE_INDEX`, `REGULATORY_CONSTRAINTS_INDEX`)
  are written as `f(...)` in the `KPIs` tab -- literally a placeholder, not
  an actual formula. Someone has to decide what the function is before any
  tool, human or AI, can compute them.

A `ALIASES` dict near the top of `calculate_derived_kpis.py` fixes a few
naming mismatches between the formula text and the tracked metric names
(e.g. the formula says `REVENUE`, the tracked metric is `CA_TOTAL_MOBILE`).
It's deliberately visible and editable -- if you disagree with a mapping,
change it there.

This whole list is directly useful as your project's "Reflexion finale"
deliverable: which KPIs can be automated end-to-end (raw extraction +
calculation), which need better source data before they can be, and which
need a human to define the formula in the first place.

## Adding a new country

Add one row to `country_config.csv`:

```csv
country,regulator,operators,region
DR Congo,ARPTC,"Vodacom RDC, Orange RDC, Airtel RDC, Africell RDC",Africa
```

Then run the same script, no other changes. This is the validation test
from Phase 2 -- if it works well on a country you never wrote a
country-specific prompt for, the "universal prompt" claim is proven.

## Checking accuracy

After running it on a report you already collected manually:

1. Open the script's output `.xlsx` next to your existing `KPI MASTER DB`.
2. For matching `attribute_id` / `Country_name` / `year` / `quarter` /
   `Operator_name`, compare the `value` columns.
3. Also check the `source_snippet` column -- it tells you exactly where in
   the PDF the model found each number, so a wrong value is quick to
   investigate.
4. Tally match / mismatch / missing per KPI -- this is your Accuracy Report
   deliverable.

## Why the model doesn't calculate anything

The script explicitly tells the model not to compute derived KPIs (ARPU,
penetration rates, HHI, etc.) itself -- only to report numbers that are
**explicitly printed** in the document. Derived KPIs get computed
afterward in Excel using the formulas already defined in your `KPIs` tab,
exactly like Phase 1. This keeps every calculation traceable to a formula
you can check, instead of trusting the model's arithmetic (and its
sourcing) at the same time.

## Known limitations

- One source file per run. If a KPI is split across two report volumes
  (e.g. a quarterly bulletin + an annual report), run the script on each and
  let them append to the same output file.
- Very large or scanned/low-quality PDFs reduce extraction accuracy --
  check `source_snippet` on anything that looks off. The same applies to a
  very large or messily-laid-out Excel sheet -- pick the smallest sheet
  that actually contains the data you need rather than a whole workbook.
- The script doesn't dedupe: running the same source file twice appends the
  same rows twice. Delete duplicate rows manually, or start fresh output
  files per batch if that's easier for your workflow.
- Observation Basis for the ~36 flagged KPIs (`observation_basis_check_source:
  true`) depends on the model correctly reading the report's wording. It's
  evidence-based, not guaranteed -- treat `confirmed` as "matches what we'd
  expect, and the report said so," not as independently audited. Rows
  marked `differs from expected — review` are the ones most worth a human
  look, since that's exactly the case (see A001 vs A004 in the Observation
  Basis section above) where the formula-based assumption is known to
  sometimes be wrong.
- CSVs exported from Excel with French accented text aren't always UTF-8 --
  the reader falls back to `cp1252` then `latin-1`, but a CSV in some other
  encoding may still fail to decode.