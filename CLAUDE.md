# Quotation Solution — Project Context for Claude

This file gives you full context so you can continue helping without needing the full conversation history.

---

## What This Project Is

An ERP data quality analysis tool for **sales quotation data**. The ERP system has quotation data that accumulates problems over time (gaps in date coverage, misaligned periods, quantity inconsistencies). The goal is to load a raw Excel export from the ERP, run automated analysis on it, and produce an output file with diagnosis columns appended — ready for human review and correction.

**Current state:** One Python script (`analyze.py`) that does fully deterministic analysis. No LLM yet. May add LLM layer later for complex/ambiguous cases.

---

## ERP Data Model

The Excel export is a flat file where each row is a **quotation line**. Header data is repeated on every row.

### Grouping Key
Groups are defined by: `(Quotation_No, Catalog_No)`

All analysis is done per group, not per individual line.

### Key Columns (see CONFIG section in analyze.py for exact names)

| Logical Name | Default Column | Description |
|---|---|---|
| Quotation_No | `Quotation_No` | Groups rows into quotation headers |
| Catalog_No | `Catalog_No` | Product; with Quotation_No defines a group |
| C_START_DATE | `C_START_DATE` | Line validity start date |
| C_END_DATE | `C_END_DATE` | Line validity end date |
| C_PRES_VALID_FROM | `C_PRES_VALID_FROM` | Header validity start (repeated per line) |
| C_PRES_VALID_TO | `C_PRES_VALID_TO` | Header validity end (repeated per line) |
| BUY_QTY_DUE | `BUY_QTY_DUE` | Quantity per line — used for qty pattern detection |
| STATE | `STATE` | Line state: released, cancelled, created, planned |
| C_PERIOD | `C_PERIOD` | "once" = single-use line, excluded from analysis |
| C_ORIG_PRES_LINE_DB | `C_ORIG_PRES_LINE_DB` | Bool: line gets +365 days on header expiry |
| CF_MDQ_PART_STA | `CF_MDQ_PART_STA` | Product status: US=active, O=obsolete |
| C_RENEWABLE_DB | `C_RENEWABLE_DB` | Bool: header renews by shifting dates (informational) |
| C_UNLIMIT_QTY_DB | `C_UNLIMIT_QTY_DB` | Bool: unlimited quantity, no limit (informational) |

---

## Business Rules (from mapping.docx)

### Active Lines
A line is **excluded** from pattern/coverage analysis if ANY of:
- `STATE == "cancelled"` — cancelled lines are never renewed
- Duration ≤ 1 day (`C_END_DATE - C_START_DATE <= 1`) — closed/placeholder line
- `C_PERIOD == "once"` — single-use, not part of repeating pattern

All other lines (released, created, planned) are **active**.

### Date Tolerance
All date comparisons use **5-day tolerance** to absorb administrative shifts.
This is defined as `TOLERANCE_DAYS = 5` in the script.

### Renewal Behaviour
When a header expires, +365 days is added to both header and lines. This means any misalignment or gap perpetuates forever unless corrected.

### Product Status
- `CF$_MDQ_PART_STA = "US"` → product exists on warehouse, renewal expected
- `CF$_MDQ_PART_STA = "O"` → product obsolete, lines usually cancelled, no renewal expected

---

## Data Quality Problems (from mapping.docx)

1. **Period inconsistency** — lines in a group should all follow the same period pattern (e.g. all quarterly). One line with a wrong duration breaks the pattern.

2. **Quantity inconsistency** — all lines in a group should have the same quantity. Outliers indicate data errors.

3. **Gaps in coverage** — some date ranges within the group's span are not covered by any line.

4. **Header-group misalignment** — the group's coverage window doesn't match the header validity window (e.g. group starts 30 days after header start, or ends 2 months before header end).

5. **Duplicate/overlapping lines** — two lines covering the same date range (fully or partially).

6. **Special cases** to always handle correctly:
   - Lines where `C_START_DATE == C_END_DATE` or differ by 1 day → closed placeholders, skip
   - `C_PERIOD == "once"` → skip from all checks
   - Cancelled lines → skip from all checks

---

## Analysis Design Decisions

### Why NOT use average/mean for period detection
If 3 lines are 90 days and 1 line is 30 days, mean = 75 days → wrong.
The correct answer is "quarterly (90 days)" with 1 outlier.

### Period Pattern: Bucket Voting
1. Compute duration in days for each active line
2. Map to nearest standard bucket (see below) within tolerance
3. Count votes per bucket — winning bucket = inferred pattern
4. Lines NOT in the winning bucket = outliers
5. Confidence = winning_votes / total_active_lines

**Standard buckets:**
```
monthly      = 30 days  ± 10
bi-monthly   = 60 days  ± 10
quarterly    = 90 days  ± 10   (covers real calendar quarters: 89–92 days)
4-month      = 120 days ± 10
semi-annual  = 180 days ± 12   (covers real half-years: 181–184 days)
annual       = 365 days ± 15
irregular    = anything else   (uses median of group's irregular durations)
```
Buckets are intentionally non-overlapping with gaps between them.

### Quantity Pattern: Mode (not mean)
Most frequent quantity among active lines = canonical quantity.
Lines deviating from it = outliers.
Confidence = mode_count / total_active_lines.

### Coverage: Three Metrics (not just max-min)
```
group_span_days      = max(end) - min(start) + 1   ← naive, ignores gaps
actual_coverage_days = interval union of all active lines  ← true coverage
internal_gap_days    = span - coverage              ← total uncovered days
```
Comparing these three reveals whether gaps exist and how large they are. Also exposes overlaps: if `sum_of_individual_durations > actual_coverage`, there are overlapping lines.

**Interval union algorithm:** sort intervals by start, merge adjacent/overlapping ones (adjacent = differ by 1 day), sum the merged intervals.

### Header Alignment
```
start_diff = group_start - header_start   (+ve = group starts late, -ve = early)
end_diff   = group_end   - header_end     (-ve = group ends early, +ve = ends late)
```
Within TOLERANCE_DAYS → "aligned". Outside → report direction and magnitude.

### Solution Suggestion (lines_to_add)
For each gap (internal + header-alignment):
- `ratio = gap_days / pattern_days`
- If `round(ratio)` within 25% tolerance → clean fit → suggest adding N lines
- Otherwise → "does not fit cleanly — manual review"
- Only calculated for non-irregular patterns

---

## What Has Been Built

### `analyze.py`
Core analysis engine. Config section at top with all column names as variables.
All functions accept an optional `col` dict so the Streamlit app can pass
UI-selected column names without touching the file.

Key exported functions:
- `analyze_dataframe(df, col=None)` — main entry point
- `get_default_col()` — returns column config dict from module-level constants
- `get_summary_stats(result_df, col=None)` — returns dict of summary metrics
- `GROUP_COLS`, `LINE_COLS` — lists of the 22 output column names

**CLI run:**
```bash
python analyze.py "path/to/export.xlsx"
# → produces path/to/export_analysis.xlsx
```

### `app.py`
Streamlit web app. Imports analysis logic from `analyze.py`.

**Run:**
```bash
streamlit run app.py
```

**UI layout:**
- **Sidebar**: file uploader + column mapping dropdowns (auto-selects defaults)
- **Main — Summary row**: 6 metric cards (total groups, gaps, misaligned, unclear pattern, qty issues, lines to add)
- **Tab 1 — All Lines**: full result dataframe with colour highlights
- **Tab 2 — Issues Only**: filterable view (checkboxes: gaps / misaligned / unclear period / qty inconsistency), red/yellow highlights
- **Tab 3 — Group Summary**: one row per group (deduped), separate download button
- **Download buttons**: full analysis Excel + group summary Excel

**Colour coding:**
- Red background: `has_gaps=YES`, `header_aligned=NO`, `lines_to_add > 0`, `period_confidence < 50%`
- Yellow background: `is_period_outlier=YES`, `is_qty_outlier=YES`, `period_confidence 50–70%`, `qty_confidence < 100%`

**22 new columns appended:**

Group-level (same value on all rows in group):
- `group_line_count` — total lines (all states)
- `group_active_line_count` — active lines only
- `inferred_period_pattern` — e.g. "quarterly"
- `inferred_period_days` — e.g. 90
- `period_confidence_pct` — % of active lines matching pattern
- `canonical_qty` — mode quantity
- `qty_confidence_pct` — % matching canonical qty
- `group_span_days` — naive max-min span
- `actual_coverage_days` — interval union
- `internal_gap_days` — uncovered days within span
- `internal_gap_count` — number of distinct gaps
- `has_gaps` — YES / NO
- `gap_details` — e.g. `"2023-08-01 → 2023-09-30 (61d)"`
- `header_aligned` — YES / NO (within tolerance)
- `start_alignment` — e.g. `"starts 30d late"` / `"aligned"`
- `end_alignment` — e.g. `"ends 30d early"` / `"aligned"`
- `lines_to_add` — integer or blank
- `solution_notes` — e.g. `"Add 1 quarterly line(s) [2023-08-01 → 2023-09-30]"`

Per-line (unique per row):
- `line_duration_days` — this line's duration
- `line_period_bucket` — its bucket (or "bucket (excluded)" for inactive)
- `is_period_outlier` — YES / NO / N/A
- `is_qty_outlier` — YES / NO / N/A

---

## What Is NOT Built Yet (future scope)

1. **Action strategy** — deciding between: add lines, shift header, close & recreate. Intentionally skipped for now to keep scope focused on analysis/diagnosis.

2. **LLM layer** — for ambiguous cases (pattern_confidence < 50%, only 1 active line, all durations differ). Would send precomputed facts + raw lines to Claude API. Key principle: Python does all arithmetic, LLM only interprets.

3. **Streamlit UI** — the user chose CLI script for now. Could add later.

4. **ERP import file generation** — outputting corrected data back into ERP format. Not in scope yet.

---

## Important: First Thing When User Loads a New File

Check `COL_QTY` in the CONFIG section of `analyze.py` — this defaults to `"QTY"` but the actual column name in the user's ERP export may differ. All other column names default to the names documented in `mapping.docx`.

---

## Files in This Directory

| File | Purpose |
|---|---|
| `analyze.py` | Core analysis engine (CLI + importable library) |
| `app.py` | Streamlit web app — upload, analyse, explore, download |
| `requirements.txt` | Python dependencies (pandas, openpyxl, xlrd, streamlit) |
| `mapping.docx` | Original ERP documentation: data model, business rules, problem descriptions |
| `CLAUDE.md` | This file — project context for Claude sessions |
