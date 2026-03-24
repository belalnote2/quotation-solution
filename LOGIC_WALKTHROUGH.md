# analyze.py — Full Logic Walkthrough

**How to use this file:**
Every section shows three things in this order:
1. The **actual Python code** from the file (in a `python` code block)
2. A **plain-English explanation** of what that code does
3. A `> REVIEWER NOTE` where I flag something worth verifying or questioning

You do not need to understand Python to review the logic. Read the explanation,
then glance at the code to confirm the explanation matches what you see.

---

## Architecture: How the Pieces Fit Together

```
analyze.py
│
├── CONFIG (top of file)
│   ├── Column name constants (COL_QUOTATION_NO, COL_START_DATE, ...)
│   ├── TOLERANCE_DAYS = 5
│   └── PERIOD_BUCKETS list
│
├── get_default_col()
│   └── Packages the constants into a dict so the Streamlit app can override them
│
├── HELPER FUNCTIONS (small, reusable tools)
│   ├── is_line_active(row, col)          ← should this line be analysed?
│   ├── map_to_bucket(duration_days)      ← what period category is this line?
│   ├── compute_interval_union(intervals) ← how many days are truly covered?
│   ├── detect_gaps(sorted_active_df, col)← find the uncovered date ranges
│   ├── build_coverage_bar(...)           ← build the visual ████░░░ bar
│   └── solve_gap(gap_days, pattern_days) ← how many lines fill a gap?
│
├── analyze_group(group_df, header_start, header_end, col)   ← THE CORE
│   │   Called once per (Quotation_No, Catalog_No) group
│   │   Returns: gm (group metrics dict) + per_line (per-row dict)
│   │
│   ├── Step 1:  Count lines (total + active + flag counts)
│   ├── Step 1b: Handle renewable lines (null dates → fill from header)
│   ├── Step 2:  Period pattern — bucket voting → winning pattern
│   ├── Step 3:  Quantity pattern — mode voting → canonical qty
│   ├── Step 4:  Coverage analysis (span vs actual vs gap)
│   ├── Step 5:  Gap detection + overlap calculation
│   ├── Step 6:  Header alignment (does group match header window?)
│   ├── Step 7:  Lines to add (solution suggestion per gap)
│   └── Step 8:  Per-line outlier flags (period + qty)
│
├── analyze_dataframe(df, col=None)       ← ORCHESTRATOR
│   │   Main entry point. Loops over all groups, writes results back.
│   └── Calls analyze_group() for every group
│
├── get_summary_stats(result_df, col=None)← SUMMARY NUMBERS
│   └── Counts groups with issues for the dashboard metric cards
│
└── CLI entry point (__main__)
    └── Reads Excel → analyze_dataframe() → writes Excel output
```

**Data flow in one sentence:**
The orchestrator splits the flat Excel file into groups, sends each group to `analyze_group`,
gets back a dict of results, and stamps those results as new columns on every row of the group.

**Who calls what:**
```
app.py (Streamlit UI)
    └── calls analyze_dataframe()
            └── calls analyze_group()  [once per group]
                    ├── calls is_line_active()
                    ├── calls map_to_bucket()
                    ├── calls compute_interval_union()
                    ├── calls detect_gaps()
                    ├── calls build_coverage_bar()
                    └── calls solve_gap()
```

---

## Part 1: Configuration

### Actual code (lines 33–66)

```python
COL_QUOTATION_NO  = "Quotation_No"
COL_CATALOG_NO    = "Catalog_No"
COL_LINE_NO       = "Line_No"
COL_REL_NO        = "Rel_No"
COL_START_DATE    = "C_START_DATE"
COL_END_DATE      = "C_END_DATE"
COL_HEADER_START  = "C_PRES_VALID_FROM"
COL_HEADER_END    = "C_PRES_VALID_TO"
COL_QTY           = "BUY_QTY_DUE"
COL_STATE         = "STATE"
COL_PERIOD        = "C_PERIOD"
COL_ORIG_DB       = "C_ORIG_PRES_LINE_DB"
COL_MDQ_STATUS    = "CF_MDQ_PART_STA"
COL_RENEWABLE     = "C_RENEWABLE_DB"
COL_UNLIMIT_QTY   = "C_UNLIMIT_QTY_DB"

TOLERANCE_DAYS = 5

PERIOD_BUCKETS = [
    ("monthly",      30,  10),   # 20–40 days
    ("bi-monthly",   60,  10),   # 50–70 days
    ("quarterly",    90,  10),   # 80–100 days
    ("4-month",     120,  10),   # 110–130 days
    ("semi-annual", 180,  12),   # 168–192 days
    ("annual",      365,  15),   # 350–380 days
]
```

### What this does

These are just named constants — the column names as they appear in the ERP Excel export.
Instead of writing `"C_START_DATE"` everywhere in the code, it writes `COL_START_DATE`.
If the column name ever changes in the ERP, you only update it in one place here.

`TOLERANCE_DAYS = 5` is a single threshold used everywhere for "close enough":
gaps smaller than 5 days are ignored, and alignment differences smaller than 5 days are OK.

`PERIOD_BUCKETS` is a list of `(name, target_days, tolerance_days)` tuples.
Each line's duration is checked against these — whichever bucket it falls into is its category.
The tolerances are intentionally non-overlapping: a 105-day line fits no bucket and is "irregular."

> **REVIEWER NOTE — Tolerance ±10/12/15 days on buckets:**
> For a quarterly pattern, ±10 means anything from 80 to 100 days is called "quarterly."
> Real calendar quarters are 89–92 days, so this is intentionally generous.
> If your data has lines that are consistently 105 days, they fall into NO bucket and are
> called "irregular." Check the `line_period_bucket` column in your output to spot this.

> **REVIEWER NOTE — TOLERANCE_DAYS = 5 is used for three separate checks:**
> (1) gap detection, (2) header alignment, (3) solve_gap fit check.
> All three use the same 5-day value. This may be correct for your business, but
> it is worth asking: should a 4-day gap in coverage really be silently ignored?

---

## Part 2: `get_default_col()` — column name dictionary

### Actual code (lines 73–96)

```python
def get_default_col():
    return {
        # ── Used in analysis ──────────────────────────────────────
        "quotation_no":  COL_QUOTATION_NO,
        "catalog_no":    COL_CATALOG_NO,
        "start_date":    COL_START_DATE,
        "end_date":      COL_END_DATE,
        "header_start":  COL_HEADER_START,
        "header_end":    COL_HEADER_END,
        "qty":           COL_QTY,
        "state":         COL_STATE,
        "period":        COL_PERIOD,
        # ── Informational (mapped, shown in output, not yet in logic) ─
        "renewable":     COL_RENEWABLE,
        "unlimit_qty":   COL_UNLIMIT_QTY,
        "orig_pres":     COL_ORIG_DB,
        "mdq_status":    COL_MDQ_STATUS,
    }
```

### What this does

Packages the column name constants into a dictionary with short logical keys.
The analysis functions always use keys like `col["start_date"]` instead of the raw string
`"C_START_DATE"`. When the Streamlit app runs, it builds a different version of this dict
from the UI dropdowns, so the user can select different column names without touching the code.

> **REVIEWER NOTE:**
> The four columns at the bottom (`renewable`, `unlimit_qty`, `orig_pres`, `mdq_status`)
> are mapped here but not used in any calculations — they are "informational."
> `orig_pres` is counted per group but never affects gap/period/qty logic.
> `mdq_status` (US=active, O=obsolete) is not used at all in the analysis.

---

## Part 3: `is_line_active()` — should this line be included?

### Actual code (lines 103–135)

```python
def is_line_active(row, col):
    # Cancelled state
    try:
        if str(row[col["state"]]).strip().lower() == "cancelled":
            return False
    except (KeyError, TypeError):
        pass

    # "Once" period — single-use line, not part of repeating pattern
    try:
        if str(row[col["period"]]).strip().lower() == "once":
            return False
    except (KeyError, TypeError, AttributeError):
        pass

    # Closed / placeholder line (duration ≤ 1 day)
    try:
        start = row[col["start_date"]]
        end   = row[col["end_date"]]
        if pd.notna(start) and pd.notna(end):
            if (end - start).days < 4:
                return False
        # Null dates → renewable line, keep active
    except (KeyError, TypeError, AttributeError):
        pass

    return True
```

### What this does

For every row, decides whether to include it in the analysis or skip it.
Three reasons to skip a line:
1. It is cancelled — cancelled lines are never renewed, so analysing them makes no sense
2. It has `C_PERIOD = "once"` — a one-time-only line, not part of a repeating pattern
3. Its date range is less than 4 days — treated as a closed placeholder/stub, not a real line

If dates are missing (null), the line is kept as active — it is treated as a "renewable" line
that covers the whole header period when it renews.

The `try/except` blocks around each check mean: if the column is missing from the data
entirely, that check is simply skipped rather than crashing.

> **REVIEWER NOTE — Comment says "≤ 1 day" but code checks `< 4`:**
> The function's own comment says "duration ≤ 1 day." The actual check is `(end - start).days < 4`.
> That means durations of 1, 2, and 3 days are ALL excluded — not just 1-day lines.
> The documentation says the same wrong thing. This is a real discrepancy.
> Do you have any 2-day or 3-day lines in your data that should be analysed?
> If yes, this threshold needs to be corrected to `<= 1`.

---

## Part 4: `map_to_bucket()` — what period category is this line?

### Actual code (lines 138–147)

```python
def map_to_bucket(duration_days):
    for name, target, tol in PERIOD_BUCKETS:
        if abs(duration_days - target) <= tol:
            return name, target
    return "irregular", duration_days
```

### What this does

Takes a number (the line's duration in days) and checks it against each bucket in order.
If `|duration - target| ≤ tolerance` → that bucket matches, return its name and target.
If nothing matches → return `"irregular"` with the raw duration as the target.

Example: a line of 91 days → `|91 - 90| = 1 ≤ 10` → returns `("quarterly", 90)`.
Example: a line of 105 days → misses quarterly (|105-90|=15 > 10) AND misses 4-month
(|105-120|=15 > 10) → returns `("irregular", 105)`.

> **REVIEWER NOTE:**
> The loop checks buckets in order (monthly first, annual last).
> Because the bucket ranges are non-overlapping (no duration fits two buckets),
> the order does not affect correctness. This is safe.

---

## Part 5: `compute_interval_union()` — how many days are truly covered?

### Actual code (lines 150–172)

```python
def compute_interval_union(intervals):
    if not intervals:
        return 0, []

    sorted_ivs = sorted(intervals, key=lambda x: x[0])
    merged = [[sorted_ivs[0][0], sorted_ivs[0][1]]]

    for start, end in sorted_ivs[1:]:
        # +1 day: adjacent intervals (e.g. Mar 31 / Apr 1) are merged
        if start <= merged[-1][1] + timedelta(days=1):
            if end > merged[-1][1]:
                merged[-1][1] = end
        else:
            merged.append([start, end])

    total = sum((e - s).days + 1 for s, e in merged)
    return total, merged
```

### What this does

Given a list of `(start_date, end_date)` pairs, calculates the total number of days
that are covered by at least one interval — removing double-counting from overlapping ranges.

Step 1: Sort all ranges by start date.
Step 2: Walk through in order. If the next range overlaps with (or is adjacent to)
the current merged range, extend it. If there is a real gap, start a new merged range.
The `+ timedelta(days=1)` means "adjacent" lines (Mar 31 end / Apr 1 start) are merged —
they are treated as continuous, not as a gap.
Step 3: Sum the days in all merged ranges.

Example:
```
Input:  [Jan 1–Mar 31], [Feb 15–Apr 30], [May 1–Jun 30]
After sort: same order (already sorted)
Merge 1: start=[Jan 1, Mar 31]. Next: Feb 15 ≤ Mar 31+1 → extend to Apr 30 → [Jan 1, Apr 30]
Merge 2: May 1 = Apr 30 + 1 day → adjacent → extend to Jun 30 → [Jan 1, Jun 30]
Total: 181 days (Jan 1 to Jun 30 inclusive)
```

> **REVIEWER NOTE — Adjacent day merge:**
> Two lines where one ends Mar 31 and the next starts Apr 1 are counted as continuous coverage.
> This is the correct behaviour for quotation lines. The only case where this could hide a
> problem is if your ERP uses a deliberate 1-day gap as a marker — but that would be unusual.

---

## Part 6: `detect_gaps()` — find the actual uncovered date ranges

### Actual code (lines 175–206)

```python
def detect_gaps(sorted_active_df, col):
    gaps = []
    rows = list(sorted_active_df.iterrows())
    if not rows:
        return gaps

    running_end = rows[0][1][col["end_date"]]

    for i in range(1, len(rows)):
        _, curr = rows[i]
        gap_days = (curr[col["start_date"]] - running_end).days - 1
        if gap_days > TOLERANCE_DAYS:
            gaps.append({
                "gap_start": running_end + timedelta(days=1),
                "gap_end":   curr[col["start_date"]] - timedelta(days=1),
                "gap_days":  gap_days,
            })
        if curr[col["end_date"]] > running_end:
            running_end = curr[col["end_date"]]

    return gaps
```

### What this does

Finds the actual date ranges that are NOT covered by any active line.

Starts from the first line's end date and walks forward through each line in date order.
For each line, checks: how many days are between the previous coverage end and this line's start?
If that gap is more than `TOLERANCE_DAYS` (5 days), it is recorded as a real gap.

The key trick is `running_end` — it tracks the furthest end date seen so far, not just the
previous line's end. This handles the case where one line contains another:

```
Line A: Jan 1 → Sep 26
Line B: Feb 1 → Aug 24   ← contained inside A, ends earlier
Line C: Sep 27 → Dec 31  ← should NOT be a gap after B

Without running_end: gap_days = Sep 27 - Aug 24 - 1 = 33 days → FALSE gap reported
With running_end:    running_end = Sep 26 (from A), gap_days = Sep 27 - Sep 26 - 1 = 0 → correct
```

> **REVIEWER NOTE — This is well implemented.**
> The running max end date is the correct approach and handles overlapping lines properly.
> No issues here.

---

## Part 7: `build_coverage_bar()` — the visual timeline

### Actual code (lines 209–268)

```python
def build_coverage_bar(bar_start, bar_end, intervals, width=48):
    if bar_start is None or bar_end is None:
        return ""
    total_days = (bar_end - bar_start).days + 1
    if total_days <= 0:
        return ""
    if not intervals:
        return "░" * width

    sorted_ivs = sorted(intervals, key=lambda x: x[0])
    bar          = []
    prev_interval = None

    for i in range(width):
        slot_s_off = int(i       * total_days / width)
        slot_e_off = max(slot_s_off, int((i + 1) * total_days / width) - 1)
        slot_s = bar_start + timedelta(days=slot_s_off)
        slot_e = bar_start + timedelta(days=slot_e_off)

        covering = [j for j, (s, e) in enumerate(sorted_ivs)
                    if s <= slot_e and e >= slot_s]

        if len(covering) == 0:
            bar.append("░")           # gap
            prev_interval = None
        elif len(covering) >= 2:
            bar.append("▓")           # overlap
            prev_interval = covering[0]
        else:
            j = covering[0]
            if prev_interval is not None and prev_interval != j:
                bar.append("|")       # boundary between two adjacent lines
            else:
                bar.append("█")       # covered
            prev_interval = j

    return "".join(bar)
```

### What this does

Produces a 48-character string like `████████████|████████████░░░░░████████████`
that shows, proportionally across time, which parts of the quotation are covered.

Divides the full header date window into 48 equal time slots.
For each slot, checks how many active line intervals cover it:
- 0 intervals → `░` (gap — nothing here)
- 2+ intervals → `▓` (overlap — more than one line covers this period)
- 1 interval, same as previous slot → `█` (solid covered)
- 1 interval, DIFFERENT from previous → `|` (boundary between two lines)

The bar spans the header period (not just the group's own dates), so a late-starting
group shows `░░░░████████` with gaps at the left edge.

> **REVIEWER NOTE — Visual only, no impact on numbers.**
> This function does not affect any calculated values. It is purely for display.
> The `|` pipe character lets you count individual periods visually.
> For a 1-year quarterly group you should see: `████████████|████████████|████████████|████████████`

---

## Part 8: `solve_gap()` — how many lines are needed to fill a gap?

### Actual code (lines 271–291)

```python
def solve_gap(gap_days, pattern_days):
    if not pattern_days or pattern_days <= 0:
        return None
    ratio = gap_days / pattern_days
    n = round(ratio)
    if n < 1:
        return None
    if abs(gap_days - n * pattern_days) <= TOLERANCE_DAYS:
        return n
    return None
```

### What this does

Given a gap size (in days) and the group's pattern (e.g. 90 days for quarterly),
calculates whether the gap can be filled cleanly by adding whole lines.

```
ratio = gap_days / pattern_days    e.g.  182 / 90 = 2.02
n     = round(ratio)               e.g.  round(2.02) = 2
check = |gap_days - n × pattern|   e.g.  |182 - 180| = 2  ≤ 5 → clean fit → return 2
```

If the gap does not fit within `TOLERANCE_DAYS` of a whole number of periods → return None
(the gap is not solvable automatically and gets marked as `✗` in the output).

> **REVIEWER NOTE — Documentation says "25% tolerance" but code uses 5 days:**
> The project documentation says "within 25% tolerance." The actual code uses
> `TOLERANCE_DAYS = 5` (a fixed 5-day window, not a percentage).
> For a 90-day pattern: 25% = 22.5 days tolerance, but the code only allows 5 days.
> This means the code is much stricter than documented. A gap of 87 days would fail
> the 5-day check (|87 - 90| = 3, actually passes), but a gap of 83 days would fail
> (|83 - 90| = 7 > 5 → "✗"). The documentation is misleading here.

---

## Part 9: `analyze_group()` — the core (8 steps)

This is the most important function. Everything else feeds into it.
It is called once for every `(Quotation_No, Catalog_No)` group.

---

### Step 1 — Count lines

#### Actual code (lines 316–340)

```python
total_lines = len(group_df)
active_mask = group_df.apply(lambda row: is_line_active(row, col), axis=1)
active_df   = group_df[active_mask].sort_values(col["start_date"])
n_active    = len(active_df)

gm["group_line_count"]        = total_lines
gm["group_active_line_count"] = n_active

_TRUTHY = {"1", "true", "yes", "y", "x"}
for _key, _out_col in [
    ("unlimit_qty", "unlimit_qty_count"),
    ("orig_pres",   "orig_pres_count"),
]:
    _src = col.get(_key, "")
    if _src and _src in group_df.columns:
        gm[_out_col] = int(sum(
            1 for v in group_df[_src]
            if not pd.isna(v) and str(v).strip().lower() in _TRUTHY
        ))
    else:
        gm[_out_col] = 0
```

#### What this does

Calls `is_line_active()` on every row in the group, builds a True/False mask,
and creates a filtered `active_df` containing only the lines that should be analysed.
`active_df` is sorted by `start_date` — this sorted order is important for gap detection later.

Then counts boolean flags: for `C_UNLIMIT_QTY_DB` and `C_ORIG_PRES_LINE_DB`,
counts how many lines in the group have those flags set to any truthy value
(the ERP may store these as `1`, `"Y"`, `"true"`, etc. — all handled).

---

### Step 1b — Handle renewable lines (null dates)

#### Actual code (lines 355–388)

```python
h_start = header_start if pd.notna(header_start) else None
h_end   = header_end   if pd.notna(header_end)   else None

fully_infinite_set = {
    idx for idx, row in active_df.iterrows()
    if pd.isna(row[col["start_date"]]) and pd.isna(row[col["end_date"]])
}
infinite_set = {
    idx for idx, row in active_df.iterrows()
    if pd.isna(row[col["start_date"]]) or pd.isna(row[col["end_date"]])
}
has_infinite         = bool(infinite_set)
has_dated            = any(i not in infinite_set for i in active_df.index)
has_infinite_overlap = bool(fully_infinite_set) and has_dated

eff_df = active_df.copy()
if has_infinite:
    for idx in infinite_set:
        if pd.isna(eff_df.at[idx, col["start_date"]]) and h_start is not None:
            eff_df.at[idx, col["start_date"]] = h_start
        if pd.isna(eff_df.at[idx, col["end_date"]]) and h_end is not None:
            eff_df.at[idx, col["end_date"]]   = h_end
    # Drop lines that still have a null (no header date available to fill)
    still_null = [
        idx for idx in infinite_set
        if pd.isna(eff_df.at[idx, col["start_date"]])
        or pd.isna(eff_df.at[idx, col["end_date"]])
    ]
    if still_null:
        eff_df = eff_df.drop(still_null)
eff_df = eff_df.sort_values(col["start_date"])
```

#### What this does

Some lines have null (missing) dates — these are "renewable" lines whose dates are
managed by the ERP at renewal time. For analysis purposes, the code fills in the header
dates as substitutes:
- Both dates null → "fully renewable" → use `header_start` and `header_end`
- Only start null → "open-start" → use `header_start` for the start
- Only end null → "open-end" → use `header_end` for the end

`eff_df` (effective dataframe) is the result: same as `active_df` but with nulls filled.
Lines that still have nulls after filling (because the header dates are also null) are dropped.

`has_infinite_overlap` is flagged when a fully-renewable line AND explicit dated lines exist
in the same group — meaning one line covers the whole year AND there are more lines on top.

> **REVIEWER NOTE — `has_infinite_overlap` is computed but never written to the output:**
> This flag is calculated here but never added to `gm` (the group metrics dict).
> It does not appear in the Excel output. If you want to see which groups have this
> situation, it would need to be added as a new column.

---

### Step 2 — Period pattern (bucket voting)

#### Actual code (lines 478–535)

```python
if not active_buckets:
    winning_bucket = "irregular"
    winning_days   = 0
    winning_count  = n_active
else:
    bucket_names = [v[0] for v in active_buckets.values()]
    vote         = Counter(bucket_names)
    winning_bucket, winning_count = vote.most_common(1)[0]

    # Tiebreaker: pick the bucket whose lines cover the most total days
    top_count     = winning_count
    tied_buckets  = [b for b, c in vote.items() if c == top_count]
    if len(tied_buckets) > 1:
        bucket_total_days = {}
        for idx, (bname, _) in active_buckets.items():
            if bname in tied_buckets:
                bucket_total_days[bname] = (
                    bucket_total_days.get(bname, 0) + (active_durations.get(idx) or 0)
                )
        winning_bucket = max(bucket_total_days, key=bucket_total_days.get)

    if winning_bucket == "irregular":
        irr_durs = sorted(
            d for idx, d in active_durations.items()
            if active_buckets[idx][0] == "irregular"
        )
        winning_days = irr_durs[len(irr_durs) // 2] if irr_durs else 0
    else:
        winning_days = next(
            target for name, target, _ in PERIOD_BUCKETS if name == winning_bucket
        )

gm["inferred_period_pattern"] = winning_bucket
gm["inferred_period_days"]    = winning_days
gm["period_confidence_pct"]   = round(winning_count / n_active * 100)
```

#### What this does

Each active line has already been assigned a bucket (quarterly, monthly, etc.) in the
per-line loop that precedes this block. This step counts the votes:

```
e.g. 4 lines: quarterly, quarterly, quarterly, monthly
→ vote = {quarterly: 3, monthly: 1}
→ winning_bucket = "quarterly", winning_count = 3
→ period_confidence_pct = 3/4 × 100 = 75%
```

**Tie-break:** if two buckets have equal votes, pick the one whose lines cover more total days.
Rationale: a single 365-day line represents the pattern better than a single 30-day line.

**Irregular groups:** if the winning bucket is "irregular," `winning_days` is set to the
median of the irregular durations (not the mean — median is less affected by outliers).

> **REVIEWER NOTE — Single active line always gets 100% confidence:**
> 1 vote out of 1 = 100%. Mathematically correct but misleading — you cannot determine
> a pattern from one data point. These groups show as "clear" (green) in the app
> even though nothing was actually verified.

> **REVIEWER NOTE — `avg_period_days` is stored separately as a simple mean:**
> `inferred_period_days` is always the bucket target (e.g. exactly 90 for quarterly).
> `avg_period_days` is the real measured mean of actual durations.
> If avg=75 but inferred=90, the bucket voting may be misclassifying lines.
> Always compare both columns when reviewing uncertain groups.

---

### Step 3 — Quantity pattern (mode voting)

#### Actual code (lines 539–582)

```python
try:
    active_qtys = [
        group_df.loc[idx, col["qty"]]
        for idx in active_df.index
        if pd.notna(group_df.loc[idx, col["qty"]])
    ]
except KeyError:
    active_qtys = []

if active_qtys:
    qty_vote      = Counter(active_qtys)
    top_qty_count = qty_vote.most_common(1)[0][1]
    # Tie-breaking: if two quantities share the top count, pick the larger value.
    tied_qtys     = [q for q, c in qty_vote.items() if c == top_qty_count]
    canonical_qty = max(tied_qtys)
    qty_count     = top_qty_count
    gm["canonical_qty"]      = canonical_qty
    gm["qty_confidence_pct"] = round(qty_count / n_active * 100)
else:
    canonical_qty = None
    gm["canonical_qty"]      = NA
    gm["qty_confidence_pct"] = NA
```

#### What this does

Finds the most frequent quantity among active lines (the statistical mode, not the average).

```
e.g. 4 lines: qty=100, qty=100, qty=100, qty=50
→ vote = {100: 3, 50: 1}
→ canonical_qty = 100
→ qty_confidence_pct = 3/4 × 100 = 75%
```

The `try/except KeyError` handles the case where the quantity column does not exist in the data.

**Tie-break:** if two quantities have equal vote counts, the larger one wins.

> **REVIEWER NOTE — Tie-break picks the LARGER quantity:**
> If lines are 50/50 between qty=10 and qty=20, the code calls 20 "canonical"
> and flags qty=10 lines as outliers. If qty=10 was actually a recent correction,
> those lines are wrongly flagged. Tie cases in qty are worth reviewing manually.

---

### Step 4 — Coverage analysis

#### Actual code (lines 584–602)

```python
if eff_df.empty:
    group_start = h_start if h_start is not None else NA
    group_end   = h_end   if h_end   is not None else NA
    _span, _coverage, _gap_days, intervals = 0, 0, 0, []
else:
    group_start  = eff_df[col["start_date"]].min()
    group_end    = eff_df[col["end_date"]].max()
    _span        = (group_end - group_start).days + 1
    intervals    = list(zip(eff_df[col["start_date"]], eff_df[col["end_date"]]))
    _coverage, _ = compute_interval_union(intervals)
    _gap_days    = _span - _coverage

gm["group_start"]          = group_start.date() if isinstance(group_start, pd.Timestamp) else group_start
gm["group_end"]            = group_end.date()   if isinstance(group_end,   pd.Timestamp) else group_end
gm["group_span_days"]      = _span
gm["actual_coverage_days"] = _coverage
```

#### What this does

```
group_start = earliest start date of any active line
group_end   = latest end date of any active line
group_span  = (group_end - group_start) + 1  ← the full window from first to last date
                                               ← this INCLUDES any gaps inside

actual_coverage = compute_interval_union(all active line ranges)
                ← this EXCLUDES gaps — only counts truly covered days

gap_days = group_span - actual_coverage
         ← if > 0: there are uncovered days within the group's own window
         ← if = 0: coverage is continuous (no internal gaps)
```

> **REVIEWER NOTE — `gap_days` does not include the header-edge gaps:**
> If the group starts 30 days after the header, that 30-day edge gap is NOT counted here.
> `gap_days` only counts gaps within the group's own start-to-end window.
> The header-edge gaps are handled separately in Steps 6 and 7.
> The `lines_to_add` column does combine all three types.

---

### Step 5 — Gap detection and overlap

#### Actual code (lines 604–645)

```python
gaps             = detect_gaps(eff_df, col) if not eff_df.empty else []
gm["gap_days"]   = _gap_days
gm["gap_count"]  = len(gaps)

if intervals:
    _sum_durs = sum((e - s).days + 1 for s, e in intervals)
    _overlap  = max(0, _sum_durs - _coverage)
    _sorted_ivs   = sorted(intervals, key=lambda x: x[0])
    _running_end  = _sorted_ivs[0][1]
    _overlap_count = 0
    for _s, _e in _sorted_ivs[1:]:
        if _s <= _running_end:
            _overlap_count += 1
        if _e > _running_end:
            _running_end = _e
else:
    _overlap = 0
    _overlap_count = 0
gm["overlap_days"]  = _overlap
gm["overlap_count"] = _overlap_count

if gaps:
    gm["gap_details"] = " | ".join(
        f"{g['gap_start'].date()} → {g['gap_end'].date()} ({g['gap_days']}d)"
        for g in gaps
    )
else:
    gm["gap_details"] = ""

_bar_s = h_start if h_start is not None else (group_start if isinstance(group_start, pd.Timestamp) else None)
_bar_e = h_end   if h_end   is not None else (group_end   if isinstance(group_end,   pd.Timestamp) else None)
gm["coverage_bar"] = build_coverage_bar(_bar_s, _bar_e, intervals)
```

#### What this does

Calls `detect_gaps()` to get the list of actual gap date ranges (already explained in Part 6).

**Overlap calculation:**
```
sum_durs  = add up each individual line's duration
overlap   = sum_durs - actual_coverage
```
If `sum_durs > actual_coverage`, some date range is covered by more than one line.
The difference is the total double-covered days.

`overlap_count` counts how many lines start before the running maximum end date of the
lines seen so far — i.e., how many lines begin while an earlier line is still open.

The coverage bar spans the header dates (if available) so edge gaps are visible at the
left/right of the bar.

> **REVIEWER NOTE — `overlap_count` counts lines, not pairs:**
> If lines A, B, C all overlap each other, overlap_count = 2 (B and C both start
> before A's end). It is not 3 (the number of overlapping pairs A-B, A-C, B-C).
> Fine as a signal, just be aware of the definition when reading the output.

---

### Step 6 — Header alignment

#### Actual code (lines 648–682)

```python
header_dates_available = pd.notna(header_start) and pd.notna(header_end)
group_dates_available  = (
    not isinstance(group_start, str) and not isinstance(group_end, str)
)

if header_dates_available and group_dates_available:
    start_diff = (group_start - header_start).days
    end_diff   = (group_end   - header_end).days

    start_ok = abs(start_diff) <= TOLERANCE_DAYS
    end_ok   = abs(end_diff)   <= TOLERANCE_DAYS

    gm["header_aligned"] = "YES" if (start_ok and end_ok) else "NO"

    if start_ok:
        gm["start_alignment"] = "aligned"
    elif start_diff > 0:
        gm["start_alignment"] = f"starts {start_diff}d late"
    else:
        gm["start_alignment"] = f"starts {abs(start_diff)}d early"

    if end_ok:
        gm["end_alignment"] = "aligned"
    elif end_diff < 0:
        gm["end_alignment"] = f"ends {abs(end_diff)}d early"
    else:
        gm["end_alignment"] = f"ends {end_diff}d late"
else:
    start_ok = end_ok = True   # skip header gap checks in solution step
    start_diff = end_diff = 0
    gm["header_aligned"]  = NA
    gm["start_alignment"] = NA
    gm["end_alignment"]   = NA
```

#### What this does

Checks whether the group's actual date coverage matches the header's declared validity window.

```
start_diff = group_start - header_start
    positive → group starts LATE  → gap at beginning
    negative → group starts EARLY → extends before header (unusual)

end_diff = group_end - header_end
    negative → group ends EARLY   → gap at end
    positive → group ends LATE    → extends after header (unusual)
```

If both diffs are within 5 days → `header_aligned = "YES"`.
Otherwise → `"NO"` with a human-readable label: `"starts 30d late"`, `"ends 45d early"`.

If either header dates or group dates are unavailable → all alignment fields set to `"N/A"`.

> **REVIEWER NOTE — This only checks the outer boundary, not internal gaps:**
> `header_aligned = "YES"` means the group's first line starts near the header start
> AND the group's last line ends near the header end. A group with perfect alignment
> but a 3-month internal gap will still show `header_aligned = "YES"`.
> Always check `gap_count` alongside `header_aligned`.

---

### Step 7 — Lines to add (solution suggestion)

#### Actual code (lines 684–740)

```python
_all_gaps = []

# Internal gaps
for g in gaps:
    _all_gaps.append({
        "label":    f"{g['gap_start'].date()} → {g['gap_end'].date()} ({g['gap_days']}d)",
        "gap_days": g["gap_days"],
    })

# Start gap: group begins after header start
if not start_ok and start_diff > TOLERANCE_DAYS:
    _all_gaps.append({
        "label":    f"{_gs} → {_ge} ({start_diff}d) [before group]",
        "gap_days": start_diff,
    })

# End gap: group ends before header end
if not end_ok and end_diff < -TOLERANCE_DAYS:
    _all_gaps.append({
        "label":    f"{_gs} → {_ge} ({abs(end_diff)}d) [after group]",
        "gap_days": abs(end_diff),
    })

total_to_add    = 0
solved_count    = 0
gap_labels      = []
solution_labels = []

for _gap_item in _all_gaps:
    gap_labels.append(_gap_item["label"])
    if winning_bucket != "irregular":
        _n = solve_gap(_gap_item["gap_days"], winning_days)
        if _n:
            total_to_add += _n
            solved_count += 1
            solution_labels.append(f"+{_n} {winning_bucket}")
        else:
            solution_labels.append("✗")
    else:
        solution_labels.append("✗")

gm["gaps_solved_ratio"] = f"{_total_gaps}gap/{total_to_add}l" if _total_gaps > 0 else ""
gm["gap_list"]          = " | ".join(gap_labels)
gm["solution_list"]     = " | ".join(solution_labels)
gm["lines_to_add"]      = total_to_add if total_to_add > 0 else ("no gap" if _total_gaps == 0 else 0)
```

#### What this does

Builds a combined list of ALL gaps: internal gaps + the two possible header-edge gaps
(start gap if group starts late, end gap if group ends early).

For each gap, calls `solve_gap()` to see if it fits a whole number of the inferred pattern.
If it fits → record `"+N quarterly"` (or whatever pattern).
If it does not fit, or if the pattern is irregular → record `"✗"`.

The parallel lists `gap_list` and `solution_list` are stored as pipe-separated strings
so you can read them in the Excel output column-by-column.

`lines_to_add` is the sum of all N values from the solvable gaps.
`"no gap"` is stored (not 0) when there are no gaps at all, so you can distinguish
"nothing to fix" from "gap exists but we cannot auto-solve it."

> **REVIEWER NOTE — "✗" entries are silent — no separate flag:**
> If a group has 3 gaps and only 1 is solvable, `lines_to_add = N` (just the solvable one).
> The other 2 are `"✗"` in `solution_list` but there is no dedicated flag column
> to filter on "has unsolvable gaps." You need to filter on `solution_list` containing `"✗"`.

---

### Step 8 — Per-line outlier flags

#### Actual code (lines 742–773)

```python
for idx in group_df.index:
    is_active_line = idx in active_df.index

    # Period outlier
    if is_active_line:
        if idx in active_buckets:
            bname, _ = active_buckets[idx]
            per_line[idx]["is_period_outlier"] = (
                "YES" if bname != winning_bucket else "NO"
            )
        else:
            per_line[idx]["is_period_outlier"] = NA
    else:
        per_line[idx]["is_period_outlier"] = NA

    # Quantity outlier
    if not is_active_line:
        per_line[idx]["is_qty_outlier"] = NA
    elif canonical_qty is None:
        per_line[idx]["is_qty_outlier"] = NA
    else:
        try:
            row_qty = group_df.loc[idx, col["qty"]]
            per_line[idx]["is_qty_outlier"] = (
                "YES" if row_qty != canonical_qty else "NO"
            )
        except Exception:
            per_line[idx]["is_qty_outlier"] = NA
```

#### What this does

For every row in the group (including inactive ones), assigns the outlier flags.

**Period outlier:**
- Inactive line → `"N/A"` (we never analysed it, cannot call it an outlier)
- Active line with no effective bucket → `"N/A"` (renewable line, no header dates to compare)
- Active line whose bucket matches the winning bucket → `"NO"`
- Active line whose bucket differs → `"YES"` ← this line breaks the group's pattern

**Quantity outlier:**
- Inactive line → `"N/A"`
- Active line, qty = canonical_qty → `"NO"`
- Active line, qty ≠ canonical_qty → `"YES"` ← wrong quantity

---

## Part 10: `analyze_dataframe()` — the orchestrator

### Actual code (lines 823–888)

```python
def analyze_dataframe(df, col=None):
    if col is None:
        col = get_default_col()

    result = df.copy()

    # Initialise all new columns to None
    for c in GROUP_COLS + LINE_COLS:
        result[c] = None

    # Parse date columns
    for dcol in [col["start_date"], col["end_date"],
                 col["header_start"], col["header_end"]]:
        if dcol in result.columns:
            result[dcol] = pd.to_datetime(result[dcol], errors="coerce")

    groups   = result.groupby([col["quotation_no"], col["catalog_no"]], sort=False)
    n_groups = len(groups)
    print(f"Processing {len(df):,} rows across {n_groups:,} groups...", flush=True)

    for i, (_, group_df) in enumerate(groups, 1):
        if i % 200 == 0 or i == n_groups:
            print(f"  {i}/{n_groups} groups done", flush=True)

        try:
            header_start = group_df[col["header_start"]].iloc[0]
            header_end   = group_df[col["header_end"]].iloc[0]
        except (KeyError, IndexError):
            header_start = header_end = pd.NaT

        gm, per_line = analyze_group(group_df, header_start, header_end, col)

        # Write group-level metrics to every row in this group
        for c, val in gm.items():
            result.loc[group_df.index, c] = val

        # Write per-line metrics to individual rows
        for orig_idx, line_vals in per_line.items():
            for c, val in line_vals.items():
                result.loc[orig_idx, c] = val

    # Post-loop: how many distinct products per quotation header
    try:
        _quot_counts = (
            result.groupby(col["quotation_no"])[col["catalog_no"]].nunique()
        )
        result["groups_in_quotation"] = result[col["quotation_no"]].map(_quot_counts)
    except Exception:
        result["groups_in_quotation"] = None

    return result
```

### What this does

This is the main entry point. It:
1. Makes a copy of the raw dataframe (never modifies the original)
2. Adds all 32 output columns, initialised to `None`
3. Parses the 4 date columns from string/Excel format to proper date objects
   (`errors="coerce"` means unparseable dates become `NaT` instead of crashing)
4. Groups the dataframe by `(Quotation_No, Catalog_No)` — one group per combination
5. For each group: calls `analyze_group()`, then writes the results back to the main dataframe
   — group-level results go to EVERY row of the group, per-line results go to individual rows
6. After all groups: computes `groups_in_quotation` — how many distinct products
   share the same quotation header. This requires all groups to be processed first,
   so it cannot be computed inside `analyze_group()`.

> **REVIEWER NOTE — Header dates taken from `iloc[0]` (first row only):**
> `group_df[col["header_start"]].iloc[0]` takes the header date from the first row of the group.
> The comment says "header data is repeated on every line" — but if any row in the group
> has a different header date (a data quality issue in the source), the code uses whichever
> happens to be first and ignores the rest silently.
> If you see unexpected alignment results, check whether header dates are consistent across
> all rows of a group.

---

## Part 11: `get_summary_stats()` — the dashboard numbers

### Actual code (lines 895–929)

```python
def get_summary_stats(result_df, col=None):
    if col is None:
        col = get_default_col()

    def count_groups(mask):
        subset = result_df[mask]
        if subset.empty:
            return 0
        return subset.groupby([col["quotation_no"], col["catalog_no"]]).ngroups

    total_groups    = result_df.groupby([col["quotation_no"], col["catalog_no"]]).ngroups
    groups_gaps     = count_groups(pd.to_numeric(result_df["gap_days"], errors="coerce") > 0)
    groups_misalign = count_groups(result_df["header_aligned"] == "NO")
    groups_low_conf = count_groups(
        pd.to_numeric(result_df["period_confidence_pct"], errors="coerce") < 70
    )
    groups_qty_issue = count_groups(
        pd.to_numeric(result_df["qty_confidence_pct"], errors="coerce") < 100
    )
    total_lines_to_add = int(
        pd.to_numeric(result_df["lines_to_add"], errors="coerce").fillna(0).sum()
    )

    return {
        "total_groups":       total_groups,
        "groups_gaps":        groups_gaps,
        "groups_misalign":    groups_misalign,
        "groups_low_conf":    groups_low_conf,
        "groups_qty_issue":   groups_qty_issue,
        "total_lines_to_add": total_lines_to_add,
    }
```

### What this does

Produces the 6 numbers shown as metric cards in the Streamlit dashboard.
The inner function `count_groups(mask)` filters the dataframe to matching rows,
then counts unique `(Quotation_No, Catalog_No)` combinations — so the counts are
**group counts, not row counts**.

`pd.to_numeric(..., errors="coerce")` converts text like `"N/A"` or `"no gap"` to
`NaN` before comparing, so those rows do not accidentally count as issues.

> **REVIEWER NOTE — `groups_gaps` only counts internal gaps, not header-edge gaps:**
> The filter is `gap_days > 0`. But `gap_days` only covers internal gaps (within the
> group's span). A group that is perfectly internally continuous but starts 30 days after
> the header (`header_aligned = "NO"`) would appear in `groups_misalign` but NOT in
> `groups_gaps`. The two counts can overlap but also diverge.

---

## Part 12: Known Gaps in the Logic

These are things the code does NOT currently handle:

| Gap | What happens instead | Impact |
|-----|----------------------|--------|
| `C_ORIG_PRES_LINE_DB` not used in analysis | Counted per group only | Lines that shift +365d on renewal are not treated differently from ordinary lines |
| `CF_MDQ_PART_STA` (US/O) not used | Ignored entirely | Obsolete products get gap/alignment alerts even when they should just be closed |
| `C_RENEWABLE_DB` not used | Ignored entirely | Whether a header auto-renews does not change how gaps are scored |
| Single active line → 100% confidence | Passes as "clear" | Groups with one line appear fully confident; cannot actually verify a pattern |
| Unsolvable gaps (`"✗"`) not separately flagged | Only visible in `solution_list` text | Cannot filter to "groups with unsolvable gaps" without text search |
| `has_infinite_overlap` computed but not stored | Discarded | Renewable-line overlap situation invisible in the Excel output |

---

## Quick Reference: All Output Columns

| Column | Level | Meaning |
|--------|-------|---------|
| `group_line_count` | Group | Total lines including cancelled |
| `group_active_line_count` | Group | Lines used in the analysis |
| `unlimit_qty_count` | Group | Lines with C_UNLIMIT_QTY_DB = true |
| `orig_pres_count` | Group | Lines with C_ORIG_PRES_LINE_DB = true |
| `groups_in_quotation` | Group | How many distinct products in this quotation header |
| `group_start` | Group | Earliest active line start date |
| `group_end` | Group | Latest active line end date |
| `group_span_days` | Group | Naive window (start to end, includes gaps) |
| `actual_coverage_days` | Group | True covered days (interval union) |
| `inferred_period_pattern` | Group | e.g. "quarterly" |
| `inferred_period_days` | Group | e.g. 90 (always the bucket target, not measured) |
| `avg_period_days` | Group | Actual mean of active line durations |
| `period_confidence_pct` | Group | % of active lines matching the pattern |
| `active_line_periods` | Group | Pipe-separated list of each active line's date range |
| `canonical_qty` | Group | Most frequent quantity among active lines |
| `qty_confidence_pct` | Group | % of active lines with the canonical qty |
| `active_line_qtys` | Group | Pipe-separated list of active line quantities |
| `coverage_bar` | Group | Visual timeline: █ covered ░ gap ▓ overlap |
| `gap_days` | Group | Uncovered days within group span (internal only) |
| `gap_count` | Group | Number of distinct internal gaps |
| `overlap_days` | Group | Days covered by 2+ lines simultaneously |
| `overlap_count` | Group | Lines that start before a prior line has ended |
| `gap_details` | Group | Human-readable list of gap date ranges |
| `header_aligned` | Group | YES/NO — does group span match header window |
| `start_alignment` | Group | e.g. "starts 30d late" / "aligned" |
| `end_alignment` | Group | e.g. "ends 45d early" / "aligned" |
| `lines_to_add` | Group | Count of lines needed to fill solvable gaps |
| `gaps_solved_ratio` | Group | e.g. "3gap/5l" = 3 gaps needing 5 new lines |
| `gap_list` | Group | Each gap's date range (internal + header-edge) |
| `solution_list` | Group | "+N quarterly" or "✗" per gap (parallel to gap_list) |
| `line_period_bucket` | Line | e.g. "quarterly", "irregular", "quarterly (excluded)" |
| `is_period_outlier` | Line | YES = breaks group pattern / NO / N/A |
| `is_qty_outlier` | Line | YES = wrong quantity / NO / N/A |
