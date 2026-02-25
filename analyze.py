"""
Quotation ERP Data Analysis Script
===================================
Reads a raw ERP Excel export of quotation lines, groups them by
(Quotation_No, Catalog_No), and appends 22 analysis columns covering:
  - Period pattern detection (robust voting, not average)
  - Quantity pattern detection (mode, not average)
  - Coverage analysis (interval union — not naive max-min)
  - Header alignment (start/end direction + days difference)
  - Solution suggestion (how many lines to add to fix gaps)

Usage:
    python analyze.py path/to/your_export.xlsx
    → produces path/to/your_export_analysis.xlsx

Configuration:
    Edit the CONFIG section below to match your actual Excel column names.
    The Streamlit app (app.py) overrides these via UI dropdowns automatically.
"""

import sys
from collections import Counter
from datetime import timedelta

import pandas as pd


# ═══════════════════════════════════════════════════════════════
#  CONFIG  —  change these to match your actual Excel column names
#  (used by the CLI; the Streamlit app overrides them via UI)
# ═══════════════════════════════════════════════════════════════

COL_QUOTATION_NO  = "Quotation_No"           # Groups rows into quotation headers
COL_CATALOG_NO    = "Catalog_No"             # Together with Quotation_No defines a group
COL_LINE_NO       = "Line_No"
COL_REL_NO        = "Rel_No"
COL_START_DATE    = "C_START_DATE"           # Line validity start
COL_END_DATE      = "C_END_DATE"             # Line validity end
COL_HEADER_START  = "C_PRES_VALID_FROM"      # Header validity start (repeated on each row)
COL_HEADER_END    = "C_PRES_VALID_TO"        # Header validity end   (repeated on each row)
COL_QTY           = "BUY_QTY_DUE"           # Quantity per line
COL_STATE         = "STATE"                  # e.g. "released", "cancelled", "planned"
COL_PERIOD        = "C_PERIOD"               # e.g. "once" means single-use line
COL_ORIG_DB       = "C_ORIG_PRES_LINE_DB"    # Boolean: line gets +365d on header expiry
COL_MDQ_STATUS    = "CF_MDQ_PART_STA"        # Product status: "US"=active, "O"=obsolete
COL_RENEWABLE     = "C_RENEWABLE_DB"         # Boolean: header renews by date-shifting
COL_UNLIMIT_QTY   = "C_UNLIMIT_QTY_DB"       # Boolean: unlimited quantity (no limit)

# Date comparison tolerance: gaps/misalignments within this many days are ignored
TOLERANCE_DAYS = 5


# ═══════════════════════════════════════════════════════════════
#  PERIOD BUCKETS
#  Each entry: (name, target_days, tolerance_days)
#  Buckets are non-overlapping. Durations outside all buckets → "irregular"
# ═══════════════════════════════════════════════════════════════

PERIOD_BUCKETS = [
    ("monthly",      30,  10),   # 20–40 days
    ("bi-monthly",   60,  10),   # 50–70 days
    ("quarterly",    90,  10),   # 80–100 days  (covers real calendar qtrs 89-92d)
    ("4-month",     120,  10),   # 110–130 days
    ("semi-annual", 180,  12),   # 168–192 days (covers real half-years 181-184d)
    ("annual",      365,  15),   # 350–380 days
]


# ═══════════════════════════════════════════════════════════════
#  COLUMN CONFIG HELPERS
# ═══════════════════════════════════════════════════════════════

def get_default_col():
    """
    Returns the column name config dict built from the module-level COL_* constants.
    Used by the CLI. The Streamlit app builds its own dict from UI dropdowns.
    Keys marked (analysis) are actively used in calculations.
    Keys marked (info) are mapped but not yet used in logic — present for display.
    """
    return {
        # ── Used in analysis ──────────────────────────────────────
        "quotation_no":  COL_QUOTATION_NO,   # grouping key
        "catalog_no":    COL_CATALOG_NO,     # grouping key
        "start_date":    COL_START_DATE,     # line period start
        "end_date":      COL_END_DATE,       # line period end
        "header_start":  COL_HEADER_START,   # header validity start
        "header_end":    COL_HEADER_END,     # header validity end
        "qty":           COL_QTY,            # quantity pattern detection
        "state":         COL_STATE,          # active/cancelled detection
        "period":        COL_PERIOD,         # "once" exclusion
        # ── Informational (mapped, shown in output, not yet in logic) ─
        "renewable":     COL_RENEWABLE,      # header renewal flag
        "unlimit_qty":   COL_UNLIMIT_QTY,    # unlimited quantity flag
        "orig_pres":     COL_ORIG_DB,        # orig prescription line flag
        "mdq_status":    COL_MDQ_STATUS,     # product warehouse status
    }


# ═══════════════════════════════════════════════════════════════
#  HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════════

def is_line_active(row, col):
    """
    Returns True if this line should be included in pattern and coverage analysis.
    Excluded lines: cancelled, closed/placeholder (duration ≤ 1 day), or once-period.
    """
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
    # Lines with null/empty dates are renewable — they stay active;
    # their effective coverage = header dates (handled in analyze_group).
    try:
        start = row[col["start_date"]]
        end   = row[col["end_date"]]
        if pd.notna(start) and pd.notna(end):
            if (end - start).days < 4:
                return False
        # Null dates → renewable line, keep active
    except (KeyError, TypeError, AttributeError):
        pass  # cannot evaluate → do not exclude on this criterion

    return True


def map_to_bucket(duration_days):
    """
    Maps a line duration (integer days) to the nearest standard period bucket.
    Returns (bucket_name: str, target_days: int).
    Falls through to ("irregular", duration_days) if no bucket matches.
    """
    for name, target, tol in PERIOD_BUCKETS:
        if abs(duration_days - target) <= tol:
            return name, target
    return "irregular", duration_days


def compute_interval_union(intervals):
    """
    Given a list of (start_date, end_date) date pairs, merges overlapping
    and adjacent intervals (adjacent = consecutive days) and returns:
        total_covered_days: int  — sum of all merged interval lengths
        merged: list of [start, end]  — the merged intervals
    """
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


def detect_gaps(sorted_active_df, col):
    """
    Given a date-sorted DataFrame of active lines, finds gaps between
    lines that exceed TOLERANCE_DAYS.

    Tracks the running maximum end date (not just the previous row) so that
    overlapping lines don't create false gaps.  Example: if line A ends Sep 26
    and line B ends Aug 24 (contained inside A), the next line starting Sep 27
    is correctly seen as adjacent to Sep 26 — not as a 33-day gap after Aug 24.

    Returns list of dicts: {gap_start, gap_end, gap_days}
    """
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


def build_coverage_bar(bar_start, bar_end, intervals, width=48):
    """
    Builds a proportional text timeline bar for a group.

    Each of the `width` characters represents an equal time slice of the
    window bar_start → bar_end:
        █  covered by exactly one active line
        |  boundary between two adjacent (non-overlapping) lines
             — lets you count individual periods at a glance
        ▓  overlap — two or more lines cover the same slot
        ░  gap — no active line covers this slot

    When header dates are available the bar spans the full header period,
    so a late-starting group shows ░ at the left edge and an early-ending
    group shows ░ at the right edge.

    Examples:
        Quarterly, full:   ████████████|████████████|████████████|████████████
        Q3 missing:        ████████████|████████████░░░░░░░░░░░░░████████████
        Overlap:           ████████████▓▓▓▓████████████|████████████████████
        Header misaligned: ░░░░████████████|████████████░░░░░░░░░░░░░░░░░░░░
        Monthly (12 ln):   ███|███|███|███|███|███|███|███|███|███|███|███
    """
    if bar_start is None or bar_end is None:
        return ""
    total_days = (bar_end - bar_start).days + 1
    if total_days <= 0:
        return ""
    if not intervals:
        return "░" * width

    sorted_ivs = sorted(intervals, key=lambda x: x[0])

    bar          = []
    prev_interval = None   # index of the interval that covered the previous slot

    for i in range(width):
        slot_s_off = int(i       * total_days / width)
        slot_e_off = max(slot_s_off, int((i + 1) * total_days / width) - 1)
        slot_s = bar_start + timedelta(days=slot_s_off)
        slot_e = bar_start + timedelta(days=slot_e_off)

        covering = [j for j, (s, e) in enumerate(sorted_ivs)
                    if s <= slot_e and e >= slot_s]

        if len(covering) == 0:
            bar.append("░")           # gap — clearly light
            prev_interval = None
        elif len(covering) >= 2:
            bar.append("▓")           # overlap — darker shade, stands out
            prev_interval = covering[0]
        else:
            j = covering[0]
            if prev_interval is not None and prev_interval != j:
                bar.append("|")       # boundary between two adjacent lines
            else:
                bar.append("█")       # solid covered slot
            prev_interval = j

    return "".join(bar)


def solve_gap(gap_days, pattern_days):
    """
    Calculates how many lines of the given pattern duration are needed to fill
    a gap of gap_days.

    Returns an integer N if the gap is within TOLERANCE_DAYS of an exact
    whole multiple of pattern_days, otherwise returns None (not a clean fit).

    Example:
        gap=182d, pattern=90d  → 182/90=2.02 → n=2 → |182−180|=2 ≤ 5 → 2
        gap=305d, pattern=365d → 305/365=0.84 → n=1 → |305−365|=60 > 5 → None
    """
    if not pattern_days or pattern_days <= 0:
        return None
    ratio = gap_days / pattern_days
    n = round(ratio)
    if n < 1:
        return None
    if abs(gap_days - n * pattern_days) <= TOLERANCE_DAYS:
        return n
    return None


# ═══════════════════════════════════════════════════════════════
#  CORE: ANALYZE ONE GROUP
# ═══════════════════════════════════════════════════════════════

def analyze_group(group_df, header_start, header_end, col):
    """
    Analyzes one (Quotation_No, Catalog_No) group.

    Args:
        group_df     : DataFrame slice for this group (original indices preserved)
        header_start : pd.Timestamp or NaT — header validity start
        header_end   : pd.Timestamp or NaT — header validity end
        col          : column name config dict (from get_default_col() or UI)

    Returns:
        gm       : dict of group-level metric values
                   (written identically to every row in the group)
        per_line : dict of {original_df_index: {per-line column values}}
    """
    NA = "N/A"
    gm = {}

    # ── 1. Line counts ────────────────────────────────────────────────────────
    total_lines = len(group_df)
    active_mask = group_df.apply(lambda row: is_line_active(row, col), axis=1)
    active_df   = group_df[active_mask].sort_values(col["start_date"])
    n_active    = len(active_df)

    gm["group_line_count"]        = total_lines
    gm["group_active_line_count"] = n_active

    # ── Flag column counts (all lines in group, regardless of active status) ──
    # Count how many lines have each boolean ERP flag set to a truthy value.
    # Typical ERP export values: 1 / "1" / True / "Y" / "YES" / "true".
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

    # ── Classify active lines by date completeness ────────────────────────────
    # Three null-date cases, each with a different effective date resolution:
    #
    #   both null    → fully renewable: covers full header period (start → end)
    #   start null   → open-start:      effective start = header start, end = given end
    #   end null     → open-end:        effective end   = header end,   start = given start
    #
    # has_infinite_overlap is only flagged when FULLY renewable lines (both null)
    # coexist with explicit dated lines — that is a true "one line covers the whole
    # year, plus more lines on top" duplicate-coverage situation.
    # Open-start / open-end lines are anchored to one explicit date so they do not
    # automatically overlap everything; any actual overlap is caught by the interval
    # union / gap detection logic.
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

    # Build effective-date DataFrame: fill each null date field with the
    # corresponding header date.  Lines still missing a date after substitution
    # (no header date available) are dropped from coverage analysis.
    eff_df = active_df.copy()
    if has_infinite:
        for idx in infinite_set:
            if pd.isna(eff_df.at[idx, col["start_date"]]) and h_start is not None:
                eff_df.at[idx, col["start_date"]] = h_start
            if pd.isna(eff_df.at[idx, col["end_date"]]) and h_end is not None:
                eff_df.at[idx, col["end_date"]]   = h_end
        # Drop lines that still have a null date (no header date available to fill)
        still_null = [
            idx for idx in infinite_set
            if pd.isna(eff_df.at[idx, col["start_date"]])
            or pd.isna(eff_df.at[idx, col["end_date"]])
        ]
        if still_null:
            eff_df = eff_df.drop(still_null)
    eff_df = eff_df.sort_values(col["start_date"])

    # ── Per-line duration + bucket (all rows in group) ────────────────────────
    # Renewable lines use header-derived effective duration.
    # active_durations / active_buckets are also collected here for voting.
    per_line         = {}
    active_durations = {}   # idx → effective duration_days
    active_buckets   = {}   # idx → (base bucket_name, target_days)

    for idx, row in group_df.iterrows():
        is_active = bool(active_mask.get(idx, False))
        is_inf    = idx in infinite_set

        if is_inf:
            # One or both dates missing — fill each null field with header date
            orig_s = row[col["start_date"]]
            orig_e = row[col["end_date"]]
            eff_s  = orig_s if pd.notna(orig_s) else h_start
            eff_e  = orig_e if pd.notna(orig_e) else h_end

            if eff_s is not None and eff_e is not None:
                dur          = (eff_e - eff_s).days + 1
                bname, bdays = map_to_bucket(dur)
                if pd.isna(orig_s) and pd.isna(orig_e):
                    b_display = f"renewable ({bname})"    # both null → full header
                elif pd.isna(orig_s):
                    b_display = f"open-start ({bname})"  # start missing → from header
                else:
                    b_display = f"open-end ({bname})"    # end missing → to header end
            else:
                dur, bname, bdays = None, "open-date", 0
                b_display         = "open-date"
        else:
            try:
                s   = row[col["start_date"]]
                e   = row[col["end_date"]]
                dur = (e - s).days + 1 if (pd.notna(s) and pd.notna(e)) else None
            except Exception:
                dur = None
            if dur is not None and dur > 0:
                bname, bdays = map_to_bucket(dur)
            else:
                bname, bdays = "invalid", 0
            b_display = bname

        per_line[idx] = {
            "line_period_bucket": b_display if is_active else f"{b_display} (excluded)",
            "is_period_outlier":  None,
            "is_qty_outlier":     None,
        }

        # Collect voting data only for active lines that have effective dates
        if is_active and idx in eff_df.index:
            active_durations[idx] = dur or 0
            active_buckets[idx]   = (bname, bdays)

    # ── Handle groups where all lines are inactive/cancelled ─────────────────
    if n_active == 0:
        gm.update({
            "group_start":             NA,
            "group_end":               NA,
            "group_span_days":         0,
            "actual_coverage_days":    0,
            "inferred_period_pattern": NA,
            "inferred_period_days":    NA,
            "avg_period_days":         NA,
            "period_confidence_pct":   NA,
            "active_line_periods":     "",
            "canonical_qty":           NA,
            "qty_confidence_pct":      NA,
            "active_line_qtys":        "",
            "coverage_bar":            "",
            "gap_days":                0,
            "gap_count":               0,
            "overlap_days":            0,
            "overlap_count":           0,
            "gap_details":             "",
            "header_aligned":          NA,
            "start_alignment":         NA,
            "end_alignment":           NA,
            "lines_to_add":            "no gap",
            "gaps_solved_ratio":       "",
            "gap_list":                "",
            "solution_list":           "",
        })
        for idx in per_line:
            per_line[idx]["is_period_outlier"] = NA
            per_line[idx]["is_qty_outlier"]    = NA
        return gm, per_line

    # ── 2. Period pattern: bucket voting ──────────────────────────────────────
    # active_durations / active_buckets were populated in the per-line loop.
    # Renewable lines vote with their header-derived bucket (e.g. "annual").
    if not active_buckets:
        # All active lines are renewable with no header dates → pattern unknown
        winning_bucket = "irregular"
        winning_days   = 0
        winning_count  = n_active
    else:
        bucket_names = [v[0] for v in active_buckets.values()]
        vote         = Counter(bucket_names)
        winning_bucket, winning_count = vote.most_common(1)[0]

        # Tiebreaker: if two or more buckets share the top vote count,
        # pick the one whose lines cover the most total days.
        # Rationale: a 365-day line represents the dominant pattern more than
        # a 60-day line even when both have exactly 1 vote.
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

    # Average period: mean of valid active line durations
    valid_durs = [d for d in active_durations.values() if d and d > 0]
    gm["avg_period_days"] = round(sum(valid_durs) / len(valid_durs)) if valid_durs else NA

    # Active line periods: sorted list of each active line's effective date range
    if not eff_df.empty:
        _periods = []
        for _, _row in eff_df.iterrows():
            _s = _row[col["start_date"]]
            _e = _row[col["end_date"]]
            _dur = (_e - _s).days + 1
            _periods.append(f"{_s.date()} → {_e.date()} ({_dur}d)")
        gm["active_line_periods"] = " | ".join(_periods)
    else:
        gm["active_line_periods"] = ""

    # ── 3. Quantity pattern: mode voting ──────────────────────────────────────
    #  The most frequent quantity among active lines = canonical quantity.
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
        # Rationale: a higher quantity is more likely to represent the "real" contract volume.
        tied_qtys     = [q for q, c in qty_vote.items() if c == top_qty_count]
        canonical_qty = max(tied_qtys)
        qty_count     = top_qty_count
        gm["canonical_qty"]      = canonical_qty
        gm["qty_confidence_pct"] = round(qty_count / n_active * 100)
    else:
        canonical_qty = None
        gm["canonical_qty"]      = NA
        gm["qty_confidence_pct"] = NA

    # Pipe-separated list of each active line's quantity, in start-date order.
    # Uses eff_df (sorted by start_date) so the order matches active_line_periods.
    if not eff_df.empty:
        def _fmt_qty(v):
            if pd.isna(v):
                return "?"
            try:
                f = float(v)
                return str(int(f)) if f == int(f) else str(f)
            except Exception:
                return str(v)
        _qty_list = []
        for _idx, _ in eff_df.iterrows():
            try:
                _qty_list.append(_fmt_qty(group_df.loc[_idx, col["qty"]]))
            except Exception:
                _qty_list.append("?")
        gm["active_line_qtys"] = " | ".join(_qty_list)
    else:
        gm["active_line_qtys"] = ""

    # ── 4. Coverage analysis (using effective-date DataFrame) ─────────────────
    # Computes gap_days = total uncovered days within the group's date span.
    # eff_df uses header dates for renewable/open-date lines.
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

    # ── 5. Gap detection + overlap days ───────────────────────────────────────
    # gap_days   : total uncovered days (0 = continuous coverage)
    # overlap_days: when a fully renewable line + explicit dated lines coexist,
    #               overlap_days = coverage of the explicit lines (all duplicated).
    gaps             = detect_gaps(eff_df, col) if not eff_df.empty else []
    gm["gap_days"]   = _gap_days
    gm["gap_count"]  = len(gaps)

    # overlap_days: total days covered by more than one active line
    # Formula: sum of individual line durations minus the interval union.
    # Any positive result means some date range is covered by 2+ lines.
    if intervals:
        _sum_durs = sum((e - s).days + 1 for s, e in intervals)
        _overlap  = max(0, _sum_durs - _coverage)
        # Count how many intervals start before the previous one ends (sorted order)
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

    # Coverage bar: use header period as the window when available so that
    # alignment gaps show at the edges; fall back to group span.
    _bar_s = h_start if h_start is not None else (group_start if isinstance(group_start, pd.Timestamp) else None)
    _bar_e = h_end   if h_end   is not None else (group_end   if isinstance(group_end,   pd.Timestamp) else None)
    gm["coverage_bar"] = build_coverage_bar(_bar_s, _bar_e, intervals)

    # ── 6. Header alignment ───────────────────────────────────────────────────
    #  start_diff: group_start - header_start  (positive = group starts late)
    #  end_diff:   group_end   - header_end    (negative = group ends early)
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

    # ── 7. Lines to add + solution detail ────────────────────────────────────
    # Build a unified list of ALL gaps (internal + header-alignment).
    # For each gap: attempt to fill with the inferred pattern unit.
    # gap_list     : every gap's date range and size
    # solution_list: parallel — "+N pattern" if it fits cleanly, "✗" if not
    # gaps_solved_ratio: "solved/total" e.g. "2/3"

    _all_gaps = []

    # Internal gaps
    for g in gaps:
        _all_gaps.append({
            "label":    f"{g['gap_start'].date()} → {g['gap_end'].date()} ({g['gap_days']}d)",
            "gap_days": g["gap_days"],
        })

    # Start gap: group begins after header start
    if not start_ok and start_diff > TOLERANCE_DAYS:
        _gs = header_start.date() if isinstance(header_start, pd.Timestamp) else header_start
        _ge = (group_start - timedelta(days=1)).date() if isinstance(group_start, pd.Timestamp) else group_start
        _all_gaps.append({
            "label":    f"{_gs} → {_ge} ({start_diff}d) [before group]",
            "gap_days": start_diff,
        })

    # End gap: group ends before header end
    if not end_ok and end_diff < -TOLERANCE_DAYS:
        _gs = (group_end + timedelta(days=1)).date() if isinstance(group_end, pd.Timestamp) else group_end
        _ge = header_end.date() if isinstance(header_end, pd.Timestamp) else header_end
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

    _total_gaps = len(_all_gaps)
    gm["gaps_solved_ratio"] = f"{_total_gaps}gap/{total_to_add}l" if _total_gaps > 0 else ""
    gm["gap_list"]          = " | ".join(gap_labels)      if gap_labels      else ""
    gm["solution_list"]     = " | ".join(solution_labels) if solution_labels else ""
    gm["lines_to_add"]      = total_to_add if total_to_add > 0 else ("no gap" if _total_gaps == 0 else 0)

    # ── 8. Per-line outlier flags ─────────────────────────────────────────────
    for idx in group_df.index:
        is_active_line = idx in active_df.index

        # Period outlier: active line whose effective bucket ≠ winning bucket.
        # Renewable lines with no header dates were excluded from voting → N/A.
        if is_active_line:
            if idx in active_buckets:
                bname, _ = active_buckets[idx]
                per_line[idx]["is_period_outlier"] = (
                    "YES" if bname != winning_bucket else "NO"
                )
            else:
                # Renewable line with no header dates — could not classify
                per_line[idx]["is_period_outlier"] = NA
        else:
            per_line[idx]["is_period_outlier"] = NA

        # Quantity outlier: active line whose qty doesn't match canonical qty
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

    return gm, per_line


# ═══════════════════════════════════════════════════════════════
#  DATAFRAME ORCHESTRATION
# ═══════════════════════════════════════════════════════════════

# Group-level columns (same value on every row in the group)
GROUP_COLS = [
    "group_line_count",
    "group_active_line_count",
    "unlimit_qty_count",      # lines with C_UNLIMIT_QTY_DB = true
    "orig_pres_count",        # lines with C_ORIG_PRES_LINE_DB = true
    "groups_in_quotation",    # how many distinct Catalog_No share this Quotation_No
    "group_start",
    "group_end",
    "group_span_days",         # max(end) - min(start) + 1 (naive, includes gaps)
    "actual_coverage_days",    # interval union: true days covered by active lines
    "inferred_period_pattern",
    "inferred_period_days",
    "avg_period_days",         # mean duration of active lines (simple average)
    "period_confidence_pct",
    "active_line_periods",     # pipe-separated list of each active line's period
    "canonical_qty",
    "qty_confidence_pct",
    "active_line_qtys",    # pipe-separated list of each active line's qty (start-date order)
    "coverage_bar",    # 48-char visual timeline: █=covered ░=gap (spans header period)
    "gap_days",        # total uncovered days within group span (0 = no gaps)
    "gap_count",       # number of distinct gaps (0 = continuous coverage)
    "overlap_days",    # days covered by more than one active line (0 = none)
    "overlap_count",   # number of overlapping line pairs (0 = no overlaps)
    "gap_details",
    "header_aligned",
    "start_alignment",
    "end_alignment",
    "lines_to_add",
    "gaps_solved_ratio",   # "solved/total" e.g. "2/3"
    "gap_list",            # each gap's date range and size
    "solution_list",       # parallel to gap_list: "+N pattern" or "✗"
]

# Per-line columns (unique value per row)
LINE_COLS = [
    "line_period_bucket",
    "is_period_outlier",
    "is_qty_outlier",
]


def analyze_dataframe(df, col=None):
    """
    Main entry point. Takes the raw ERP dataframe, returns it with
    analysis columns appended (original data is not modified).

    Args:
        df  : raw pandas DataFrame from Excel export
        col : column name config dict. If None, uses get_default_col()
              (module-level COL_* constants). Pass a custom dict from the
              Streamlit UI to override column names at runtime.
    """
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

        # Header dates are repeated on every line — take from the first row
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

    # ── Post-loop: cross-group metric ─────────────────────────────────────────
    # groups_in_quotation: how many distinct Catalog_No values share the same
    # Quotation_No (i.e. how many product groups are under one quotation header).
    # Computed here because it requires seeing all groups at once.
    try:
        _quot_counts = (
            result.groupby(col["quotation_no"])[col["catalog_no"]]
            .nunique()
        )
        result["groups_in_quotation"] = result[col["quotation_no"]].map(_quot_counts)
    except Exception:
        result["groups_in_quotation"] = None

    return result


# ═══════════════════════════════════════════════════════════════
#  SUMMARY STATS (used by both CLI and Streamlit)
# ═══════════════════════════════════════════════════════════════

def get_summary_stats(result_df, col=None):
    """
    Returns a dict of summary statistics about the analysis results.
    Used by CLI (print_summary) and Streamlit (metric cards).
    """
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


def print_summary(result_df, col=None):
    s = get_summary_stats(result_df, col)
    print()
    print("══════════════════════════════════════════════")
    print("  ANALYSIS SUMMARY")
    print("══════════════════════════════════════════════")
    print(f"  Total groups analysed     : {s['total_groups']:,}")
    print(f"  Groups with gaps          : {s['groups_gaps']:,}")
    print(f"  Groups misaligned header  : {s['groups_misalign']:,}")
    print(f"  Groups unclear period     : {s['groups_low_conf']:,}  (< 70% confidence)")
    print(f"  Groups qty inconsistency  : {s['groups_qty_issue']:,}")
    print(f"  Total lines to add        : {s['total_lines_to_add']:,}")
    print("══════════════════════════════════════════════")
    print()


# ═══════════════════════════════════════════════════════════════
#  CLI ENTRY POINT
# ═══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        print("Usage: python analyze.py path/to/your_export.xlsx")
        sys.exit(1)

    input_path = sys.argv[1]
    print(f"Reading: {input_path}")

    df = pd.read_excel(input_path)
    print(f"Loaded {len(df):,} rows, {len(df.columns)} columns.")

    result_df = analyze_dataframe(df)   # uses default col config
    print_summary(result_df)

    # Build output path: insert _analysis before .xlsx extension
    if input_path.lower().endswith(".xlsx"):
        out_path = input_path[:-5] + "_analysis.xlsx"
    elif input_path.lower().endswith(".xls"):
        out_path = input_path[:-4] + "_analysis.xlsx"
    else:
        out_path = input_path + "_analysis.xlsx"

    result_df.to_excel(out_path, index=False)
    print(f"Output saved: {out_path}")
