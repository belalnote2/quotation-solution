"""
Quotation Analyser — Streamlit App
====================================
Upload your ERP Excel export, configure column names, run analysis,
explore results in-browser, and download the annotated file.

Run:
    streamlit run app.py
"""

import io
import pandas as pd
import streamlit as st

from analyze import (
    analyze_dataframe,
    get_summary_stats,
    get_default_col,
    GROUP_COLS,
    LINE_COLS,
    PERIOD_BUCKETS,
    TOLERANCE_DAYS,
)

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Quotation Analyser",
    page_icon="📊",
    layout="wide",
)

# ── Session state ─────────────────────────────────────────────────────────────
for key, default in [("result_df", None), ("col_config", None)]:
    if key not in st.session_state:
        st.session_state[key] = default

ALL_ANALYSIS_COLS = GROUP_COLS + LINE_COLS

# ── Field definitions (key, display label, help text) ────────────────────────
# Order matters: determines the display order in the mapping table + overrides.
FIELD_DEFS = [
    ("quotation_no", "Quotation No",         "Groups rows into quotation headers"),
    ("catalog_no",   "Catalog No",           "Product/article — with Quotation No defines a group"),
    ("start_date",   "Line Start Date",      "C_START_DATE — line validity start"),
    ("end_date",     "Line End Date",        "C_END_DATE — line validity end"),
    ("header_start", "Header Valid From",    "C_PRES_VALID_FROM — header start (repeated per row)"),
    ("header_end",   "Header Valid To",      "C_PRES_VALID_TO — header end (repeated per row)"),
    ("qty",          "Quantity",             "BUY_QTY_DUE — line quantity for pattern detection"),
    ("state",        "State",                "STATE — cancelled lines are excluded from analysis"),
    ("period",       "Period type",          "C_PERIOD — lines with value 'once' are excluded"),
    ("renewable",    "Renewable flag",       "C_RENEWABLE_DB — header renews by date-shifting"),
    ("unlimit_qty",  "Unlimited qty flag",   "C_UNLIMIT_QTY_DB — no quantity limit on this line"),
    ("orig_pres",    "Orig pres line",       "C_ORIG_PRES_LINE_DB — line shifts +365d on renewal"),
    ("mdq_status",   "MDQ product status",   "CF_MDQ_PART_STA — O=obsolete, US=active on warehouse"),
]

# ── Column preset views (used for reference; tabs build their own views) ──────
PRESET_VIEWS = {
    "All analysis columns": ALL_ANALYSIS_COLS,
    "Coverage": [
        "group_start", "group_end",
        "group_span_days", "actual_coverage_days",
        "gap_days", "gap_count",
        "gap_details",
    ],
    "Period pattern (per line)": [
        "line_period_bucket", "is_period_outlier",
        "inferred_period_pattern", "inferred_period_days", "period_confidence_pct",
    ],
    "Quantity (per line)": [
        "canonical_qty", "qty_confidence_pct", "is_qty_outlier",
    ],
    "Header alignment": [
        "group_start", "group_end",
        "header_aligned", "start_alignment", "end_alignment",
    ],
    "Issues overview": [
        "gap_days", "gap_count", "gap_details",
        "header_aligned", "start_alignment", "end_alignment",
        "period_confidence_pct", "qty_confidence_pct",
        "lines_to_add", "gaps_solved_ratio",
    ],
    "Custom": None,
}


# ── Helpers ───────────────────────────────────────────────────────────────────

def auto_map_columns(df_columns, defaults):
    """
    Auto-detect column mappings from the uploaded file.
    Tries exact match first, then case-insensitive fallback.

    Returns:
        mapped    : dict {key → actual_column_name_in_file}  (all keys present)
        not_found : list of keys where no match was found
    """
    cols_lower = {c.lower(): c for c in df_columns}
    fallback   = df_columns[0] if df_columns else ""

    mapped, not_found = {}, []
    for key, default_name in defaults.items():
        if default_name in df_columns:
            mapped[key] = default_name                          # exact match ✅
        elif default_name.lower() in cols_lower:
            mapped[key] = cols_lower[default_name.lower()]      # case-insensitive ✅
        else:
            mapped[key] = fallback                              # not found ⚠️
            not_found.append(key)

    return mapped, not_found


def to_excel_bytes(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Analysis")
    return buf.getvalue()


def style_table(df, analysis_cols=None, key_cols=None):
    """
    Three-layer visual styling:
      Layer 1 (base)  — analysis columns: light blue tint (#eef5fb)
      Layer 2         — key columns: stronger blue + bold (#cee4f5)
      Layer 3 (top)   — issue cells: red/yellow overrides (same as before)
    ERP source columns keep the default white background.
    """
    styles = pd.DataFrame("", index=df.index, columns=df.columns)

    # Layer 1 — analysis column tint (all result columns as a visual packet)
    for c in (analysis_cols or []):
        if c in df.columns:
            styles[c] = "background-color:#eef5fb"

    # Layer 2 — key column emphasis (most important indicator per tab)
    for c in (key_cols or []):
        if c in df.columns:
            styles[c] = "background-color:#cee4f5;font-weight:600"

    # Layer 2b — period analysis columns: green tint to group them visually
    _PERIOD_COLS = {"inferred_period_pattern", "inferred_period_days",
                    "avg_period_days", "period_confidence_pct"}
    for c in df.columns:
        if c in _PERIOD_COLS:
            styles[c] = "background-color:#e8f5e9"

    # Layer 2c — solution columns: amber/orange tint to group them visually
    _SOLUTION_COLS = {"lines_to_add", "gaps_solved_ratio", "gap_list", "solution_list"}
    for c in df.columns:
        if c in _SOLUTION_COLS:
            styles[c] = "background-color:#fff8e1"

    # Layer 3 — issue highlights (override everything above)
    for c in df.columns:
        if c in ("gap_days", "gap_count", "overlap_days", "overlap_count"):
            num = pd.to_numeric(df[c], errors="coerce")
            styles.loc[num > 0, c] = "background-color:#ffd6d6;color:#900;font-weight:600"
        elif c == "header_aligned":
            mask = df[c] == "NO"
            styles.loc[mask, c] = "background-color:#ffd6d6;color:#900;font-weight:600"
        elif c in ("is_period_outlier", "is_qty_outlier"):
            mask = df[c] == "YES"
            styles.loc[mask, c] = "background-color:#fff3cc;color:#664d00;font-weight:600"
        elif c == "period_confidence_pct":
            num = pd.to_numeric(df[c], errors="coerce")
            styles.loc[num < 50,  c] = "background-color:#ffd6d6"
            styles.loc[(num >= 50) & (num < 70), c] = "background-color:#fff3cc"
        elif c == "qty_confidence_pct":
            num = pd.to_numeric(df[c], errors="coerce")
            styles.loc[num < 100, c] = "background-color:#fff3cc;color:#664d00"
        elif c == "lines_to_add":
            num = pd.to_numeric(df[c], errors="coerce")
            styles.loc[num > 0, c] = "background-color:#ffd6d6"
        elif c in ("unlimit_qty_count", "orig_pres_count"):
            num = pd.to_numeric(df[c], errors="coerce")
            styles.loc[num > 0, c] = "background-color:#cee4f5;font-weight:600"

    return styles


def build_issues_mask(df):
    """Boolean mask: rows whose group has at least one issue."""
    return (
        (pd.to_numeric(df.get("gap_days",     pd.Series(0, index=df.index)), errors="coerce") > 0)
        | (pd.to_numeric(df.get("overlap_days", pd.Series(0, index=df.index)), errors="coerce") > 0)
        | (df.get("header_aligned", pd.Series("", index=df.index)) == "NO")
        | (pd.to_numeric(df.get("period_confidence_pct", pd.Series()), errors="coerce") < 70)
        | (pd.to_numeric(df.get("qty_confidence_pct",    pd.Series()), errors="coerce") < 100)
    )


def build_group_summary(df, col_config):
    """One row per (Quotation_No, Catalog_No) — group-level columns only."""
    key_cols = [col_config["quotation_no"], col_config["catalog_no"]]
    keep     = key_cols + [c for c in GROUP_COLS if c in df.columns]
    return df[keep].drop_duplicates(subset=key_cols).reset_index(drop=True)


def render_table(df, label="", analysis_cols=None, key_cols=None):
    """Render a styled dataframe, fall back to plain if styling fails."""
    # Strip time component from all datetime columns (show 2025-06-27 not 2025-06-27 00:00:00)
    display_df = df.copy()
    for c in display_df.columns:
        if pd.api.types.is_datetime64_any_dtype(display_df[c]):
            display_df[c] = display_df[c].dt.date

    if label:
        st.caption(
            label + "   ·   "
            "⬜ ERP source   "
            "🟦 Analysis result   "
            "🔷 Key indicator   "
            "🟡 Warning   "
            "🔴 Issue"
        )
    try:
        styled = display_df.style.apply(
            lambda _: style_table(display_df, analysis_cols, key_cols),
            axis=None,
        )
        st.dataframe(styled, use_container_width=True, height=520)
    except Exception:
        st.dataframe(display_df, use_container_width=True, height=520)


def focused_view(df, col_config, orig_keys, analysis_cols):
    """
    Build a focused column view for an analysis-type tab.
    orig_keys     : list of col_config keys for original ERP columns to include
    analysis_cols : list of analysis column names (fixed strings) to include
    Returns a dataframe with only the relevant columns, in order.
    """
    orig = [col_config.get(k, "") for k in orig_keys]
    orig = [c for c in orig if c and c in df.columns]
    anly = [c for c in analysis_cols if c in df.columns]
    seen, final = set(), []
    for c in orig + anly:
        if c not in seen:
            seen.add(c)
            final.append(c)
    return df[final] if final else df


def issues_only_toggle(df, mask, key):
    """Checkbox to filter to issue rows only. Returns filtered df."""
    c1, c2 = st.columns([5, 1])
    with c2:
        toggle = st.checkbox("Issues only", value=False, key=f"tog_{key}")
    return df[mask] if toggle else df, toggle


# ═══════════════════════════════════════════════════════════════
#  SIDEBAR
# ═══════════════════════════════════════════════════════════════

with st.sidebar:
    st.header("📂 Upload")

    uploaded_file = st.file_uploader(
        "ERP Excel export",
        type=["xlsx", "xls"],
        help="The raw quotation lines export from your ERP system.",
    )

    if uploaded_file:
        try:
            raw_df = pd.read_excel(uploaded_file)
            cols   = raw_df.columns.tolist()
        except Exception as e:
            st.error(f"Could not read file: {e}")
            st.stop()

        # ── Auto-detect column mapping ────────────────────────────────────────
        auto_mapped, not_found = auto_map_columns(cols, get_default_col())

        st.divider()
        st.subheader("⚙️ Column mapping")

        if not not_found:
            st.success(f"All {len(FIELD_DEFS)} columns detected automatically")
        else:
            missing_labels = [
                label for key, label, _ in FIELD_DEFS if key in not_found
            ]
            st.warning(
                f"{len(not_found)} column(s) not found in file:\n\n"
                + "\n".join(f"• {l}" for l in missing_labels)
            )

        # ── Mapping status table ──────────────────────────────────────────────
        mapping_rows = [
            {
                " ": "✅" if key not in not_found else "⚠️",
                "Field": label,
                "Column in file": auto_mapped[key] if key not in not_found else "— not found",
            }
            for key, label, _ in FIELD_DEFS
        ]
        st.dataframe(
            pd.DataFrame(mapping_rows),
            hide_index=True,
            use_container_width=True,
            height=min(37 * len(FIELD_DEFS) + 38, 430),
        )

        # ── Override expander (collapsed unless something is missing) ─────────
        override_label = (
            "Override mappings ⚠️ — action required"
            if not_found else
            "Override mappings"
        )
        col_config = {}
        with st.expander(override_label, expanded=bool(not_found)):
            st.caption(
                "Columns are pre-filled from auto-detection. "
                "Only change if a mapping is wrong."
            )
            for key, label, help_text in FIELD_DEFS:
                mapped_col = auto_mapped.get(key, cols[0])
                idx = cols.index(mapped_col) if mapped_col in cols else 0
                col_config[key] = st.selectbox(
                    label, cols, index=idx,
                    key=f"ov_{key}",
                    help=help_text,
                )

        st.divider()
        run_btn = st.button("▶ Run Analysis", type="primary", use_container_width=True)


# ═══════════════════════════════════════════════════════════════
#  MAIN AREA
# ═══════════════════════════════════════════════════════════════

st.title("📊 Quotation Data Analyser")
st.caption(
    "Analyses ERP quotation lines for period pattern issues, coverage gaps, "
    "quantity inconsistencies, and header alignment problems."
)

if not uploaded_file:
    st.info("👈 Upload your Excel file in the sidebar to get started.")
    st.stop()

# ── File preview ──────────────────────────────────────────────────────────────
with st.expander("📄 File preview (first 10 rows)", expanded=False):
    st.dataframe(raw_df.head(10), use_container_width=True)
    st.caption(f"{len(raw_df):,} rows · {len(raw_df.columns)} columns")

# ── Run analysis ──────────────────────────────────────────────────────────────
if run_btn:
    with st.spinner("Analysing groups…"):
        try:
            result = analyze_dataframe(raw_df, col=col_config)
            st.session_state.result_df  = result
            st.session_state.col_config = col_config
        except Exception as e:
            st.error(f"Analysis failed: {e}")
            st.exception(e)
            st.stop()
    st.success("Analysis complete!")

# ── Stop here if no results yet ───────────────────────────────────────────────
if st.session_state.result_df is None:
    st.stop()

result_df  = st.session_state.result_df
col_config = st.session_state.col_config
stats      = get_summary_stats(result_df, col_config)
raw_cols   = raw_df.columns.tolist()    # original uploaded columns

# ── Summary metric cards ──────────────────────────────────────────────────────
st.subheader("Summary")
m1, m2, m3, m4, m5, m6 = st.columns(6)
m1.metric("Total groups",      f"{stats['total_groups']:,}")
m2.metric("Groups with gaps",  f"{stats['groups_gaps']:,}",      delta_color="inverse",
          delta=f"{stats['groups_gaps']} need attention" if stats['groups_gaps'] else None)
m3.metric("Header misaligned", f"{stats['groups_misalign']:,}",  delta_color="inverse")
m4.metric("Unclear period",    f"{stats['groups_low_conf']:,}",  delta_color="inverse",
          help="Groups where < 70% of active lines match the inferred period pattern")
m5.metric("Qty inconsistency", f"{stats['groups_qty_issue']:,}", delta_color="inverse")
m6.metric("Lines to add",      f"{stats['total_lines_to_add']:,}",
          help="Total lines needed to fill all cleanly-calculable gaps")

st.divider()

# ── Download full result ──────────────────────────────────────────────────────
dl_col, _ = st.columns([2, 5])
with dl_col:
    st.download_button(
        label="⬇️ Download full analysis (Excel)",
        data=to_excel_bytes(result_df),
        file_name="quotation_analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.divider()

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab_solution, tab_align, tab_period, tab_about = st.tabs([
    "🔧 Solution Analysis",
    "📌 Header Alignment",
    "📋 Line Level Analysis",
    "📖 Analysis Logic",
])

# Shared column keys for identifiers shown in every analysis tab
ID_KEYS = ["quotation_no", "catalog_no"]


# ═══════════════════════════════════════════════════════════════
#  TAB 1 — PERIOD CONSISTENCY
#  Shows: how well each line matches the group's inferred period
# ═══════════════════════════════════════════════════════════════
with tab_period:
    period_mask = (
        (pd.to_numeric(result_df.get("period_confidence_pct", pd.Series()),
                       errors="coerce") < 100)
        | (result_df.get("is_period_outlier", pd.Series()) == "YES")
    )
    n_grp = result_df[period_mask].groupby(
        [col_config["quotation_no"], col_config["catalog_no"]]
    ).ngroups if period_mask.any() else 0
    n_lines = int(period_mask.sum())

    pm1, pm2 = st.columns(2)
    pm1.metric("Groups with period issues", n_grp)
    pm2.metric("Outlier lines", n_lines,
               help="Lines whose duration doesn't match the group's dominant pattern")

    df_p = result_df

    # Analysis columns (group-level values repeat on every line of the group)
    _period_tab_cols = [
        "line_period_bucket", "is_period_outlier",
        "group_line_count", "group_active_line_count", "group_start", "group_end",
        "inferred_period_pattern", "inferred_period_days",
        "avg_period_days", "period_confidence_pct",
        "canonical_qty", "qty_confidence_pct", "active_line_qtys",
    ]

    # Build flat view (qty added to raw ERP source keys)
    view_p = focused_view(
        df_p, col_config,
        orig_keys     = ID_KEYS + ["start_date", "end_date", "state",
                                   "header_start", "header_end", "qty"],
        analysis_cols = _period_tab_cols,
    )

    # ── MultiIndex column headers ──────────────────────────────────────────────
    _PERIOD_SECTIONS = [
        ("Identifiers",      [col_config.get("quotation_no", ""),
                              col_config.get("catalog_no", "")]),
        ("Raw Data",         [col_config.get("start_date", ""),
                              col_config.get("end_date", ""),
                              col_config.get("state", ""),
                              col_config.get("header_start", ""),
                              col_config.get("header_end", ""),
                              col_config.get("qty", "")]),
        ("Period (line)",    ["line_period_bucket", "is_period_outlier"]),
        ("Group Info",       ["group_line_count", "group_active_line_count",
                              "group_start", "group_end"]),
        ("Period Analysis",  ["inferred_period_pattern", "inferred_period_days",
                              "avg_period_days", "period_confidence_pct"]),
        ("Quantity Analysis",["canonical_qty", "qty_confidence_pct", "active_line_qtys"]),
    ]

    _p_flat_to_section = {}
    for _sec, _cols in _PERIOD_SECTIONS:
        for _c in _cols:
            _p_flat_to_section[_c] = _sec

    _p_mi_tuples = [(_p_flat_to_section.get(c, "Other"), c) for c in view_p.columns]
    _p_mi_cols   = pd.MultiIndex.from_tuples(_p_mi_tuples)

    _display_p = view_p.copy()
    for _c in _display_p.columns:
        if pd.api.types.is_datetime64_any_dtype(_display_p[_c]):
            _display_p[_c] = _display_p[_c].dt.date

    _p_flat_styles = style_table(
        _display_p,
        analysis_cols=_period_tab_cols,
        key_cols=["is_period_outlier", "period_confidence_pct"],
    )

    _display_p.columns    = _p_mi_cols
    _p_flat_styles.columns = _p_mi_cols

    st.caption(
        f"{len(df_p):,} rows   ·   "
        "Group-level columns (Group Info, Period Analysis, Quantity Analysis) "
        "repeat the same value on every line of that group   ·   "
        "⬜ Raw Data   🟦 Analysis result   🔷 Key   🟡 Warning   🔴 Issue"
    )
    try:
        _styled_p = _display_p.style.apply(lambda _: _p_flat_styles, axis=None)
        st.dataframe(_styled_p, use_container_width=True, height=520)
    except Exception:
        st.dataframe(_display_p, use_container_width=True, height=520)


# ═══════════════════════════════════════════════════════════════
#  TAB 2 — HEADER ALIGNMENT
#  One row per group — shows coverage metrics vs header validity
# ═══════════════════════════════════════════════════════════════
with tab_align:
    # Deduplicate to one row per group (header alignment is a group-level metric)
    _align_keys = [col_config["quotation_no"], col_config["catalog_no"]]
    group_align = result_df.drop_duplicates(subset=_align_keys).reset_index(drop=True)

    _align_issue_mask = group_align.get("header_aligned", pd.Series()) == "NO"
    n_grp_ali = int(_align_issue_mask.sum())

    am1, am2, am3 = st.columns(3)
    am1.metric("Groups misaligned with header", n_grp_ali)
    am2.metric("Aligned groups", stats["total_groups"] - n_grp_ali)
    am3.metric("Total groups", stats["total_groups"])

    # Build column list: ID → coverage → gaps → alignment → pattern + solution
    _align_orig = []
    for _k in ["quotation_no", "catalog_no", "header_start", "header_end"]:
        _c = col_config.get(_k, "")
        if _c and _c in group_align.columns and _c not in _align_orig:
            _align_orig.append(_c)

    _align_analysis = [
        "group_line_count", "group_active_line_count", "orig_pres_count",
        "group_start", "group_end",
        "actual_coverage_days", "group_span_days", "overlap_days", "overlap_count",
        "gap_days", "gap_count",
        "header_aligned", "start_alignment", "end_alignment",
        "inferred_period_pattern", "inferred_period_days", "lines_to_add",
    ]
    _align_all = _align_orig + [c for c in _align_analysis if c in group_align.columns]
    view_align = group_align[[c for c in _align_all if c in group_align.columns]]

    render_table(
        view_align,
        label=f"{len(view_align):,} groups",
        analysis_cols=_align_analysis,
        key_cols=["header_aligned", "start_alignment", "end_alignment", "lines_to_add"],
    )

    _dl_align, _ = st.columns([2, 5])
    with _dl_align:
        st.download_button(
            label="⬇️ Download header alignment (Excel)",
            data=to_excel_bytes(view_align),
            file_name="quotation_header_alignment.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_align",
        )

    st.divider()
    with st.expander("📖 How this analysis works — definitions, exclusions & calculations"):

        st.markdown("### Which lines are included?")
        st.markdown("""
Each group contains one or more quotation lines. **Not all lines are used** in the analysis.
A line is **excluded** if any of the following is true:

| Condition | Why it is skipped |
|---|---|
| `STATE = cancelled` | Cancelled lines are never renewed — they no longer affect coverage |
| Duration **< 5 days** (`end − start < 5`) | Zero- or near-zero-day lines are administrative placeholders, not real coverage periods |
| `C_PERIOD = once` | Single-use lines are not part of a repeating schedule |

All remaining lines — regardless of state (released, created, planned) — are counted as **active**.

> **Special case — empty dates:**
> Some lines have no start or end date. These are *renewable* lines whose coverage is tied to the quotation header.
> - Both dates empty → line covers the **full header period** (from `C_PRES_VALID_FROM` to `C_PRES_VALID_TO`)
> - Only start date empty → effective start = header start, end date stays as-is
> - Only end date empty → effective end = header end, start date stays as-is
        """)

        st.markdown("### Column definitions")
        st.markdown("""
| Column | What it means |
|---|---|
| **group_line_count** | Total number of lines in the group, including cancelled and excluded ones |
| **group_active_line_count** | Lines that pass all exclusion rules and are used in calculations |
| **orig_pres_count** | How many lines in the group have `C_ORIG_PRES_LINE_DB = true` — these lines get +365 days added when the quotation header renews |
| **group_start** | The earliest start date among all active lines (after substituting header dates for empty fields) |
| **group_end** | The latest end date among all active lines |
| **actual_coverage_days** | The true number of days covered by at least one active line — calculated by merging all line periods that touch or overlap, then summing the merged result |
| **group_span_days** | The simple max-minus-min span: `group_end − group_start + 1`. This includes any gaps inside. Compare to `actual_coverage_days` to see if gaps exist |
| **overlap_days** | Days where two or more active lines cover the same date. Formula: sum of all individual line durations minus the merged coverage |
| **overlap_count** | How many times a line starts before the previous one has ended (in date order). One overlap event = one such instance |
| **gap_days** | Total days inside the group's span that are not covered by any active line. `gap_days = group_span_days − actual_coverage_days` |
| **gap_count** | How many separate uncovered periods exist. A group can have 0 gaps, 1 large gap, or several small gaps |
| **header_aligned** | YES if the group's coverage window matches the quotation header validity window within ±5 days on both ends. NO otherwise |
| **start_alignment** | How far off the group's first active date is from the header start. "aligned" = within 5 days. "starts 30d late" = group begins 30 days after the header opens |
| **end_alignment** | How far off the group's last active date is from the header end. "ends 30d early" = group stops 30 days before the header closes |
| **inferred_period_pattern** | The dominant repeating period of the lines in this group (e.g. quarterly, monthly). Determined by voting — the most common duration bucket wins |
| **inferred_period_days** | The number of days that pattern represents (e.g. quarterly = 90 days) |
| **lines_to_add** | Estimated number of new lines needed to fill all gaps cleanly. Only calculated when the pattern is regular (not irregular) and the gap size divides evenly into the pattern |
        """)

        st.markdown("### How each calculation is made")
        st.markdown(f"""
**actual_coverage_days — interval union**
```
1. Take all active lines, sorted by start date
2. Start with the first line's period as the current merged block
3. For each next line:
     if it starts before or on the day after the current block ends
         → extend the current block if needed (merge)
     else
         → save the current block, start a new one
4. Sum the lengths of all merged blocks → actual_coverage_days
```
*Example:*
```
Line A: Jan 1 → Mar 31  (90 days)
Line B: Apr 1 → Jun 30  (91 days)   ← starts the day after A ends → merged
Line C: Sep 1 → Nov 30  (91 days)   ← gap before this one → new block

Merged blocks:  [Jan 1 → Jun 30]  +  [Sep 1 → Nov 30]
actual_coverage_days = 181 + 91 = 272 days
group_span_days      = Nov 30 − Jan 1 + 1 = 334 days
gap_days             = 334 − 272 = 62 days  (the Jul–Aug gap)
```

**gap_days & gap_count**
```
gap_days  = group_span_days − actual_coverage_days

gap_count:
  For each consecutive pair of active lines (sorted by date):
    gap = next_line_start − previous_line_end − 1 day
    if gap > {TOLERANCE_DAYS} days → count it as a gap
```

**overlap_days & overlap_count**
```
overlap_days  = sum(each active line's own duration) − actual_coverage_days
             → positive means some days are covered twice or more

overlap_count:
  Sort lines by start date, track the furthest end date seen so far
  For each next line:
    if it starts before the furthest end → overlap_count + 1
    update furthest end if this line goes further
```

**header_aligned, start_alignment, end_alignment**
```
start_diff = group_start − header_start   (positive = group starts late)
end_diff   = group_end   − header_end     (negative = group ends early)

if |start_diff| ≤ {TOLERANCE_DAYS} days AND |end_diff| ≤ {TOLERANCE_DAYS} days
    → header_aligned = YES
else
    → header_aligned = NO

start_alignment = "aligned"           if |start_diff| ≤ {TOLERANCE_DAYS}
                = "starts Nd late"    if start_diff > {TOLERANCE_DAYS}
                = "starts Nd early"   if start_diff < −{TOLERANCE_DAYS}
```

**inferred_period_pattern**
```
For each active line, compute duration in days
Map that duration to the nearest standard bucket:
  monthly = 30d ± 10,  quarterly = 90d ± 10,  annual = 365d ± 15, etc.
Count votes per bucket → the bucket with the most lines wins
Confidence = winning votes ÷ total active lines × 100%
```
*Example:*
```
Group has 5 active lines with durations: 91d, 89d, 90d, 30d, 92d

Bucket mapping:
  91d → quarterly   (90 ± 10) ✓
  89d → quarterly   (90 ± 10) ✓
  90d → quarterly   (90 ± 10) ✓
  30d → monthly     (30 ± 10) ✓
  92d → quarterly   (90 ± 10) ✓

Votes:  quarterly = 4,  monthly = 1
Winner: quarterly
inferred_period_pattern = "quarterly"
inferred_period_days    = 90
period_confidence_pct   = 4 ÷ 5 × 100 = 80%
The 30-day line is flagged as is_period_outlier = YES
```

**lines_to_add**
```
For each gap (internal gaps + start/end header gaps):
  n = round(gap_days / pattern_days)
  if |gap_days − n × pattern_days| ≤ {TOLERANCE_DAYS} days → add n lines for this gap
lines_to_add = sum of n across all gaps
(0 / blank if pattern is irregular or gap does not fit cleanly)
```
*Example:*
```
Pattern = quarterly (90 days), TOLERANCE_DAYS = {TOLERANCE_DAYS}

Internal gap: 182 days
  n = round(182 ÷ 90) = round(2.02) = 2
  |182 − 2×90| = |182 − 180| = 2 ≤ {TOLERANCE_DAYS} → fits → add 2 lines

Header end gap: 88 days (group ends before header closes)
  n = round(88 ÷ 90) = round(0.98) = 1
  |88 − 1×90| = 2 ≤ {TOLERANCE_DAYS} → fits → add 1 line

Gap of 305 days with annual pattern (365 days):
  n = round(305 ÷ 365) = round(0.84) = 1
  |305 − 1×365| = 60 > {TOLERANCE_DAYS} → does NOT fit cleanly → skip (✗)

lines_to_add = 2 + 1 = 3
```
        """)


# ═══════════════════════════════════════════════════════════════
#  TAB 5 — SOLUTION ANALYSIS
#  One row per group — coverage structure + period pattern for fixing
# ═══════════════════════════════════════════════════════════════
with tab_solution:
    _sol_keys = [col_config["quotation_no"], col_config["catalog_no"]]
    group_sol = result_df.drop_duplicates(subset=_sol_keys).reset_index(drop=True)

    # Summary metrics
    _sol_gap_mask     = pd.to_numeric(group_sol.get("gap_days",     0), errors="coerce") > 0
    _sol_overlap_mask = pd.to_numeric(group_sol.get("overlap_days", 0), errors="coerce") > 0
    _sol_conf_mask    = pd.to_numeric(group_sol.get("period_confidence_pct", 100), errors="coerce") < 70

    sm1, sm2, sm3, sm4 = st.columns(4)
    sm1.metric("Groups with gaps",            int(_sol_gap_mask.sum()))
    sm2.metric("Groups with overlap",         int(_sol_overlap_mask.sum()))
    sm3.metric("Unclear period pattern",      int(_sol_conf_mask.sum()),
               help="Groups where < 70% of active lines agree on the period pattern")
    sm4.metric("Total lines to add",          stats["total_lines_to_add"])

    # Column list: ID → header dates → group dates → coverage → period analysis → solution
    _sol_orig = []
    for _k in ["quotation_no", "catalog_no", "header_start", "header_end"]:
        _c = col_config.get(_k, "")
        if _c and _c in group_sol.columns and _c not in _sol_orig:
            _sol_orig.append(_c)

    _sol_analysis = [
        "groups_in_quotation",
        "group_line_count", "group_active_line_count",
        "group_start", "group_end",
        "actual_coverage_days", "group_span_days",
        "gap_days", "gap_count",
        "overlap_days", "overlap_count",
        "header_aligned", "start_alignment", "end_alignment",
        "coverage_bar",
        "active_line_periods",
        "inferred_period_pattern", "inferred_period_days",
        "avg_period_days",
        "period_confidence_pct",
        "canonical_qty",
        "qty_confidence_pct",
        "active_line_qtys",
        "lines_to_add",
        "gaps_solved_ratio",
        "gap_list",
        "solution_list",
    ]
    _sol_all = _sol_orig + [c for c in _sol_analysis if c in group_sol.columns]
    view_sol = group_sol[[c for c in _sol_all if c in group_sol.columns]]

    # ── Grouped column headers (MultiIndex) ───────────────────────────────────
    # Maps each flat column name to a section label.
    # The section label appears as a spanning group header above related columns.
    _SOL_SECTIONS = [
        ("Identifiers",      [col_config.get("quotation_no",""), col_config.get("catalog_no","")]),
        ("Header Dates",     [col_config.get("header_start",""), col_config.get("header_end",""),
                              "groups_in_quotation"]),
        ("Group Info",       ["group_line_count", "group_active_line_count",
                              "group_start", "group_end"]),
        ("Coverage",         ["actual_coverage_days", "group_span_days",
                              "gap_days", "gap_count",
                              "overlap_days", "overlap_count"]),
        ("Header Alignment", ["header_aligned", "start_alignment", "end_alignment"]),
        ("Visual",           ["coverage_bar", "active_line_periods"]),
        ("Period Analysis",    ["inferred_period_pattern", "inferred_period_days",
                                "avg_period_days", "period_confidence_pct"]),
        ("Quantity Analysis",  ["canonical_qty", "qty_confidence_pct", "active_line_qtys"]),
        ("Solution",           ["lines_to_add", "gaps_solved_ratio",
                                "gap_list", "solution_list"]),
    ]

    # Build MultiIndex tuples in the exact column order of view_sol
    _flat_to_section = {}
    for _section, _cols in _SOL_SECTIONS:
        for _c in _cols:
            _flat_to_section[_c] = _section

    _mi_tuples = [
        (_flat_to_section.get(c, "Other"), c)
        for c in view_sol.columns
    ]
    _mi_cols = pd.MultiIndex.from_tuples(_mi_tuples)

    # Strip datetime columns for display
    _display_sol = view_sol.copy()
    for _c in _display_sol.columns:
        if pd.api.types.is_datetime64_any_dtype(_display_sol[_c]):
            _display_sol[_c] = _display_sol[_c].dt.date

    # Apply cell styling while column names are still flat (style_table uses string names)
    _flat_styles = style_table(
        _display_sol,
        analysis_cols=_sol_analysis,
        key_cols=["gap_days", "gap_count", "lines_to_add", "period_confidence_pct"],
    )

    # Rename both the data and the styles DataFrame to the same MultiIndex
    _display_sol.columns = _mi_cols
    _flat_styles.columns  = _mi_cols

    st.caption(
        f"{len(view_sol):,} groups   ·   "
        "Column sections: Identifiers | Header Dates | Group Info | "
        "Coverage | Header Alignment | Visual | Period Analysis | Solution   ·   "
        "⬜ ERP source   🟦 Analysis result   🔷 Key   🟡 Warning   🔴 Issue"
    )
    try:
        _styled_sol = _display_sol.style.apply(lambda _: _flat_styles, axis=None)
        st.dataframe(_styled_sol, use_container_width=True, height=520)
    except Exception:
        st.dataframe(_display_sol, use_container_width=True, height=520)

    _dl_sol, _ = st.columns([2, 5])
    with _dl_sol:
        st.download_button(
            label="⬇️ Download solution analysis (Excel)",
            data=to_excel_bytes(view_sol),
            file_name="quotation_solution_analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_sol",
        )


# ═══════════════════════════════════════════════════════════════
#  TAB 4 — ANALYSIS LOGIC (About / Metadata)
# ═══════════════════════════════════════════════════════════════
with tab_about:
    st.header("How the Analysis Works")
    st.caption(
        "Reference documentation for every calculated column. "
        "Use this to understand, verify, or challenge the results."
    )

    # ── Current parameters ────────────────────────────────────────────────────
    with st.expander("⚙️ Current analysis parameters", expanded=True):
        pc1, pc2 = st.columns(2)
        with pc1:
            st.markdown(f"**Date tolerance:** `{TOLERANCE_DAYS}` days")
            st.caption(
                "Gaps or misalignments smaller than this are treated as aligned. "
                "Absorbs small administrative date shifts."
            )
        with pc2:
            st.markdown("**Period buckets:**")
            bucket_rows = [
                {"Pattern": name, "Target days": target,
                 "Tolerance ±": tol, "Valid range": f"{target-tol}–{target+tol} days"}
                for name, target, tol in PERIOD_BUCKETS
            ]
            st.dataframe(
                pd.DataFrame(bucket_rows),
                hide_index=True,
                use_container_width=True,
            )
            st.caption(
                "Any duration outside all bucket ranges is labelled **irregular**. "
                "Buckets are non-overlapping by design."
            )

    # ── Groups ────────────────────────────────────────────────────────────────
    with st.expander("🗂 Groups — how they are formed"):
        st.markdown("""
Lines are grouped by the combination of **Quotation_No + Catalog_No**.

Every analysis metric is calculated **per group**, not per individual line.
The same metric value is written to all rows that belong to the group
(so you can filter on any row and see the full group context).
        """)

    # ── Active lines ──────────────────────────────────────────────────────────
    with st.expander("✅ Active lines — what gets included in calculations"):
        st.markdown("""
A line is **excluded** from pattern and coverage analysis if **any** of these is true:

| Condition | Column | Reason |
|---|---|---|
| State is cancelled | `STATE = cancelled` | Cancelled lines are not renewed |
| Closed / placeholder | `C_END_DATE − C_START_DATE < 5 days` | Very short lines (< 5 days) are administrative markers or placeholders |
| Once-period | `C_PERIOD = once` | Single-use lines are not part of a repeating pattern |

All other states (released, created, planned) are treated as **active**.

> `group_active_line_count` tells you how many active lines a group has.
> `group_line_count` counts everything including excluded lines.
        """)

    # ── Period pattern ────────────────────────────────────────────────────────
    with st.expander("📅 Period pattern detection — bucket voting"):
        st.markdown("""
**Why not use average duration?**
If 3 lines are 90 days and 1 is 30 days, the mean = 75 days → wrong.
Bucket voting correctly identifies the pattern as *quarterly* and flags the 30-day line as an outlier.

**Algorithm:**
1. Calculate `duration_days = C_END_DATE − C_START_DATE + 1` for each active line.
2. Map each duration to the nearest period bucket (see parameters above).
3. Count votes per bucket — the bucket with the most lines wins.
4. `period_confidence_pct` = winning votes ÷ total active lines × 100.
5. Active lines **not** in the winning bucket → `is_period_outlier = YES`.

**Irregular pattern:**
When no standard bucket fits (e.g. all lines are 45 days), the pattern is labelled *irregular*
and `inferred_period_days` is set to the median of those durations.
A solution suggestion cannot be auto-calculated for irregular patterns.

**Per-line columns:** `line_duration_days` · `line_period_bucket` · `is_period_outlier`
**Group columns:** `inferred_period_pattern` · `inferred_period_days` · `period_confidence_pct`
        """)

    # ── Quantity pattern ──────────────────────────────────────────────────────
    with st.expander("🔢 Quantity pattern detection — mode voting"):
        st.markdown("""
All active lines in a group should carry the same quantity.

**Algorithm:**
1. Collect quantities (`BUY_QTY_DUE`) from active lines only.
2. Find the most frequent value (mode).
3. `qty_confidence_pct` = mode count ÷ active lines × 100.
4. Lines with a different quantity → `is_qty_outlier = YES`.

**Why not average?**
If 4 lines have qty = 100 and 1 has qty = 200, the mean = 120 → misleading.
The mode correctly identifies 100 as the canonical quantity.

**Per-line columns:** `is_qty_outlier`
**Group columns:** `canonical_qty` · `qty_confidence_pct`
        """)

    # ── Coverage ──────────────────────────────────────────────────────────────
    with st.expander("📐 Coverage analysis — three metrics explained"):
        st.markdown("""
Three metrics are calculated intentionally so differences become visible:

| Column | Formula | What it means |
|---|---|---|
| `group_span_days` | max(end) − min(start) + 1 | Naive total span — **ignores gaps** |
| `actual_coverage_days` | Interval union of active lines | True days actually covered by at least one line |
| `gap_days` | span − coverage | Total uncovered days within the span |
| `gap_count` | — | Number of distinct gaps > 5 days |

**Interval union** merges adjacent or overlapping lines before summing:
- Lines ending Mar 31 and starting Apr 1 are treated as adjacent and merged.
- If `group_span_days > actual_coverage_days` → there are gaps (`gap_days > 0`).
- If lines overlap, `actual_coverage_days < sum of individual durations` (`overlap_days > 0`).

`group_start` and `group_end` are the earliest and latest active line dates.
They represent the **window the group actually covers**, independent of the header.

**Columns:** `group_start` · `group_end` · `group_span_days` · `actual_coverage_days`
· `gap_days` · `gap_count` · `overlap_days` · `overlap_count` · `gap_details`
        """)

    # ── Header alignment ──────────────────────────────────────────────────────
    with st.expander("📌 Header alignment"):
        st.markdown(f"""
The group's coverage window is compared against the quotation header validity window
(`C_PRES_VALID_FROM` → `C_PRES_VALID_TO`).

| Metric | Formula | Interpretation |
|---|---|---|
| Start difference | `group_start − header_start` | Positive = group starts **late**; negative = starts early |
| End difference | `group_end − header_end` | Negative = group ends **early**; positive = ends late |

Within **±{TOLERANCE_DAYS} days** on both sides → `header_aligned = YES`.

**`start_alignment` examples:**
- `aligned` — within tolerance
- `starts 30d late` — group coverage begins 30 days after the header opens
- `starts 15d early` — group coverage begins before the header

**`end_alignment` examples:**
- `ends 30d early` — group coverage ends 30 days before the header closes
- `ends 15d late` — group coverage extends past the header end date

**Columns:** `header_aligned` · `start_alignment` · `end_alignment`
        """)

    # ── Solution suggestion ───────────────────────────────────────────────────
    with st.expander("💡 Solution suggestion — lines to add"):
        st.markdown("""
For each detected gap (internal gaps + header-alignment gaps), the script calculates
how many lines of the inferred pattern would be needed to fill it.

**Algorithm:**
```
n = round(gap_days / pattern_days)
if |gap_days − n × pattern_days| ≤ TOLERANCE_DAYS:
    → "Add N [pattern] lines"   (clean fit)
else:
    → "✗ does not fit cleanly"  (manual review needed)
```

**Example:**
- Pattern = quarterly (90d), gap = 182d → n=2 → |182−180|=2 ≤ 5d → "Add 2 quarterly lines"
- Pattern = quarterly (90d), gap = 62d → n=1 → |62−90|=28 > 5d → manual review
- Pattern = annual (365d), gap = 305d → n=1 → |305−365|=60 > 5d → manual review

**Scope:**
1. Internal gaps (between consecutive active lines)
2. Start gap (group starts after header start)
3. End gap (group ends before header end)

`lines_to_add` = total count across all clean gaps in this group.
`solution_notes` = plain-text description of each gap and its fix.

> Irregular patterns cannot be auto-calculated — manual review required.

**Columns:** `lines_to_add` · `solution_notes`
        """)

    # ── Column reference ──────────────────────────────────────────────────────
    with st.expander("📋 Full column reference"):
        col_ref = [
            ("group_line_count",       "Group",    "Total rows in group (all states incl. cancelled)"),
            ("group_active_line_count","Group",    "Active lines only (excl. cancelled / closed / once)"),
            ("unlimit_qty_count",      "Group",    "Lines with C_UNLIMIT_QTY_DB = true"),
            ("orig_pres_count",        "Group",    "Lines with C_ORIG_PRES_LINE_DB = true (get +365d on renewal)"),
            ("group_start",            "Group",    "Earliest start date among active lines (after header-date substitution)"),
            ("group_end",              "Group",    "Latest end date among active lines"),
            ("group_span_days",        "Group",    "max(end) − min(start) + 1 (naive, includes gaps)"),
            ("actual_coverage_days",   "Group",    "Interval union — true days covered by at least one active line"),
            ("inferred_period_pattern","Group",    "Dominant period pattern (bucket voting)"),
            ("inferred_period_days",   "Group",    "Target days for the inferred pattern (e.g. quarterly = 90)"),
            ("avg_period_days",        "Group",    "Mean duration of active lines (simple average)"),
            ("period_confidence_pct",  "Group",    "% of active lines matching the inferred pattern"),
            ("active_line_periods",    "Group",    "Pipe-separated list of each active line's effective date range"),
            ("canonical_qty",          "Group",    "Most frequent quantity among active lines"),
            ("qty_confidence_pct",     "Group",    "% of active lines with canonical quantity"),
            ("coverage_bar",           "Group",    "48-char visual timeline: █=covered ░=gap ▓=overlap |=line boundary"),
            ("gap_days",               "Group",    "Total uncovered days within the group span (0 = no gaps)"),
            ("gap_count",              "Group",    "Number of distinct gaps detected (0 = continuous)"),
            ("overlap_days",           "Group",    "Days covered by more than one active line simultaneously"),
            ("overlap_count",          "Group",    "Number of overlapping line pairs"),
            ("gap_details",            "Group",    "Each internal gap's start date, end date, and size"),
            ("header_aligned",         "Group",    "YES if group coverage aligns with header (±5d on both ends)"),
            ("start_alignment",        "Group",    "'aligned' or 'starts Nd late/early'"),
            ("end_alignment",          "Group",    "'aligned' or 'ends Nd early/late'"),
            ("lines_to_add",           "Group",    "Total lines needed to fill all cleanly-calculable gaps"),
            ("gaps_solved_ratio",      "Group",    "e.g. '3gap/2l' = 3 gaps found, 2 could be filled by pattern"),
            ("gap_list",               "Group",    "All gaps (internal + header-alignment) with date ranges"),
            ("solution_list",          "Group",    "Parallel to gap_list: '+N pattern' if fits cleanly, '✗' if not"),
            ("line_period_bucket",     "Per line", "This line's mapped period bucket (or 'excluded' if inactive)"),
            ("is_period_outlier",      "Per line", "YES if this line's bucket ≠ group's inferred pattern"),
            ("is_qty_outlier",         "Per line", "YES if this line's quantity ≠ canonical quantity"),
        ]
        st.dataframe(
            pd.DataFrame(col_ref, columns=["Column", "Level", "Description"]),
            hide_index=True,
            use_container_width=True,
        )
