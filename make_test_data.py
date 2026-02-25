"""
Generates test_data.xlsx — a synthetic ERP quotation export covering
35 groups that test every edge case and business scenario in the analysis tool.

Groups by quotation:
  TEST-001 (3 groups): perfect quarterly / Q3 gap / period outlier
  TEST-002 (3 groups): qty inconsistency / overlap / header misalignment
  TEST-003 (3 groups): fully renewable / open-start / renewable+dated mix
  TEST-004 (3 groups): cancelled+placeholder mix / all cancelled / two gaps
  TEST-005 (3 groups): tie-breaking / monthly 12 lines / orig_pres flags
  TEST-006 (3 groups): unlimit_qty flags / annual single / semi-annual
  TEST-007 (2 groups): irregular / mixed exclusions
  TEST-008 (3 groups): 4-month pattern / bi-monthly / starts before header
  TEST-009 (3 groups): 5d gap (tolerated) / 6d gap (detected) / +2 quarterly
  TEST-010 (3 groups): multiple qty outliers / all PLANNED / mixed states
  TEST-011 (3 groups): only once-period / all placeholders / open-end line
  TEST-012 (3 groups): 2 solvable gaps / multi-year annual / ends late
"""
import pandas as pd

NA = None
H     = ('2025-01-01', '2025-12-31')   # standard 1-year header (Jan–Dec 2025)
H_2Y  = ('2025-01-01', '2026-12-31')   # 2-year header
H_APR = ('2025-04-01', '2025-12-31')   # header starting April (for early-start test)
H_HALF = ('2025-01-01', '2025-06-30')  # 6-month header (for ends-late test)
H_2426 = ('2025-01-01', '2026-12-31')  # alias for 2-year (used in tie-break group)


def r(quotation_no, catalog_no, qty, start, end, h_start, h_end,
      state='RELEASED', orig_pres=False, unlimit_qty=NA, renewable=True,
      period=NA, mdq_status='US'):
    return {
        'Quotation_No':        quotation_no,
        'Catalog_No':          catalog_no,
        'Buy_Qty_Due':         qty,
        'Discount':            0,
        'Wanted_Delivery_Date': NA,
        'Price_Freeze_Db':     'FREE',
        'C_Pres_Comment':      NA,
        'C_PERIOD':            period,          # matches COL_PERIOD default
        'C_Period_No':         NA,
        'C_Max_Per_Period':    NA,
        'C_Unlimit_Qty_Db':    unlimit_qty,
        'Tax_Code':            'DK25',
        'Tax_Class_Id':        'DKN',
        'C_Start_Date':        start,
        'C_End_Date':          end,
        'C_Remain_Line_Qty':   qty,
        'C_Rule':              NA,
        'C_First_Delivery_Date': NA,
        'C_Orig_Pres_Line_Db': orig_pres,
        'STATE':               state,
        'C_PRES_VALID_FROM':   h_start,
        'C_PRES_VALID_TO':     h_end,
        'C_RENEWABLE_DB':      renewable,
        'CF_MDQ_PART_STA':     mdq_status,
    }


rows = []

# ════════════════════════════════════════════════════════════════════════════
#  ORIGINAL 20 GROUPS (unchanged)
# ════════════════════════════════════════════════════════════════════════════

# ── GROUP 1: TEST-001 / PROD-A ───────────────────────────────────────────────
# Perfect quarterly, fully aligned — lines added in SHUFFLED order
# Tests: order independence, quarterly 100%, no gap, no overlap, aligned
rows += [
    r('TEST-001', 'PROD-A', 50, '2025-07-01', '2025-09-30', *H),  # Q3 first
    r('TEST-001', 'PROD-A', 50, '2025-01-01', '2025-03-31', *H),  # Q1
    r('TEST-001', 'PROD-A', 50, '2025-10-01', '2025-12-31', *H),  # Q4
    r('TEST-001', 'PROD-A', 50, '2025-04-01', '2025-06-30', *H),  # Q2
]

# ── GROUP 2: TEST-001 / PROD-B ───────────────────────────────────────────────
# Quarterly with Q3 missing — 1 internal gap solvable with +1 quarterly
# Tests: gap detection, lines_to_add=1, gaps_solved_ratio="1gap/1l"
rows += [
    r('TEST-001', 'PROD-B', 30, '2025-01-01', '2025-03-31', *H),  # Q1
    r('TEST-001', 'PROD-B', 30, '2025-04-01', '2025-06-30', *H),  # Q2
    # Q3 (Jul–Sep) missing -> 92d gap
    r('TEST-001', 'PROD-B', 30, '2025-10-01', '2025-12-31', *H),  # Q4
]

# ── GROUP 3: TEST-001 / PROD-C ───────────────────────────────────────────────
# Period outlier: 3 quarterly + 1 monthly-length line (Oct only, 31d)
# Tests: period_confidence=75%, is_period_outlier, gap Nov–Dec
rows += [
    r('TEST-001', 'PROD-C', 10, '2025-01-01', '2025-03-31', *H),  # Q1 quarterly
    r('TEST-001', 'PROD-C', 10, '2025-04-01', '2025-06-30', *H),  # Q2 quarterly
    r('TEST-001', 'PROD-C', 10, '2025-07-01', '2025-09-30', *H),  # Q3 quarterly
    r('TEST-001', 'PROD-C', 10, '2025-10-01', '2025-10-31', *H),  # 31d -> monthly OUTLIER
]

# ── GROUP 4: TEST-002 / PROD-A ───────────────────────────────────────────────
# Qty inconsistency: 3 lines qty=100, Q3 has qty=250
# Tests: canonical_qty=100, qty_confidence=75%, is_qty_outlier on Q3
rows += [
    r('TEST-002', 'PROD-A', 100, '2025-01-01', '2025-03-31', *H),
    r('TEST-002', 'PROD-A', 100, '2025-04-01', '2025-06-30', *H),
    r('TEST-002', 'PROD-A', 250, '2025-07-01', '2025-09-30', *H),  # WRONG QTY
    r('TEST-002', 'PROD-A', 100, '2025-10-01', '2025-12-31', *H),
]

# ── GROUP 5: TEST-002 / PROD-B ───────────────────────────────────────────────
# Overlapping lines: line 2 starts mid-Q1 and overlaps into Q2
# Tests: overlap_days > 0, overlap_count = 1
rows += [
    r('TEST-002', 'PROD-B', 20, '2025-01-01', '2025-03-31', *H),  # Q1
    r('TEST-002', 'PROD-B', 20, '2025-03-01', '2025-05-31', *H),  # OVERLAPS Q1 by 31d
    r('TEST-002', 'PROD-B', 20, '2025-07-01', '2025-09-30', *H),  # Q3
    r('TEST-002', 'PROD-B', 20, '2025-10-01', '2025-12-31', *H),  # Q4
]

# ── GROUP 6: TEST-002 / PROD-C ───────────────────────────────────────────────
# Header misalignment: group covers only Q2+Q3, header is full year
# Tests: starts 89d late, ends 92d early, header_aligned=NO
rows += [
    r('TEST-002', 'PROD-C', 15, '2025-04-01', '2025-06-30', *H),  # Q2
    r('TEST-002', 'PROD-C', 15, '2025-07-01', '2025-09-30', *H),  # Q3
]

# ── GROUP 7: TEST-003 / PROD-A ───────────────────────────────────────────────
# Fully renewable line: both dates null -> covers full header period
# Tests: null-date handling, effective dates = header dates, annual pattern
rows += [
    r('TEST-003', 'PROD-A', 5, NA, NA, *H),  # renewable — full year
]

# ── GROUP 8: TEST-003 / PROD-B ───────────────────────────────────────────────
# Open-start line: start null, end given -> effective start = header start
# Tests: partial null date, effective start fills from header
rows += [
    r('TEST-003', 'PROD-B', 5, NA,           '2025-06-30', *H),  # open-start -> Jan–Jun
    r('TEST-003', 'PROD-B', 5, '2025-07-01', '2025-12-31', *H),  # normal H2
]

# ── GROUP 9: TEST-003 / PROD-C ───────────────────────────────────────────────
# Renewable + dated mix: 1 renewable (full year) + 4 quarterly dated lines
# Tests: overlap_days large (renewable duplicates everything), overlap_count
rows += [
    r('TEST-003', 'PROD-C', 8, NA,           NA,           *H),  # renewable (full year)
    r('TEST-003', 'PROD-C', 8, '2025-01-01', '2025-03-31', *H),  # Q1 (duplicate coverage)
    r('TEST-003', 'PROD-C', 8, '2025-04-01', '2025-06-30', *H),  # Q2
    r('TEST-003', 'PROD-C', 8, '2025-07-01', '2025-09-30', *H),  # Q3
    r('TEST-003', 'PROD-C', 8, '2025-10-01', '2025-12-31', *H),  # Q4
]

# ── GROUP 10: TEST-004 / PROD-A ──────────────────────────────────────────────
# Cancelled + short placeholder lines mixed with real active lines
# Tests: exclusion logic — cancelled and <5d lines excluded from analysis
rows += [
    r('TEST-004', 'PROD-A', 25, '2025-01-01', '2025-03-31', *H, state='cancelled'),   # excluded
    r('TEST-004', 'PROD-A', 25, '2025-04-01', '2025-06-30', *H, state='cancelled'),   # excluded
    r('TEST-004', 'PROD-A', 25, '2025-01-02', '2025-01-03', *H),  # 2d placeholder  -> excluded
    r('TEST-004', 'PROD-A', 25, '2025-04-01', '2025-04-02', *H),  # 2d placeholder  -> excluded
    r('TEST-004', 'PROD-A', 25, '2025-07-01', '2025-09-30', *H),  # ACTIVE Q3
    r('TEST-004', 'PROD-A', 25, '2025-10-01', '2025-12-31', *H),  # ACTIVE Q4
]

# ── GROUP 11: TEST-004 / PROD-B ──────────────────────────────────────────────
# All lines cancelled — nothing to analyse
# Tests: n_active=0 code path, all metrics show N/A
rows += [
    r('TEST-004', 'PROD-B', 12, '2025-01-01', '2025-03-31', *H, state='cancelled'),
    r('TEST-004', 'PROD-B', 12, '2025-04-01', '2025-06-30', *H, state='cancelled'),
    r('TEST-004', 'PROD-B', 12, '2025-07-01', '2025-09-30', *H, state='cancelled'),
    r('TEST-004', 'PROD-B', 12, '2025-10-01', '2025-12-31', *H, state='cancelled'),
]

# ── GROUP 12: TEST-004 / PROD-C ──────────────────────────────────────────────
# Two gaps: one NOT solvable (45d ÷ 90 = 0.5 -> rounds to 0, or 1 but |45-90|=45 > 5),
# one solvable (from Aug 15 to Dec 31 via Q4)
# Tests: gap_count=2, solution_list mixes ✗ and +1 quarterly
rows += [
    r('TEST-004', 'PROD-C', 40, '2025-01-01', '2025-03-31', *H),  # Q1
    r('TEST-004', 'PROD-C', 40, '2025-04-01', '2025-06-30', *H),  # Q2
    # 45d gap: Jul 1 -> Aug 14  (|45-90|=45 > 5 -> not solvable)
    r('TEST-004', 'PROD-C', 40, '2025-08-15', '2025-09-30', *H),  # 47d filler
    # 0d gap after Sep 30 (Oct 1) within tolerance
    r('TEST-004', 'PROD-C', 40, '2025-10-01', '2025-12-31', *H),  # Q4
]

# ── GROUP 13: TEST-005 / PROD-A ──────────────────────────────────────────────
# Tie-breaking test: annual (365d) vs bi-monthly (60d) — annual must win
# Tests: tiebreaker by total coverage days (365 > 60)
rows += [
    r('TEST-005', 'PROD-A', 3, '2025-01-01', '2025-03-01', *H_2426),  # ~60d -> bi-monthly
    r('TEST-005', 'PROD-A', 3, '2026-01-01', '2026-12-31', *H_2426),  # 365d -> annual
]

# ── GROUP 14: TEST-005 / PROD-B ──────────────────────────────────────────────
# Monthly pattern: 12 consecutive monthly lines, full year
# Tests: monthly 100%, no gaps, aligned, coverage_bar shows 12 segments
rows += [
    r('TEST-005', 'PROD-B', 60, '2025-01-01', '2025-01-31', *H),
    r('TEST-005', 'PROD-B', 60, '2025-02-01', '2025-02-28', *H),
    r('TEST-005', 'PROD-B', 60, '2025-03-01', '2025-03-31', *H),
    r('TEST-005', 'PROD-B', 60, '2025-04-01', '2025-04-30', *H),
    r('TEST-005', 'PROD-B', 60, '2025-05-01', '2025-05-31', *H),
    r('TEST-005', 'PROD-B', 60, '2025-06-01', '2025-06-30', *H),
    r('TEST-005', 'PROD-B', 60, '2025-07-01', '2025-07-31', *H),
    r('TEST-005', 'PROD-B', 60, '2025-08-01', '2025-08-31', *H),
    r('TEST-005', 'PROD-B', 60, '2025-09-01', '2025-09-30', *H),
    r('TEST-005', 'PROD-B', 60, '2025-10-01', '2025-10-31', *H),
    r('TEST-005', 'PROD-B', 60, '2025-11-01', '2025-11-30', *H),
    r('TEST-005', 'PROD-B', 60, '2025-12-01', '2025-12-31', *H),
]

# ── GROUP 15: TEST-005 / PROD-C ──────────────────────────────────────────────
# C_Orig_Pres_Line_Db flag: 2 of 4 quarterly lines flagged
# Tests: orig_pres_count = 2
rows += [
    r('TEST-005', 'PROD-C', 75, '2025-01-01', '2025-03-31', *H, orig_pres=True),
    r('TEST-005', 'PROD-C', 75, '2025-04-01', '2025-06-30', *H, orig_pres=True),
    r('TEST-005', 'PROD-C', 75, '2025-07-01', '2025-09-30', *H, orig_pres=False),
    r('TEST-005', 'PROD-C', 75, '2025-10-01', '2025-12-31', *H, orig_pres=False),
]

# ── GROUP 16: TEST-006 / PROD-A ──────────────────────────────────────────────
# C_Unlimit_Qty_Db flag: 3 of 4 quarterly lines flagged
# Tests: unlimit_qty_count = 3
rows += [
    r('TEST-006', 'PROD-A', 90, '2025-01-01', '2025-03-31', *H, unlimit_qty='1'),
    r('TEST-006', 'PROD-A', 90, '2025-04-01', '2025-06-30', *H, unlimit_qty='1'),
    r('TEST-006', 'PROD-A', 90, '2025-07-01', '2025-09-30', *H, unlimit_qty='1'),
    r('TEST-006', 'PROD-A', 90, '2025-10-01', '2025-12-31', *H, unlimit_qty=NA),
]

# ── GROUP 17: TEST-006 / PROD-B ──────────────────────────────────────────────
# Annual pattern: single line covers full year
# Tests: annual 100%, no gap, aligned, single active line
rows += [
    r('TEST-006', 'PROD-B', 200, '2025-01-01', '2025-12-31', *H),
]

# ── GROUP 18: TEST-006 / PROD-C ──────────────────────────────────────────────
# Semi-annual pattern: 2 lines each ~180d
# Tests: semi-annual 100%, no gaps, aligned
rows += [
    r('TEST-006', 'PROD-C', 500, '2025-01-01', '2025-06-30', *H),  # 181d
    r('TEST-006', 'PROD-C', 500, '2025-07-01', '2025-12-31', *H),  # 184d
]

# ── GROUP 19: TEST-007 / PROD-A ──────────────────────────────────────────────
# Irregular pattern: durations 45d, 55d, 70d — none fit standard buckets
# Tests: inferred_period_pattern='irregular', lines_to_add=0 (can't auto-solve)
H3 = ('2025-01-01', '2025-09-30')
rows += [
    r('TEST-007', 'PROD-A', 11, '2025-01-01', '2025-02-14', *H3),  # 45d
    r('TEST-007', 'PROD-A', 11, '2025-02-15', '2025-04-10', *H3),  # 55d
    r('TEST-007', 'PROD-A', 11, '2025-04-11', '2025-06-19', *H3),  # 70d
]

# ── GROUP 20: TEST-007 / PROD-B ──────────────────────────────────────────────
# Mixed exclusion reasons: 1 cancelled, 1 once-period, 2 active quarterly
# Tests: group_line_count=4, group_active_line_count=2, once-period excluded
rows += [
    r('TEST-007', 'PROD-B', 33, '2025-01-01', '2025-03-31', *H, state='cancelled'),
    r('TEST-007', 'PROD-B', 33, '2025-04-01', '2025-04-30', *H, period='once'),
    r('TEST-007', 'PROD-B', 33, '2025-07-01', '2025-09-30', *H),
    r('TEST-007', 'PROD-B', 33, '2025-10-01', '2025-12-31', *H),
]


# ════════════════════════════════════════════════════════════════════════════
#  NEW GROUPS 21–35
# ════════════════════════════════════════════════════════════════════════════

# ── GROUP 21: TEST-008 / PROD-A ──────────────────────────────────────────────
# 4-month (120d) pattern: 3 lines covering full year
#   Jan 1 -> Apr 30 = 120d  (within 120 ± 10 ✓)
#   May 1 -> Aug 31 = 123d  (within 120 ± 10 ✓)
#   Sep 1 -> Dec 31 = 122d  (within 120 ± 10 ✓)
# Tests: 4-month bucket, 100% confidence, no gap, aligned
rows += [
    r('TEST-008', 'PROD-A', 80, '2025-01-01', '2025-04-30', *H),  # 120d
    r('TEST-008', 'PROD-A', 80, '2025-05-01', '2025-08-31', *H),  # 123d
    r('TEST-008', 'PROD-A', 80, '2025-09-01', '2025-12-31', *H),  # 122d
]

# ── GROUP 22: TEST-008 / PROD-B ──────────────────────────────────────────────
# Bi-monthly (60d) pattern: 6 lines covering full year
#   Jan 1 -> Feb 28 = 59d, Mar 1 -> Apr 30 = 61d, May 1 -> Jun 30 = 61d
#   Jul 1 -> Aug 31 = 62d, Sep 1 -> Oct 31 = 61d, Nov 1 -> Dec 31 = 61d
# Tests: bi-monthly bucket, 100% confidence, no gap, aligned
rows += [
    r('TEST-008', 'PROD-B', 45, '2025-01-01', '2025-02-28', *H),  # 59d
    r('TEST-008', 'PROD-B', 45, '2025-03-01', '2025-04-30', *H),  # 61d
    r('TEST-008', 'PROD-B', 45, '2025-05-01', '2025-06-30', *H),  # 61d
    r('TEST-008', 'PROD-B', 45, '2025-07-01', '2025-08-31', *H),  # 62d
    r('TEST-008', 'PROD-B', 45, '2025-09-01', '2025-10-31', *H),  # 61d
    r('TEST-008', 'PROD-B', 45, '2025-11-01', '2025-12-31', *H),  # 61d
]

# ── GROUP 23: TEST-008 / PROD-C ──────────────────────────────────────────────
# Group starts BEFORE header (early start misalignment)
# Header: Apr 1 – Dec 31.  Q1 (Jan–Mar) is before the header opens.
# start_diff = Jan 1 − Apr 1 = −90 -> "starts 90d early"
# end_diff   = Dec 31 − Dec 31 = 0  -> aligned
# Tests: start_alignment="starts 90d early", header_aligned=NO
rows += [
    r('TEST-008', 'PROD-C', 20, '2025-01-01', '2025-03-31', *H_APR),  # Q1 — before header
    r('TEST-008', 'PROD-C', 20, '2025-04-01', '2025-06-30', *H_APR),  # Q2
    r('TEST-008', 'PROD-C', 20, '2025-07-01', '2025-09-30', *H_APR),  # Q3
    r('TEST-008', 'PROD-C', 20, '2025-10-01', '2025-12-31', *H_APR),  # Q4
]

# ── GROUP 24: TEST-009 / PROD-A ──────────────────────────────────────────────
# Gap EXACTLY at TOLERANCE_DAYS (5d) -> should NOT be flagged as a gap
# Q1 ends Mar 31; next line starts Apr 6 -> gap = (Apr 6 − Mar 31) − 1 = 5 days
# 5 > TOLERANCE_DAYS(5) = False -> gap ignored, has_gaps = NO
# Tests: tolerance boundary — 5d gap is within tolerance
rows += [
    r('TEST-009', 'PROD-A', 55, '2025-01-01', '2025-03-31', *H),  # Q1
    r('TEST-009', 'PROD-A', 55, '2025-04-06', '2025-06-30', *H),  # 5d gap before (tolerated)
    r('TEST-009', 'PROD-A', 55, '2025-07-01', '2025-09-30', *H),  # Q3
    r('TEST-009', 'PROD-A', 55, '2025-10-01', '2025-12-31', *H),  # Q4
]

# ── GROUP 25: TEST-009 / PROD-B ──────────────────────────────────────────────
# Gap JUST OUTSIDE TOLERANCE (6d) -> IS flagged as a gap
# Q1 ends Mar 31; next line starts Apr 7 -> gap = (Apr 7 − Mar 31) − 1 = 6 days
# 6 > TOLERANCE_DAYS(5) = True -> gap detected
# Tests: tolerance boundary — 6d gap is outside tolerance, gap detected
rows += [
    r('TEST-009', 'PROD-B', 55, '2025-01-01', '2025-03-31', *H),  # Q1
    r('TEST-009', 'PROD-B', 55, '2025-04-07', '2025-06-30', *H),  # 6d gap before (detected)
    r('TEST-009', 'PROD-B', 55, '2025-07-01', '2025-09-30', *H),  # Q3
    r('TEST-009', 'PROD-B', 55, '2025-10-01', '2025-12-31', *H),  # Q4
]

# ── GROUP 26: TEST-009 / PROD-C ──────────────────────────────────────────────
# Q2 + Q3 both missing -> single internal gap of 183 days
# gap_days = Apr 1 -> Sep 30 = 183d
# solve_gap: n = round(183/90) = 2, |183 − 2×90| = |183−180| = 3 ≤ 5 -> +2 quarterly
# Tests: lines_to_add=2, gaps_solved_ratio="1gap/2l", solution_list="+2 quarterly"
rows += [
    r('TEST-009', 'PROD-C', 35, '2025-01-01', '2025-03-31', *H),  # Q1
    # Q2 + Q3 both missing (Apr 1 – Sep 30 = 183d gap)
    r('TEST-009', 'PROD-C', 35, '2025-10-01', '2025-12-31', *H),  # Q4
]

# ── GROUP 27: TEST-010 / PROD-A ──────────────────────────────────────────────
# Multiple qty outliers: 8 quarterly lines over 2 years, 6×qty=100, 2×qty=50
# canonical_qty=100, qty_confidence=75%, 2 lines is_qty_outlier=YES
# Tests: more than one outlier, confidence 75%
rows += [
    r('TEST-010', 'PROD-A', 100, '2025-01-01', '2025-03-31', *H_2Y),  # Q1 2025
    r('TEST-010', 'PROD-A',  50, '2025-04-01', '2025-06-30', *H_2Y),  # WRONG QTY
    r('TEST-010', 'PROD-A', 100, '2025-07-01', '2025-09-30', *H_2Y),  # Q3 2025
    r('TEST-010', 'PROD-A', 100, '2025-10-01', '2025-12-31', *H_2Y),  # Q4 2025
    r('TEST-010', 'PROD-A', 100, '2026-01-01', '2026-03-31', *H_2Y),  # Q1 2026
    r('TEST-010', 'PROD-A',  50, '2026-04-01', '2026-06-30', *H_2Y),  # WRONG QTY
    r('TEST-010', 'PROD-A', 100, '2026-07-01', '2026-09-30', *H_2Y),  # Q3 2026
    r('TEST-010', 'PROD-A', 100, '2026-10-01', '2026-12-31', *H_2Y),  # Q4 2026
]

# ── GROUP 28: TEST-010 / PROD-B ──────────────────────────────────────────────
# All lines in PLANNED state — should all be treated as active (not cancelled)
# Tests: PLANNED state = active, quarterly 100%, no issues, state exclusion only for 'cancelled'
rows += [
    r('TEST-010', 'PROD-B', 70, '2025-01-01', '2025-03-31', *H, state='PLANNED'),
    r('TEST-010', 'PROD-B', 70, '2025-04-01', '2025-06-30', *H, state='PLANNED'),
    r('TEST-010', 'PROD-B', 70, '2025-07-01', '2025-09-30', *H, state='PLANNED'),
    r('TEST-010', 'PROD-B', 70, '2025-10-01', '2025-12-31', *H, state='PLANNED'),
]

# ── GROUP 29: TEST-010 / PROD-C ──────────────────────────────────────────────
# Mixed states: CREATED / RELEASED / PLANNED — all active (none are 'cancelled')
# Tests: mixed states all active, quarterly 100%, no issues
rows += [
    r('TEST-010', 'PROD-C', 90, '2025-01-01', '2025-03-31', *H, state='CREATED'),
    r('TEST-010', 'PROD-C', 90, '2025-04-01', '2025-06-30', *H, state='RELEASED'),
    r('TEST-010', 'PROD-C', 90, '2025-07-01', '2025-09-30', *H, state='PLANNED'),
    r('TEST-010', 'PROD-C', 90, '2025-10-01', '2025-12-31', *H, state='RELEASED'),
]

# ── GROUP 30: TEST-011 / PROD-A ──────────────────────────────────────────────
# Only once-period lines -> ALL excluded -> 0 active lines
# Tests: C_PERIOD='once' exclusion, n_active=0, all group metrics = N/A
rows += [
    r('TEST-011', 'PROD-A', 15, '2025-01-01', '2025-03-31', *H, period='once'),
    r('TEST-011', 'PROD-A', 15, '2025-04-01', '2025-06-30', *H, period='once'),
]

# ── GROUP 31: TEST-011 / PROD-B ──────────────────────────────────────────────
# All lines < 5 days -> ALL excluded as placeholders -> 0 active lines
# Durations: 2d, 1d, 3d, 2d — all below the 5-day threshold
# Tests: duration exclusion, n_active=0, all group metrics = N/A
rows += [
    r('TEST-011', 'PROD-B', 5, '2025-01-01', '2025-01-02', *H),  # 2d -> excluded
    r('TEST-011', 'PROD-B', 5, '2025-02-01', '2025-02-01', *H),  # 1d -> excluded
    r('TEST-011', 'PROD-B', 5, '2025-04-01', '2025-04-03', *H),  # 3d -> excluded
    r('TEST-011', 'PROD-B', 5, '2025-07-01', '2025-07-02', *H),  # 2d -> excluded
]

# ── GROUP 32: TEST-011 / PROD-C ──────────────────────────────────────────────
# Open-end line (start given, end null) mixed with a dated semi-annual line
# Line 1: Jan 1, end=null -> effective end = header Dec 31 -> covers full year (365d -> annual)
# Line 2: Jan 1 -> Jun 30 (181d -> semi-annual) -> overlaps Line 1
# Tie: 1 annual vs 1 semi-annual -> tiebreaker: annual has 365d > 181d -> annual wins
# Tests: open-end handling, overlap from open-end + dated, tie-breaking
rows += [
    r('TEST-011', 'PROD-C', 22, '2025-01-01', NA,           *H),  # open-end -> full year
    r('TEST-011', 'PROD-C', 22, '2025-01-01', '2025-06-30', *H),  # 181d -> semi-annual, overlaps
]

# ── GROUP 33: TEST-012 / PROD-A ──────────────────────────────────────────────
# Two solvable gaps: Q1 + Q3 present, Q2 (internal) and Q4 (end) both missing
# Internal gap: Apr 1 -> Jun 30 = 91d -> |91−90| = 1 ≤ 5 -> +1 quarterly
# End gap:      Oct 1 -> Dec 31 = 92d -> |92−90| = 2 ≤ 5 -> +1 quarterly
# Tests: gap_count=1 (internal only), lines_to_add=2, gaps_solved_ratio="2gap/2l"
rows += [
    r('TEST-012', 'PROD-A', 60, '2025-01-01', '2025-03-31', *H),  # Q1
    # Q2 missing (Apr 1 – Jun 30 = 91d internal gap)
    r('TEST-012', 'PROD-A', 60, '2025-07-01', '2025-09-30', *H),  # Q3
    # Q4 missing (Oct 1 – Dec 31 = 92d end gap vs header)
]

# ── GROUP 34: TEST-012 / PROD-B ──────────────────────────────────────────────
# Multi-year header (2026): 2 annual lines perfectly aligned
# Header: Jan 2025 – Dec 2026 (2 years).  Two annual lines, one per year.
# Tests: multi-year scenario, annual 100%, no gaps, header fully aligned
rows += [
    r('TEST-012', 'PROD-B', 300, '2025-01-01', '2025-12-31', *H_2Y),  # Year 1 (365d)
    r('TEST-012', 'PROD-B', 300, '2026-01-01', '2026-12-31', *H_2Y),  # Year 2 (365d)
]

# ── GROUP 35: TEST-012 / PROD-C ──────────────────────────────────────────────
# Group ends AFTER header (late end misalignment)
# Header: Jan 1 – Jun 30 (6-month).  Third quarterly line extends into Q3 (past header).
# end_diff = Sep 30 − Jun 30 = +92 -> "ends 92d late"
# start_diff = Jan 1 − Jan 1 = 0 -> aligned
# Tests: end_alignment="ends 92d late", header_aligned=NO
rows += [
    r('TEST-012', 'PROD-C', 110, '2025-01-01', '2025-03-31', *H_HALF),  # Q1
    r('TEST-012', 'PROD-C', 110, '2025-04-01', '2025-06-30', *H_HALF),  # Q2 (last in header)
    r('TEST-012', 'PROD-C', 110, '2025-07-01', '2025-09-30', *H_HALF),  # Q3 -> PAST header end
]


# ════════════════════════════════════════════════════════════════════════════
#  WRITE FILE + SUMMARY
# ════════════════════════════════════════════════════════════════════════════

df = pd.DataFrame(rows)
df.to_excel('test_data.xlsx', index=False)

groups = list(df.groupby(['Quotation_No', 'Catalog_No']))
print(f'Created test_data.xlsx: {len(df)} rows, {len(groups)} groups')
print()

descriptions = {
    ('TEST-001', 'PROD-A'): 'Perfect quarterly, shuffled order',
    ('TEST-001', 'PROD-B'): 'Q3 gap -> +1 quarterly',
    ('TEST-001', 'PROD-C'): 'Period outlier (monthly in quarterly group)',
    ('TEST-002', 'PROD-A'): 'Qty outlier (Q3 qty=250, rest=100)',
    ('TEST-002', 'PROD-B'): 'Overlapping lines (line 2 overlaps Q1)',
    ('TEST-002', 'PROD-C'): 'Header misaligned (only Q2+Q3, full-year header)',
    ('TEST-003', 'PROD-A'): 'Fully renewable (both dates null)',
    ('TEST-003', 'PROD-B'): 'Open-start line (start null -> fills from header)',
    ('TEST-003', 'PROD-C'): 'Renewable + dated mix (large overlap)',
    ('TEST-004', 'PROD-A'): 'Cancelled + placeholder mix (2 active lines remain)',
    ('TEST-004', 'PROD-B'): 'All cancelled -> n_active=0, all N/A',
    ('TEST-004', 'PROD-C'): '2 gaps: one unsolvable (45d), one solvable (Q4)',
    ('TEST-005', 'PROD-A'): 'Tie-breaking: annual (365d) beats bi-monthly (60d)',
    ('TEST-005', 'PROD-B'): 'Monthly pattern, 12 lines',
    ('TEST-005', 'PROD-C'): 'orig_pres_count=2 (2 of 4 lines flagged)',
    ('TEST-006', 'PROD-A'): 'unlimit_qty_count=3 (3 of 4 lines flagged)',
    ('TEST-006', 'PROD-B'): 'Annual, single line, aligned',
    ('TEST-006', 'PROD-C'): 'Semi-annual, 2 lines, aligned',
    ('TEST-007', 'PROD-A'): 'Irregular pattern (45d/55d/70d)',
    ('TEST-007', 'PROD-B'): 'Mixed exclusions: 1 cancelled + 1 once + 2 active',
    ('TEST-008', 'PROD-A'): '4-month pattern (120d), 3 lines',
    ('TEST-008', 'PROD-B'): 'Bi-monthly pattern (60d), 6 lines',
    ('TEST-008', 'PROD-C'): 'Starts 90d BEFORE header (early start)',
    ('TEST-009', 'PROD-A'): '5d gap = exactly TOLERANCE -> not flagged',
    ('TEST-009', 'PROD-B'): '6d gap = TOLERANCE+1 -> detected as gap',
    ('TEST-009', 'PROD-C'): 'Q2+Q3 missing -> 183d gap -> +2 quarterly',
    ('TEST-010', 'PROD-A'): '2 qty outliers (6×100 + 2×50, 8 lines, 2-year header)',
    ('TEST-010', 'PROD-B'): 'All PLANNED state -> all active, quarterly 100%',
    ('TEST-010', 'PROD-C'): 'Mixed CREATED/RELEASED/PLANNED -> all active',
    ('TEST-011', 'PROD-A'): 'Only once-period lines -> n_active=0',
    ('TEST-011', 'PROD-B'): 'All lines <5 days -> n_active=0 (all placeholders)',
    ('TEST-011', 'PROD-C'): 'Open-end line + dated -> overlap, annual wins tie',
    ('TEST-012', 'PROD-A'): 'Q1+Q3 only -> 2 solvable gaps -> lines_to_add=2',
    ('TEST-012', 'PROD-B'): 'Multi-year header, 2 annual lines, aligned',
    ('TEST-012', 'PROD-C'): 'Ends 92d AFTER header (late end)',
}

for (q, c), g in groups:
    desc = descriptions.get((q, c), '')
    print(f'  {q} / {c:8s}  {len(g):2d} lines   {desc}')
