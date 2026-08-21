"""Pivot All Region output writer (Phase 1).

Template round-trip: load a copy of the prior-month/golden workbook and stamp
the computed values into three sheets — Data, Pivot Data, Commission — then save
to a new path. Embedded pivots/formulas become static values (the workbook is
NOT refreshed by Excel here), per the PROJECT_CONTEXT §1 decision.

openpyxl loads the ~4 MB workbook slowly (~85 s); a full round-trip is ~3 min.
"""
from __future__ import annotations
import os
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string as ci

import par_pipeline as par
from par_pipeline import DATA_COLUMNS

DATA_HEADER_ROW = 1          # Data: header row 1, values from row 2
COMMISSION_LABEL_ROW = 3     # Commission: labels row 3, values from row 4
COMMISSION_DATA_START = 4
LAST_COMMISSION_COL = "CV"
PIVOT_HEADER_ROW = 4         # Pivot Data: header row 4, values from row 5
# Pivot Data measure columns, in the golden sheet's order (NO Total Weight).
PIVOT_MEASURES = ["Total Parcel", "Total Delivery Fee", "Total Commission",
                  "Drop off to SH Parcel", "Drop off to SH Commission", "Insurance Fee"]


# --- Adjustment & Penalty sheet layout (verified vs golden M06Y2026) -----------
ADJ_SHEET = "Adjustment&Penalty"
ADJ_HEADER_ROW = 3               # labels row 3, values from row 4
ADJ_DATA_START = 4
ADJ_SUMMARY_ROW = 2              # SUM row above the headers
# main table A:J, outstanding L:Y, and three pivot blocks (start-col, measure count).
ADJ_MAIN_COLS = list(range(1, 11))            # A..J
ADJ_OS_COLS = list(range(12, 26))             # L..Y
ADJ_PIVOT_DPMS_C0 = ci("AA")                  # AA..AF  (label + 5 measures)
ADJ_PIVOT_DPID_C0 = ci("AH")                  # AH..AM
ADJ_PIVOT_OS_C0 = ci("AO")                    # AO..AR  (label + 3 measures)


def _cell_val(v):
    """Coerce a value to something openpyxl can store (no numpy scalars / NaN)."""
    if v is None:
        return None
    if isinstance(v, float) and np.isnan(v):
        return None
    if isinstance(v, np.integer):
        return int(v)
    if isinstance(v, np.floating):
        f = float(v)
        return None if np.isnan(f) else f
    if isinstance(v, float) and v.is_integer():
        return v
    return v


def _isnum(x) -> bool:
    try:
        float(x)
        return True
    except (TypeError, ValueError):
        return False


def _set(ws, row, col, value):
    """Assign a cell value. NOTE: ws.cell(row, col, value=None) is a no-op in
    openpyxl (it only assigns when value is not None), so always assign via the
    .value attribute — this is what makes clears (value=None) actually stick."""
    ws.cell(row=row, column=col).value = value


def write_pivot_all_region(template_path, out_path, data_df, commission_df,
                           adj_result=None, verbose=True) -> str:
    """Write Data / Pivot Data / Commission (and, if given, Adjustment&Penalty) as
    values into a copy of the template.

    `data_df`       16-col Data frame (par.DATA_COLUMNS).
    `commission_df` one row per DPID, columns keyed by Excel letter (A..CV).
    `adj_result`    optional dict from par_adjustment.build_adjustment_penalty.
    """
    if verbose:
        print(f"[writer] loading template: {template_path}")
    wb = openpyxl.load_workbook(template_path)
    _write_data(wb, data_df, verbose)
    _write_pivot(wb, data_df, verbose)
    _write_commission(wb, commission_df, verbose)
    if adj_result is not None:
        _write_adjustment(wb, adj_result, verbose)
    if verbose:
        print(f"[writer] saving: {out_path}")
    wb.save(out_path)
    return out_path


def _write_data(wb, data_df, verbose):
    ws = wb["Data"]
    r0 = DATA_HEADER_ROW + 1
    n = len(data_df)
    ncol = len(DATA_COLUMNS)
    clear_to = max(ws.max_row, r0 + n - 1)
    for r in range(r0, clear_to + 1):               # clear old body first
        for j in range(1, ncol + 1):
            _set(ws, r, j, None)
    for i, (_, row) in enumerate(data_df.iterrows()):
        r = r0 + i
        for j, col in enumerate(DATA_COLUMNS, start=1):
            _set(ws, r, j, _cell_val(row[col]))
    if verbose:
        print(f"[writer] Data: cleared rows {r0}..{clear_to}, wrote {n} rows (rows {r0}..{r0 + n - 1})")


def _write_pivot(wb, data_df, verbose):
    """Replace the live PivotTable with a static flat table (DPID x Mapping Type)."""
    ws = wb["Pivot Data"]
    try:                                            # drop the pivot definition so Excel won't reshape it
        ws._pivots = []
    except Exception:
        pass
    old_max = ws.max_row
    r0 = PIVOT_HEADER_ROW + 1

    d = data_df.copy()
    d["_k"] = par._norm_id(d["DPID"])
    for m in PIVOT_MEASURES:
        d[m] = pd.to_numeric(d[m], errors="coerce").fillna(0.0)
    g = (d.groupby(["_k", "Mapping Type"], sort=False)[PIVOT_MEASURES]
           .sum().reset_index())
    g["_sort"] = pd.to_numeric(g["_k"], errors="coerce")
    g = g.sort_values(["_sort", "_k", "Mapping Type"]).reset_index(drop=True)

    # clear old body
    clear_to = max(old_max, r0 + len(g) - 1)
    for r in range(r0, clear_to + 1):
        for c in range(1, len(PIVOT_MEASURES) + 3):
            _set(ws, r, c, None)

    r = r0
    prev_k = object()
    for _, row in g.iterrows():
        k = row["_k"]
        show = (k != prev_k)
        _set(ws, r, 1, (float(k) if _isnum(k) else k) if show else None)   # DPID once per group
        _set(ws, r, 2, row["Mapping Type"])
        for j, m in enumerate(PIVOT_MEASURES, start=3):
            _set(ws, r, j, _cell_val(row[m]))
        prev_k = k
        r += 1
    if verbose:
        print(f"[writer] Pivot Data: wrote {len(g)} rows (rows {r0}..{r - 1})")


def _find_last_row(ws, r0, key_col=1, max_blank=5):
    last, r, blanks = r0 - 1, r0, 0
    while r <= ws.max_row and blanks < max_blank:
        if ws.cell(row=r, column=key_col).value not in (None, ""):
            last, blanks = r, 0
        else:
            blanks += 1
        r += 1
    return last


def write_adjustment_penalty(template_path, out_path, adj_result, verbose=True) -> str:
    """Stand-alone round-trip writer for just the Adjustment&Penalty sheet.

    `adj_result` is the dict from par_adjustment.build_adjustment_penalty
    (keys: main, outstanding, pivot_dpms, pivot_dpid, pivot_os).
    """
    if verbose:
        print(f"[writer] loading template: {template_path}")
    wb = openpyxl.load_workbook(template_path)
    _write_adjustment(wb, adj_result, verbose)
    if verbose:
        print(f"[writer] saving: {out_path}")
    wb.save(out_path)
    return out_path


def _clear_block(ws, r0, cols, extra_blank_rows=0):
    """Clear cols from row r0 down to the block's current last populated row."""
    last = _find_last_row(ws, r0, key_col=cols[0])
    clear_to = last + extra_blank_rows
    for r in range(r0, clear_to + 1):
        for c in cols:
            _set(ws, r, c, None)
    return clear_to


def _write_table(ws, df, r0, cols):
    """Write df (in column order) starting at row r0 across the given 1-based cols."""
    for i, (_, row) in enumerate(df.iterrows()):
        r = r0 + i
        for c, val in zip(cols, row.tolist()):
            _set(ws, r, c, _cell_val(val))
    return r0 + len(df) - 1


def _write_adjustment(wb, res, verbose):
    if ADJ_SHEET not in wb.sheetnames:
        raise KeyError(f"sheet '{ADJ_SHEET}' not in template")
    ws = wb[ADJ_SHEET]
    try:                                             # neutralise the 3 embedded pivots
        ws._pivots = []
    except Exception:
        pass

    main = res["main"]; ost = res["outstanding"]
    pv_dpms = res["pivot_dpms"]; pv_dpid = res["pivot_dpid"]; pv_os = res["pivot_os"]

    # main table A:J
    _clear_block(ws, ADJ_DATA_START, ADJ_MAIN_COLS)
    _write_table(ws, main, ADJ_DATA_START, ADJ_MAIN_COLS)

    # outstanding block L:Y
    _clear_block(ws, ADJ_DATA_START, ADJ_OS_COLS)
    _write_table(ws, ost, ADJ_DATA_START, ADJ_OS_COLS)

    # three pivots (label + measures), each in its own column block
    for df, c0 in ((pv_dpms, ADJ_PIVOT_DPMS_C0),
                   (pv_dpid, ADJ_PIVOT_DPID_C0),
                   (pv_os, ADJ_PIVOT_OS_C0)):
        cols = list(range(c0, c0 + df.shape[1]))
        _clear_block(ws, ADJ_DATA_START, cols)
        _write_table(ws, df, ADJ_DATA_START, cols)

    # summary SUM row (row 2): main E:I and outstanding Q,R,S,V
    def _sum(df, col):
        return float(pd.to_numeric(df[col], errors="coerce").fillna(0).sum())
    _set(ws, ADJ_SUMMARY_ROW, 5, _sum(main, "Adjustment Komisi"))     # E
    _set(ws, ADJ_SUMMARY_ROW, 6, _sum(main, "Adjustment Bruto"))      # F
    _set(ws, ADJ_SUMMARY_ROW, 7, _sum(main, "Adjustment PPh"))        # G
    _set(ws, ADJ_SUMMARY_ROW, 8, _sum(main, "Penalty Parcel"))        # H
    _set(ws, ADJ_SUMMARY_ROW, 9, _sum(main, "Penalty PPDP"))          # I
    if len(ost):
        _set(ws, ADJ_SUMMARY_ROW, 17, _sum(ost, "Total Outstanding")) # Q
        _set(ws, ADJ_SUMMARY_ROW, 18, _sum(ost, "Denda"))             # R
        _set(ws, ADJ_SUMMARY_ROW, 19, _sum(ost, "Total Penalty"))     # S
        _set(ws, ADJ_SUMMARY_ROW, 22, -_sum(ost, "Total Outstanding"))# V (OS Gsheet)

    if verbose:
        print(f"[writer] Adjustment&Penalty: main={len(main)} rows, outstanding={len(ost)} rows, "
              f"pivots dpms={len(pv_dpms)}/dpid={len(pv_dpid)}/os={len(pv_os)}")


def _write_commission(wb, commission_df, verbose):
    ws = wb["Commission"]
    letters = [get_column_letter(c) for c in range(1, ci(LAST_COMMISSION_COL) + 1)]
    r0 = COMMISSION_DATA_START
    old_last = _find_last_row(ws, r0)
    n = len(commission_df)
    clear_to = max(old_last, r0 + n - 1)

    missing = [L for L in letters if L not in commission_df.columns]
    if missing and verbose:
        print(f"[writer] WARNING: Commission columns not produced (left blank): {missing}")

    for r in range(r0, clear_to + 1):               # clear old body first
        for L in letters:
            _set(ws, r, ci(L), None)
    for i, (_, row) in enumerate(commission_df.iterrows()):
        r = r0 + i
        for L in letters:
            if L in commission_df.columns:
                _set(ws, r, ci(L), _cell_val(row.get(L)))
    if verbose:
        print(f"[writer] Commission: cleared rows {r0}..{clear_to} (old_last={old_last}), "
              f"wrote {n} rows (rows {r0}..{r0 + n - 1})")
