"""Phase 2a+ — automate the `Format SR Penalty` compile (pandas).

Replaces the manual step of hand-assembling `Format SR Penalty - {Mon}.xlsx` from the
raw stakeholder files. Produces the tracking-level `Penalty Parcel` detail (the sheet
that drives columns H/I of the Adjustment&Penalty sheet), classified by `Case Pivot`.

The four penalty CASES and their verified rules (reconciled to June's Format SR /
golden M06 exactly — see PROJECT_CONTEXT §10 Phase 2a+):

  Case                              Case Pivot       Source file -> sheet            Amount rule
  --------------------------------  ---------------  ------------------------------  -------------------------------
  Denda PPDP                        Penalty PPDP     PPDP -> RAW                     -Total Penalty Amount
  Fake PODA Mitra Last Mile         Penalty Parcel   Fake POD&PODA -> Penalty PODA  -Penalty  (where Penalty > 0)
  Fake POD Mitra Last Mile          Penalty Parcel   Fake POD&PODA -> Penalty POD   -Penalty  (where Penalty > 0)
  Penalty Missing Mitra Last Mile   Penalty Parcel   LM -> periode ...              -Est claim amount
                                                                                    (where >0 AND Final Check CL != 'Take out')

Missing rows carry no DP ID; it is mapped from `Investigating Hub Name` via the
hub->DP lookup built from the POD/PODA data (which carries both Hub Name and dp_id).

The human-review decisions (PPDP sanggahan outcome, Missing `Final Check CL`) live IN
the source data as columns — the compiler reads them, it does not make them.

Two input surfaces feed the SAME core (`_compile_detail`), so the verified rules are
identical either way:
  * `build_penalty_detail(...)`             — the 3 raw stakeholder exports (Penalty SR/);
  * `build_penalty_detail_from_intake(...)` — one standardized "Penalty Intake" workbook
                                              with the 4 near-raw tabs (see PART C).
"""
from __future__ import annotations
import pandas as pd

import par_pipeline as par

# Format SR `Penalty Parcel` detail schema (order matches the manual worksheet).
DETAIL_COLS = ["DP ID", "DPMS", "Nama Mitra", "region", "Tracking ID", "Shipper Name",
               "Granular Status", "Case", "Case Pivot", "Deduction Amount",
               "Invoice Number", "Note"]

CASE_PPDP = "Denda PPDP"
CASE_PODA = "Fake PODA Mitra Last Mile"
CASE_POD = "Fake POD Mitra Last Mile"
CASE_MISS = "Penalty Missing Mitra Last Mile"
PIVOT_PPDP = "Penalty PPDP"
PIVOT_PARCEL = "Penalty Parcel"

# Standardized "Penalty Intake" tab names (PART C).
INTAKE_TAB_PPDP = "Penalty PPDP"
INTAKE_TAB_PODPODA = "Penalty Fake POD-PODA"
INTAKE_TAB_MISSING = "Penalty Missing LM"
INTAKE_TAB_OUTSTANDING = "Penalty Outstanding"


def _col(df, *names):
    """First matching column (case-insensitive, trimmed) or None."""
    low = {str(c).strip().lower(): c for c in df.columns}
    for n in names:
        if n.strip().lower() in low:
            return low[n.strip().lower()]
    return None


def _rows(df, case, pivot, c_dp, c_mitra, c_tid, c_shipper, c_status, amount):
    """Assemble detail rows for one case from parallel column selections."""
    out = pd.DataFrame({
        "DP ID": df[c_dp] if c_dp else None,
        "DPMS": None,
        "Nama Mitra": df[c_mitra] if c_mitra else None,
        "region": None,
        "Tracking ID": df[c_tid] if c_tid else None,
        "Shipper Name": df[c_shipper] if c_shipper else None,
        "Granular Status": df[c_status] if c_status else None,
        "Case": case,
        "Case Pivot": pivot,
        "Deduction Amount": amount,
        "Invoice Number": None,
        "Note": None,
    })
    return out


def _clean_hub(h):
    """Normalise a hub code: drop the '(Inactive) ' prefix and surrounding whitespace."""
    s = str(h).strip()
    if s.lower().startswith("(inactive)"):
        s = s[len("(inactive)"):].strip()
    return s


def _hub_to_dp(*dfs):
    """Build {clean Hub Name -> (dp_id, dp_name)} from sheets that carry both."""
    mapping = {}
    for df in dfs:
        c_hub = _col(df, "Hub Name")
        c_dp = _col(df, "dp_id", "DP ID")
        c_name = _col(df, "dp_name", "Nama Mitra", "DP Name")
        if not (c_hub and c_dp):
            continue
        for _, r in df[[c_hub, c_dp] + ([c_name] if c_name else [])].iterrows():
            hub = _clean_hub(r[c_hub])
            if hub and hub not in mapping and pd.notna(r[c_dp]):
                mapping[hub] = (r[c_dp], r[c_name] if c_name else None)
    return mapping


def _casetype_of(v):
    """Map a Type cell ('POD'/'PODA'/…) to its Case label (check PODA first — it contains 'pod')."""
    s = str(v).lower()
    if "poda" in s:
        return CASE_PODA
    if "pod" in s:
        return CASE_POD
    return None


def _compile_detail(ppdp_df, podpoda_df, missing_df,
                    hub_dp_overrides=None, verbose=True) -> pd.DataFrame:
    """Core compile: three DataFrames -> the tracking-level penalty detail.

    Shared by both input surfaces (raw files and the Penalty Intake workbook). The rules
    are the verified Phase 2a+ ones. `podpoda_df` is a single frame carrying a Type column
    (POD/PODA) plus `Hub Name`+`DP ID` (the source of the hub->DP map for Missing).
    """
    overrides = {}
    for hub, val in (hub_dp_overrides or {}).items():
        overrides[_clean_hub(hub)] = val if isinstance(val, (tuple, list)) else (val, None)
    parts = []

    # 1) Denda PPDP (-> I) --------------------------------------------------------
    if ppdp_df is not None and len(ppdp_df):
        c_amt = _col(ppdp_df, "Total Penalty Amount")
        d = ppdp_df[pd.to_numeric(ppdp_df[c_amt], errors="coerce").fillna(0) != 0]
        parts.append(_rows(
            d, CASE_PPDP, PIVOT_PPDP,
            _col(d, "DP ID", "dp_id"), _col(d, "Nama Mitra", "dp_name"), _col(d, "Tracking ID"),
            _col(d, "Shipper Name"), _col(d, "Granular Status"),
            -pd.to_numeric(d[c_amt], errors="coerce"),
        ))

    # 2) Fake POD / PODA (-> H), split by Type ------------------------------------
    hubmap = {}
    if podpoda_df is not None and len(podpoda_df):
        hubmap = _hub_to_dp(podpoda_df)                      # clean Hub Name -> (dp_id, dp_name)
        c_amt = _col(podpoda_df, "Penalty")
        c_type = _col(podpoda_df, "Type", "Type (POD/PODA)", "Type (POD / PODA)")
        base = podpoda_df[pd.to_numeric(podpoda_df[c_amt], errors="coerce").fillna(0) > 0].copy()
        base["_case"] = base[c_type].map(_casetype_of) if c_type else CASE_PODA
        base["_case"] = base["_case"].fillna(CASE_PODA)      # blank Type -> PODA (same H pivot anyway)
        for case in (CASE_PODA, CASE_POD):
            d = base[base["_case"] == case]
            if len(d):
                parts.append(_rows(
                    d, case, PIVOT_PARCEL,
                    _col(d, "dp_id", "DP ID"), _col(d, "dp_name", "Nama Mitra"), _col(d, "Tracking ID"),
                    _col(d, "Shipper Name"), _col(d, "Granular Status"),
                    -pd.to_numeric(d[c_amt], errors="coerce"),
                ))

    # 3) Penalty Missing (-> H) ---------------------------------------------------
    unmapped = pd.DataFrame()
    if missing_df is not None and len(missing_df):
        c_amt = _col(missing_df, "Est claim amount", "Est Claim Amount")
        c_cl = _col(missing_df, "Final Check CL")
        c_hub = _col(missing_df, "Investigating Hub Name")
        keep = pd.to_numeric(missing_df[c_amt], errors="coerce").fillna(0) > 0
        if c_cl:
            keep &= missing_df[c_cl].astype(str).str.strip().str.lower() != "take out"
        d = missing_df[keep].copy()

        def _resolve(h):
            k = _clean_hub(h)
            return overrides.get(k) or hubmap.get(k) or (None, None)

        resolved = d[c_hub].map(_resolve) if c_hub else [(None, None)] * len(d)
        d["_dp"] = [x[0] for x in resolved]
        d["_mitra"] = [x[1] for x in resolved]
        if c_hub:
            unmapped = d.loc[d["_dp"].isna(),
                             [c for c in (c_hub, _col(d, "Tracking ID"), c_amt) if c]].copy()
        if len(unmapped) and verbose:
            print(f"[format_sr] WARNING: {len(unmapped)} Missing rows have no hub->DP mapping "
                  f"-> surfaced for manual assignment (add to hub_dp_overrides): "
                  f"{sorted(unmapped[c_hub].astype(str).unique())[:8]}")
        parts.append(_rows(
            d, CASE_MISS, PIVOT_PARCEL,
            "_dp", "_mitra", _col(d, "Tracking ID"),
            _col(d, "Shipper Name"), _col(d, "Order Granular Status", "Granular Status"),
            -pd.to_numeric(d[c_amt], errors="coerce"),
        ))

    detail = (pd.concat(parts, ignore_index=True)[DETAIL_COLS]
              if parts else pd.DataFrame(columns=DETAIL_COLS))
    detail.attrs["unmapped_missing"] = unmapped              # surfaced, never dropped

    if verbose:
        g = detail.groupby("Case")["Deduction Amount"].agg(["size", "sum"])
        print("[format_sr] compiled penalty detail:")
        for case, row in g.iterrows():
            print(f"    {case:34} n={int(row['size']):5d}  sum={row['sum']:,.1f}")
        print(f"    {'TOTAL':34} n={len(detail):5d}  sum={detail['Deduction Amount'].sum():,.1f}")
    return detail


def _fill_dpms_region(detail, dpms_region_lookup):
    if dpms_region_lookup is not None and len(detail):
        dr = detail["DP ID"].map(par._norm_one).map(lambda k: dpms_region_lookup(k) or (None, None))
        detail["DPMS"] = [x[0] for x in dr]
        detail["region"] = [x[1] for x in dr]
    return detail


def build_penalty_detail(ppdp_path, podpoda_path, missing_path,
                         dpms_region_lookup=None, hub_dp_overrides=None,
                         verbose=True) -> pd.DataFrame:
    """Compile the penalty detail from the THREE RAW stakeholder exports (Penalty SR/).

    `dpms_region_lookup` (optional): callable dp_key -> (DPMS, region) to fill the detail's
    DPMS/region columns for a faithful Format SR reproduction (not needed by
    build_adjustment_penalty, which re-derives them via the Data sheet).
    `hub_dp_overrides` (optional): {hub code -> dp_id | (dp_id, mitra_name)} for inactive
    hubs absent from the POD/PODA data (e.g. June 'DP-KOI-GDCEK'->50231, 'DP-MAC-GDTLN'->2621).
    """
    ppdp = pd.read_excel(ppdp_path, sheet_name="RAW")
    poda = pd.read_excel(podpoda_path, sheet_name="Penalty PODA"); poda["Type"] = "PODA"
    pod = pd.read_excel(podpoda_path, sheet_name="Penalty POD"); pod["Type"] = "POD"
    podpoda = pd.concat([poda, pod], ignore_index=True)       # one frame w/ a Type column
    missing = pd.read_excel(missing_path, sheet_name=0)        # 'periode ...' is the first sheet
    detail = _compile_detail(ppdp, podpoda, missing, hub_dp_overrides=hub_dp_overrides, verbose=verbose)
    return _fill_dpms_region(detail, dpms_region_lookup)


def build_penalty_detail_from_intake(penalty_intake_path,
                                     dpms_region_lookup=None, hub_dp_overrides=None,
                                     verbose=True) -> pd.DataFrame:
    """Compile the penalty detail from the standardized "Penalty Intake" workbook (PART C).

    Reads the near-raw tabs `Penalty PPDP` / `Penalty Fake POD-PODA` / `Penalty Missing LM`
    and delegates to the same core as the raw-file path. (The `Penalty Outstanding` tab is
    consumed separately by par_adjustment for the L:Y block.)
    """
    xls = pd.ExcelFile(penalty_intake_path)

    def _tab(name):
        try:
            return pd.read_excel(xls, sheet_name=name)
        except Exception:
            return None

    detail = _compile_detail(_tab(INTAKE_TAB_PPDP), _tab(INTAKE_TAB_PODPODA), _tab(INTAKE_TAB_MISSING),
                             hub_dp_overrides=hub_dp_overrides, verbose=verbose)
    return _fill_dpms_region(detail, dpms_region_lookup)


def read_intake_outstanding(penalty_intake_path):
    """Return the `Penalty Outstanding` tab as a DataFrame (or None), for par_adjustment L:Y."""
    try:
        return pd.read_excel(penalty_intake_path, sheet_name=INTAKE_TAB_OUTSTANDING)
    except Exception:
        return None
