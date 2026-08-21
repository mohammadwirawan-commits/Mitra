"""
Pivot All Region pipeline — Phase 1 (Data -> Pivot Data -> Commission).

Pure, environment-agnostic functions. The Colab notebook and the local
verification harness both call these; only the *loading* of inputs differs.

Conventions mirror the existing notebook:
  - reference dataframes use the notebook's names (df_reg_type, df_lipoh, ...)
  - "Mapping Type" is the Pivot/Commission category vocabulary
"""
from __future__ import annotations
import numpy as np
import pandas as pd

# The 16 columns of the workbook "Data" sheet, in order.
DATA_COLUMNS = [
    "Global Shipper ID", "Shipper Name", "DPID", "DPMS", "Mitra Name", "Region",
    "Shipper Tier", "Mapping Type", "Total Parcel", "Total Weight",
    "Total Delivery Fee", "% Commission", "Total Commission",
    "Drop off to SH Parcel", "Drop off to SH Commission", "Insurance Fee",
]

DEFAULT_REGION = "All Region - Non POH"


# --------------------------------------------------------------------------
# Data sheet
# --------------------------------------------------------------------------
def _blank_data_frame(n: int) -> pd.DataFrame:
    return pd.DataFrame({c: [np.nan] * n for c in DATA_COLUMNS})


def _map_akus_alo_unreg(df: pd.DataFrame) -> pd.DataFrame:
    out = _blank_data_frame(len(df))
    out["Global Shipper ID"] = df["shipper_id"].values
    out["Shipper Name"] = df["shipper_name"].values
    out["DPID"] = df["dp_id"].values
    out["Mitra Name"] = df["dp_name"].values
    out["Mapping Type"] = df["mapping_type"].values
    out["Total Parcel"] = df["total_tracking_id"].values
    out["Total Weight"] = df["total_nv_weight"].values
    out["Total Commission"] = df["total_commission"].values
    out["Drop off to SH Parcel"] = df["total_parcel_drop_off"].values
    out["Drop off to SH Commission"] = df["total_drop_off_commission"].values
    return out


def _map_regular_cod(df: pd.DataFrame) -> pd.DataFrame:
    out = _blank_data_frame(len(df))
    out["Global Shipper ID"] = df["shipper_id"].values
    out["Shipper Name"] = df["shipper_name"].values
    out["DPID"] = df["dp_id"].values
    out["Mitra Name"] = df["dp_name"].values
    # Mapping Type is the Regular/COD remark, NOT the literal "Registered"
    out["Mapping Type"] = df["remarks_cod_regular"].values
    out["Shipper Tier"] = df["tier"].values
    out["Total Parcel"] = df["total_tracking_id"].values
    out["Total Weight"] = df["total_nv_weight"].values
    out["Total Delivery Fee"] = df["total_delivery_fee"].values
    out["% Commission"] = df["commission_pct"].values
    out["Total Commission"] = df["commission"].values
    out["Drop off to SH Parcel"] = df["total_parcel_drop_off"].values
    out["Drop off to SH Commission"] = df["total_drop_off_commission"].values
    out["Insurance Fee"] = df["total_insurance_fee"].values
    return out


def _map_simple(df: pd.DataFrame, mapping_type: str, sid, sname, dpid, mitra) -> pd.DataFrame:
    """LEX / MPM / Last Mile: only shipper, dp, parcel count, commission."""
    out = _blank_data_frame(len(df))
    out["Global Shipper ID"] = df[sid].values
    out["Shipper Name"] = df[sname].values
    out["DPID"] = df[dpid].values
    out["Mitra Name"] = df[mitra].values
    out["Mapping Type"] = mapping_type
    out["Total Parcel"] = df["total_tracking_id"].values
    out["Total Commission"] = df["total_commission"].values
    return out


def build_data_sheet(shipper: dict[str, pd.DataFrame], df_reg_type: pd.DataFrame,
                     df_cross_border: pd.DataFrame | None = None,
                     df_disputes: pd.DataFrame | None = None) -> pd.DataFrame:
    """Consolidate the 5 shipper-level result frames into the 16-col Data sheet.

    `shipper` keys: 'akus_alo_unreg', 'regular_cod', 'last_mile', 'lex', 'mpm'.
    `df_reg_type` = intake tab '3. Region Split' (DP ID, DPMS ID, Region Split).
    """
    parts = [
        _map_akus_alo_unreg(shipper["akus_alo_unreg"]),
        _map_regular_cod(shipper["regular_cod"]),
        _map_simple(shipper["last_mile"], "Last Mile",
                    "shipper_id", "shipper_name", "dp_id", "dp_name"),
        _map_simple(shipper["lex"], "LEX Hub Pickup",
                    "global_shipper_id", "shipper_name", "DPID", "Mitra Name"),
        _map_simple(shipper["mpm"], "Mitra Pickup Mitra",
                    "shipper_id", "shipper_name", "DPID (Pickup by)", "Mitra Name (Pickup by)"),
    ]
    data = pd.concat(parts, ignore_index=True)

    # ---- Append manual rows BEFORE the lookup so they get DPMS/Region too ----
    # Cross Border (30% x delivery fee) + prior-month dispute/claim rows.
    extra = []
    if df_cross_border is not None and len(df_cross_border):
        cb = _blank_data_frame(len(df_cross_border))
        cb["DPID"] = df_cross_border.get("DPID", df_cross_border.get("dp_id")).values
        cb["Mapping Type"] = "Cross Border"
        cb["Total Parcel"] = df_cross_border.get("parcel", 1)
        fee = pd.to_numeric(df_cross_border["delivery_fee"], errors="coerce")
        cb["Total Delivery Fee"] = fee.values
        cb["Total Commission"] = (fee * 0.30).values
        if "DPMS ID" in df_cross_border:            # carry provided DPMS/Region (fallbacks below)
            cb["DPMS"] = df_cross_border["DPMS ID"].values
        if "Region Split" in df_cross_border:
            cb["Region"] = df_cross_border["Region Split"].values
        extra.append(cb)
    if df_disputes is not None and len(df_disputes):
        extra.append(df_disputes.reindex(columns=DATA_COLUMNS))
    if extra:
        data = pd.concat([data] + extra, ignore_index=True)

    # ---- DPMS + Region lookup by DPID over ALL rows (base + manual) ----
    # Priority: Region Split file -> value already on the row (manual tabs) -> fallback
    # (DPMS -> DPID, Region -> DEFAULT_REGION).
    reg = df_reg_type.rename(columns={
        "DP ID": "_rs_dpid", "DPMS ID": "_rs_dpms", "Region Split": "_rs_region"})
    reg = reg[["_rs_dpid", "_rs_dpms", "_rs_region"]].copy()
    reg["_key"] = _norm_id(reg["_rs_dpid"])
    reg = reg.drop_duplicates("_key", keep="first")

    data["_key"] = _norm_id(data["DPID"])
    data = data.merge(reg[["_key", "_rs_dpms", "_rs_region"]], on="_key", how="left")

    def _pick(primary, secondary, fallback):
        s = primary.where(primary.notna() & (primary.astype(str).str.strip() != ""), secondary)
        return s.where(s.notna() & (s.astype(str).str.strip() != ""), fallback)

    data["DPMS"] = _pick(data["_rs_dpms"], data["DPMS"], data["DPID"])
    data["Region"] = _pick(data["_rs_region"], data["Region"], DEFAULT_REGION)
    data = data.drop(columns=["_key", "_rs_dpms", "_rs_region"])

    return data[DATA_COLUMNS]


def _norm_one(v) -> str:
    """Normalise a single id value to a comparable string key (strip .0 float tails)."""
    if pd.isna(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    sv = str(v).strip()
    if sv.endswith(".0") and sv[:-2].isdigit():
        return sv[:-2]
    return sv


def _norm_id(s: pd.Series) -> pd.Series:
    """Normalise an id column to a comparable string key (strip .0 float tails)."""
    return s.map(_norm_one)


# --------------------------------------------------------------------------
# Commission — derived-column engine
#
# Reproduces the workbook formulas exactly. Input columns (pivot-driven and
# pasted KPI) are read as-is; only the FORMULA columns below are computed.
# Columns are addressed by their Excel letter to avoid label ambiguity.
# --------------------------------------------------------------------------

# Commission Scheme lookup tables (from the 'Commission Scheme' sheet).
_POST_VOL_BREAKS = [1, 501, 1001, 3001, 10001]          # I28:I32
_POST_VOL_PCT    = [0.15, 0.35, 0.40, 0.45, 0.50]        # K28:K32
_DROPOFF_BREAKS  = [0.0, 0.90, 0.95]                     # J42:J44 / J47:J49
_DROPOFF_JABO    = [200, 225, 250]                       # K42:K44
_DROPOFF_LUAR    = [250, 275, 300]                       # K47:K49
_ACQ_SCORE_BREAKS = [0.01, 0.90, 0.99]                   # I12:I14
_ACQ_SCORE_KPI    = [-0.05, 0.00, 0.05]                  # J12:J14


def _match_le(value, breaks):
    """Excel MATCH(value, breaks, 1): index of largest break <= value (or None)."""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return None
    idx = None
    for i, b in enumerate(breaks):
        if value >= b:
            idx = i
        else:
            break
    return idx


def _num(v):
    try:
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return 0.0
        return float(v)
    except (TypeError, ValueError):
        return 0.0


# Commission pivot columns -> (measure in Data, Mapping Type filter).
COMMISSION_PIVOT_MAP = {
    "E": ("Total Parcel", "MP Registered Claim"), "F": ("Total Delivery Fee", "MP Registered Claim"),
    "H": ("Total Commission", "MP Registered Claim"), "I": ("Drop off to SH Parcel", "MP Registered Claim"),
    "J": ("Drop off to SH Commission", "MP Registered Claim"),
    "K": ("Total Parcel", "Acquisition Claim"), "L": ("Total Commission", "Acquisition Claim"),
    "M": ("Drop off to SH Parcel", "Acquisition Claim"), "N": ("Drop off to SH Commission", "Acquisition Claim"),
    "Q": ("Total Parcel", "Acquisition"), "R": ("Total Commission", "Acquisition"),
    "S": ("Drop off to SH Parcel", "Acquisition"), "T": ("Drop off to SH Commission", "Acquisition"),
    "AF": ("Total Parcel", "Allocation Claim"), "AG": ("Total Commission", "Allocation Claim"),
    "AH": ("Drop off to SH Parcel", "Allocation Claim"), "AI": ("Drop off to SH Commission", "Allocation Claim"),
    "AL": ("Total Parcel", "Allocation"), "AM": ("Total Commission", "Allocation"),
    "AN": ("Drop off to SH Parcel", "Allocation"), "AO": ("Drop off to SH Commission", "Allocation"),
    "AT": ("Total Parcel", "Unregistered Eligible Claim"), "AU": ("Total Commission", "Unregistered Eligible Claim"),
    "AV": ("Drop off to SH Parcel", "Unregistered Eligible Claim"), "AW": ("Drop off to SH Commission", "Unregistered Eligible Claim"),
    "AZ": ("Total Parcel", "Unregistered Eligible"), "BA": ("Total Commission", "Unregistered Eligible"),
    "BB": ("Drop off to SH Parcel", "Unregistered Eligible"), "BC": ("Drop off to SH Commission", "Unregistered Eligible"),
    "BH": ("Total Parcel", "Regular"), "BI": ("Total Delivery Fee", "Regular"), "BK": ("Total Commission", "Regular"),
    "BL": ("Insurance Fee", "Regular"), "BM": ("Drop off to SH Parcel", "Regular"), "BN": ("Drop off to SH Commission", "Regular"),
    "BO": ("Total Parcel", "COD"), "BP": ("Total Delivery Fee", "COD"), "BR": ("Total Commission", "COD"),
    "BS": ("Drop off to SH Parcel", "COD"), "BT": ("Drop off to SH Commission", "COD"),
    "BU": ("Total Parcel", "Cross Border"), "BV": ("Total Delivery Fee", "Cross Border"), "BX": ("Total Commission", "Cross Border"),
    "BY": ("Drop off to SH Parcel", "Cross Border"), "BZ": ("Drop off to SH Commission", "Cross Border"),
    "CA": ("Total Parcel", "LEX Hub Pickup"), "CB": ("Total Commission", "LEX Hub Pickup"),
    "CI": ("Total Parcel", "Mitra Pickup Mitra"), "CJ": ("Total Commission", "Mitra Pickup Mitra"),
    "CK": ("Total Parcel", "Last Mile"), "CL": ("Total Commission", "Last Mile"),
}

# Pasted-KPI columns the assembly fills from intake / prev-pivot (via kpi_row_fn).
# Z = Acquisition 'Total Score' read DIRECTLY from the Acquisition KPI tab (the
# per-component weights in the sheet labels are stale; the file's Total Score is
# authoritative and is what payout keys off).
COMMISSION_KPI_COLS = ["G", "O", "AJ", "AX",                       # claim KPI % (prev pivot -100%)
                       "U", "V", "W", "X", "Y", "Z", "AC",         # Acquisition components + Total Score + Tag
                       "AP", "AQ", "BD", "BE", "CC", "CD",         # Allocation / Unreg / drop-off
                       "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV"]  # Last Mile


def build_pivot(data: pd.DataFrame):
    """Per-(DPID, Mapping Type) sums of every measure. Returns a lookup dict:
    pivot[(dpid_key, mapping_type)][measure] = sum."""
    measures = ["Total Parcel", "Total Weight", "Total Delivery Fee", "Total Commission",
                "Drop off to SH Parcel", "Drop off to SH Commission", "Insurance Fee"]
    d = data.copy()
    d["_k"] = _norm_id(d["DPID"])
    for m in measures:
        d[m] = pd.to_numeric(d[m], errors="coerce").fillna(0.0)
    # Excel GETPIVOTDATA matches item names case-insensitively -> fold the key.
    d["_mt"] = d["Mapping Type"].astype(str).str.strip().str.casefold()
    grp = d.groupby(["_k", "_mt"])[measures].sum()
    pivot = {}
    for (k, mt), r in grp.iterrows():
        pivot[(k, mt)] = {m: float(r[m]) for m in measures}
    return pivot


def assemble_commission(data: pd.DataFrame, kpi_row_fn) -> pd.DataFrame:
    """Build the Commission table (one row per DPID), keyed by Excel letter.

    kpi_row_fn(dpid_key, region, mitra) -> dict of the pasted-KPI columns
    (any of COMMISSION_KPI_COLS); missing keys default to 0.
    """
    pivot = build_pivot(data)

    # DPID -> (DPMS, Region, Mitra Name), first occurrence; sorted by DPID like the sheet.
    keyinfo = (data.assign(_k=_norm_id(data["DPID"]))
                   .groupby("_k", sort=False)
                   .agg(DPMS=("DPMS", "first"), Region=("Region", "first"),
                        Mitra=("Mitra Name", "first"), DPID=("DPID", "first")))
    keyinfo = keyinfo.reset_index()
    keyinfo["_sort"] = pd.to_numeric(keyinfo["DPID"], errors="coerce")
    keyinfo = keyinfo.sort_values(["_sort", "_k"]).reset_index(drop=True)

    rows = []
    for _, ki in keyinfo.iterrows():
        k = ki["_k"]
        row = {"A": ki["DPID"], "B": ki["DPMS"], "C": ki["Mitra"], "D": ki["Region"]}
        # pivot-driven columns (case-insensitive Mapping Type match, like GETPIVOTDATA)
        for L, (measure, mt) in COMMISSION_PIVOT_MAP.items():
            row[L] = pivot.get((k, mt.strip().casefold()), {}).get(measure, 0.0)
        # pasted KPI columns
        row.setdefault("CC", "")
        kpi = kpi_row_fn(k, ki["Region"], ki["Mitra"]) or {}
        for L in COMMISSION_KPI_COLS:
            row[L] = kpi.get(L, row.get(L, 0.0))
        # derived
        row.update(compute_commission_derived(row))
        rows.append(row)
    return pd.DataFrame(rows)


def compute_commission_derived(row: dict) -> dict:
    """Given a dict of Commission columns keyed by Excel letter (input values
    present), return the computed FORMULA columns. Mirrors the workbook."""
    g = lambda L: _num(row.get(L))
    out = {}

    # --- Acquisition score / payout ---
    out["P"] = g("O") * g("L")                                   # claim KPI amt
    # Prefer the Acquisition file's precomputed Total Score (authoritative);
    # fall back to the weighted components only if it is absent.
    z_in = row.get("Z")
    if z_in is not None and str(z_in).strip() != "" and not (isinstance(z_in, float) and np.isnan(z_in)):
        Z = _num(z_in)
    else:
        Z = 0.40*g("U") + 0.10*g("V") + 0.35*g("W") + 0.05*g("X") + 0.10*g("Y")
    out["Z"] = Z
    kidx = _match_le(Z, _ACQ_SCORE_BREAKS)
    kpi = _ACQ_SCORE_KPI[kidx] if kidx is not None else -0.05
    out["AA"] = 1.0 + kpi
    out["AB"] = {-0.05: "Deduct 5%", 0.0: "Commission As Is", 0.05: "Bonus 5%"}[kpi]
    out["AD"] = out["AA"] - 1.0
    out["AE"] = g("R") * out["AD"]

    # --- Allocation ---
    out["AK"] = g("AJ") * g("AG")
    out["AR"] = 1.0 if (g("AP") >= 0.95 and g("AQ") >= 0.99) else 0.95
    out["AS"] = (out["AR"] - 1.0) * g("AM")

    # --- Unregistered eligible ---
    out["AY"] = g("AX") * g("AU")
    out["BF"] = 1.0 if (g("BD") >= 0.87 and g("BE") >= 0.90) else 0.95
    out["BG"] = (out["BF"] - 1.0) * g("BA")

    # --- Regular / COD volume % (shared bracket on BH+BO) ---
    vol = g("BH") + g("BO")
    vidx = _match_le(vol, _POST_VOL_BREAKS)
    vpct = _POST_VOL_PCT[vidx] if vidx is not None else 0.0
    out["BJ"] = vpct if g("BH") > 0 else 0.0
    out["BQ"] = vpct if g("BO") > 0 else 0.0

    # --- Cross border % ---
    out["BW"] = 0.30 if g("BU") > 0 else 0.0

    # --- Drop off to SH ---
    out["CF"] = sum(g(L) for L in ["I","M","S","AH","AN","AV","BB","BM","BS","BY"])
    out["CG"] = sum(g(L) for L in ["J","N","T","AI","AO","AW","BC","BN","BT","BZ"])
    if out["CF"] == 0:
        out["CE"] = 0.0
    else:
        cd = g("CD")
        didx = _match_le(cd, _DROPOFF_BREAKS)
        didx = 0 if didx is None else didx
        table = _DROPOFF_JABO if str(row.get("CC")).strip() == "Jabodetabek" else _DROPOFF_LUAR
        out["CE"] = table[didx]
    out["CH"] = (out["CG"]
                 - sum(g(L) for L in ["S","AN","BB","BM","BS","BY"]) * out["CE"]
                 - sum(g(L) for L in ["J","N","AI","AW"]))
    return out


# --------------------------------------------------------------------------
# Pasted-KPI builder
#
# Produces the kpi_row_fn(dpid_key, region, mitra) -> dict that
# assemble_commission() consumes, from the intake / first-stage / prev-pivot
# dataframes. All sources are optional; an absent source simply contributes no
# columns and the engine defaults them to 0. See PROJECT_CONTEXT.md §6.
# --------------------------------------------------------------------------

# Last Mile: golden Commission CM:CV -> Last Mile Tracker column labels.
LM_TRACKER_COLS = {
    "CM": "Hub Name", "CN": "Attempt Rate N0", "CO": "Attempt Rate N1",
    "CP": "Terminal Day 6", "CQ": "Prior Attempt N0", "CR": "RTS Rate",
    "CS": "Valid PODA", "CT": "Valid POD", "CU": "Final Score", "CV": "Final Tier",
}
# Which CM:CV columns are percentages (Hub Name / Final Tier are text/label).
_LM_PCT_COLS = {"CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU"}


def _pct(v) -> float:
    """Parse a KPI value to a fraction.
    '93.18%' -> 0.9318 ; 0.998 -> 0.998 ; 95.67 -> 0.9567 ; blank/NaN -> 0.0.
    """
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return 0.0
    if isinstance(v, str):
        s = v.strip()
        if s == "":
            return 0.0
        had_pct = "%" in s
        s = s.replace("%", "").replace(",", "")
        try:
            x = float(s)
        except ValueError:
            return 0.0
        return x / 100.0 if had_pct else (x / 100.0 if x > 1.5 else x)
    try:
        x = float(v)
    except (TypeError, ValueError):
        return 0.0
    return x / 100.0 if x > 1.5 else x


def _clean_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Strip trailing weight annotations / newlines from column names
    ('LI_N0\\n(40%)' -> 'LI_N0')."""
    return df.rename(columns={c: str(c).split("\n")[0].split("(")[0].strip() for c in df.columns})


def _first_col(df: pd.DataFrame, names) -> str | None:
    lower = {str(c).strip().lower(): c for c in df.columns}
    for n in names:
        if n.lower() in lower:
            return lower[n.lower()]
    return None


def _is_poh_region(region) -> bool:
    r = str(region).lower()
    return ("poh" in r) and ("non poh" not in r) and ("non-poh" not in r)


def _is_poh_type(t) -> bool:
    s = str(t).lower()
    return ("poh" in s) and ("non" not in s)


def build_kpi_row_fn(df_acq=None, df_allo=None, df_unreg=None,
                     df_lipoh=None, df_lm_tracker=None, prev_claim_kpi=None,
                     lm_col_map=None):
    """Return kpi_row_fn(dpid_key, region, mitra) -> dict of pasted-KPI columns.

      df_acq         Acquisition KPI (intake tab / file)  -> U,V,W,X,Y,Z,AC (Z=Total Score direct)
      df_allo        Allocation KPI                        -> AP,AQ  (join DPID + POH flag)
      df_unreg       Unreg KPI                             -> BD,BE  (BE=LI_POH if POH else PU Scan_N0)
      df_lipoh       first-stage '2. LIPOH DropOff-SH'     -> CC,CD
      df_lm_tracker  first-stage '6. Last Mile Tracker'    -> CM:CV
      prev_claim_kpi dict {dpid_key: {'G':raw,'O':raw,'AJ':raw,'AX':raw}} from the
                     previous month's pivot; the closure returns raw-1.0 (i.e. -100%).
    """
    lm_map = dict(LM_TRACKER_COLS)
    if lm_col_map:
        lm_map.update(lm_col_map)

    # ---- Acquisition: U..Z + AC(Tag), by DPID, dedup keep-first ----
    acq = {}
    if df_acq is not None and len(df_acq):
        a = _clean_cols(df_acq).copy()
        a["_k"] = _norm_id(a["DP ID"])
        a = a.drop_duplicates("_k", keep="first")
        for _, r in a.iterrows():
            acq[r["_k"]] = {
                "U": _pct(r.get("LI_N0")), "V": _pct(r.get("LI_N1")),
                "W": _pct(r.get("PU Scan_N0")), "X": _pct(r.get("Update_RAR")),
                "Y": _pct(r.get("PoPA")), "Z": _pct(r.get("Total Score")),
                "AC": ("" if pd.isna(r.get("Tag")) else r.get("Tag")),
            }

    # ---- Allocation: AP,AQ keyed by (DPID, POH) ----
    allo = {}
    if df_allo is not None and len(df_allo):
        al = _clean_cols(df_allo)
        for _, r in al.iterrows():
            k = _norm_one(r["DP ID"])
            allo[(k, _is_poh_type(r.get("Type")))] = {
                "AP": _pct(r.get("RoT_Prior_B2B")), "AQ": _pct(r.get("N0_Attempt"))}

    # ---- Unreg: BD,BE keyed by (DPID, POH) ----
    unreg = {}
    if df_unreg is not None and len(df_unreg):
        un = _clean_cols(df_unreg)
        for _, r in un.iterrows():
            k = _norm_one(r["DP ID"])
            poh = _is_poh_type(r.get("Type"))
            unreg[(k, poh)] = {
                "BD": _pct(r.get("LI_N0")),
                "BE": _pct(r.get("LI_POH")) if poh else _pct(r.get("PU Scan_N0")),
            }

    # ---- Drop-off KPI: CC(region text), CD(LI-POH) ----
    lipoh = {}
    if df_lipoh is not None and len(df_lipoh):
        lp = df_lipoh
        dp_c = _first_col(lp, ["DP ID", "DPID", "dp_id"])
        reg_c = _first_col(lp, ["Clean Up Region", "Region", "Clean Region"])
        poh_c = _first_col(lp, ["LI-POH", "LI_POH", "LiPOH", "LI POH"])
        if dp_c:
            for _, r in lp.iterrows():
                k = _norm_one(r[dp_c])
                lipoh[k] = {
                    "CC": ("" if not reg_c or pd.isna(r.get(reg_c)) else r.get(reg_c)),
                    "CD": _pct(r.get(poh_c)) if poh_c else 0.0,
                }

    # ---- Last Mile: CM:CV by DPID ----
    lmt = {}
    if df_lm_tracker is not None and len(df_lm_tracker):
        tr = df_lm_tracker
        dp_c = _first_col(tr, ["DP ID", "DPID", "dp_id"])
        if dp_c:
            resolved = {L: _first_col(tr, [name]) for L, name in lm_map.items()}
            for _, r in tr.iterrows():
                k = _norm_one(r[dp_c])
                d = {}
                for L, src in resolved.items():
                    if src is None:
                        continue
                    v = r.get(src)
                    d[L] = _pct(v) if L in _LM_PCT_COLS else ("" if pd.isna(v) else v)
                lmt[k] = d

    def kpi_row_fn(dpid_key, region, mitra):
        poh = _is_poh_region(region)
        out = {}
        if dpid_key in acq:
            out.update(acq[dpid_key])
        a = allo.get((dpid_key, poh), allo.get((dpid_key, not poh)))
        if a:
            out.update(a)
        u = unreg.get((dpid_key, poh), unreg.get((dpid_key, not poh)))
        if u:
            out.update(u)
        if dpid_key in lipoh:
            out.update(lipoh[dpid_key])
        if dpid_key in lmt:
            out.update(lmt[dpid_key])
        if prev_claim_kpi and dpid_key in prev_claim_kpi:
            for L, v in prev_claim_kpi[dpid_key].items():
                out[L] = _num(v) - 1.0
        return out

    return kpi_row_fn
