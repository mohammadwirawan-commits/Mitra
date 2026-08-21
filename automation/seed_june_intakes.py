"""Seed the two intake templates with real JUNE data + the combined Metabase file, so the
pipeline can be run end-to-end for June as a self-serve test (PART D).

Produces a ready-to-upload tree under  <repo>/_ready-to-run-Jun-2026/ :
    Intake/Mitra Commission Intake Jun-2026.xlsx   (all master-intake tabs, filled)
    Intake/Penalty Intake Jun-2026.xlsx            (4 penalty tabs, filled)
    Raw Metabase/mitra_commission_2026-06.xlsx     (the 3 batch files combined)

The blank TEMPLATES/ files are left untouched — these are separate filled copies.

Data sources (all verified to match the template schemas):
  * first-stage tabs   <- Mitra/Resources/*  +  Mitra/Mitra Origin City & SSB File.xlsx
  * KPI tabs           <- sample-data-june-2026/*.csv
  * Cross Border/Claims/Adjustment <- reconstructed from golden reference/M06 (reproduces June)
  * SSB tab            <- headers only (no June routing-SSB export on hand; nil effect under
                          raw-fee mode — see PROJECT_CONTEXT / the run notes)
  * penalty tabs       <- sample-data-june-2026/Penalty SR/* + Format SR Penalty outstanding
"""
from __future__ import annotations
import os
import sys
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import par_format_sr as fsr

BASE = os.path.dirname(HERE)                                   # Mitra All Region
DOCS = os.path.dirname(BASE)
MITRA = os.path.join(DOCS, "Mitra")
RES = os.path.join(MITRA, "Resources")
DATA = os.path.join(BASE, "sample-data-june-2026")
REF = os.path.join(BASE, "reference")
TEMPLATES = os.path.join(BASE, "TEMPLATES")
PENALTY_SR = os.path.join(DATA, "Penalty SR")

OUT = os.path.join(BASE, "_ready-to-run-Jun-2026")
OUT_INTAKE = os.path.join(OUT, "Intake")
OUT_META = os.path.join(OUT, "Raw Metabase")

MASTER_TMPL = os.path.join(TEMPLATES, "Mitra Commission Intake - MASTER TEMPLATE.xlsx")
GOLDEN = os.path.join(REF, "M06Y2026 Pivot All Region.xlsx")

# master tab -> (source file, sheet)  [direct copy; schemas confirmed identical]
TAB_SOURCES = {
    "1a. Fake Order-Shipper": (os.path.join(RES, "1. List Fake Order.xlsx"), "Shipper level"),
    "1b. Fake Order-TrID":    (os.path.join(RES, "1. List Fake Order.xlsx"), "Trids level"),
    "2. LIPOH DropOff-SH":    (os.path.join(RES, "2. Mitra Drop Off to SH Calculation.xlsx"), 0),
    "3. Region Split":        (os.path.join(RES, "3. Mitra Region Split.xlsx"), 0),
    "4. Manifest LEX":        (os.path.join(RES, "4. Manifest LEX Mitra.xlsx"), 0),
    "5. Last Mile Volume":    (os.path.join(RES, "5. Last Mile Volume.xlsx"), 0),
    "6. Last Mile Tracker":   (os.path.join(RES, "6. Last Mile Tracker.xlsx"), 0),
    "7. List Drop SH":        (os.path.join(RES, "List Mitra Drop SH.xlsx"), "Drop Off"),
    "8. List Program MPM":    (os.path.join(RES, "List Program MPM.xlsx"), "MPM"),
}
KPI_SOURCES = {
    "Acquisition KPI": os.path.join(DATA, "Acquisition Jun 26.csv"),
    "Allocation KPI":  os.path.join(DATA, "[2606] Mitra KPI Calculation_Jun 26 - Allocation KPI Compile (1).csv"),
    "Unreg KPI":       os.path.join(DATA, "[2606] Mitra KPI Calculation_Jun 26 - Unreg KPI Compile (1).csv"),
}
CLAIM_TYPES = ("Acquisition Claim", "ALLOCATION Claim", "MP Registered Claim")


def _template_header(tab):
    df = pd.read_excel(MASTER_TMPL, sheet_name=tab, nrows=0)
    return list(df.columns)


def _reconstruct_manual():
    """Cross Border / Claims / Adjustment reconstructed from golden M06 (reproduces June)."""
    data = pd.read_excel(GOLDEN, sheet_name="Data")
    mt = data["Mapping Type"].astype(str)

    # Claims: golden Data claim rows -> the Claims tab columns.
    claims = data[mt.isin(CLAIM_TYPES)].copy()
    claims_out = pd.DataFrame({
        "Mapping Type": claims["Mapping Type"], "Global Shipper ID": claims["Global Shipper ID"],
        "Shipper Name": claims["Shipper Name"], "DPID": claims["DPID"], "DPMS": claims["DPMS"],
        "Mitra Name": claims["Mitra Name"], "Region": claims["Region"],
        "Total Parcel": claims["Total Parcel"], "Total Weight": claims["Total Weight"],
        "Total Commission": claims["Total Commission"],
    })

    # Cross Border: golden aggregates 1 row/DPID; expand to tracking-level rows (parcel count),
    # each carrying an equal share of the delivery fee, so build_data_sheet re-aggregates to golden.
    xb = data[mt == "Cross Border"]
    xb_rows = []
    for _, r in xb.iterrows():
        n = int(r["Total Parcel"]) or 1
        fee_each = float(r["Total Delivery Fee"]) / n
        for i in range(n):
            xb_rows.append({"tracking_id": f"SPXXB-{int(r['DPID'])}-{i+1:04d}",
                            "DPID": r["DPID"], "DPMS ID": r["DPMS"],
                            "Region Split": r["Region"], "delivery_fee": fee_each})
    xb_out = pd.DataFrame(xb_rows, columns=_template_header("Cross Border"))

    # Adjustment: the E/F-filled rows of the golden Adjustment&Penalty main table (A:J).
    adj = pd.read_excel(GOLDEN, sheet_name="Adjustment&Penalty", header=2)  # labels on row 3
    adj = adj.rename(columns={adj.columns[0]: "DP ID"})
    def _nz(x):
        return pd.notna(x) and x not in (0, "-", "")
    keep = adj.apply(lambda r: (_nz(r.get("Adjustment Komisi")) or _nz(r.get("Adjustment Bruto")))
                     and not (_nz(r.get("Penalty Parcel")) or _nz(r.get("Penalty PPDP"))), axis=1)
    a = adj[keep]
    adj_out = pd.DataFrame({
        "DP ID": a["DP ID"], "DPMS ID": a.get("DPMS ID"), "Nama Mitra": a.get("Nama Mitra"),
        "Region": a.get("Region"), "Adjustment Komisi": a.get("Adjustment Komisi"),
        "Adjustment Bruto": a.get("Adjustment Bruto"), "Adjustment PPh": a.get("Adjustment PPh"),
        "Note": a.get("Note"),
    })
    return {"Cross Border": xb_out, "Claims": claims_out, "Adjustment": adj_out}


def build_master(path):
    sheets = {}
    # first-stage tabs (direct copy)
    for tab, (src, sheet) in TAB_SOURCES.items():
        sheets[tab] = pd.read_excel(src, sheet_name=sheet)
    # KPI tabs
    for tab, src in KPI_SOURCES.items():
        sheets[tab] = pd.read_csv(src)
    # SSB: headers only (no June routing-SSB export; nil effect under raw-fee mode)
    sheets["SSB"] = pd.DataFrame(columns=_template_header("SSB"))
    # manual tabs reconstructed from golden
    sheets.update(_reconstruct_manual())
    # Data Mitra Changes: keep the template's header + example row (not read by the pipeline yet)
    sheets["Data Mitra Changes"] = pd.read_excel(MASTER_TMPL, sheet_name="Data Mitra Changes")

    # write in the template's tab order (README first, then the rest)
    order = pd.ExcelFile(MASTER_TMPL).sheet_names
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        pd.DataFrame({"README": ["Filled with June 2026 data by seed_june_intakes.py. "
                                 "SSB tab intentionally empty (see run notes)."]}).to_excel(
            w, sheet_name="README", index=False)
        for tab in order:
            if tab in sheets:
                sheets[tab].to_excel(w, sheet_name=tab, index=False)
    return {t: len(df) for t, df in sheets.items()}


def build_penalty(path):
    ppdp = pd.read_excel(os.path.join(PENALTY_SR, "Penalty PPDP Mitra June 2026.xlsx"), sheet_name="RAW")
    fp = os.path.join(PENALTY_SR, "Penalty Mitra Last Mile (Fake POD & Fake PODA) 06.2026.xlsx")

    def _fp(sheet, typ):
        d = pd.read_excel(fp, sheet_name=sheet)
        return pd.DataFrame({
            "Type (POD/PODA)": typ, "Hub Name": d[fsr._col(d, "Hub Name")],
            "DP ID": d[fsr._col(d, "dp_id")], "Nama Mitra": d[fsr._col(d, "dp_name")],
            "Tracking ID": d[fsr._col(d, "Tracking ID")], "Granular Status": d[fsr._col(d, "Granular Status")],
            "Penalty": d[fsr._col(d, "Penalty")],
        })

    miss = pd.read_excel(os.path.join(PENALTY_SR, "Penalty LM Mitra June 2026 (2).xlsx"), sheet_name=0)
    t_ppdp = pd.DataFrame({
        "DP ID": ppdp[fsr._col(ppdp, "DP ID")], "Nama Mitra": ppdp[fsr._col(ppdp, "Nama Mitra")],
        "Tracking ID": ppdp[fsr._col(ppdp, "Tracking ID")], "Shipper Name": ppdp[fsr._col(ppdp, "Shipper Name")],
        "Granular Status": ppdp[fsr._col(ppdp, "Granular Status")],
        "Total Penalty Amount": ppdp[fsr._col(ppdp, "Total Penalty Amount")],
    })
    t_podpoda = pd.concat([_fp("Penalty PODA", "PODA"), _fp("Penalty POD", "POD")], ignore_index=True)
    t_miss = pd.DataFrame({
        "Tracking ID": miss[fsr._col(miss, "Tracking ID")],
        "Investigating Hub Name": miss[fsr._col(miss, "Investigating Hub Name")],
        "Order Granular Status": miss[fsr._col(miss, "Order Granular Status")],
        "Type": miss[fsr._col(miss, "Type")], "Est Claim Amount": miss[fsr._col(miss, "Est claim amount")],
        "Final Check CL": miss[fsr._col(miss, "Final Check CL")],
    })
    osw = pd.read_excel(os.path.join(DATA, "Format SR Penalty - Jun'26.xlsx"),
                        sheet_name="Penalty Outstanding", header=4)
    t_os = pd.DataFrame({
        "DP ID": osw[fsr._col(osw, "DP ID")], "DPMS ID": osw[fsr._col(osw, "DPMS")],
        "Nama Mitra": osw[fsr._col(osw, "Nama Mitra")], "Region": osw[fsr._col(osw, "region")],
        "Invoice": osw[fsr._col(osw, "Invoice")], "Total Outstanding": osw[fsr._col(osw, "Total Outstanding")],
        "Denda": osw[fsr._col(osw, "Penalty")], "Total Penalty": osw[fsr._col(osw, "Total Penalty")],
        "Periode": osw[fsr._col(osw, "Periode")], "Note": osw[fsr._col(osw, "Note")],
    }).dropna(how="all")
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        t_ppdp.to_excel(w, sheet_name="Penalty PPDP", index=False)
        t_podpoda.to_excel(w, sheet_name="Penalty Fake POD-PODA", index=False)
        t_miss.to_excel(w, sheet_name="Penalty Missing LM", index=False)
        t_os.to_excel(w, sheet_name="Penalty Outstanding", index=False)
    return {"Penalty PPDP": len(t_ppdp), "Penalty Fake POD-PODA": len(t_podpoda),
            "Penalty Missing LM": len(t_miss), "Penalty Outstanding": len(t_os)}


def build_metabase(path):
    batches = [os.path.join(DATA, f"mitra_commission_2026-06_batch_{i}.xlsx") for i in (1, 2, 3)]
    frames, cols0 = [], None
    for b in batches:
        df = pd.read_excel(b)
        if cols0 is None:
            cols0 = list(df.columns)
        elif list(df.columns) != cols0:
            print(f"  WARNING: {os.path.basename(b)} columns differ from batch_1")
        frames.append(df)
    combined = pd.concat(frames, ignore_index=True)
    combined.to_excel(path, index=False)
    return len(combined), [len(f) for f in frames]


def main():
    os.makedirs(OUT_INTAKE, exist_ok=True)
    os.makedirs(OUT_META, exist_ok=True)
    master_path = os.path.join(OUT_INTAKE, "Mitra Commission Intake Jun-2026.xlsx")
    penalty_path = os.path.join(OUT_INTAKE, "Penalty Intake Jun-2026.xlsx")
    meta_path = os.path.join(OUT_META, "mitra_commission_2026-06.xlsx")

    print("[1/3] master intake ...")
    mrows = build_master(master_path)
    for t, n in mrows.items():
        print(f"    {t:26} {n:>7} rows")
    print("[2/3] penalty intake ...")
    prows = build_penalty(penalty_path)
    for t, n in prows.items():
        print(f"    {t:26} {n:>7} rows")
    print("[3/3] metabase combine ...")
    total, parts = build_metabase(meta_path)
    print(f"    combined {parts} -> {total} rows")
    print("\nWrote ready-to-upload tree:")
    print(f"  {master_path}")
    print(f"  {penalty_path}")
    print(f"  {meta_path}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    main()
