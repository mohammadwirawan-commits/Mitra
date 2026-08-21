"""Validate a filled 'Mitra Intake All Region' workbook against the columns the
Pivot-All-Region stage requires. Mirrors validate_intake_columns.py.

    python validate_all_region_intake.py "Mitra Intake All Region Jun-2026.xlsx"

Defaults to the TEMPLATE in this folder.
"""
import os, sys
import pandas as pd

_DIR = os.path.dirname(os.path.abspath(__file__))
FILE = sys.argv[1] if len(sys.argv) > 1 else os.path.join(_DIR, "Mitra Intake All Region TEMPLATE.xlsx")

# tab -> required columns
REQUIRED = {
    "Acquisition KPI":       ["DP ID", "Total Score"],  # Total Score is read directly
    "Allocation KPI":        ["DP ID", "Type", "N0_Attempt", "RoT_Prior_B2B"],
    "Unreg KPI":             ["DP ID", "Type", "LI_N0", "PU Scan_N0", "LI_POH"],
    "Cross Border":          ["tracking_id", "DPID", "delivery_fee"],
    "Penalty PPDP":          ["DPMS ID", "Total Penalty Amount"],
    "Penalty Fake POD-PODA": ["DPMS ID", "Amount"],
    "Penalty Missing LM":    ["DPMS ID", "Estimasi Klaim"],
    "Adjustment":            ["DP ID", "Region", "Adjustment Komisi", "Adjustment Bruto", "Adjustment PPh", "Note"],
    "Penalty Outstanding":   ["DP ID", "Invoice", "Total Outstanding", "Periode"],
    "Data Mitra Changes":    ["DPMS ID", "Owner Name", "Tax ID", "Jenis Usaha"],
}

xls = pd.ExcelFile(FILE)
present_tabs = set(xls.sheet_names)
all_ok = True
for tab, cols in REQUIRED.items():
    if tab not in present_tabs:
        print(f"[FAIL] tab missing: '{tab}'")
        all_ok = False
        continue
    df = pd.read_excel(xls, sheet_name=tab)
    missing = [c for c in cols if c not in df.columns]
    status = "OK " if not missing else "FAIL"
    if missing:
        all_ok = False
    print(f"[{status}] {tab:22s} needs {len(cols)} cols  missing={missing}")
    if missing:
        print(f"        actual columns: {list(df.columns)}")

print()
print("ALL INTAKE COLUMNS PRESENT" if all_ok else "MISSING COLUMNS - stage would KeyError")
sys.exit(0 if all_ok else 1)
