"""Build the single MASTER intake template (one Google-Sheet-ready workbook).

Consolidates the first-stage intake tabs and the all-region KPI/penalty tabs into
ONE workbook, grouped and colour-coded by the stakeholder who fills each tab, with
a README / assignment tab up front. Upload it to Drive, "Open as Google Sheets",
and share ONE link; each stakeholder pastes into their own tab; then
File -> Download -> .xlsx into the month's Intake folder.

Headers + the grey example row are copied verbatim from the two existing templates
so the column schemas the notebook reads stay exactly correct. Tab names are NOT
changed (the pipeline reads them by name) — ownership lives in the README + tab colour.
"""
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

DOCS = r"C:/Users/mohammad.wirawan_nin/Documents"
FIRST = os.path.join(DOCS, "Mitra", "Mitra_Commission_Input_Template.xlsx")
ALLR = os.path.join(DOCS, "Mitra All Region", "automation", "Mitra Intake All Region TEMPLATE.xlsx")
OUT = os.path.join(DOCS, "Mitra All Region", "TEMPLATES", "Mitra Commission Intake - MASTER TEMPLATE.xlsx")

# Tabs defined inline (not present in either source template). Header + one example row.
INLINE = "INLINE"
INLINE_TABS = {
    # Claims: prior-month dispute rows, pasted at Data-sheet granularity (columns match the
    # Data sheet so build_data_sheet's df_disputes hook consumes them directly).
    "Claims": (
        ["Mapping Type", "Global Shipper ID", "Shipper Name", "DPID", "DPMS", "Mitra Name",
         "Region", "Total Parcel", "Total Weight", "Total Commission"],
        ["Acquisition Claim", 11327547, "Lazada (TulOjeTy)", 3165, 3165,
         "Mitra JKT @ Pademangan Barat", "All Region - Non POH", 2, 2.0, 1700],
    ),
}

HDR_FILL = PatternFill("solid", fgColor="305496")
HDR_FONT = Font(bold=True, color="FFFFFF")
REF_FILL = PatternFill("solid", fgColor="595959")            # reference (do-not-edit) header
EX_FILL = PatternFill("solid", fgColor="F2F2F2")
EX_FONT = Font(italic=True, color="808080")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# group -> (colour, [ (tab, source, "what to paste", "what it drives / function") ])
GROUPS = [
    ("Wira (You)", "2E75B6", [
        ("Acquisition KPI", ALLR, "Marketplace Acquisition KPI per DP ID (Total Score authoritative).",
         "Sets the Acquisition payout %: Total Score >=0.99 -> +5% bonus, >=0.90 -> as-is, else -5%."),
        ("Unreg KPI", ALLR, "Unregistered Shipper KPI (LI-N+0, LI-PU / LI-POH) per DP ID + Type.",
         "Unregistered-Eligible KPI multiplier: LI-N+0 >=0.87 AND LI-PU/POH >=0.90 -> full, else -5%."),
        ("2. LIPOH DropOff-SH", FIRST, "Drop Off to SH calc: Clean Up Region + LI-POH per DP ID.",
         "Sets the drop-off-to-SH rate (Rp/parcel) by region group (Jabodetabek vs Luar) x LI-POH band."),
        ("5. Last Mile Volume", FIRST, "Last Mile per-parcel volume export.",
         "The base parcel list for Last Mile commission (each parcel priced by status x size)."),
        ("6. Last Mile Tracker", FIRST, "Last Mile hub KPI + PPP rates (keyed by hub_name).",
         "Supplies PPP rates (Reg/Bulky/Superbulky) + hub Final Tier/Score used for LM commission & KPI."),
    ]),
    ("Kamal", "548235", [
        ("Allocation KPI", ALLR, "Allocation Shipper KPI: ROT Prior + N+0 Attempt per DP ID + Type.",
         "Allocation KPI multiplier: ROT Prior >=0.95 AND N+0 >=0.99 -> full commission, else -5%."),
    ]),
    ("Davy / Niko", "C55A11", [
        ("1a. Fake Order-Shipper", FIRST, "Fake orders at DP x Shipper level.",
         "Flags fraud parcels: removed from the commission base + drive the -Rp10k/parcel fake penalty."),
        ("1b. Fake Order-TrID", FIRST, "Fake orders at tracking-id level.",
         "Same as 1a but matched per tracking id (catches fakes the shipper-level list misses)."),
    ]),
    ("Deshy", "7030A0", [
        ("4. Manifest LEX", FIRST, "LEX Hub Pickup manifest (tracking id, DPID, commission).",
         "Basis for LEX Hub Pickup commission (Rp 500 per pickup parcel)."),
        ("Cross Border", ALLR, "XB parcels: tracking_id, DPID, delivery_fee (comm = 30%).",
         "Cross-Border commission = 30% x delivery fee for each XB parcel."),
        ("Data Mitra Changes", ALLR, "New/changed Mitra master data + ownership / SKB.",
         "Feeds Data Sources + PPh: sets entity type (tax rate), SKB, bank details, hold-commission flag."),
    ]),
    ("Fanny", "158A8A", [
        ("7. List Drop SH", FIRST, "Mitra Drop-off-to-SH participation list + join/withdraw dates.",
         "Sets drop-off eligibility window: only parcels within the active period earn drop-off commission."),
        ("8. List Program MPM", FIRST, "Mitra Pickup Mitra program (pickup-for / pickup-by).",
         "Identifies MPM parcels and routes their pickup commission to the pickup-by Mitra."),
    ]),
    # NOTE: the 4 penalty inputs (PPDP, Fake POD-PODA, Missing LM, Outstanding) moved OUT of this
    # master sheet into the separate "Penalty Intake" template (make_penalty_intake_template.py) —
    # they come from different external stakeholders and feed Adjustment&Penalty via par_format_sr.
    ("You / Finance (compiled)", "BF9000", [
        ("Claims", INLINE, "Prior-month dispute/claim rows (Acquisition/Allocation/etc. Claim).",
         "Added to the Data sheet as claim rows -> Commission claim columns (K/L, AF/AG, ...)."),
        ("Adjustment", ALLR, "Approved commission adjustments (compiled from objections).",
         "Added/subtracted on Adjustment & Penalty (only objections already validated by the team)."),
    ]),
    ("Reference — DO NOT EDIT (pre-filled)", "808080", [
        ("3. Region Split", FIRST, "DP ID -> DPMS ID + Region Split. Maintained centrally.",
         "Master mapping: assigns every Mitra to a region (POH / Non-POH / By Owner) used by ALL sheets."),
        ("SSB", FIRST, "SSB export: Tracking ID -> L2 Name + Script ID. System file.",
         "Registered pricing: sets destination (to-city L2) + triggers the /1.01 delivery-fee adjustment."),
    ]),
]


def _read_tab(path, tab):
    wb = openpyxl.load_workbook(path)
    ws = wb[tab]
    header = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    example = [ws.cell(row=2, column=c).value for c in range(1, ws.max_column + 1)] if ws.max_row >= 2 else []
    return header, example


def _add_tab(wb, tab, header, example, colour, is_ref):
    ws = wb.create_sheet(title=tab)
    ws.sheet_properties.tabColor = colour
    for c, h in enumerate(header, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = HDR_FONT
        cell.fill = REF_FILL if is_ref else HDR_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[cell.column_letter].width = max(12, min(30, len(str(h)) + 4))
    for c, v in enumerate(example, start=1):
        cell = ws.cell(row=2, column=c, value=v)
        cell.font = EX_FONT
        cell.fill = EX_FILL
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 30
    return ws


def _build_readme(wb):
    ws = wb.create_sheet(title="README", index=0)
    ws.sheet_properties.tabColor = "000000"
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 70
    title = Font(bold=True, size=14)
    h = Font(bold=True, color="FFFFFF")
    hf = PatternFill("solid", fgColor="305496")

    ws.column_dimensions["D"].width = 78
    ws["A1"] = "MITRA COMMISSION — MONTHLY INTAKE (Master)"
    ws["A1"].font = title
    ws["A2"] = "Month: ____________          Deadline: by the 3rd working day"
    ws["A2"].font = Font(bold=True, italic=True)

    lines = [
        "",
        "HOW TO USE",
        "1. Find the tab(s) with your name below (tabs are colour-coded by owner).",
        "2. Paste your data UNDER the header row. Keep the column headers exactly as-is.",
        "3. DELETE the grey example row once you've pasted your data.",
        "4. Do NOT rename tabs, add/remove/reorder columns, or edit the Reference tabs.",
        "5. When everyone is done: File → Download → Microsoft Excel (.xlsx),",
        "   name it  \"Mitra Commission Intake {Mon-YYYY}.xlsx\"  (e.g. Jul-2026),",
        "   and upload it to that month's  Intake  folder on Drive.",
        "",
        "PENALTIES ARE SEPARATE: the 4 penalty inputs (PPDP, Fake POD-PODA, Missing LM, Outstanding)",
        "live in the separate \"Penalty Intake\" sheet — Desy / Arbeta / Opi / Finance fill that one.",
        "",
        "TAB ASSIGNMENTS",
    ]
    r = 3
    for text in lines:
        ws.cell(row=r, column=1, value=text).font = Font(bold=(text in ("HOW TO USE", "TAB ASSIGNMENTS")))
        r += 1

    # assignment table header
    for c, lbl in enumerate(["Owner", "Tab", "What to paste", "What it drives (function)"], start=1):
        cell = ws.cell(row=r, column=c, value=lbl)
        cell.font = h
        cell.fill = hf
        cell.alignment = Alignment(horizontal="left", vertical="center")
    r += 1
    for owner, colour, tabs in GROUPS:
        for tab, _src, wtp, func in tabs:
            ws.cell(row=r, column=1, value=owner).font = Font(bold=True, color=colour)
            ws.cell(row=r, column=2, value=tab)
            ws.cell(row=r, column=3, value=wtp).alignment = Alignment(wrap_text=True, vertical="top")
            ws.cell(row=r, column=4, value=func).alignment = Alignment(wrap_text=True, vertical="top")
            r += 1
    ws.freeze_panes = "A2"
    return ws


def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)                      # drop default sheet
    for owner, colour, tabs in GROUPS:
        is_ref = owner.startswith("Reference")
        for tab, src, _wtp, _func in tabs:
            header, example = INLINE_TABS[tab] if src == INLINE else _read_tab(src, tab)
            _add_tab(wb, tab, header, example, colour, is_ref)
    _build_readme(wb)                         # inserted at index 0 (first tab)
    wb.save(OUT)
    print(f"Saved master intake template: {OUT}")
    print(f"Tabs ({len(wb.sheetnames)}): {wb.sheetnames}")


if __name__ == "__main__":
    main()
