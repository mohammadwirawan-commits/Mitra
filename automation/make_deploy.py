"""Assemble a ready-to-copy `deploy/` folder — the exact files a monthly Colab run needs.

deploy/ is a generated snapshot (git-ignored); the source of truth stays in automation/,
TEMPLATES/, and the notebook. Re-run this after any code change to refresh the snapshot:

    python automation/make_deploy.py

Layout produced:
    deploy/
      README.md                              <- what goes where
      automation/                            <- copy INTO  {DRIVE_ROOT}/automation  on Drive
        par_pipeline.py  par_writer.py  par_adjustment.py  par_format_sr.py
      Mitra Commission Intake - <notebook>   <- open in Colab
      intake-templates/                      <- share with the team (one-time / when schema changes)
        Mitra Commission Intake - MASTER TEMPLATE.xlsx
        Penalty Intake TEMPLATE.xlsx
"""
import os
import glob
import shutil

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # repo root
AUTO = os.path.join(BASE, "automation")
DEPLOY = os.path.join(BASE, "deploy")

# The 4 modules the notebook imports at runtime (cell 9e7ae9d6).
RUNTIME_MODULES = ["par_pipeline.py", "par_writer.py", "par_adjustment.py", "par_format_sr.py"]
TEMPLATES = [
    os.path.join(BASE, "TEMPLATES", "Mitra Commission Intake - MASTER TEMPLATE.xlsx"),
    os.path.join(BASE, "TEMPLATES", "Penalty Intake TEMPLATE.xlsx"),
    os.path.join(AUTO, "intake_cleanup.gs"),   # Apps Script: monthly "clear last month" button
]

README = """# Deploy — what to copy for a monthly run

This folder is a generated snapshot (regenerate with `python automation/make_deploy.py`).

## 1. Runtime code  ->  Drive `{DRIVE_ROOT}/automation/`
Copy the 4 files in **`automation/`** here into your Drive's
`Mitra Commission/automation/` folder (overwrite the old ones). The notebook imports
them from there (`sys.path` -> `{DRIVE_ROOT}/automation`).

  - par_pipeline.py   par_writer.py   par_adjustment.py   par_format_sr.py

## 2. Notebook  ->  Colab
Open **`{notebook}`** in Colab (or upload to Drive and open). Set `MONTH_OVERRIDE`
(top cell) to the target month like `2026-07`, then Run all.

## 3. Intake templates  ->  the team (one-time, or when the schema changes)
The two files in **`intake-templates/`** are the stakeholder-facing intakes. Upload
`Mitra Commission Intake - MASTER TEMPLATE.xlsx` to Drive, "Open as Google Sheets",
and share one link; the team fills their tabs and exports it as
`Mitra Commission Intake {{Mon-YYYY}}.xlsx` into the month's `Intake/` folder.
`Penalty Intake TEMPLATE.xlsx` is the separate penalty intake.

### Monthly reset button (one-time install)
`intake-templates/intake_cleanup.gs` is a Google Apps Script for the shared intake
Sheet. In the Sheet: Extensions ▸ Apps Script ▸ paste the file ▸ Save ▸ reload.
It adds "Intake Tools ▸ Clear last month's entries" — clears the pasted data in the
owner tabs (keeps headers, formatting, the reference tabs, and README). Run it each
month after that month's workbook has been generated.

## 4. After the run — REQUIRED
**Open the downloaded `M##Y#### Pivot All Region.xlsx` in Excel once and Save it.**
That refreshes the pivots (set to refresh-on-open), recalculates every formula
(Commission / Rekap / By-Owner / PPh), and caches the values so next month's PPh
prior-month freeze works.

## Prereqs already on Drive (unchanged)
Prior-month `Pivot All Region.xlsx` (opened+saved in Excel), the month's `Intake/`,
`Raw Metabase/`, penalty inputs, and the Origin City / Rate Card reference files.
"""


def _copy(src, dst):
    """Overwrite-in-place copy that tolerates a locked destination (e.g. the file is
    open in Excel) — warns and skips instead of aborting the whole snapshot."""
    try:
        shutil.copy2(src, dst)
    except PermissionError:
        print(f"  [skip] in use, left as-is: deploy/{os.path.relpath(dst, DEPLOY)}")


def main():
    # Don't rmtree (a file open in Excel would lock the whole tree) — overwrite in place.
    os.makedirs(os.path.join(DEPLOY, "automation"), exist_ok=True)
    os.makedirs(os.path.join(DEPLOY, "intake-templates"), exist_ok=True)

    for m in RUNTIME_MODULES:
        _copy(os.path.join(AUTO, m), os.path.join(DEPLOY, "automation", m))

    nb = glob.glob(os.path.join(BASE, "Mitra Commission Intake*.ipynb"))
    nb_name = os.path.basename(nb[0]) if nb else "Mitra Commission Intake.ipynb"
    if nb:
        _copy(nb[0], os.path.join(DEPLOY, nb_name))

    for t in TEMPLATES:
        if os.path.exists(t):
            _copy(t, os.path.join(DEPLOY, "intake-templates", os.path.basename(t)))

    with open(os.path.join(DEPLOY, "README.md"), "w", encoding="utf8") as f:
        f.write(README.replace("{notebook}", nb_name))

    print(f"Assembled: {DEPLOY}")
    for root, _dirs, files in os.walk(DEPLOY):
        for fn in sorted(files):
            rel = os.path.relpath(os.path.join(root, fn), DEPLOY)
            print(f"  deploy/{rel}")


if __name__ == "__main__":
    main()
