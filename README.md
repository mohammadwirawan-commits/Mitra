# Mitra Commission — "Pivot All Region" Automation

Python automation that rebuilds the monthly **Mitra Commission "Pivot All Region"**
workbook (`M##Y#### Pivot All Region.xlsx`), replacing a manual, multi-file
spreadsheet process. Raw Metabase exports + KPI intake + penalty inputs flow
through a pipeline and are stamped into a copy of the prior-month template.

> ⚠️ **Data policy:** This repository contains **code and blank templates only**.
> All real commission/penalty data (mitra names, financial figures, tracking IDs,
> PII) is excluded via [`.gitignore`](.gitignore) and must never be committed.

## Pipeline stages

```
Data → Pivot Data → Commission → Adjustment & Penalty
     → Data Sources → Rekap Commission → By Owner → PPh → Generate
```

## Status

| Phase | Scope | State |
|-------|-------|-------|
| **1**  | Data → Pivot Data → Commission            | ✅ Complete & validated |
| **2a** | Adjustment & Penalty (+ Penalty compile)  | ✅ Complete & verified |
| **2b** | Data Sources → Rekap Commission → By Owner | ⬜ Next |
| **3**  | PPh → Generate → Generate By Owner         | ⬜ Planned |

June was validated end-to-end against the golden workbook to within
**−0.009%** of grand-total commission and accepted as validated.

## Repository layout

```
automation/            # Pipeline code (par_*.py), builders (make_*.py),
                       #   local harnesses, and PROJECT_CONTEXT.md (full spec)
TEMPLATES/             # The 2 deliverable intake files (blank templates)
Mitra Commission Intake - Jul-2026.ipynb   # Notebook that drives a monthly run
```

**Excluded from git** (present locally, needed to run/validate): `reference/`
(golden workbooks), `sample-data-june-2026/`, `_ready-to-run-Jun-2026/`,
`Results/`, `July Actual/`, and generated output workbooks.

## Key modules (`automation/`)

- `par_pipeline.py` — engine + KPI row builder (Data → Pivot Data → Commission).
- `par_writer.py` — template round-trip writer (values into a copy of the prior workbook).
- `par_adjustment.py` — Adjustment & Penalty sheet + its embedded pivots.
- `par_format_sr.py` — compiles the penalty detail from raw SR files / Penalty Intake.
- `make_master_intake.py`, `make_penalty_intake_template.py` — regenerate the templates in `TEMPLATES/`.
- `run_local_june.py`, `reconcile_june.py`, `verify_commission_engine.py` — local verification harnesses.
- `PROJECT_CONTEXT.md` — full working spec, mappings, and per-phase detail.

## Running a month

1. Owners fill the two intake files (from `TEMPLATES/`) plus the shared
   Penalty Intake sheet for that month.
2. Drop the raw Metabase export(s) into the month's `Raw Metabase/` folder.
3. Run the notebook (`MONTH_OVERRIDE="YYYY-MM"`, Run all).
4. Verify by diffing the generated sheet against the golden
   `Pivot All Region.xlsx` using the engine-isolation method (feed golden
   inputs, compare computed outputs, float tolerance).

See `automation/PROJECT_CONTEXT.md` for the authoritative detail.
