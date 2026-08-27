# Mitra Commission — "Pivot All Region" Automation — Project Context

> Handoff doc for continuing in a new chat. Everything needed to resume is here:
> goal, file locations, environment, what's built + verified, exact schemas/rules,
> gotchas, and what's left.

---

## 1. Goal & scope

Automate the manual Excel work the Handover PDF calls **"Pivot All Region (Below)"
through "Generate by Owner"** — i.e. building the monthly workbook
`M##Y#### Pivot All Region.xlsx` from the Python `Results/` output + external supporting data.

**Sheets in scope, in order:** Data → Pivot Data → Commission → Adjustment&Penalty →
Data Sources → Rekap Commission → By Owner per Mitra → PPh → Generate → Generate By Owner.

### Decisions already made (with the user)
- **Output mode:** *values into a template copy* (pure Python; embedded pivots/formulas
  become computed static values). No Excel/xlwings needed.
- **Intake:** a single workbook, one tab per external data source, stored in Google Drive at
  `Mitra Commission/{Mon-YYYY}/Mitra Intake All Region {Mon-YYYY}.xlsx` (month is dynamic).
- **Delivery shape:** append a new "Pivot All Region" section to the existing Colab notebook,
  reusing its month/path config and helpers.
- **Phasing (recommended & adopted):**
  - **Phase 1 — Commission core:** Data → Pivot Data → Commission ✅ *COMPLETE* —
    engine + KPI wiring + output writer + notebook port all built & verified.
  - **Phase 2 — Charges & summary:** Adjustment&Penalty → Data Sources → Rekap → By Owner
  - **Phase 3 — Tax & outputs:** PPh → Generate → Generate By Owner

---

## 2. Environment

- **Working dir:** `C:\Users\mohammad.wirawan_nin\Documents\Mitra All Region`
- **Python (local):** `C:/Users/mohammad.wirawan_nin/AppData/Local/Programs/Python/Python312/python.exe`
  (3.12.7, openpyxl 3.1.5, pandas 3.0.5). ⚠ The `WindowsApps\python.exe` stub is non-functional — use the full path above.
- **Production runtime:** the existing pipeline is a **Google Colab** notebook with Drive
  mounted at `/content/drive/MyDrive/Mitra Commission`.
- **openpyxl performance note:** loading the 4 MB `M06Y2026 Pivot All Region.xlsx` takes ~85 s;
  a full round-trip is ~3 min and may drop pivot caches. The output writer should either accept
  that (single round-trip, values mode) or write lean value-only sheets.

---

## 3. Key files & locations

### The manual workbook to reproduce (golden reference)
- `Documents/Mitra All Region/M06Y2026 Pivot All Region.xlsx` — known-good **June 2026** output.
- `Documents/Mitra All Region/Mitra - Handover.pdf` — the step-by-step manual SOP.
- `Documents/Mitra All Region/Results/` — the Python pipeline output (a *fresh* run; differs
  slightly from the June deliverable, so it is NOT a row-exact oracle).

### The existing (already-automated) first-stage pipeline
- **Notebook:** `Documents/Mitra/Mitra Commission Intake - Jul-2026.ipynb` (114 cells). Pulls
  Metabase, reads the first-stage intake, applies rate-card/insurance/registered overrides,
  writes `Results/`. **Ends exactly where our automation begins.**
- **Config cell** `id=8p7d90yN2rTt`: `MONTH_OVERRIDE` (""=prev month), `MONTH_FOLDER`,
  `DRIVE_ROOT="/content/drive/MyDrive/Mitra Commission"`, `MONTH_PATH`, `RESULTS_DIR`,
  `INTAKE_FILE = f"{MONTH_PATH}/Intake/Mitra Commission Intake {MONTH_FOLDER}.xlsx"`,
  `ORIGIN_CITY_FILE`, `RATE_CARD_FILE`. **Reuse these.**
- **Helpers:** `save_versioned(df, base_path, base_name, ext)` (cell `g2jfgmkfaDPM`),
  folder-tree cell (`jnTHEr27ifz5`), intake-load cell (`X-T9SgHGtth6`).
- **First-stage intake template:** `Documents/Mitra/Mitra_Commission_Input_Template.xlsx`.
  Tabs (header row 1): `1a. Fake Order-Shipper`, `1b. Fake Order-TrID`, `2. LIPOH DropOff-SH`,
  `3. Region Split`, `4. Manifest LEX`, `5. Last Mile Volume`, `6. Last Mile Tracker`,
  `7. List Drop SH`, `8. List Program MPM`, `SSB`. **These dataframes are already loaded in the
  notebook** (`df_reg_type`, `df_lipoh`, `df_last_mile_tracker`, …) — our stage consumes them.
- Reference/resources: `Documents/Mitra/Resources/3. Mitra Region Split.xlsx` (Sheet1: `DP ID,
  DPMS ID, Region Split`), `Mitra Origin City & SSB File.xlsx`, `Rate Card.xlsx`.
- Existing validators (pattern to mirror): `Documents/Mitra/validate_intake_columns.py`,
  `validate_registered_logic.py`.

### Local June KPI/support files (for verification & schema locking)
- `Documents/Mitra All Region/Acquisition Jun 26.csv` — **Acquisition KPI** (see §6).
- `Downloads/Allocation KPI Compile - Jun 2026.csv`, `Downloads/Unreg KPI Compile Jun 26 [New].csv`.
- `Downloads/[2606] Mitra KPI Calculation_Jun 26 - {Allocation|Unreg} KPI Compile.csv`.
- `Desktop/Mitra Commission - Jun 2026/` — raw KPI component CSVs (LiPOH, LiPU, OCLI, Valid
  POD/PODA, RTS, Terminal D6, Attempt Rate, Drop Off to SH Calc, Unreg KPI Compile).

### Our code (all under `Documents/Mitra All Region/automation/`)
- `par_pipeline.py` — pipeline logic: `build_data_sheet`, `build_pivot`, `assemble_commission`,
  `compute_commission_derived`, **and `build_kpi_row_fn`** (reads intake/first-stage/prev-pivot
  frames into the pasted-KPI columns). **Core module** (pandas only, openpyxl-free).
- `par_writer.py` — **output writer** (openpyxl). `write_pivot_all_region(template, out, data_df,
  commission_df)`: template round-trip, stamps Data / Pivot Data / Commission as VALUES into a copy
  of the prior-month workbook. NB: openpyxl `cell(r,c,value=None)` is a no-op — clears assign via
  `.value = None` (see `_set`).
- `verify_commission_engine.py` — verifies the engine vs the golden file (passes 0 mismatches).
- `run_local_june.py` — local end-to-end harness: builds Data + Commission from June inputs, writes
  the workbook, and diffs pasted-KPI columns vs golden (splits absent-vs-differs). `--no-write` skips
  the slow write. Writes `M06Y2026 Pivot All Region_GENERATED.xlsx` in this folder.
- `make_intake_template.py` — generates `Mitra Intake All Region TEMPLATE.xlsx`.
- `validate_all_region_intake.py` — validates a filled intake's columns.
- `Mitra Intake All Region TEMPLATE.xlsx` — the generated intake template (deliverable).
- `PROJECT_CONTEXT.md` — this file.
- Full plan: `C:\Users\mohammad.wirawan_nin\.claude\plans\i-want-to-automate-gentle-clover.md`.

---

## 4. Data sheet (16 cols) — build logic  ✅ verified

Concatenate the 5 `Results/Shipper ID Level/*.csv` files into this schema, then look up
DPMS+Region by DPID, then append manual Cross-Border + prior-month dispute rows.

**Columns:** `Global Shipper ID, Shipper Name, DPID, DPMS, Mitra Name, Region, Shipper Tier,
Mapping Type, Total Parcel, Total Weight, Total Delivery Fee, % Commission, Total Commission,
Drop off to SH Parcel, Drop off to SH Commission, Insurance Fee`.

**Per-file mapping:**
- `Shipper Level Akuisisi, Alokasi, dan Unregistered Parcel.csv` → Mapping Type = `mapping_type`
  (ACQUISITION/ALLOCATION/Unregistered Eligible); parcels=`total_tracking_id`,
  weight=`total_nv_weight`, comm=`total_commission`, drop-off from `total_parcel_drop_off` /
  `total_drop_off_commission`.
- `Shipper Level Regular & COD Parcel.csv` → Mapping Type = **`remarks_cod_regular`** (Regular/COD,
  NOT the literal "Registered"); adds tier, `commission_pct`, `total_delivery_fee`,
  `total_insurance_fee`, `commission`.
- `Shipper Level Last Mile Parcel.csv` → Mapping Type `Last Mile`; parcels+comm only.
- `Shipper Level Lex Hub Pickup Parcel.csv` → `LEX Hub Pickup`; keys `global_shipper_id`, `DPID`, `Mitra Name`.
- `Shipper Level Mitra Pickup Mitra.csv` → `Mitra Pickup Mitra`; keys `DPID (Pickup by)`, `Mitra Name (Pickup by)`.

**DPMS/Region lookup:** join DPID → `3. Region Split` (`DP ID, DPMS ID, Region Split`). Missing
DPMS → DPID; missing Region → `All Region - Non POH`.

**Mapping Type vocabulary (for the pivot):** `ACQUISITION, ALLOCATION, Unregistered Eligible,
Regular, COD, LEX Hub Pickup, Mitra Pickup Mitra, Last Mile`, plus manual
`Acquisition Claim, Allocation Claim, MP Registered Claim, Unregistered Eligible Claim, Cross Border`.

**Verification:** LM/LEX/MPM aggregates match June exactly; Regular matches rows + 15%; region
lookup 99.4% (the 0.6% miss equals our Region Split file's own value → file-version drift, not a
bug). ACQ/ALLO/fee deltas are from the Results run differing from the June deliverable (expected).

---

## 5. Pivot Data & Commission structure

- **Pivot Data** = pivot of Data by **DPID × Mapping Type**, summing every measure. Commission
  pulls from it via `GETPIVOTDATA`. ⚠ **Excel `GETPIVOTDATA` matches item names
  case-insensitively** — "Acquisition" matches Data's "ACQUISITION". (Our code case-folds keys.)
- **Commission** = **one row per DPID (169 rows)**. **Label row = row 3**; data starts **row 4**.
  Key cols: `A=DP ID, B=DPMS, C=Mitra Name, D=Region Split`.
- Each column is one of: **pivot-driven** (read from pivot), **pasted KPI** (from intake / prev
  pivot), or **computed formula**. Full column→(measure, mapping type) map and the computed-column
  formulas are encoded in `par_pipeline.py` (`COMMISSION_PIVOT_MAP`, `compute_commission_derived`).

### Computed-column formulas (verified against golden)
- `P = O*L`; `AK = AJ*AG`; `AY = AX*AU`.
- **Acquisition:** `Z` = Total Score (read directly, see §6). `AA` (% payout) from Z via scheme
  (below). `AB` status text. `AD = AA-1`; `AE = R*AD`.
- **Allocation:** `AR = IF(AP>=0.95 AND AQ>=0.99, 1, 0.95)`; `AS = (AR-1)*AM`.
- **Unreg:** `BF = IF(BD>=0.87 AND BE>=0.90, 1, 0.95)`; `BG = (BF-1)*BA`.
- **Regular/COD %:** `BJ`/`BQ` = volume-tier % on `BH+BO` (0 if that channel has no parcels).
- **Cross Border %:** `BW = IF(BU>0, 0.30, 0)`.
- **Drop-off:** `CF = SUM(I,M,S,AH,AN,AV,BB,BM,BS,BY)`; `CG = SUM(J,N,T,AI,AO,AW,BC,BN,BT,BZ)`;
  `CE` = drop-off tier rate by region group (Jabodetabek vs Luar) and `CD` (LI-POH), 0 if CF=0;
  `CH = CG - SUM(S,AN,BB,BM,BS,BY)*CE - SUM(J,N,AI,AW)`.

### Commission Scheme constants (from the 'Commission Scheme' sheet)
- **Post volume tier %** (on `BH+BO`, MATCH-≤): breaks `[1,501,1001,3001,10001]` →
  `[0.15,0.35,0.40,0.45,0.50]`.
- **Drop-off tier** (MATCH-≤ on LI-POH `[0,0.90,0.95]`): Jabodetabek `[200,225,250]`,
  Luar Jabodetabek `[250,275,300]`.
- **Acquisition payout** (MATCH-≤ on Total Score `[0.01,0.90,0.99]`) → KPI `[-0.05,0,+0.05]`;
  `AA = 1 + KPI` → `0.95 / 1.00 / 1.05`; `AB = Deduct 5% / Commission As Is / Bonus 5%`.
- Others: LEX Hub Pickup Rp 500/parcel; Cross Border 30%; Last Mile tiers by Final KPI + size.

---

## 6. KPI sources (the pasted Commission columns)

| Commission cols | Source | Mapping |
|---|---|---|
| `U:Y` + `Z` (Acquisition) | **`Acquisition {Mon} {YY}` file** | **Read `Total Score` directly into Z** (see note). Cols: `DP ID, Mitra Name, LI_N0, LI_N1, PU Scan_N0, Update_RAR, PoPA, Total Score, % Payout Commission (Ori), Status (Ori), Tag`. Values may be `%` strings. 428 duplicate DP IDs in the June file → dedup keep-first. |
| `AP:AQ` (Allocation) | **Allocation KPI Compile** | `DP ID, Mitra Name, Type, N0_Attempt, RoT_Prior_B2B` → `AP=RoT_Prior_B2B, AQ=N0_Attempt`. Join on (DP ID + Type POH/non-POH). |
| `BD:BE` (Unreg) | **Unreg KPI Compile** | `DP ID, Mitra Name, Type, LI_N0, PU Scan_N0, LI_POH` → `BD=LI_N0`, `BE = LI_POH if POH else PU Scan_N0`. |
| `CC:CD` (drop-off) | first-stage intake `2. LIPOH DropOff-SH` | `CC=Clean Up Region, CD=LI-POH`. |
| `CM:CV` (last mile) | first-stage intake `6. Last Mile Tracker` | Attempt Rate N0/N1, Terminal Day 6, Prior Attempt N0, RTS Rate, Valid PODA/POD, Final Score, Final Tier. |
| `G,O,AJ,AX` (claim KPI %) | **previous month's `Pivot All Region.xlsx`** | prev-month value − 100%. Mostly 0. |

**⚠ Acquisition Total Score — important:** the per-component weight *labels* in the Commission
sheet are **stale** (they read PU 20% / RAR 20%; the real file is PU 35% / RAR 5%, and even those
weights do **not** reproduce the file's own Total Score). So **always read `Total Score` directly**,
never recompute. Total Score == golden Z on 100/151 shared rows; the 51 misses are a **fill-down
error in the manual June file** (golden Z stuck at 0.9265 for a block of DPIDs) — the automation
computes them correctly and thus fixes that manual error.

---

## 7. Verification status

- **Commission calc engine — 100% verified.** `verify_commission_engine.py` feeds the engine the
  golden file's own inputs and gets **0 mismatches on all 19 computed columns** (Z, AA, AB, AD, AE,
  P, AK, AR, AS, AY, BF, BG, BJ, BQ, BW, CE, CF, CG, CH).
- **Data — verified** by aggregate reconciliation (see §4).
- **Pivot — verified**: per-(DPID×MappingType) sums reconcile to Data to the rupiah.
- The `Results/` folder is a fresh run, so it is NOT row-exact vs the June deliverable — the oracle
  is the *engine* test (golden inputs→outputs), not row counts.
- **KPI wiring — verified** (`run_local_june.py`): on DPIDs present in both our KPI file and golden,
  Allocation AP = 66/66 exact, Regular/COD % = 168/168; Acquisition U/Z differ on exactly the 51-row
  golden fill-down block (§6) which the automation corrects.
- **Output writer — verified** (`par_writer.py`): produces a valid workbook — label rows 1–3 intact,
  Data from row 2 / Pivot Data from row 5 / Commission from row 4, leftover template rows cleared,
  `AC`(Tag) populated, live PivotTable neutralised, 0 leftover formula cells in the data region.

---

## 8. Gotchas / lessons

1. `GETPIVOTDATA` item matching is **case-insensitive** ("Acquisition" ↔ "ACQUISITION").
2. Commission label row is **row 3**, data starts **row 4**; **169 rows = one per DPID** (not DPMS).
3. Regular & COD Mapping Type comes from `remarks_cod_regular`, not the literal "Registered".
4. Acquisition **Total Score is authoritative**; component weight labels are stale.
5. The manual June file contains a **fill-down error** in Total Score for a block of mitras.
6. `openpyxl` load of the 4 MB workbook is slow (~85 s) — design the writer accordingly.
7. The connected Google Drive in-tool is the user's **personal** account; the Mitra work Drive is
   not reachable via the Drive connector — work from local copies.

---

## 9. Intake template (deliverable) — `Mitra Intake All Region TEMPLATE.xlsx`

Tabs (header row 1, one example row in grey italic): `README`, **Phase 1:** `Acquisition KPI`
(DP ID, Mitra Name, LI_N0, LI_N1, PU Scan_N0, Update_RAR, PoPA, Total Score, % Payout Commission
(Ori), Status (Ori), Tag), `Allocation KPI`, `Unreg KPI`, `Cross Border` (tracking_id, DPID,
DPMS ID, Region Split, delivery_fee). **Phase 2:** `Penalty PPDP`, `Penalty Fake POD-PODA`,
`Penalty Missing LM`, `Adjustment`, `Penalty Outstanding`. **Phase 3:** `Data Mitra Changes`.
Regenerate with `make_intake_template.py`; validate with `validate_all_region_intake.py`.

Region Split, LIPOH (drop-off KPI) and Last Mile Tracker are **not** duplicated here — they come
from the first-stage intake.

### Master intake (single stakeholder-facing workbook) — `Mitra Commission Intake - MASTER TEMPLATE.xlsx`
Built by `automation/make_master_intake.py`. **One** workbook holding ALL tabs (first-stage sources
+ all-region KPI/penalty), grouped and tab-colour-coded by the stakeholder who fills each, with a
README/assignment tab. Upload once to Drive → "Open as Google Sheets" → share one link; each owner
pastes into their tab; monthly, File→Download→.xlsx as `Mitra Commission Intake {Mon-YYYY}.xlsx` into
the Intake folder. The notebook now reads BOTH stages from this one file
(`ALL_REGION_INTAKE_FILE = INTAKE_FILE`). Headers/example rows are copied verbatim from the two prior
templates, so schemas stay exact; tab names are unchanged (pipeline reads them by name). Owner map:
Wira=Acq/Unreg KPI+LIPOH+LM Vol/Tracker; Kamal=Alloc KPI; Davy/Niko=Fake Order; Deshy=Manifest LEX+
Cross Border+Data Mitra Changes; Fanny=Drop SH+Program MPM; Desy=PPDP; Arbeta/Opi=Fake POD-PODA+
Missing LM; Finance=Adjustment+Outstanding; Reference(do-not-edit)=Region Split+SSB.

---

## 10. What's left / next steps

### June end-to-end run — ✅ VALIDATED (2026-07-29)
User ran the full notebook with the seeded June intakes → `M06Y2026 Pivot All Region - Python.xlsx`.
vs golden: **grand-total commission −Rp 44,804 (−0.009%)** (was −347k before the raw-fee + manual-row
fixes). Fully attributed — ACQ↔ALLO reclass net −9,600 (6 DPIDs, first-stage classification), Regular
delivery-fee residual −33,154 (Regular is 0.24% of total), Unreg −2,050 (incl. DP 5739 & 7436 with no
unregistered data rows this run). Adjustment&Penalty E/F/G/H/I + Claims/Cross Border/LEX/Last Mile/MPM
**EXACT**; Regular/COD % 168/168. Data sheet is finer-grained (35,402 vs 16,344 rows) but the same
16,341 unique Shipper×DPID×MappingType keys — aggregates identically. Automation also corrects golden's
manual Acquisition fill-down. **User accepted as validated; no residual-chasing.**

### Phase 1 — ✅ COMPLETE (2026-07)
- **Output writer** ✅ `par_writer.py` — template round-trip, values (decision: template round-trip).
- **KPI wiring** ✅ `build_kpi_row_fn` — Acquisition U..Z + AC(Tag), Allocation AP/AQ (by DPID+POH),
  Unreg BD/BE, drop-off CC/CD, Last Mile CM:CV; all sources optional (tolerant).
- **Notebook port** ✅ "Pivot All Region" section (3 cells) appended after the last Results write,
  reusing the in-memory `agg_*` frames + `df_reg_type`/`df_lipoh`/`df_last_mile_tracker` + config.
- **Verification** ✅ engine 0 mismatches; local harness proves wiring — Allocation AP 66/66 exact
  on shared DPIDs, Regular/COD % 168/168; remaining diffs are the documented golden fill-down block
  (U/Z, 51 rows — automation fixes it) and partial local KPI files (66-DPID Allocation/Unreg).

**Small Phase-1 follow-ups (non-blocking):**
- `prev_claim_kpi` (Commission G/O/AJ/AX = prev-pivot value − 100%) is stubbed `None` in the notebook
  (defaults to 0, "mostly 0" per §6). Wire a prev-month reader when **May's `Pivot All Region.xlsx`**
  is available; that also enables a fully-reproduced June end-to-end run.
- Validate the LM-Tracker CM:CV column-name mapping (`par.LM_TRACKER_COLS`) against a real intake —
  these are display-only KPI (not used in any commission calc), so mismatches only leave blanks.
- Cross-Border rows don't get DPMS/Region from the intake tab (only DPID/fee/30% comm); fine for the
  tiny XB set, revisit if needed.

### Phase 2a — ✅ Adjustment & Penalty COMPLETE (2026-07-29)
- **Module** ✅ `par_adjustment.py` — `build_adjustment_penalty(data, df_adjustment, format_sr_path)`
  returns `{main, outstanding, pivot_dpms, pivot_dpid, pivot_os}`.
- **Sources** (mirror the manual workflow): the `Adjustment` intake tab → 72 adjustment rows
  (E/F/G/J); the **`Format SR Penalty - {Mon}.xlsx`** worksheet the team compiles from several
  external files → penalty rows + outstanding block. Split by the detail sheet's **`Case Pivot`**
  column: `"Penalty Parcel"` → **H** (= Fake POD + Fake PODA + Missing LM combined), `"Penalty
  PPDP"` → **I**. Outstanding block ← Format SR `Penalty Outstanding` sheet.
- **Layout** (verified vs golden M06): main table A:J STACKED (72 adj then 53 penalty, penalty rows
  in `Pivot Parcel` order); B(DPMS)/D(Region) = VLOOKUP into Data by DP ID → `#N/A` if absent;
  outstanding L:Y (V=−Q, W=invoice suffix, X=`-NN` seq, Y=Indonesian month); 3 value-pivots AA:AF
  (by DPMS), AH:AM (by DPID), AO:AR (Outstanding by DPMS).
- **Writer** ✅ `par_writer.write_adjustment_penalty` (+ `write_pivot_all_region(..., adj_result=)`).
- **Notebook** ✅ config cell globs `Format SR Penalty*.xlsx`; build cell reads the `Adjustment` tab,
  builds `adj_result`, and passes it to the writer.
- **Verification** ✅ per-DP H **0 mismatch**, I **0 mismatch**; main totals E/F/H/I exact; 125 rows;
  DP 5707 → `#N/A`; outstanding + helper cols exact; pivot grand totals exact (Denda −31,661.9 &
  Total Penalty −23,621,631.9 match golden row-2). Cosmetic-only: pivots sort DPMS/DPID ascending
  (golden Excel pivot order differs, values identical); outstanding is the 15-row Format SR snapshot
  vs golden's 16 (one DP 6217 April row, −37,929).
- **Intake note:** the `Format SR Penalty` drop-in supersedes the master intake's granular penalty
  tabs (Penalty PPDP / Fake POD-PODA / Missing LM / Penalty Outstanding). Trim them in
  `make_master_intake.py` when convenient; keep `Adjustment`.

### Phase 2a+ — ✅ `Format SR Penalty` compile AUTOMATED (2026-07-29)
The manual worksheet the team assembled from external files is now generated in code.
- **Module** ✅ `par_format_sr.py` — `build_penalty_detail(ppdp_path, podpoda_path, missing_path,
  hub_dp_overrides=…)` → the tracking-level penalty detail (feeds Adjustment&Penalty H/I).
- **Verified rules** (reconciled to golden M06 — per-DP **0 mismatch** on H and I):
  | Case | Source → sheet | Amount rule | June |
  |---|---|---|---|
  | Denda PPDP (→I) | `Penalty PPDP…xlsx` → `RAW` | −`Total Penalty Amount` | 448 / −17,450,412 |
  | Fake PODA (→H) | `…Fake POD & PODA….xlsx` → `Penalty PODA` | −`Penalty` where >0 | 3,725 / −1,984,712.5 |
  | Fake POD (→H) | same → `Penalty POD` | −`Penalty` where >0 | 3 / −332,960 |
  | Missing (→H) | `Penalty LM….xlsx` → sheet 0 | −`Est claim amount` where >0 & `Final Check CL`≠"Take out" | 117 / −11,625,121 |
- Missing DP ← `Investigating Hub Name` via hub→DP map built from the POD/PODA sheets (strips the
  `(Inactive) ` prefix). **Residual:** inactive hubs with no LM penalties aren't in that map and aren't
  derivable from any provided file (Region Split is DP/DPMS/Region only) — resolved via a small
  maintained `hub_dp_overrides` dict (June: `DP-KOI-GDCEK`→50231, `DP-MAC-GDTLN`→2621). Unmapped rows
  are surfaced on `detail.attrs['unmapped_missing']`, never dropped.
- **Not automated (separate inputs):** the Outstanding block (L:Y) source (finance SOA) — still read via
  `build_adjustment_penalty(outstanding_path=…)`; and `fake orders…csv` is a Data/Commission fake-order
  input, not an Adjustment&Penalty case.
- **Design decision (user):** the 4 raw penalty files come from different external stakeholders → keep
  them as **file drop-ins in a `Penalty SR/` folder**, NOT combined into master-intake tabs.
- **Wiring** ✅ `par_adjustment.build_adjustment_penalty` accepts `penalty_detail=` OR `format_sr_path=`
  (+ `outstanding_path=`); notebook globs `Penalty SR/`, compiles, and stamps. Chain: raw files →
  `build_penalty_detail` → `build_adjustment_penalty` → `par_writer`.

### Phase 2a+ (PART C) — ✅ Standardized "Penalty Intake" templates (2026-07-29)
Replaced the arbitrary raw-file drops with a standardized intake, per user decision (delivery = a
**separate shared "Penalty Intake" Google Sheet**, 4 near-raw tabs; the pipeline still filters/maps).
- `par_format_sr.py` refactored to a DataFrame core `_compile_detail(ppdp_df, podpoda_df, missing_df,
  hub_dp_overrides)` fed by **two readers**: `build_penalty_detail(...)` (3 raw files) and
  `build_penalty_detail_from_intake(path)` (the 4-tab workbook). `read_intake_outstanding(path)`
  returns the `Penalty Outstanding` tab. Same verified rules either way (POD/PODA split by a `Type`
  column; POD↔PODA misclassification is harmless — both land in the H "Penalty Parcel" pivot).
- `par_adjustment.build_adjustment_penalty` now also takes `outstanding_df=` (precedence:
  `outstanding_df` > `outstanding_path` > `format_sr_path` outstanding sheet).
- **New `make_penalty_intake_template.py`** → `Penalty Intake TEMPLATE.xlsx` (README + 4 tabs, colour
  by owner). Tabs (do NOT rename — read by `par_format_sr.INTAKE_TAB_*`):
  `Penalty PPDP` (DP ID, Nama Mitra, Tracking ID, Shipper Name, Granular Status, Total Penalty Amount),
  `Penalty Fake POD-PODA` (Type (POD/PODA), Hub Name, DP ID, Nama Mitra, Tracking ID, Granular Status,
  Penalty), `Penalty Missing LM` (Tracking ID, Investigating Hub Name, Order Granular Status, Type,
  Est Claim Amount, Final Check CL), `Penalty Outstanding` (DP ID, DPMS ID, Nama Mitra, Region,
  Invoice, Total Outstanding, Denda, Total Penalty, Periode, Note).
- `make_master_intake.py` — dropped the 4 obsolete penalty tabs (now 18 tabs); README points penalty
  owners to the separate sheet.
- Notebook config/build cells prefer `Penalty Intake*.xlsx` (Intake folder), fall back to `Penalty SR/`
  raw files, then compiled Format SR.
- **Verification:** converted June raw → `Penalty Intake - Jun'26.xlsx` and ran the intake path →
  per-DP H/I **0 mismatch** vs golden (H=−13,942,793.5, I=−17,450,412, 53 rows, outstanding matches).
  Raw-file path still 0-mismatch after the refactor.

### Phase 2b — ✅ DONE (2026-08-21, branch `phase2b-live-pivots`)
**Writer re-architected to a LIVE round-trip** (`par_writer.py`): stop flattening — preserve all
pivots (`Pivot Data`, `Commission!CX`, 3 Adjustment pivots), set every cache `refreshOnLoad=True` +
rewrite its source ref, `fullCalcOnLoad=True`; stamp Commission value-cols only (keep GETPIVOTDATA/
computed formulas, fill-down/clear to match rows). Commission sort = ascending DPID + "By Owner" rows
(Region=="By Owner") at the bottom. `Rekap Commission` / `By Owner per Mitra` = stamp col-A key list
only (formulas live): `build_rekap_keys` (unique DPMS + owner codes) / `build_byowner_keys`. Data
Sources = expanded `Data Mitra Changes` intake (28 cols, 26 match DS headers) upserted by DP ID
(`upsert_data_sources`). Notebook cell `2cebdce7` wired. **Manual step:** open in Excel to refresh.

### Phase 3a — PPh Mitra Ninja 2026 (income-tax sheet) — in progress (2026-08-21)
`par_writer._write_pph(wb, month, active_dpids, byowner_dpids)` — verified vs golden:
- 12-month rolling ledger, keyed by DP ID (col B). **Bruto** (current month, June=col X) = live VLOOKUP
  `Rekap` col 48 (Total Commission) + col 49 (KPI Pickup) − `GETPIVOTDATA` Adjustment Bruto; installed
  only for `active_dpids` (Commission DP set), inactive → 0. Prior month **frozen to values** (read the
  template's cached column, `read_only,data_only`). Calc pointer `AR = $X` repointed.
- **PPh calc `AZ` is live** (`IF(SKB="Ada",0.5%,IF(PT/CV,2%,progressive PPh-21 TER on 50%×gross))`) —
  rates verified. Month PPh cols (`AE:AP`) are the tax team's lagging ledger — **left untouched**.
- New mitras: append only **active, non-by-owner, missing** DP IDs (from written `Data Sources`),
  seeding master A:Q + filling down formulas. By-owner taken from the commission frame (`byowner_dpid_set`)
  since the DS Region col is a formula in formula-view. Month→col via `par.pph_bruto_label`.
- Verified: June-from-golden = 165 active X-formulas, 0 appended (5 active-missing are by-owner);
  May→June (`M05` template) freezes W's 182 formulas to values, installs X. **Excel open-check pending.**

### Next
- **Phase 3b** — Generate + Generate By Owner (transpose selected DPIDs into the per-column PDF layout
  the existing Generate-PDF Python consumes).

**Verification approach for every phase:** diff generated sheet vs `M06Y2026 Pivot All Region.xlsx`
using the engine-isolation method (feed golden inputs, compare computed outputs), with float tolerance.
