# Atlas Report — Session Handoff Prompt

## Context

Building an Atlas Report system for 8020REI, a B2B SaaS real estate data platform.
Working directory: `C:\Users\danpo\Downloads\Python 8020\Atlas Report New`

---

## What Has Been Built

### Step 1 — Fulfillment Compilation
**Script:** `build_population.py`
**Output:** `Fulfillment_Compilation.xlsx`
All 5 fulfillment files compiled into one dataset (77,000 rows, 37,338 unique properties).
- Source folder: `Fulfillments/`
- Added columns: `Month`, `Marketing_Channel`, `Source_File`
- Property identifier: `FOLIO`

---

### Step 2 — Column C Aggregated Metrics (MAX and AVERAGE)
**Script:** `build_reports.py`
**Outputs:** `Expected_Result_Max.xlsx`, `Expected_Result_Average.xlsx`
Populates only Column C ("Properties in the fulfillment").

**MAX logic:**
- Likely Deal Score → max across months → range bucket
- Total Score → max across months → range bucket
- Action Plan → min days (30 < 60 < 90 = best)
- Marketing Count → sum of DM + SMS from latest month only; 0 contacts treated as 1
- Distress → unique FOLIOs per distress type (MAIN DISTRESS #1–4 + binary VACANT + Default Risk cross-ref)

**AVERAGE logic:** avg score, mode action plan, same mkt/distress.

---

### Step 3 — Population / Audit File
**Script:** `build_population.py`
**Output:** `Population_For_Calculation.xlsx`
One row per unique property (37,338 rows). Every value used in MAX and AVERAGE calculations.

---

### Step 4 — Deals and Leads Integration (Columns F + G)
**Script:** `build_deals_leads.py`
**Source:** `Documents/Deals and Leads Atlas.xlsx` (999 rows: 848 Leads, 151 Deals)
**Outputs:**
- `Expected_Result_Max_FG.xlsx` — Columns C, F, G populated
- `Expected_Result_Max_FG_InFulfillment.xlsx` — same but F/G limited to fulfillment-matched only
- `Documents/Deals and Leads Atlas - Enriched.xlsx` — original + `Data_Source` column
- `Documents/Deals and Leads Atlas - In Fulfillment.xlsx` — 349 fulfillment-matched rows only

**Data resolution logic:**
1. **Primary:** Join on `FOLIO` to `Fulfillment_Compilation.xlsx` → use MAX values
   - 349 matched (345 Leads, 4 Deals) → `Data_Source = "Fulfillment"`
2. **Fallback:** Use values directly from Deals and Leads Atlas file
   - 650 not in fulfillment (503 Leads, 147 Deals) → `Data_Source = "Deals and Leads Atlas"`

---

### Step 5 — Full Report: Columns B, C, D, E, F, G (CURRENT MAIN SCRIPT)
**Script:** `build_domain_report.py`
**Output:** `Expected_Result_Max_BCDFG_InFulfillment.xlsx`

This is the most complete report. All 6 data columns populated:

| Column | Description | Source | Count |
|---|---|---|---|
| B | Total Properties | Domain (BUYBOX SCORE > 0) | ~136,669 |
| C | Properties in the fulfillment | Fulfillment MAX | 37,338 |
| D | Sold since Oct 1 2025 | Domain LAST SALE DATE ≥ Oct 1, in fulfillment | ~390 |
| E | Market Deals | Market deals ∩ fulfillment-sold (BUYBOX ID overlap) | ~105 |
| F | Clients Lead | Fulfillment-matched leads only | 345 |
| G | Client Deals | Fulfillment-matched deals only | 4 |
| H–K | Concentration % | Excel formulas (D/C, E/C, G/C, F/C) | — |

**Domain file:** Loaded from `Domain Full Data/domain.parquet` (cached from two xlsx parts via `compile_domain.py`).
**Default Risk:** Cross-referenced via `Domain Full Data/default_risk.parquet`.
**Market Deals:** `Market Deals/Atlas Market Deals.xlsx` joined on `PropertyID` = BUYBOX ID.

---

## Distress Logic — How Each Column Gets It

| Column | Method |
|---|---|
| B (Total Properties) | `domain_distress_counts()` — reads binary columns from domain; ABSENTEE: 1=in-state, 2=out-of-state, 0=owner occupied |
| C (Fulfillment) | `dist_long` from MAIN DISTRESS #1–4 + binary VACANT + Default Risk (FOLIO cross-ref) |
| D (Sold) | `domain_distress_counts()` on sold subset of domain |
| E (Market Deals) | `domain_distress_counts()` on market-deal-overlap subset |
| F (Client Leads) | `ff_distress_map` — same as Column C (fulfillment distress only) |
| G (Client Deals) | Same as F |

**Critical:** Owner Occupied is derived from `ABSENTEE = 0` in domain data. It is **not** available in the fulfillment distress columns (MAIN DISTRESS #1–4). This means:
- Columns **C, F, G** will always show **Owner Occupied = 0**
- Columns **B, D, E** will show actual Owner Occupied counts from domain ABSENTEE

---

## Open Issues — Fix Next Session

### Issue 1 — Owner Occupied: fillna(0) bug in Columns B, D, E
**Location:** `build_domain_report.py` → `domain_distress_counts()` function, line ~111
**Problem:**
```python
abs_vals = pd.to_numeric(df_sub['ABSENTEE'], errors='coerce').fillna(0)
result['Owner Occupied'] = int((abs_vals == 0).sum())
```
`fillna(0)` treats properties with **null/missing ABSENTEE** as Owner Occupied. This inflates the count.
Current Column E shows **72** for Owner Occupied — some or all may be null ABSENTEE, not genuine owner-occupied.
**Fix:** Change `fillna(0)` to `fillna(np.nan)` and filter only confirmed 0 values:
```python
abs_vals = pd.to_numeric(df_sub['ABSENTEE'], errors='coerce')
result['Owner Occupied'] = int((abs_vals == 0).sum())
```
(Do not fillna — let NaN stay NaN so it doesn't count.)

### Issue 2 — Owner Occupied = 0 for Columns C, F, G (by design — decide if acceptable)
**Problem:** Fulfillment distress data (MAIN DISTRESS #1–4) does not carry an Owner Occupied signal. Client Leads (F) and Client Deals (G) always show 0 for Owner Occupied.
**Decision needed:** Should we cross-reference fulfillment FOLIOs + deals/leads FOLIOs against the domain's ABSENTEE column to populate Owner Occupied for C, F, and G? This would require joining `ff_agg.index` against the domain parquet.

### Issue 3 — Action Plan nearly empty for Deals (only 4 of 151)
**Root cause:** ALL 151 Deal records in `Deals and Leads Atlas.xlsx` have `ACTION PLANS = NaN`. Only the 4 deals matched to fulfillment have action plan data.
**Decision needed:** Leave blanks for fallback deals, or use another Atlas column as proxy?

### Issue 4 — Marketing Count bins include 0 (legacy — may already be fixed in build_domain_report.py)
**Status:** `build_domain_report.py` uses `mkt_bins = [0.5, 5, 10, 19, 1e9]` (lower bound 0.5, so 0-contact properties fall below all bins and are excluded). Verify `build_deals_leads.py` and `build_reports.py` also use 0.5, not -0.5.

---

## Files Reference

| File | Purpose |
|---|---|
| `build_domain_report.py` | **MAIN SCRIPT** — generates full BCDEFG report |
| `build_reports.py` | Generates Column C only (Max and Average) |
| `build_population.py` | Compiles fulfillment files + generates audit file |
| `build_deals_leads.py` | Generates Columns F and G + enriched Atlas files |
| `compile_domain.py` | One-time: reads two domain xlsx parts → saves parquet |
| `domain_check.py` | Diagnostic: inspects domain columns/structure |
| `market_deals_check.py/.2/.3/.4` | Diagnostics: match market deals to domain/fulfillment |
| `Fulfillment_Compilation.xlsx` | Master compiled dataset (77k rows) |
| `Expected_Result_Max_BCDFG_InFulfillment.xlsx` | **LATEST REPORT OUTPUT** |
| `Expected_Result_Max_FG.xlsx` | Older report — Columns C, F, G only (all leads/deals) |
| `Expected_Result_Max_FG_InFulfillment.xlsx` | Older report — Columns C, F, G (fulfillment-matched only) |
| `Expected_Result_Max.xlsx` | Older report — Column C only (MAX logic) |
| `Expected_Result_Average.xlsx` | Older report — Column C only (AVERAGE logic) |
| `Population_For_Calculation.xlsx` | Property-level audit file (37,338 rows) |
| `Documents/Deals and Leads Atlas.xlsx` | Client-reported leads and deals (999 rows) |
| `Documents/Deals and Leads Atlas - Enriched.xlsx` | Same + Data_Source column |
| `Documents/Deals and Leads Atlas - In Fulfillment.xlsx` | Fulfillment-matched subset (349 rows) |
| `Domain Full Data/domain.parquet` | Cached full domain (536k rows) |
| `Domain Full Data/default_risk.parquet` | Default Risk properties |
| `Market Deals/Atlas Market Deals.xlsx` | Investor market deals (joined via PropertyID = BUYBOX ID) |
| `Fulfillments/` | 5 source fulfillment xlsx files (Oct–Dec 2025) |
| `Expected Result/Atlas - Expected Result.csv` | Template defining report structure |

---

## Key Numbers (Latest Run of build_domain_report.py)

| Metric | Value |
|---|---|
| Domain properties (BUYBOX > 0) | ~136,669 |
| Fulfillment unique FOLIOs | 37,338 |
| Sold since Oct 1 (in fulfillment) | ~390 |
| Market Deals (overlap with fulfillment-sold) | ~105 |
| Client Leads (fulfillment-matched) | 345 |
| Client Deals (fulfillment-matched) | 4 |
| Owner Occupied in Column E | 72 (may be inflated — see Issue 1) |
| Owner Occupied in Columns C, F, G | 0 (by design — see Issue 2) |

---

## Pending / Not Yet Started

- Fix Issue 1 (fillna bug for Owner Occupied in B/D/E)
- Decide Issue 2 (populate Owner Occupied for C/F/G via domain cross-ref)
- Decide Issue 3 (Action Plan proxy for fallback Deals)
- Verify Issue 4 (mkt_bins lower bound in older scripts)
- Column D "Sold" — confirm final total with user (currently ~390 fulfillment-recommended sold properties)
- Column H–K formulas — confirm denominator logic with user (currently all use Column C as base)
