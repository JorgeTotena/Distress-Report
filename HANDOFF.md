# Atlas Report — Session Handoff
**Last updated:** 2026-03-25

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
- 3 months: Oct 2025 (31k rows), Nov 2025 (16k rows), Dec 2025 (30k rows)
- 2 channels: SMS (52k rows), DM (25k rows)
- Every row is a unique (FOLIO × Month × Channel) combination — no raw duplicates

---

### Step 2 — Column C Aggregated Metrics (MAX and AVERAGE)
**Script:** `build_reports.py`
**Outputs:** `Expected_Result_Max.xlsx`, `Expected_Result_Average.xlsx`
Populates only Column C ("Properties in the fulfillment").

**MAX logic:**
- Likely Deal Score → max across months → range bucket
- Total Score → max across months → range bucket
- Action Plan → min days (30 < 60 < 90 = best)
- Marketing Count → max(total DM across all months, total SMS across all months); 0 contacts treated as 1
- Distress → unique FOLIOs per distress type (MAIN DISTRESS #1–4 + binary VACANT + Default Risk cross-ref)

**AVERAGE logic:** avg score, mode action plan, same mkt/distress.

---

### Step 3 — Population / Audit File
**Script:** `build_population.py`
**Output:** `Population_For_Calculation.xlsx`
One row per unique property (37,338 rows). Every value used in MAX and AVERAGE calculations.

---

### Step 4 — Deals, Leads, and Appointments Integration (Columns E + F)
**Script:** `build_domain_report.py` (inline, STEP 2)
**Source:** `Documents/Deals and Leads Freedom.xlsx`

**Status types handled:**
- `Lead` → filtered by `LEAD DATE >= CLIENT_START_DATE (2021-10-26)` → counted as Lead (Column E)
- `Appointment` → filtered by `APPOINTMENT DATE >= CLIENT_START_DATE` → reclassified as Lead and counted in Column E
- `Deal` → filtered by `LAST SALE DATE >= CLIENT_START_DATE` → counted as Deal (Column F)
- `Contract` → **excluded** (no date column available)

**Data resolution logic:**
1. **Primary:** Join on `FOLIO` to fulfillment → use MAX values → `Data_Source = "Fulfillment"`
2. **Fallback:** Use values directly from Freedom file → `Data_Source = "Deals and Leads Atlas"` (excluded from Column E/F — only fulfillment-matched rows go into the report)

**Note:** `build_deals_leads.py` is an older standalone script that still references `Deals and Leads Atlas.xlsx`. It is not used for the current main output.

---

### Step 5 — Full Report: Columns B, C, D, E, F + Formulas G–I (CURRENT MAIN SCRIPT)
**Script:** `build_domain_report.py`
**Output:** `Expected_Result_Max_BCDFG_InFulfillment.xlsx`

All data columns populated:

| Column | Description | Source | Count |
|---|---|---|---|
| B | Total Properties | Domain (BUYBOX SCORE > 0) | ~136,669 |
| C | Properties in the fulfillment | Fulfillment MAX | 37,338 |
| D | Sold since Oct 1 2025 | Domain LAST SALE DATE ≥ Oct 1, in fulfillment | ~390 |
| E | Clients Lead | Fulfillment-matched leads + appointments (>= 2021-10-26) | 345+ |
| F | Client Deals | Fulfillment-matched deals (>= 2021-10-26) | 4+ |
| G | Sold Concentration % | Formula: =D/C | — |
| H | Client Deals Concentration | Formula: =F/C | — |
| I | Client Leads Concentration | Formula: =E/C | — |
| [Market Deals] | *Disabled* — column removed from COLS; code in STEP 5 is commented out | — | — |

**Domain file:** Loaded from `Domain Full Data/domain.parquet` (cached from two xlsx parts via `compile_domain.py`).
**Default Risk:** Now a direct column in the domain file (`DEFAULT RISK`); no separate parquet needed.
**Market Deals:** `Market Deals/Atlas Market Deals.xlsx` — data and logic intact, just disabled. See re-enable instructions below.

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

**Owner Occupied for C/F/G:** Derived by joining fulfillment/leads/deals FOLIOs against domain ABSENTEE column (ABSENTEE == 0). FOLIO match rate is 100% for fulfillment, 99.9% for leads, 97.5% for deals — no coverage concern.

---

## Marketing Count Logic (ALL columns — consistent)

**Rule:** `max(total DM across all months, total SMS across all months)` per property.

- Example: 7 SMS contacts + 4 DM contacts → count = **7** (not 11)
- Measures strongest channel, not combined total
- Properties with 0 contacts in both channels are treated as 1 (recommended at least once)
- `MKT COUNT` column in fulfillment is empty (76,518 nulls) — not used
- No CC or other channel columns exist in the data

**Column B (domain):** Uses `max(DM, SMS)` on the single domain snapshot row per property.
**Columns C/D/E/F/G (fulfillment-based):** Uses `max(sum of all DM rows, sum of all SMS rows)` per FOLIO across all months.

---

## Join Key Logic

| Join | Key | Why |
|---|---|---|
| Fulfillment → domain (Owner Occupied) | FOLIO | Fulfillment primary key is FOLIO |
| Deals/Leads Atlas → fulfillment | FOLIO | Atlas file also uses FOLIO |
| Default Risk → fulfillment/domain | PROPERTY ID (BUYBOX) | Default risk parquet has no FOLIO |
| Market Deals → domain (Column E) | PROPERTY ID (BUYBOX) | Market deals file has no FOLIO |

---

## Key Numbers (Latest Run)

| Metric | Value |
|---|---|
| Domain properties (BUYBOX > 0) | 136,669 |
| Fulfillment unique FOLIOs | 37,338 |
| Sold since Oct 1 (in fulfillment) | 390 |
| Market Deals (overlap with fulfillment-sold) | 105 (disabled) |
| Client Leads incl. Appointments (Column E, fulfillment-matched, >= 2021-10-26) | 345+ (rerun to confirm) |
| Client Deals (Column F, fulfillment-matched, >= 2021-10-26) | 4+ (rerun to confirm) |
| Owner Occupied — Column C | 25,300 |
| Owner Occupied — Column E (Leads) | 175 |
| Owner Occupied — Column F (Deals) | 3 |

---

## Fixes Applied This Session

### Fix 1 — Owner Occupied fillna(0) bug (Columns B, D, E)
**Location:** `build_domain_report.py` → `domain_distress_counts()`, ABSENTEE block
**Problem:** `fillna(0)` was treating properties with null ABSENTEE as Owner Occupied, inflating the count.
**Fix:** Removed `fillna(0)` — NaN stays NaN and does not count as ABSENTEE == 0.

### Fix 2 — Owner Occupied for Columns C, F, G (cross-reference via domain)
**Location:** `build_domain_report.py` — new block after domain load, before Column B
**Problem:** Fulfillment distress data (MAIN DISTRESS #1–4) has no Owner Occupied signal; C/F/G always showed 0.
**Fix:** After loading domain, build a FOLIO→ABSENTEE lookup and patch `C_dist`, `F_dist`, `G_dist` with correct Owner Occupied counts.

### Fix 3 — Marketing Count logic (all columns)
**Location:** `build_domain_report.py` — ff_mkt computation and dom['total_mkt']
**Problem:** Old logic summed DM + SMS from latest month only for fulfillment; summed DM + SMS for domain. Both overcounted when a property was contacted via multiple channels.
**Fix:**
- Fulfillment: `max(sum of all DM rows, sum of all SMS rows)` per FOLIO across all months
- Domain: `max(DM, SMS)` per property row

### Fix 4 — mkt_bins lower bound (verified, no change needed)
All three scripts (`build_domain_report.py`, `build_deals_leads.py`, `build_reports.py`) already use `[0.5, 5, 10, 19, 1e9]` — properties with 0 contacts fall below all bins and are excluded correctly (then clipped to 1 for fulfillment-based columns).

---

## Open / Decided Issues

| Issue | Status | Decision |
|---|---|---|
| Owner Occupied bug in B/D/E | ✅ Fixed | Remove fillna(0) |
| Owner Occupied for C/F/G | ✅ Fixed | Cross-reference via domain ABSENTEE join on FOLIO |
| Action Plan for fallback Deals (147 of 151) | ✅ No action | Column G only includes 4 FF-matched deals; 147 fallbacks excluded. All 151 deals have ACTION PLANS = NaN anyway. |
| Marketing Count bins lower bound | ✅ Verified | Already correct in all scripts |
| Marketing Count sum vs max | ✅ Fixed | Changed to max(DM, SMS) across all columns |
| Market Deals (Column E) | ⏸ Disabled | Removed from COLS; logic commented out in STEP 5 — see re-enable instructions below |
| Appointments counted as Leads | ✅ Implemented | PROPERTY STATUS == 'Appointment' filtered by APPOINTMENT DATE, then reclassified as Lead before Column E count |
| Date filter for leads/deals | ✅ Implemented | CLIENT_START_DATE = 2021-10-26; Lead → LEAD DATE, Appointment → APPOINTMENT DATE, Deal → LAST SALE DATE; Contracts excluded |
| Column letters shifted | ✅ Updated | Market Deals removed from COLS → E=Clients Lead, F=Client Deals, G/H/I=formulas (3 formulas, not 4) |
| Data source file | ✅ Updated | Now reads `Deals and Leads Freedom.xlsx` (not Atlas.xlsx) |

---

## Market Deals Column — Re-enable Instructions

Column E (Market Deals) and Column I (Sold to Investors Concentration) are currently disabled in `build_domain_report.py`. The data file and all logic are intact. To re-enable:

**1. Uncomment STEP 5** (~line 340) — the full block that loads `MARKET_PATH`, computes the overlap, and produces `E_likely`, `E_score`, `E_action`, `E_mkt`, `E_dist`.

**2. Restore `'Market Deals'` and `'Sold to investors concentration'` in `COLS`** (~line 70):
```python
COLS = [
    'Category', 'Total Properties', 'Properties in the fulfillment',
    'Sold', 'Market Deals', 'Clients Lead', 'Client Deals',
    'Sold Concentration %Sold', 'Sold to investors concentration',
    'Client Deals Concentration', 'Client Leads Concentration',
]
```

**3. Restore `build_rows` signature** — uncomment the E params line:
```python
E_lik, E_sco, E_act, E_mkt, E_dis,
```

**4. Restore `data_row` inner function** — uncomment `r[4] = e` and shift F/G back to indices 5 and 6.

**5. Restore all `data_row()` calls** inside `build_rows` — add `E_*[lbl]` / `sum(E_*.values())` back as the 5th argument in every call.

**6. Restore `COUNT_COLS`** to `[1, 2, 3, 4, 5, 6]` and formula indices to `[7, 8, 9, 10]` with:
```python
ws[r][7].value  = f'=IFERROR(D{r}/C{r},"")'   # H: Sold Conc
ws[r][8].value  = f'=IFERROR(E{r}/C{r},"")'   # I: Market Deals Conc
ws[r][9].value  = f'=IFERROR(G{r}/C{r},"")'   # J: Client Deals Conc
ws[r][10].value = f'=IFERROR(F{r}/C{r},"")'   # K: Client Leads Conc
```

**7. Restore header comments** for `'E'` (Market Deals) and `'I'` (Sold to Investors), and shift `'F'`→`'F'`, `'G'`→`'G'`, etc. back to their original letters.

**8. Restore column widths** to include `'E'` through `'K'`.

**9. Restore the `build_rows()` call** at the bottom — uncomment the `E_likely, E_score, E_action, E_mkt, E_dist,` line.

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
| `Documents/Deals and Leads Freedom.xlsx` | **ACTIVE** — leads, deals, appointments for Freedom REI (used by main script) |
| `Documents/Deals and Leads Atlas.xlsx` | Older file — used by `build_deals_leads.py` only, not the main script |
| `Documents/Deals and Leads Atlas - Enriched.xlsx` | Same + Data_Source column (older, from build_deals_leads.py) |
| `Documents/Deals and Leads Atlas - In Fulfillment.xlsx` | Fulfillment-matched subset (349 rows) (older) |
| `Domain Full Data/domain.parquet` | Cached full domain (536,941 rows); DEFAULT RISK now a column here |
| `Domain Full Data/default_risk.parquet` | Legacy — no longer used; DEFAULT RISK is now in domain.parquet |
| `Market Deals/Atlas Market Deals.xlsx` | Investor market deals (joined via PropertyID = BUYBOX ID) |
| `Fulfillments/` | 5 source fulfillment xlsx files (Oct–Dec 2025) |
| `Expected Result/Atlas - Expected Result.csv` | Template defining report structure |

---

## Client Info

| Field | Value |
|---|---|
| Client name | Freedom REI |
| Start date with 8020REI | **2021-10-26** |
| Deals/Leads file | `Documents/Deals and Leads Freedom.xlsx` |

The `CLIENT_START_DATE` is hardcoded in `build_domain_report.py`. When running the report for a different client, update that value at the top of the script.

---

## Pending / Not Yet Started

- Column D "Sold" — confirm final total with user (currently 390 fulfillment-recommended sold properties)
- Column G–I formulas — confirm denominator logic with user (currently all use Column C as base)
- Rerun `build_domain_report.py` after appointment/lead changes to confirm updated Column E count
- `build_reports.py` and `build_deals_leads.py` — marketing count logic and appointment handling **not synced** to these older scripts; they are not used for the current main output but may need updating if reused
