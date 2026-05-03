# Atlas Report — Session Handoff
**Last updated:** 2026-05-03 — fixed domain duplicate-FOLIO inflation in Columns B / D / E (deduped after load, keep most recent LAST SALE DATE)

## Context
Building an Atlas Report system for 8020REI, a B2B SaaS real estate data platform.
Working directory: `E:\Claude projects\Test project\Distress Report`

---

## Fulfillment Window Policy

**The fulfillment window is defined entirely by the .xlsx files present in `Fulfillments/`.** The user controls the window by adding or removing files. There is no hidden 6-month filter — whatever is in the folder is the window.

- `WINDOW_START` = first day of the earliest fulfillment file's month (parsed from the `YYYY-MM` filename prefix)
- `WINDOW_END`   = first day of the latest fulfillment file's month
- `SOLD_SINCE`   = `WINDOW_START` — the Column D "sold since" cutoff

**Examples:**
- Earliest file is `2025-11-08 ...xlsx` → `SOLD_SINCE` = 2025-11-01
- Earliest file is `2025-09-08 ...xlsx` → `SOLD_SINCE` = 2025-09-01

**Why a folder-driven window:**
1. **County data validation lag** — county records take time to reflect recent sales, liens, ownership changes; the user excludes too-recent months by not putting them in the folder.
2. **Client marketing cycles** — clients need time to work leads from a fulfillment.
3. **Disposition timelines** — deals take time to close; the window captures active-lead outcomes.
4. **No surprises** — earlier versions enforced a 6-month "intended window" that silently dropped files outside it. The folder is now the single source of truth.

**Current window:** 2025-11 – 2026-01 (3 files in folder)

---

## What Has Been Built

### Step 1 — Fulfillment Compilation
**Script:** `build_domain_report.py` (inline, STEP 1; final write at end of STEP 5)
**Output:** `Fulfillment_Compilation.xlsx`
All fulfillment files in `Fulfillments/` compiled into one dataset.
- Source folder: `Fulfillments/`
- Added columns: `Month`, `Marketing_Channel`, `Source_File`
- Property identifier: `FOLIO`
- **Validation columns appended at end of STEP 5** (used to spot-check the report):
  - `LAST SALE DATE` — most recent sale per FOLIO from domain → validates Column D
  - `PROPERTY STATUS` — `Lead` / `Deal` from deals/leads file (Deal wins if both) → validates F/G
  - `MARKET DEAL` — `Yes`/`No` flag from market-deals overlap → validates Column E
- The save is wrapped with `engine_kwargs={'options': {'strings_to_urls': False, 'tmpdir': str(BASE)}}` because (a) fulfillment data has >65,530 property URLs (Excel's per-sheet hyperlink limit) and (b) the dev machine's `C:` drive is often full, so xlsxwriter's temp files must go on the project drive.

---

### Step 2 — Column C Aggregated Metrics (MAX logic)
**Script:** `build_domain_report.py` (inline, STEP 1 aggregation)

**MAX logic:**
- Likely Deal Score → max across months → range bucket
- Total Score → max across months → range bucket
- Action Plan → min days (30 < 60 < 90 = best)
- Marketing Count → max(total DM across all months, total SMS across all months); 0 contacts treated as 1
- Distress → MAIN DISTRESS #1–4 + binary VACANT + Default Risk cross-referenced from domain by FOLIO

---

### Step 3 — Domain Compilation
**Script:** `compile_domain.py`
**Output:** `Domain Full Data/domain.parquet`
Reads all xlsx parts in `Domain Full Data/` and concatenates into a single parquet. Re-run whenever domain files change.

**Current domain:** `COO config_712.6K_412_part_1.xlsx` + `COO config_712.6K_412_part_2.xlsx`
- 716,089 rows, 709,929 unique FOLIOs
- 254,915 properties with BUYBOX SCORE > 0
- Contains `DEFAULT RISK` column (273,042 FOLIOs flagged)

**Note:** Some older domain exports omit `DEFAULT RISK`. The script handles this gracefully — if the column is missing, Default Risk counts will be 0 and a warning is printed. No crash.

---

### Step 4 — Deals, Leads, and Appointments Integration (Columns F + G)
**Script:** `build_domain_report.py` (inline, STEP 2)
**Source:** Any `.xlsx` in `Documents/` whose filename contains "leads" or "deals" (case-insensitive).

**Status types handled:**
- `Lead` → filtered by `LEAD DATE >= CLIENT_START_DATE (2023-09-26)` → counted as Lead (Column F). **Blank LEAD DATE → included**
- `Appointment` → filtered by `APPOINTMENT DATE >= CLIENT_START_DATE` → reclassified as Lead (Column F). **Blank APPOINTMENT DATE → included**
- `Dead Lead` → filtered by `LEAD DATE >= CLIENT_START_DATE` → reclassified as Lead (Column F). **Blank LEAD DATE → included**
- `Contract` → uses `LEAD DATE` if available → reclassified as Lead (Column F). **Blank LEAD DATE → included**
- `Deal` → filtered by `LAST SALE DATE >= CLIENT_START_DATE` → counted as Deal (Column G). Must have a valid date.

**Data resolution logic:**
1. **Primary:** Join on `FOLIO` to fulfillment → use MAX values → `Data_Source = "Fulfillment"`
2. **Fallback:** Use values directly from file → `Data_Source = "Deals and Leads Atlas"` (excluded from Column F/G — only fulfillment-matched rows go into the report)

---

### Step 5 — Full Report: Columns B, C, D, E, F, G + Formulas H–K (CURRENT MAIN SCRIPT)
**Script:** `build_domain_report.py`
**Output:** `{CLIENT_NAME} - Distress Report - YYYY-MM.xlsx` (dynamic — e.g. `SBD - Distress Report - 2026-04.xlsx`)

All data columns populated:

| Column | Description | Source | Count |
|---|---|---|---|
| B | Total Properties | Domain (BUYBOX SCORE > 0) | 254,915 |
| C | Properties in the fulfillment | Fulfillment MAX | 233,266 |
| D | Sold since Aug 1 2025 | Domain LAST SALE DATE ≥ Aug 1, in fulfillment | 4,887 |
| E | Market Deals | Properties bought at a significant discount and resold for profit in a short period — overlap of Market Deals file ∩ fulfillment-sold (join on PropertyID = BUYBOX ID) | 3,284 |
| F | Clients Lead | Fulfillment-matched leads + appointments (>= 2023-09-26) | 4,690 |
| G | Client Deals | Fulfillment-matched deals (>= 2023-09-26) | 41 |
| H | Sold Concentration % | Formula: =D/C | — |
| I | Sold to Investors Concentration | Formula: =E/C | — |
| J | Client Deals Concentration | Formula: =G/C | — |
| K | Client Leads Concentration | Formula: =F/C | — |

**Domain file:** Loaded from `Domain Full Data/domain.parquet` (cached from two xlsx parts via `compile_domain.py`).
**Default Risk:** Direct column in the domain file (`DEFAULT RISK`); gracefully handled if absent.
**Market Deals:** `Market Deals/SBD Market Deals.xlsx` — join via `PropertyID` → `PROPERTY ID (BUYBOX)`.

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

**Owner Occupied for C/F/G:** Derived by joining fulfillment/leads/deals FOLIOs against domain ABSENTEE column (ABSENTEE == 0).

---

## Marketing Count Logic (ALL columns — consistent)

**Rule:** `max(total DM across all months, total SMS across all months)` per property.

- Example: 7 SMS contacts + 4 DM contacts → count = **7** (not 11)
- Measures strongest channel, not combined total
- Properties with 0 contacts in both channels are treated as 1 (recommended at least once)
- `MKT COUNT` column in fulfillment is empty — not used

**Column B (domain):** Uses `max(DM, SMS)` on the single domain snapshot row per property.
**Columns C/D/E/F/G (fulfillment-based):** Uses `max(sum of all DM rows, sum of all SMS rows)` per FOLIO across all months.

---

## Join Key Logic

| Join | Key | Why |
|---|---|---|
| Fulfillment → domain (Owner Occupied) | FOLIO | Fulfillment primary key is FOLIO |
| Deals/Leads → fulfillment | FOLIO | Deals file also uses FOLIO |
| Default Risk → fulfillment/domain | FOLIO | DEFAULT RISK is now a direct domain column, joined via FOLIO |
| Market Deals → domain (Column E) | PROPERTY ID (BUYBOX) | Market deals file uses PropertyID = BUYBOX ID |

---

## Key Numbers (Latest Run — April 2026)

| Metric | Value |
|---|---|
| Domain properties (BUYBOX > 0) | 254,915 |
| Fulfillment unique FOLIOs | 233,266 |
| Sold since Aug 1 2025 (in fulfillment) | 4,887 |
| Market Deals (overlap with fulfillment-sold) | 3,284 |
| Client Leads (Column F, fulfillment-matched, >= 2023-09-26) | 4,690 |
| Client Deals (Column G, fulfillment-matched, >= 2023-09-26) | 41 |
| Owner Occupied — Column C | 191,905 |
| Owner Occupied — Column F (Leads) | 3,915 |
| Owner Occupied — Column G (Deals) | 30 |
| FOLIOs with Default Risk | 273,042 |

---

## Fixes Applied

### Fix 1 — Owner Occupied fillna(0) bug (Columns B, D, E)
**Location:** `build_domain_report.py` → `domain_distress_counts()`, ABSENTEE block
**Fix:** Removed `fillna(0)` — NaN stays NaN and does not count as ABSENTEE == 0.

### Fix 2 — Owner Occupied for Columns C, F, G (cross-reference via domain)
**Location:** `build_domain_report.py` — block after domain load, before Column B
**Fix:** After loading domain, build a FOLIO→ABSENTEE lookup and patch `C_dist`, `F_dist`, `G_dist` with correct Owner Occupied counts.

### Fix 3 — Marketing Count logic (all columns)
**Fix:** Fulfillment: `max(sum of all DM rows, sum of all SMS rows)` per FOLIO; Domain: `max(DM, SMS)` per row.

### Fix 4 — DEFAULT RISK column missing from some domain exports
**Location:** `build_domain_report.py` — early DEFAULT RISK loading block
**Fix:** Script now checks parquet schema before attempting to load the column. If absent, `dr_folios = set()` and a warning is printed. No crash.

### Fix 5 — Header comment truncation
**Location:** `build_domain_report.py` → `write_excel()`, header_comments block
**Fix:** Comments now explicitly sized at 400×200px so full text is visible on hover.

### Fix 6 — Compilation validation columns + write robustness
**Location:** `build_domain_report.py` — STEP 5b (after market deals)
**Fix:** Compilation save moved from STEP 1 to STEP 5b so it can be enriched with `LAST SALE DATE`, `PROPERTY STATUS`, and `MARKET DEAL` for audit. xlsxwriter is configured with `strings_to_urls=False` (avoids Excel's 65,530 URL/sheet cap and the resulting memory blowup) and `tmpdir=BASE` (keeps spill files off the system C: drive).

### Fix 7 — Domain duplicate FOLIO inflation (Columns B / D / E)
**Location:** `build_domain_report.py` — right after domain load and date coercion (STEP 3)
**Symptom:** Manual filter on `Fulfillment_Compilation.xlsx` (e.g. SCORE 801–1000, LAST SALE DATE >= 2025-11-01) returned 527 unique rows; the report's Column D / 1000-801 bucket showed 561. The 34-row gap was domain rows with duplicate FOLIOs being counted multiple times.
**Root cause:** Columns B, D, and E count rows from the domain DataFrame directly. The domain had 166,782 duplicate FOLIO rows (2,345,875 total → 2,179,093 unique), so any FOLIO appearing N times was counted N times.
**Fix:** Deduplicate `dom` by FOLIO immediately after the date / numeric coercions, sorting by `LAST SALE DATE` ascending with `na_position='first'` and keeping the last row — newer sale info wins, NaT dates are dropped when a real date exists.
**Verified:** Column D / 1000-801 bucket now reads 527, exact match to the manual audit.

---

## Open / Decided Issues

| Issue | Status | Decision |
|---|---|---|
| Owner Occupied bug in B/D/E | ✅ Fixed | Remove fillna(0) |
| Owner Occupied for C/F/G | ✅ Fixed | Cross-reference via domain ABSENTEE join on FOLIO |
| Marketing Count sum vs max | ✅ Fixed | Changed to max(DM, SMS) across all columns |
| DEFAULT RISK missing from new domain exports | ✅ Fixed | Graceful fallback — counts default to 0 with warning |
| Header comment truncation | ✅ Fixed | Comments sized 400×200px |
| Market Deals (Column E) | ✅ Re-enabled | Active — uses `SBD Market Deals.xlsx`, join on PropertyID |
| Market Deals definition | ✅ Updated | Properties bought at a significant discount, resold for profit in a short period |
| Output filename | ✅ Updated | Dynamic: `{CLIENT_NAME} - Distress Report - YYYY-MM.xlsx` |
| Appointments counted as Leads | ✅ Implemented | PROPERTY STATUS == 'Appointment' → reclassified as Lead |
| Dead Leads counted as Leads | ✅ Implemented | PROPERTY STATUS == 'Dead Lead' → reclassified as Lead |
| Contracts counted as Leads | ✅ Implemented | PROPERTY STATUS == 'Contract' → reclassified as Lead |
| Blank lead/appointment dates | ✅ Implemented | Lead-type statuses with no date are included |
| Domain duplicate FOLIOs inflating B/D/E | ✅ Fixed | Dedupe after load, keep row with most recent LAST SALE DATE |

---

## Files Reference

| File | Purpose |
|---|---|
| `build_domain_report.py` | **MAIN SCRIPT** — generates full BCDEFG report |
| `compile_domain.py` | Run whenever domain xlsx files change → rebuilds `domain.parquet` |
| `build_reports.py` | Older — generates Column C only (Max and Average); not used for main output |
| `build_population.py` | Older — compiles fulfillment + generates audit file; not used for main output |
| `build_deals_leads.py` | Older — generates Columns F and G only; not used for main output |
| `Fulfillment_Compilation.xlsx` | Compiled fulfillment dataset, written each run; includes `LAST SALE DATE`, `PROPERTY STATUS`, `MARKET DEAL` validation columns |
| `{CLIENT_NAME} - Distress Report - YYYY-MM.xlsx` | **LATEST REPORT OUTPUT** — e.g. `SBD - Distress Report - 2026-04.xlsx` |
| `Population_For_Calculation.xlsx` | Older audit file |
| `Documents/Deals and Leads SBD.xlsx` | **ACTIVE** — leads, deals, appointments for SBD (auto-discovered) |
| `Domain Full Data/domain.parquet` | Cached full domain; rebuilt by `compile_domain.py` |
| `Domain Full Data/COO config_712.6K_412_part_*.xlsx` | Current domain source files |
| `Market Deals/SBD Market Deals.xlsx` | Investor market deals — join via PropertyID = BUYBOX ID |
| `Fulfillments/` | 6 source fulfillment xlsx files (Aug 2025 – Jan 2026) |
| `Expected Result/Atlas - Expected Result.csv` | Template defining report structure |

---

## Client Info

| Field | Value |
|---|---|
| Client name | SBD / SBDHOUSING |
| Start date with 8020REI | **2023-09-26** |
| Deals/Leads file | `Documents/Deals and Leads SBD.xlsx` (auto-discovered by name pattern) |
| Market Deals file | `Market Deals/SBD Market Deals.xlsx` |

Three variables at the top of `build_domain_report.py` must be updated when switching clients:
- `CLIENT_NAME` — used in the output filename (`{CLIENT_NAME} - Distress Report - YYYY-MM.xlsx`)
- `CLIENT_START_DATE` — date the client started with 8020REI
- `MARKET_PATH` — path to the client's market deals file
