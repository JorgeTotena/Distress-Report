# Atlas Report — Session Handoff
**Last updated:** 2026-06-17 — Leverage Companies; Columns F/G now count unique property (FOLIO) + exact-row dedup of the leads file; COUNTY restored to domain (companion fix); LEAD DATE / APPOINTMENT DATE added to compilation; ZIP section removed; parquet caching for fulfillments + domain column filtering; CSV domain source support

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

**Current window:** 2025-01 – 2026-03 (11 months, 31 files in folder)

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
  - `LEAD DATE` — earliest (creation) lead date per FOLIO from the window-filtered leads → lets Power BI reproduce the F window date filter
  - `APPOINTMENT DATE` — earliest appointment date per FOLIO from the window-filtered leads
  - `PROPERTY STATUS` — `Lead` / `Deal` from deals/leads file (Deal wins if both) → validates F/G
  - `MARKET DEAL` — `Yes`/`No` flag from market-deals overlap → validates Column E
  - `CLIENT LEAD` / `CLIENT DEAL` — `Yes`/`No`; unique-FOLIO membership of Columns F / G (these now equal the report's F/G totals — see Fix 9)
  - **Caveat:** leads with a blank `LEAD DATE` are kept (can't be ruled out of the window), so a FOLIO can show `CLIENT LEAD=Yes` with a blank `LEAD DATE`. Dropping blank-dated leads in Power BI undercounts vs. the report.
- The save is wrapped with `engine_kwargs={'options': {'strings_to_urls': False, 'tmpdir': str(BASE)}}` because (a) fulfillment data has >65,530 property URLs (Excel's per-sheet hyperlink limit) and (b) the dev machine's `C:` drive is often full, so xlsxwriter's temp files must go on the project drive.

---

### Step 2 — Column C Aggregated Metrics (MAX logic)
**Script:** `build_domain_report.py` (inline, STEP 1 aggregation)

**MAX logic:**
- Likely Deal Score → max across months → range bucket
- Total Score → max across months → range bucket
- Action Plan → min days (30 < 60 < 90 = best)
- Marketing Count → max(total DM across all months, total SMS across all months); 0 contacts treated as 1
- Distress → fulfillment binary distress columns per FOLIO (PRE-FORECLOSURE > 0, VACANT > 0, TAXES > 0, …) + Default Risk cross-referenced from domain by FOLIO

---

### Step 3 — Domain Compilation
**Script:** `compile_domain.py`
**Output:** `Domain Full Data/domain.parquet`
Reads all xlsx **or csv** parts in `Domain Full Data/` and saves a trimmed parquet (only the ~28 columns in `_DOM_COLS_NEEDED`). Re-run whenever domain files change. **`COUNTY` is included** — the companion's county-level Distress Universe page needs it (omitting it silently breaks the companion; see Fix 10). If you add a column to `_DOM_COLS_NEEDED`, delete `domain.parquet` so it rebuilds from the CSVs.

**Current domain:** 4 CSV parts — `COO config_1.6M_386_part_1.csv` through `_part_4.csv`
- ~2,581,394 rows, ~1,444,626 unique FOLIOs (1,444,627 after dedupe)
- 348,735 properties with BUYBOX SCORE > 0
- Contains `DEFAULT RISK` column (681,352 FOLIOs flagged)

**Column filtering:** Both `compile_domain.py` and `build_domain_report.py` share `_DOM_COLS_NEEDED` — only those ~28 columns are loaded/cached. Add any new field to both files and re-run `compile_domain.py`.

**Note:** Some older domain exports omit `DEFAULT RISK`. The script handles this gracefully — if the column is missing, Default Risk counts will be 0 and a warning is printed. No crash.

---

### Step 4 — Deals, Leads, and Appointments Integration (Columns F + G)
**Script:** `build_domain_report.py` (inline, STEP 2)
**Source:** Any `.xlsx` in `Documents/` whose filename contains "leads" or "deals" (case-insensitive).

**Exact-row dedup on load (Fix 9a):** immediately after loading the file, fully-identical duplicate rows are dropped (`dl.drop_duplicates()`). Some client exports repeat the same lead/deal row 2–4× (Leverage had ~9.6k); the report counts rows, so these literal repeats would inflate F/G. Clean files are unaffected.

**Date filter — from `WINDOW_START`, not `CLIENT_START_DATE`:** leads/deals are filtered to the **start of the fulfillment window** (`LEADS_DEALS_SINCE = WINDOW_START`, e.g. 2025-01-01), so F/G align with Column D. `CLIENT_START_DATE` is informational only.

**Status types handled:**
- `Lead` → filtered by `LEAD DATE >= WINDOW_START` → counted as Lead (Column F). **Blank LEAD DATE → included**
- `Appointment` → filtered by `APPOINTMENT DATE >= WINDOW_START` → reclassified as Lead (Column F). **Blank APPOINTMENT DATE → included**
- `Dead Lead` → filtered by `LEAD DATE >= WINDOW_START` → reclassified as Lead (Column F). **Blank LEAD DATE → included**
- `Contract` → uses `LEAD DATE` if available → reclassified as Lead (Column F). **Blank LEAD DATE → included**
- `Deal` → filtered by `LAST SALE DATE >= WINDOW_START` → counted as Deal (Column G). Must have a valid date.

**Data resolution logic:**
1. **Primary:** Join on `FOLIO` to fulfillment → use MAX values → `Data_Source = "Fulfillment"`
2. **Fallback:** Use values directly from file → `Data_Source = "Deals and Leads Atlas"` (excluded from Column F/G — only fulfillment-matched rows go into the report)

**Counting grain — unique property (Fix 9b):** `leads_ff` / `deals_ff` are deduped by `FOLIO_key` before counting, so Columns F/G count **unique properties**, consistent with B/C/D/E (and making the `F/C` / `G/C` concentration ratios valid). In the Fulfillment source every row of a FOLIO resolves to identical values, so the dedup is lossless and propagates to every F/G sub-count, **including the Distress section**.

---

### Step 5 — Full Report: Columns B, C, D, E, F, G + Formulas H–K (CURRENT MAIN SCRIPT)
**Script:** `build_domain_report.py`
**Output:** `{CLIENT_NAME} - Distress Report - YYYY-MM.xlsx` (dynamic — e.g. `Leverage Companies - Distress Report - 2026-06.xlsx`)

**Sections in the Excel:** Likely Deal Score, Total Score, Action Plan, Mkt Count, Distress. ZIP Code section removed at client request.

All data columns populated:

| Column | Description | Source | Count (Jun 2026 run) |
|---|---|---|---|
| B | Total Properties | Domain (BUYBOX SCORE > 0) | 348,735 |
| C | Properties in the fulfillment | Fulfillment MAX | 434,462 |
| D | Sold since Jan 1 2025 | Domain LAST SALE DATE ≥ Jan 1 2025, in fulfillment | 20,044 |
| E | Market Deals | Properties bought at a significant discount and resold for profit in a short period — overlap of Market Deals file ∩ fulfillment-sold (join on PropertyID = BUYBOX ID) | 938 |
| F | Clients Lead | Fulfillment-matched leads + appointments (>= 2025-01-01), **unique property (FOLIO)** | 22,282 |
| G | Client Deals | Fulfillment-matched deals (>= 2025-01-01), **unique property (FOLIO)** | 48 |
| H | Sold Concentration % | Formula: =D/C | — |
| I | Sold to Investors Concentration | Formula: =E/C | — |
| J | Client Deals Concentration | Formula: =G/C | — |
| K | Client Leads Concentration | Formula: =F/C | — |

**Domain file:** Loaded from `Domain Full Data/domain.parquet` (cached from 4 CSV parts via `compile_domain.py`).
**Default Risk:** Direct column in the domain file (`DEFAULT RISK`); gracefully handled if absent.
**Market Deals:** `Market Deals/Market Deals Leverage.xlsx` — this file has no `PropertyID`, so the script falls back to a normalized `ADDRESS + ZIP` join (see Column E logic). Files that do carry `PropertyID` join via `PropertyID` → `PROPERTY ID (BUYBOX)`.

---

## Distress Logic — How Each Column Gets It

**Source of truth for binary distresses (Pre-foreclosure, Vacant, Taxes, Estate, Inter family transfer, Probate, Judgement, Liens, High Equity, 55+, Downsizing) for fulfillment-related columns = the fulfillment file's binary columns directly** (e.g. `PRE-FORECLOSURE > 0` in any month → flagged). Default Risk comes from the domain DR cross-ref. ABSENTEE / Owner Occupied stay domain-derived (3-way classification).

| Column | Method |
|---|---|
| B (Total Properties) | `domain_distress_counts()` — reads binary columns from domain; ABSENTEE: 1=in-state, 2=out-of-state, 0=owner occupied |
| C (Fulfillment) | `dist_long` built from fulfillment binary columns + Default Risk (FOLIO cross-ref) |
| D (Sold) | `fulfillment_distress_counts(sold, dist_long)` — same source as C, restricted to sold ∩ fulfillment FOLIOs |
| E (Market Deals) | `fulfillment_distress_counts(mkt_sold, dist_long)` — same source as C, restricted to market-deal-overlap |
| F (Client Leads) | `ff_distress_map` — same source as C (fulfillment binary columns); counted on the FOLIO-deduped `leads_ff`, so per-property |
| G (Client Deals) | Same as F (per-property, on FOLIO-deduped `deals_ff`) |

**Owner Occupied for C/F/G:** Derived by joining fulfillment/leads/deals FOLIOs against domain ABSENTEE column (ABSENTEE == 0).

**Why fulfillment binary, not MAIN DISTRESS #1-4 ranking:** the ranking only carries the top 4 distresses per property. Pre-foreclosure (and similarly-mid-ranked flags) routinely fell off when properties had >4 signals, undercounting by 70x in some buckets. Reading the binary columns picks up every flagged FOLIO. See "Fix 8" below.

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

## Key Numbers (Latest Run — June 2026, Leverage Companies)

| Metric | Value |
|---|---|
| Domain properties (BUYBOX > 0) | 348,735 |
| Fulfillment unique FOLIOs | 434,462 |
| Sold since Jan 1 2025 (in fulfillment) | 20,044 |
| Market Deals (overlap with fulfillment-sold) | 938 |
| Client Leads (Column F, fulfillment-matched, >= 2025-01-01) | 22,282 |
| Client Deals (Column G, fulfillment-matched, >= 2025-01-01) | 48 |
| Owner Occupied — Column C | 306,214 |
| Owner Occupied — Column F (Leads) | 15,406 |
| Owner Occupied — Column G (Deals) | 30 |
| FOLIOs with Default Risk | 681,352 |
| Counties with matched sales | 8 |

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

### Fix 8 — Distress source = fulfillment binary columns (Columns C / D / E / F / G)
**Location:** `build_domain_report.py` — STEP 1 `dist_long` build, plus D/E call sites
**Symptom:** Manual filter on `Fulfillment_Compilation.xlsx` for `PRE-FORECLOSURE > 0` AND `LAST SALE DATE >= 2025-11-01` returned 350 rows. Column D / Pre-foreclosure showed **5**. Same kind of undercount across other binary distresses for any column compared against fulfillment.
**Root cause (two-fold):**
- Columns C / F / G derived distress from `MAIN DISTRESS #1–4` (a top-4 ranking, not a flag list). Properties with >4 signals routinely had Pre-foreclosure ranked off, dropping the count.
- Columns D / E read binary distress from the **domain** file (`domain_distress_counts(sold)`), but the domain has fewer flags set than the fulfillment file for the same FOLIOs.
**Fix:**
- `dist_long` is now built from the fulfillment binary columns directly: for each `col` in `DOMAIN_DIST_MAP`, take FOLIOs where `pd.to_numeric(df[col]) > 0` in any month and tag them with the corresponding label. Default Risk still injected from the domain DR cross-ref. ABSENTEE / Owner Occupied still domain-derived.
- New helper `fulfillment_distress_counts(df_sub, dist_long)` counts distress for any subset of FOLIOs against `dist_long`. Used by Columns D and E in place of `domain_distress_counts()`.
- Columns F / G inherit the change automatically (they read `ff_distress_map` which is derived from `dist_long`).
- Column B is intentionally untouched — its rows are mostly *not* in fulfillment, so `domain_distress_counts(dom_b)` is the only viable source.
**Verified:** Pre-foreclosure now reads C=4,487 / D=350 / E=20 / F=18 / G=1, matching manual audits on the fulfillment file.

### Fix 9 — Columns F/G grain: exact-row dedup + count by unique property
**Location:** `build_domain_report.py` — STEP 2 (after `dl` load, and after `leads_ff`/`deals_ff` are built)
**Symptom:** Client's Power BI (deduped by property) showed Col F = 22,282 / Col G = 48; the report showed **26,534 / 59**. B/C/D/E/Market-Deals all reconciled — only F/G were off.
**Root cause (two-fold):**
- The Leverage `Deals and Leads` export contained **9,617 fully-identical duplicate rows** (the same lead repeated 2–4×). The report counts rows, so the dupes passed straight through.
- Even without dupes, F/G counted lead/deal **events** (rows), while B/C/D/E count **unique properties** — so a property with several legitimate lead events was counted several times. This also broke the `F/C` and `G/C` concentration ratios (event ÷ property). The behavior was latent in every client; it only surfaced for Leverage (dirty export + a property-level audit). Earlier clients reconciled because their files had ~one row per lead.
**Fix:**
- **9a** — drop fully-identical rows on load: `dl = dl.drop_duplicates()`.
- **9b** — dedupe `leads_ff` / `deals_ff` by `FOLIO_key` before `compute_col_counts`. Lossless (Fulfillment-source rows of a FOLIO resolve to identical values) and makes every F/G sub-count, incl. Distress, per-property.
**Verified:** Col F 26,534 → **22,282**, Col G 59 → **48** (matches client PBI). Distress binary F/G counts dropped accordingly (e.g. High Equity F 24,294 → 20,526); the 3-way Absentee split was already per-property and now sums **exactly** to the F/G column totals (15,406 + 6,125 + 751 = 22,282; 30 + 14 + 4 = 48).

### Fix 10 — COUNTY restored to the domain (companion HTML/PDF)
**Location:** `build_domain_report.py` — `_DOM_COLS_NEEDED`
**Symptom:** Companion "Fulfillment Distress Analysis" silently failed (`KeyError: 'COUNTY'`, caught by try/except) — the `.html`/`.pdf` on disk were stale.
**Root cause:** the earlier column-filtered-parquet perf change dropped `COUNTY` from `_DOM_COLS_NEEDED`, but the companion's `_build_overview()` groups the Distress Universe by `COUNTY`.
**Fix:** add `COUNTY` back to `_DOM_COLS_NEEDED`, delete `domain.parquet`, and re-run so the parquet rebuilds with COUNTY (now 28 columns).
**Verified:** companion HTML + PDF regenerate; headline "matched sold properties" = 20,044 = Column D; 8 counties.

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
| Pre-foreclosure (and other binary distresses) undercounted in C/D/E/F/G | ✅ Fixed | Read binary distress from fulfillment, not MAIN DISTRESS #1-4 ranking nor domain |
| Columns F/G inflated vs client PBI (duplicate rows + event-vs-property grain) | ✅ Fixed | Drop exact-dup rows on load + dedupe leads_ff/deals_ff by FOLIO → F/G per-property (Fix 9) |
| Companion HTML/PDF silently failing (KeyError COUNTY) | ✅ Fixed | Restore COUNTY to `_DOM_COLS_NEEDED`, rebuild parquet (Fix 10) |

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
| Client name | Leverage Companies |
| Start date with 8020REI | **2024-05-24** |
| Deals/Leads file | `Documents/Deals and Leads Leverage Companies.xlsx` (auto-discovered by name pattern) |
| Market Deals file | `Market Deals/Market Deals Leverage.xlsx` (no PropertyID → address+ZIP fallback) |
| Domain source | 4 CSV parts in `Domain Full Data/` (COO config_1.6M_386_part_1–4.csv) |
| Fulfillment window | Jan 2025 – Mar 2026 (31 xlsx files) |
| Excel sections | Distress only (ZIP section removed at client request) |

Three variables at the top of `build_domain_report.py` must be updated when switching clients:
- `CLIENT_NAME` — used in the output filename (`{CLIENT_NAME} - Distress Report - YYYY-MM.xlsx`)
- `CLIENT_START_DATE` — date the client started with 8020REI
- `MARKET_PATH` — path to the client's market deals file
