# Distress Report

Automated Excel report generator for 8020REI clients. Given a set of fulfillment files, a domain export, a deals/leads file, and a market deals file, this project produces a fully formatted distress report (`{CLIENT} - Distress Report - YYYY-MM.xlsx`) with all distress metrics, score breakdowns, and concentration formulas.

---

## Prerequisites

### Python packages

Install all dependencies before running any script:

```bash
pip install pandas openpyxl pyarrow fastparquet
```

### Python version
Python 3.9 or higher is recommended.

---

## Required Files & Folder Structure

Before running the report, make sure the following files are in place:

```
Distress Report/
│
├── Fulfillments/                        ← Monthly fulfillment exports (xlsx)
│   ├── 2025-01-07 CLIENT 75K Direct Mail.xlsx
│   ├── 2025-02-04 CLIENT 120K Cold Calling.xlsx
│   └── ...                             ← all months you want in the analysis window
│
├── Domain Full Data/                    ← Domain export (xlsx or csv, can be split in multiple parts)
│   ├── COO config_1.6M_386_part_1.csv
│   └── COO config_1.6M_386_part_2.csv  ← csv and xlsx both supported
│
├── Documents/                           ← Deals & Leads file from client CRM
│   └── Deals and Leads CLIENT.xlsx
│
├── Market Deals/                        ← Market deals file (investor activity)
│   └── Market Deals CLIENT.xlsx
│
├── compile_domain.py                    ← Run once when domain files change
└── build_domain_report.py              ← MAIN SCRIPT — generates the full report
```

> **Important:** File names do not need to match exactly. The scripts auto-discover files by folder and partial name match. Deals/Leads: any `.xlsx` in `Documents/` whose name contains "leads" or "deals". Domain parts: all `.xlsx` or `.csv` files in `Domain Full Data/`.

---

## Client Configuration

At the top of `build_domain_report.py`, set three variables before each client run:

```python
CLIENT_NAME       = "Leverage Companies"           # Used in the output filename
CLIENT_START_DATE = pd.to_datetime("2024-05-24")   # When the client started with 8020REI
MARKET_PATH       = BASE / "Market Deals" / "Markeet Deals Leverage.xlsx"  # Path to market deals file
```

These are the only values that change between clients.

---

## How to Generate the Report (Step by Step)

### Step 1 — Place all source files

Make sure all required files are in their correct folders (see structure above). You need:

- At least 1 fulfillment xlsx in `Fulfillments/`
- At least 1 domain xlsx in `Domain Full Data/`
- The deals & leads xlsx in `Documents/`
- The market deals xlsx in `Market Deals/`

### Step 2 — Compile the domain (run once, or when domain files change)

This step reads the domain source files (xlsx or csv) and saves a trimmed `.parquet` cache containing only the columns the report needs. Only needs to be re-run when the domain source files are replaced or updated.

```bash
python compile_domain.py
```

**Output:** `Domain Full Data/domain.parquet`

You will see output like:
```
Found 4 file(s): ['COO config_1.6M_386_part_1.csv', ...]
Reading part 1: COO config_1.6M_386_part_1.csv...
  644836 rows, 27 columns
...
Combined: 2581394 rows, 1444626 unique FOLIOs
Saved: Domain Full Data/domain.parquet (27 columns)
```

### Step 3 — Run the main report script

```bash
python build_domain_report.py
```

This script does everything automatically:
- Uses every `.xlsx` in `Fulfillments/` as the analysis window (folder defines the window)
- Compiles all fulfillment files — caches each as a `.parquet` sidecar on first run; subsequent runs load from parquet (much faster)
- Loads the domain parquet (column-filtered for speed)
- Loads the deals & leads file
- Loads the market deals file
- Builds all columns (B through G) with distress metrics and score buckets
- Writes the final formatted Excel report

**Output:** `Leverage Companies - Distress Report - YYYY-MM.xlsx` (e.g. `Leverage Companies - Distress Report - 2026-06.xlsx`)

A side audit file is also written each run: `Fulfillment_Compilation.xlsx`. The compilation contains every fulfillment row plus three validation columns appended at the end so you can spot-check the report:

| Validation column | What it shows | Validates |
|---|---|---|
| `LAST SALE DATE` | Most recent sale per FOLIO from the domain | Column D (Sold) |
| `PROPERTY STATUS` | `Lead` or `Deal` from the deals/leads file (Deal wins if both); blank if no match | Columns F / G |
| `MARKET DEAL` | `Yes` if the FOLIO is in the Market Deals overlap, else `No` | Column E |

---

## What the Report Contains

| Column | Description |
|--------|-------------|
| B | Total Properties — all domain properties with BUYBOX SCORE > 0 |
| C | Properties in the fulfillment — MAX aggregated metrics across the actual fulfillment window |
| D | Sold since start of fulfillment window — properties sold from the first month with a fulfillment file onward |
| E | Market Deals — investor purchases that overlap with Column D (sold properties) |
| F | Client Leads — leads, appointments, dead leads, and contracts matched to fulfillment |
| G | Client Deals — closed deals matched to fulfillment |
| H | Sold Concentration % — formula =D/C |
| I | Sold to Investors Concentration — formula =E/C |
| J | Client Deals Concentration — formula =G/C |
| K | Client Leads Concentration — formula =F/C |

Each column includes sub-rows for distress breakdowns (Pre-foreclosure, Vacant, Taxes, Estate, Probate, Liens, High Equity, 55+, Default Risk, Absentee, Owner Occupied, …) and score/action plan distribution. Binary distresses for fulfillment-related columns (C/D/E/F/G) are read directly from the fulfillment file's per-month flag columns (e.g. `PRE-FORECLOSURE > 0`); MAIN DISTRESS #1–4 ranking is no longer used.

---

## Fulfillment Window Logic

**The window is defined entirely by the `.xlsx` files in `Fulfillments/`.** Drop in the months you want analyzed; remove the months you don't. Whatever is in the folder is what the report uses.

- `WINDOW_START` / `SOLD_SINCE` = first day of the earliest file's month (e.g. `2025-11-08 ...xlsx` → `2025-11-01`)
- `WINDOW_END` = first day of the latest file's month
- Column D ("Sold") counts properties whose `LAST SALE DATE >= SOLD_SINCE` and that were also recommended in a fulfillment

| Earliest file in folder | `SOLD_SINCE` |
|---|---|
| `2025-09-08 ...xlsx` | 2025-09-01 |
| `2025-11-08 ...xlsx` | 2025-11-01 |
| `2026-01-08 ...xlsx` | 2026-01-01 |

The script no longer enforces a 6-month "intended window" — there is no silent filter. To change the window, change what is in the folder.

---

## Using Claude Code to Generate the Report

If you are using [Claude Code](https://claude.ai/code) (the AI assistant in the terminal or IDE), you can simply describe what you need and it will guide you through the process. Recommended prompts:

**To run a fresh report for the current month:**
> "Run the distress report for [CLIENT NAME]. The domain files and fulfillments are already in place."

**To switch to a new client:**
> "Switch the distress report to client [CLIENT NAME], start date [DATE], and market deals file [FILENAME]. Update the variables and run the report."

**If the domain files were updated:**
> "The domain files in Domain Full Data/ were replaced. Recompile the domain and regenerate the report."

Claude Code will read the scripts, update the necessary variables, run the steps in order, and report any issues it encounters.

---

## Switching to a New Client

1. Drop the client's fulfillment files into `Fulfillments/`
2. Drop the client's domain files into `Domain Full Data/`
3. Drop the client's deals & leads file into `Documents/`
4. Drop the client's market deals file into `Market Deals/`
5. Open `build_domain_report.py` and update the three client variables at the top (lines 35–41):
   - `CLIENT_NAME`
   - `CLIENT_START_DATE`
   - `MARKET_PATH`
6. Run `python compile_domain.py` (domain files changed)
7. Run `python build_domain_report.py`

---

## Troubleshooting

| Problem | Cause | Fix |
|---|---|---|
| `FileNotFoundError: No xlsx files found in Domain Full Data` | Domain files missing or wrong folder | Place domain xlsx files in `Domain Full Data/` |
| `domain.parquet not found` | `compile_domain.py` was never run | Run `python compile_domain.py` first |
| `No fulfillment files found in <Fulfillments path>` | Folder is empty | Add the month files you want analyzed to `Fulfillments/` |
| `No deals/leads file found` | No xlsx with "deals" or "leads" in name in `Documents/` | Rename or move the file into `Documents/` |
| DEFAULT RISK shows 0 (with warning) | Domain export is missing the `DEFAULT RISK` column | Normal — script handles this gracefully; no action needed |
| Output filename has wrong month | Script ran on a different date | The filename reflects the month the script was run |

---

## Files Reference

| File | Purpose |
|---|---|
| `build_domain_report.py` | **Main script** — generates the full report |
| `compile_domain.py` | Converts domain xlsx parts to parquet cache — run once per domain update |
| `build_reports.py` | Legacy — Column C only; not used for main output |
| `build_population.py` | Legacy — fulfillment audit file; not used for main output |
| `build_deals_leads.py` | Legacy — Columns F and G only; not used for main output |
| `Fulfillment_Compilation.xlsx` | Audit file written each run — all fulfillment rows combined, plus `LAST SALE DATE`, `PROPERTY STATUS`, `MARKET DEAL` columns for validating D/F/G/E |
| `{CLIENT} - Distress Report - YYYY-MM.xlsx` | **Report output** |
| `Domain Full Data/domain.parquet` | Cached domain — rebuilt by `compile_domain.py` |
