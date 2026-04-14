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
│   ├── 2025-08-08 SBDHOUSING 75K Direct Mail.xlsx
│   ├── 2025-09-08 SBDHOUSING 116K Direct Mail.xlsx
│   ├── 2025-10-08 SBDHOUSING 116K Direct Mail.xlsx
│   ├── 2025-11-08 SBDHOUSING 116K Direct Mail.xlsx
│   ├── 2025-12-07 SBDHOUSING 116K Direct Mail.xlsx
│   └── 2026-01-07 SBDHOUSING 129K Direct Mail.xlsx
│
├── Domain Full Data/                    ← Domain export (can be split in multiple parts)
│   ├── COO config_712.6K_412_part_1.xlsx
│   └── COO config_712.6K_412_part_2.xlsx
│
├── Documents/                           ← Deals & Leads file from client CRM
│   └── Deals and Leads SBD.xlsx
│
├── Market Deals/                        ← Market deals file (investor activity)
│   └── SBD Market Deals.xlsx
│
├── compile_domain.py                    ← Run once when domain files change
└── build_domain_report.py              ← MAIN SCRIPT — generates the full report
```

> **Important:** File names do not need to match exactly. The scripts auto-discover files by folder and partial name match. Deals/Leads: any `.xlsx` in `Documents/` whose name contains "leads" or "deals". Domain parts: all `.xlsx` files in `Domain Full Data/`.

---

## Client Configuration

At the top of `build_domain_report.py`, set three variables before each client run:

```python
CLIENT_NAME       = "SBD"                   # Used in the output filename
CLIENT_START_DATE = pd.to_datetime("2023-09-26")  # When the client started with 8020REI
MARKET_PATH       = BASE / "Market Deals" / "SBD Market Deals.xlsx"  # Path to market deals file
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

This step converts the domain xlsx files into a fast-loading `.parquet` cache. Only needs to be re-run when the domain source files are replaced or updated.

```bash
python compile_domain.py
```

**Output:** `Domain Full Data/domain.parquet`

You will see output like:
```
Found 2 file(s): ['COO config_712.6K_412_part_1.xlsx', 'COO config_712.6K_412_part_2.xlsx']
Reading part 1: COO config_712.6K_412_part_1.xlsx...
  412000 rows
Reading part 2: COO config_712.6K_412_part_2.xlsx...
  304089 rows
Combined: 716089 rows, 709929 unique FOLIOs
Saved: Domain Full Data/domain.parquet
```

### Step 3 — Run the main report script

```bash
python build_domain_report.py
```

This script does everything automatically:
- Selects the correct 6-month fulfillment window based on today's date
- Compiles all fulfillment files
- Loads the domain parquet
- Loads the deals & leads file
- Loads the market deals file
- Builds all columns (B through G) with distress metrics and score buckets
- Writes the final formatted Excel report

**Output:** `SBD - Distress Report - YYYY-MM.xlsx` (e.g. `SBD - Distress Report - 2026-04.xlsx`)

A side audit file is also written each run: `Fulfillment_Compilation.xlsx`

---

## What the Report Contains

| Column | Description |
|--------|-------------|
| B | Total Properties — all domain properties with BUYBOX SCORE > 0 |
| C | Properties in the fulfillment — MAX aggregated metrics across the 6-month window |
| D | Sold since start of fulfillment window — properties sold after the window opens |
| E | Market Deals — investor purchases that overlap with Column D (sold properties) |
| F | Client Leads — leads, appointments, dead leads, and contracts matched to fulfillment |
| G | Client Deals — closed deals matched to fulfillment |
| H | Sold Concentration % — formula =D/C |
| I | Sold to Investors Concentration — formula =E/C |
| J | Client Deals Concentration — formula =G/C |
| K | Client Leads Concentration — formula =F/C |

Each column includes sub-rows for distress breakdowns (Distress #1–4, Vacant, Default Risk, Owner Occupied) and score/action plan distribution.

---

## Fulfillment Window Logic

The script automatically calculates which fulfillment files to include based on today's date:

- **Most recent month included** = Today's month − 3 months
- **Window** = 6 months ending on that month
- If fewer than 6 files are available, it uses whatever is there (minimum 1)

| Report Month | Most Recent File | 6-Month Window |
|---|---|---|
| April 2026 | January 2026 | Aug 2025 – Jan 2026 |
| May 2026 | February 2026 | Sep 2025 – Feb 2026 |
| June 2026 | March 2026 | Oct 2025 – Mar 2026 |

Nothing needs to be changed in the script — the window updates automatically each month.

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
| `No fulfillment files found in window` | Fulfillment files are outside the current 6-month window | Add the correct month files to `Fulfillments/` |
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
| `domain_check.py` | Diagnostic — inspect domain column structure |
| `market_deals_check.py` | Diagnostic — match market deals to domain/fulfillment |
| `Fulfillment_Compilation.xlsx` | Audit file written each run (all fulfillment rows combined) |
| `{CLIENT} - Distress Report - YYYY-MM.xlsx` | **Report output** |
| `Domain Full Data/domain.parquet` | Cached domain — rebuilt by `compile_domain.py` |
