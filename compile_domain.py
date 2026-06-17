"""
compile_domain.py
─────────────────────────────────────────────────────────────────────────────
One-time script: reads the domain xlsx/csv parts and saves a single parquet.
Run this whenever the source files change.

Only the columns consumed by build_domain_report.py are kept — the parquet is
a trimmed cache, not a full copy. Add columns to _DOM_COLS_NEEDED if a new
field is required downstream.

Output: Domain Full Data/domain.parquet
"""

import pandas as pd
from pathlib import Path

BASE       = Path(__file__).parent
DOMAIN_DIR = BASE / "Domain Full Data"
OUT        = DOMAIN_DIR / "domain.parquet"

# Columns used by build_domain_report.py — only these are written to parquet.
_DOM_COLS_NEEDED = [
    'FOLIO', 'BUYBOX SCORE', 'LIKELY DEAL SCORE', 'SCORE',
    'MARKETING DM COUNT', 'MARKETING SMS COUNT', 'ACTION PLANS',
    'LAST SALE DATE', 'MARKETING FIRST RECOMMENDATION', 'PROPERTY ID (BUYBOX)', 'ZIP', 'ADDRESS',
    'ABSENTEE', 'HIGH EQUITY', 'DOWNSIZING', 'PRE-FORECLOSURE', 'VACANT',
    '55+', 'ESTATE', 'INTER FAMILY TRANSFER', 'TAXES', 'PROBATE',
    'JUDGEMENT', 'LIENS CITY/COUNTY', 'LIENS OTHER', 'LIENS MECHANIC',
    'DEFAULT RISK',
]
_DOM_COLS_SET = set(_DOM_COLS_NEEDED)

# Discover xlsx and csv files in the folder (skip temp lock files starting with ~$)
xlsx_files = sorted(f for f in DOMAIN_DIR.glob("*.xlsx") if not f.name.startswith("~$"))
csv_files  = sorted(f for f in DOMAIN_DIR.glob("*.csv"))
all_files  = xlsx_files + csv_files
if not all_files:
    raise FileNotFoundError(f"No xlsx or csv files found in {DOMAIN_DIR}")
print(f"Found {len(all_files)} file(s): {[f.name for f in all_files]}")

parts = []
for i, path in enumerate(all_files, 1):
    print(f"Reading part {i}: {path.name}...")
    if path.suffix.lower() == '.csv':
        # Peek at headers first so usecols only requests columns that exist.
        _header = pd.read_csv(path, nrows=0).columns.tolist()
        _use    = [c for c in _DOM_COLS_NEEDED if c in _header]
        part    = pd.read_csv(path, usecols=_use, low_memory=False)
    else:
        part = pd.read_excel(path, usecols=lambda c: c in _DOM_COLS_SET)
    print(f"  {len(part):,} rows, {len(part.columns)} columns")
    parts.append(part)

dom = pd.concat(parts, ignore_index=True)
print(f"Combined: {len(dom):,} rows, {dom['FOLIO'].nunique():,} unique FOLIOs")

# Coerce mixed-type object columns to string so pyarrow can serialize them.
# Common offenders: ZIP codes that arrive as a mix of "12345" and 12345.0 / NaN.
obj_cols = dom.select_dtypes(include='object').columns
for c in obj_cols:
    dom[c] = dom[c].where(dom[c].notna(), None).astype('string')

dom.to_parquet(OUT, index=False)
print(f"\nSaved: {OUT} ({len(dom.columns)} columns)")
