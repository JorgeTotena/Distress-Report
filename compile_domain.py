"""
compile_domain.py
─────────────────────────────────────────────────────────────────────────────
One-time script: reads the two domain xlsx parts and saves a single parquet.
Run this whenever the source files change.

Output: Domain Full Data/domain.parquet
"""

import pandas as pd
from pathlib import Path

BASE       = Path(__file__).parent
DOMAIN_DIR = BASE / "Domain Full Data"
OUT        = DOMAIN_DIR / "domain.parquet"

# Discover all xlsx files in the folder (skip temp lock files starting with ~$)
xlsx_files = sorted(f for f in DOMAIN_DIR.glob("*.xlsx") if not f.name.startswith("~$"))
if not xlsx_files:
    raise FileNotFoundError(f"No xlsx files found in {DOMAIN_DIR}")
print(f"Found {len(xlsx_files)} file(s): {[f.name for f in xlsx_files]}")

parts = []
for i, path in enumerate(xlsx_files, 1):
    print(f"Reading part {i}: {path.name}...")
    part = pd.read_excel(path)
    print(f"  {len(part):,} rows")
    parts.append(part)

dom = pd.concat(parts, ignore_index=True)
print(f"Combined: {len(dom):,} rows, {dom['FOLIO'].nunique():,} unique FOLIOs")

# Coerce mixed-type object columns to string so pyarrow can serialize them.
# Common offenders: ZIP codes that arrive as a mix of "12345" and 12345.0 / NaN.
obj_cols = dom.select_dtypes(include='object').columns
for c in obj_cols:
    dom[c] = dom[c].where(dom[c].notna(), None).astype('string')

dom.to_parquet(OUT, index=False)
print(f"\nSaved: {OUT}")
