import pandas as pd
import numpy as np

BASE       = r"C:\Users\danpo\Downloads\Python 8020\Atlas Report New"
DOMAIN_DIR = rf"{BASE}\Domain Full Data"

# Load market deals
print("Loading Atlas Market Deals...")
md = pd.read_excel(rf"{BASE}\Market Deals\Atlas Market Deals.xlsx")
print(f"  Rows: {len(md):,}")
print(f"  Columns: {list(md.columns)}")
print()

# Identify FOLIO column (may be named differently)
folio_col = None
for c in md.columns:
    if 'folio' in c.lower() or 'parcel' in c.lower() or 'property id' in c.lower():
        folio_col = c
        break
print(f"  FOLIO column detected: {folio_col}")

if folio_col:
    md_folios = md[folio_col].astype(str).str.strip()
    md_folios = md_folios[md_folios != 'nan']
    print(f"  Unique FOLIOs in market deals: {md_folios.nunique():,}")
else:
    print("  WARNING: No FOLIO column found — printing first 5 rows:")
    print(md.head())
print()

# Load domain sold properties (since Oct 1 2025)
print("Loading domain sold properties...")
dom1 = pd.read_excel(rf"{DOMAIN_DIR}\COO config_536.9K_296_part_1.xlsx",
                     usecols=['FOLIO', 'LAST SALE DATE'])
dom2 = pd.read_excel(rf"{DOMAIN_DIR}\COO config_536.9K_296_part_2.xlsx",
                     usecols=['FOLIO', 'LAST SALE DATE'])
dom  = pd.concat([dom1, dom2], ignore_index=True)
dom['LAST SALE DATE'] = pd.to_datetime(dom['LAST SALE DATE'], errors='coerce')
sold = dom[dom['LAST SALE DATE'] >= '2025-10-01'].copy()
sold_folios = sold['FOLIO'].astype(str).str.strip()
print(f"  Sold since Oct 1 2025: {len(sold):,} rows, {sold_folios.nunique():,} unique FOLIOs")
print()

if folio_col:
    md_unique  = set(md_folios.unique())
    sold_unique = set(sold_folios.unique())

    in_both        = md_unique & sold_unique
    md_not_sold    = md_unique - sold_unique
    sold_not_md    = sold_unique - md_unique

    print("=== OVERLAP ANALYSIS ===")
    print(f"  Market deals unique FOLIOs:          {len(md_unique):,}")
    print(f"  Sold (since Oct 1) unique FOLIOs:    {len(sold_unique):,}")
    print(f"  In BOTH (market deal AND sold):      {len(in_both):,}")
    print(f"  Market deals NOT in sold:            {len(md_not_sold):,}  ({len(md_not_sold)/len(md_unique)*100:.1f}% of market deals)")
    print(f"  Sold NOT in market deals:            {len(sold_not_md):,}  ({len(sold_not_md)/len(sold_unique)*100:.1f}% of sold)")
