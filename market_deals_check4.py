import pandas as pd
import numpy as np

BASE       = r"C:\Users\danpo\Downloads\Python 8020\Atlas Report New"
DOMAIN_DIR = rf"{BASE}\Domain Full Data"

# Load market deals — use PropertyID = BUYBOX ID
md = pd.read_excel(rf"{BASE}\Market Deals\Atlas Market Deals.xlsx")
md['SaleDate'] = pd.to_datetime(md['SaleDate'], errors='coerce')
md_unique_ids = set(md['PropertyID'].dropna().astype(int).unique())

# Filter to properties actually sold (SaleDate >= Oct 1 2025)
md_sold_since_oct = md[md['SaleDate'] >= '2025-10-01']
md_sold_ids = set(md_sold_since_oct['PropertyID'].dropna().astype(int).unique())

print(f"Market deals total unique properties: {len(md_unique_ids):,}")
print(f"Market deals with SaleDate >= Oct 1 2025: {md_sold_since_oct['PropertyID'].nunique():,} unique properties")
print()

# Load domain — join on PROPERTY ID (BUYBOX) and LAST SALE DATE
dom1 = pd.read_excel(rf"{DOMAIN_DIR}\COO config_536.9K_296_part_1.xlsx",
                     usecols=['FOLIO','PROPERTY ID (BUYBOX)','LAST SALE DATE'])
dom2 = pd.read_excel(rf"{DOMAIN_DIR}\COO config_536.9K_296_part_2.xlsx",
                     usecols=['FOLIO','PROPERTY ID (BUYBOX)','LAST SALE DATE'])
dom = pd.concat([dom1, dom2], ignore_index=True)
dom['LAST SALE DATE'] = pd.to_datetime(dom['LAST SALE DATE'], errors='coerce')
dom['buybox_int'] = pd.to_numeric(dom['PROPERTY ID (BUYBOX)'], errors='coerce').dropna().astype('Int64')

sold_dom = dom[dom['LAST SALE DATE'] >= '2025-10-01'].copy()
sold_dom_ids = set(sold_dom['buybox_int'].dropna().astype(int).unique())

print(f"Domain sold since Oct 1 2025: {len(sold_dom):,} unique FOLIOs")
print(f"  of which have a BUYBOX ID:  {sold_dom['buybox_int'].notna().sum():,}")
print()

# Overlap
in_both      = md_unique_ids & sold_dom_ids
md_not_sold  = md_unique_ids - sold_dom_ids
sold_not_md  = sold_dom_ids - md_unique_ids

print("=== OVERLAP: Market Deals vs Domain Sold ===")
print(f"  Market deals unique properties:          {len(md_unique_ids):,}")
print(f"  Domain sold since Oct 1 (with BUYBOX):   {len(sold_dom_ids):,}")
print(f"  IN BOTH (market deal AND sold):          {len(in_both):,}  ({len(in_both)/len(md_unique_ids)*100:.1f}% of market deals)")
print(f"  Market deals NOT in domain sold:         {len(md_not_sold):,}  ({len(md_not_sold)/len(md_unique_ids)*100:.1f}% of market deals)")
print(f"  Domain sold NOT in market deals:         {len(sold_not_md):,}  ({len(sold_not_md)/len(sold_dom_ids)*100:.1f}% of sold)")
print()

# Also check: md properties with SaleDate >= Oct 1 vs domain sold
in_both_oct = md_sold_ids & sold_dom_ids
print("=== REFINED: Market deals sold since Oct 1 vs Domain sold since Oct 1 ===")
print(f"  Market deals sold since Oct 1 (unique):  {len(md_sold_ids):,}")
print(f"  Domain sold since Oct 1 (with BUYBOX):   {len(sold_dom_ids):,}")
print(f"  In both:                                 {len(in_both_oct):,}")
print(f"  Market sold NOT in domain sold:          {len(md_sold_ids - sold_dom_ids):,}")
print(f"  Domain sold NOT in market sold:          {len(sold_dom_ids - md_sold_ids):,}")
