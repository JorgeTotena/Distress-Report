import pandas as pd
import numpy as np

BASE       = r"C:\Users\danpo\Downloads\Python 8020\Atlas Report New"
DOMAIN_DIR = rf"{BASE}\Domain Full Data"

md = pd.read_excel(rf"{BASE}\Market Deals\Atlas Market Deals.xlsx")
md['SaleDate'] = pd.to_datetime(md['SaleDate'], errors='coerce')

print("=== FIPS CODES IN MARKET DEALS ===")
print(md['FIPS'].value_counts())
print()

print("=== PROPERTYID RANGE ===")
print(f"  Min: {md['PropertyID'].min():,}")
print(f"  Max: {md['PropertyID'].max():,}")
print()

# Try integer comparison with BUYBOX ID
dom1 = pd.read_excel(rf"{DOMAIN_DIR}\COO config_536.9K_296_part_1.xlsx",
                     usecols=['FOLIO','PROPERTY ID (BUYBOX)','PROPERTY ID (DOMAIN)','ADDRESS'])
dom2 = pd.read_excel(rf"{DOMAIN_DIR}\COO config_536.9K_296_part_2.xlsx",
                     usecols=['FOLIO','PROPERTY ID (BUYBOX)','PROPERTY ID (DOMAIN)','ADDRESS'])
dom = pd.concat([dom1, dom2], ignore_index=True)

# Try numeric BUYBOX match
dom['buybox_int'] = pd.to_numeric(dom['PROPERTY ID (BUYBOX)'], errors='coerce')
md_ids = set(md['PropertyID'].dropna().astype(int).unique())
dom_buybox_int = set(dom['buybox_int'].dropna().astype(int).unique())
print(f"  PropertyID vs BUYBOX (numeric): {len(md_ids & dom_buybox_int):,} matches")
print()

# Address match attempt
print("=== ADDRESS SAMPLE (market deals) ===")
print(md['FullStreetAddress'].head(10).tolist())
print()
print("=== ADDRESS SAMPLE (domain) ===")
print(dom['ADDRESS'].head(10).tolist())
print()

# Try uppercase normalized address match
md['addr_norm'] = md['FullStreetAddress'].astype(str).str.upper().str.strip()
dom['addr_norm'] = dom['ADDRESS'].astype(str).str.upper().str.strip()

md_addrs  = set(md['addr_norm'].unique())
dom_addrs = set(dom['addr_norm'].unique())
addr_matches = md_addrs & dom_addrs
print(f"  Address matches (normalized): {len(addr_matches):,} out of {len(md_addrs):,} market deal addresses")
print()
if addr_matches:
    print("  Sample matching addresses:")
    for a in list(addr_matches)[:5]:
        print(f"    {a}")
