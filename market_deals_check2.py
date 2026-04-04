import pandas as pd
import numpy as np

BASE       = r"C:\Users\danpo\Downloads\Python 8020\Atlas Report New"
DOMAIN_DIR = rf"{BASE}\Domain Full Data"

# Load market deals
md = pd.read_excel(rf"{BASE}\Market Deals\Atlas Market Deals.xlsx")

# Inspect PropertyID and SaleDate
print("=== MARKET DEALS FILE ===")
print(f"Rows: {len(md):,}")
print(f"PropertyID sample: {md['PropertyID'].head(10).tolist()}")
print(f"PropertyID dtype: {md['PropertyID'].dtype}")
print(f"Unique PropertyIDs: {md['PropertyID'].nunique():,}")
print()
md['SaleDate'] = pd.to_datetime(md['SaleDate'], errors='coerce')
print(f"SaleDate range: {md['SaleDate'].min()} to {md['SaleDate'].max()}")
print(f"SaleDate non-null: {md['SaleDate'].notna().sum():,}")
print(f"Sold since Oct 1 2025 (SaleDate): {(md['SaleDate'] >= '2025-10-01').sum():,}")
print()
md['rec_date'] = pd.to_datetime(md['currentsalerecordingdate_deal'], errors='coerce')
print(f"currentsalerecordingdate_deal non-null: {md['rec_date'].notna().sum():,}")
print(f"Recording date range: {md['rec_date'].min()} to {md['rec_date'].max()}")
print()
print(f"Period values: {sorted(md['Period'].unique())}")
print()

# Load a sample of domain to compare identifier formats
print("=== DOMAIN IDENTIFIERS (sample) ===")
dom_sample = pd.read_excel(rf"{DOMAIN_DIR}\COO config_536.9K_296_part_1.xlsx",
                           usecols=['FOLIO','PROPERTY ID (BUYBOX)','PROPERTY ID (DOMAIN)'], nrows=10)
print(dom_sample.to_string())
print()

# Check if PropertyID from market deals matches any domain column
print("Checking PropertyID match against domain...")
dom1 = pd.read_excel(rf"{DOMAIN_DIR}\COO config_536.9K_296_part_1.xlsx",
                     usecols=['FOLIO','PROPERTY ID (BUYBOX)','PROPERTY ID (DOMAIN)'])
dom2 = pd.read_excel(rf"{DOMAIN_DIR}\COO config_536.9K_296_part_2.xlsx",
                     usecols=['FOLIO','PROPERTY ID (BUYBOX)','PROPERTY ID (DOMAIN)'])
dom = pd.concat([dom1, dom2], ignore_index=True)

md_ids = set(md['PropertyID'].astype(str).str.strip().unique())

dom_folio  = set(dom['FOLIO'].astype(str).str.strip().unique())
dom_buybox = set(dom['PROPERTY ID (BUYBOX)'].astype(str).str.strip().unique())
dom_domain = set(dom['PROPERTY ID (DOMAIN)'].astype(str).str.strip().unique())

print(f"  Market PropertyIDs vs FOLIO:             {len(md_ids & dom_folio):,} matches")
print(f"  Market PropertyIDs vs PROPERTY ID (BUYBOX): {len(md_ids & dom_buybox):,} matches")
print(f"  Market PropertyIDs vs PROPERTY ID (DOMAIN): {len(md_ids & dom_domain):,} matches")
