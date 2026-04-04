import pandas as pd
import numpy as np

print("Loading domain files...")
p1 = pd.read_excel('Domain Full Data/COO config_536.9K_296_part_1.xlsx')
p2 = pd.read_excel('Domain Full Data/COO config_536.9K_296_part_2.xlsx')
dom = pd.concat([p1, p2], ignore_index=True)
print(f'Total rows: {len(dom):,}')
print(f'Unique FOLIOs: {dom["FOLIO"].nunique():,}')
print()
print('PROPERTY STATUS:')
print(dom['PROPERTY STATUS'].value_counts(dropna=False))
print()

dom['LAST SALE DATE'] = pd.to_datetime(dom['LAST SALE DATE'], errors='coerce')
print('LAST SALE DATE range:')
print('  Min:', dom['LAST SALE DATE'].min())
print('  Max:', dom['LAST SALE DATE'].max())
sold_oct = dom[dom['LAST SALE DATE'] >= '2025-10-01']
print(f'  Sold since Oct 1 2025: {len(sold_oct):,}')
print()

dom['LIKELY DEAL SCORE'] = pd.to_numeric(dom['LIKELY DEAL SCORE'], errors='coerce')
dom['SCORE']             = pd.to_numeric(dom['SCORE'],             errors='coerce')
print(f'LIKELY DEAL SCORE non-null: {dom["LIKELY DEAL SCORE"].notna().sum():,}')
print(f'SCORE non-null:             {dom["SCORE"].notna().sum():,}')
print(f'ACTION PLANS non-null:      {dom["ACTION PLANS"].notna().sum():,}')
dom['total_mkt'] = pd.to_numeric(dom['MARKETING DM COUNT'], errors='coerce').fillna(0) + pd.to_numeric(dom['MARKETING SMS COUNT'], errors='coerce').fillna(0)
print(f'Marketing count > 0:        {(dom["total_mkt"] > 0).sum():,}')
print(f'Marketing count = 0:        {(dom["total_mkt"] == 0).sum():,}')
print()

# Check distress binary columns available
dist_cols = ['PRE-FORECLOSURE','VACANT','TAXES','ESTATE','INTER FAMILY TRANSFER',
             'PROBATE','JUDGEMENT','LIENS CITY/COUNTY','LIENS OTHER','LIENS MECHANIC',
             'HIGH EQUITY','55+','ABSENTEE','DOWNSIZING','DEFAULT RISK']
print('Distress binary columns present in domain:')
for c in dist_cols:
    present = c in dom.columns
    if present:
        count = pd.to_numeric(dom[c], errors='coerce').fillna(0)
        print(f'  {c}: present, {(count > 0).sum():,} properties')
    else:
        print(f'  {c}: MISSING')
print("Done.")
