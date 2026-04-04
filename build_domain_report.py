"""
build_domain_report.py
─────────────────────────────────────────────────────────────────────────────
Generates Expected_Result_Max_BCDFG_InFulfillment.xlsx with:
  Column B — Total Properties        (all 536k domain properties)
  Column C — Properties in the fulfillment (MAX logic, all 37k)
  Column D — Sold since Oct 1 2025   (pre-sale scores from fulfillment where available)
  Column E — Clients Lead            (fulfillment-matched leads only: 345)
  Column F — Client Deals            (fulfillment-matched deals only: 4)
  Column G — Sold Concentration %    =D/C  (formula)
  Column H — Client Deals Conc.      =F/C  (formula)
  Column I — Client Leads Conc.      =E/C  (formula)
  [Column E Market Deals — temporarily disabled, see STEP 5 below]
"""

import pandas as pd
import numpy as np
import re
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

BASE              = Path(__file__).parent
FULFILLMENTS_DIR  = BASE / "Fulfillments"
COMPILED_PATH     = BASE / "Fulfillment_Compilation.xlsx"   # written each run for audit
DEALS_PATH        = BASE / "Documents" / "Deals and Leads Freedom.xlsx"
DOMAIN_DIR        = BASE / "Domain Full Data"
MARKET_PATH       = BASE / "Market Deals" / "Atlas Market Deals.xlsx"
OUTPUT            = BASE / "Expected_Result_Max_BCDFG_InFulfillment.xlsx"
# DR_PATH removed — DEFAULT RISK is now a column in the domain file

# ── Client start date (used to filter leads and deals) ───────────────────────
# Freedom REI started doing business with 8020REI on 2021-10-26
CLIENT_START_DATE = pd.to_datetime("2021-10-26")
print(f"Filtering leads/deals to on or after: {CLIENT_START_DATE.date()}")

# ── Bucket definitions ────────────────────────────────────────────────────────
likely_bins   = [-0.5, 20, 40, 60, 80, 100.5]
likely_labels = ['20-0', '40-21', '60-41', '80-61', '100-81']
score_bins    = [-0.5, 200, 400, 600, 800, 1000.5]
score_labels  = ['200-0', '400-201', '600-401', '800-601', '1000-801']
mkt_bins      = [0.5, 5, 10, 19, 1e9]
mkt_labels    = ['1 to 5', '6 to 10', '11 to 19', '20+']
action_map    = {30: '30 day', 60: '60 day', 90: '90 day'}

distress_order = [
    'Pre-foreclosure', 'Vacant', 'Taxes (Tax Delinquent)', 'Estate (Pre-Probate)',
    'Inter family transfer', 'Probate', 'Judgement',
    'Liens city/county', 'Liens other', 'Liens Mechanic',
    'Default Risk', 'High Equity', '55+ (Senior)',
    'Absentee', 'Absentee Out of State', 'Downsizing',
    'Owner Occupied',
]

DOMAIN_DIST_MAP = {
    'PRE-FORECLOSURE':       'Pre-foreclosure',
    'VACANT':                'Vacant',
    'TAXES':                 'Taxes (Tax Delinquent)',
    'ESTATE':                'Estate (Pre-Probate)',
    'INTER FAMILY TRANSFER': 'Inter family transfer',
    'PROBATE':               'Probate',
    'JUDGEMENT':             'Judgement',
    'LIENS CITY/COUNTY':     'Liens city/county',
    'LIENS OTHER':           'Liens other',
    'LIENS MECHANIC':        'Liens Mechanic',
    'DEFAULT RISK':          'Default Risk',
    'HIGH EQUITY':           'High Equity',
    '55+':                   '55+ (Senior)',
    # ABSENTEE handled separately: 1=in-state, 2=out-of-state, 0=owner occupied
    'DOWNSIZING':            'Downsizing',
}

COLS = [
    'Category', 'Total Properties', 'Properties in the fulfillment',
    'Sold', 'Clients Lead', 'Client Deals',
    'Sold Concentration %Sold',
    # 'Market Deals' — disabled, see STEP 5
    # 'Sold to investors concentration' — disabled (depends on Market Deals)
    'Client Deals Concentration', 'Client Leads Concentration',
]
SECTION_NAMES = {'Likely Deal Score', 'Total Score', 'Action Plan', 'Mkt Count', 'Distress'}

# ── Default Risk lookup — derived from domain (DEFAULT RISK column) ───────────
# Loaded early so dr_folios is available for fulfillment injection (STEP 1)
print("Loading Default Risk from domain...")
_parquet_early = DOMAIN_DIR / "domain.parquet"
if _parquet_early.exists():
    _dr = pd.read_parquet(_parquet_early, columns=['PROPERTY ID (BUYBOX)', 'FOLIO', 'DEFAULT RISK'])
else:
    _xlsx = sorted(f for f in DOMAIN_DIR.glob("*.xlsx") if not f.name.startswith("~$"))
    _dr   = pd.concat([pd.read_excel(f, usecols=['PROPERTY ID (BUYBOX)', 'FOLIO', 'DEFAULT RISK']) for f in _xlsx], ignore_index=True)
_dr_risk = _dr[pd.to_numeric(_dr['DEFAULT RISK'], errors='coerce').fillna(0) > 0]
dr_folios = set(_dr_risk['FOLIO'].dropna().astype(str).str.strip().unique())
print(f"  {len(dr_folios):,} unique FOLIOs with Default Risk")
del _dr, _dr_risk

# ── Helpers ───────────────────────────────────────────────────────────────────
def extract_days(val):
    if pd.isna(val): return np.nan
    m = re.match(r'(\d+)\s*DAY', str(val).upper())
    return int(m.group(1)) if m else np.nan

def count_by_bin(series, bins, labels):
    bucketed = pd.cut(series.dropna(), bins=bins, labels=labels)
    counts = bucketed.value_counts()
    return {lbl: int(counts.get(lbl, 0)) for lbl in labels}

def count_by_action(series, mapping):
    mapped = series.dropna().map(mapping)
    counts = mapped.value_counts()
    return {c: int(counts.get(c, 0)) for c in ['30 day', '60 day', '90 day']}

def domain_distress_counts(df_sub):
    result = {dtype: 0 for dtype in distress_order}
    for dom_col, label in DOMAIN_DIST_MAP.items():
        if dom_col in df_sub.columns:
            vals = pd.to_numeric(df_sub[dom_col], errors='coerce').fillna(0)
            result[label] = int((vals > 0).sum())
    # ABSENTEE: 1 = in-state, 2 = out-of-state, 0 = owner occupied
    # Do NOT fillna — NaN means unknown, not owner-occupied
    if 'ABSENTEE' in df_sub.columns:
        abs_vals = pd.to_numeric(df_sub['ABSENTEE'], errors='coerce')
        result['Absentee']              = int((abs_vals == 1).sum())
        result['Absentee Out of State'] = int((abs_vals == 2).sum())
        result['Owner Occupied']        = int((abs_vals == 0).sum())
    # DEFAULT RISK is now a direct column in the domain file — handled by DOMAIN_DIST_MAP loop above
    return result


# ══════════════════════════════════════════════════════════════════════════════
# STEP 1: Compile fulfillment files → MAX aggregations per FOLIO
# Reads all xlsx files from Fulfillments/ folder — no pre-built file needed
# ══════════════════════════════════════════════════════════════════════════════
print("Compiling fulfillment files...")
_ff_files = sorted(f for f in FULFILLMENTS_DIR.glob("*.xlsx") if not f.name.startswith("~$"))
if not _ff_files:
    raise FileNotFoundError(f"No fulfillment files found in {FULFILLMENTS_DIR}")

_parts = []
for _f in _ff_files:
    _part = pd.read_excel(_f)
    _part['Month']              = _f.name[:7]   # YYYY-MM from filename prefix
    _part['Marketing_Channel']  = 'DM' if 'Direct Mail' in _f.name else ('SMS' if 'SMS' in _f.name else 'Other')
    _part['Source_File']        = _f.name
    _parts.append(_part)
    print(f"  {_f.name}: {len(_part):,} rows")

df = pd.concat(_parts, ignore_index=True)
print(f"  Total: {len(df):,} rows, {df['FOLIO'].nunique():,} unique FOLIOs")

print("  Saving Fulfillment_Compilation.xlsx for audit...")
df.to_excel(COMPILED_PATH, index=False)
print(f"  Saved.")

df['LIKELY DEAL SCORE']   = pd.to_numeric(df['LIKELY DEAL SCORE'],   errors='coerce')
df['SCORE']               = pd.to_numeric(df['SCORE'],               errors='coerce')
for _col in ['MARKETING DM COUNT', 'MARKETING SMS COUNT']:
    df[_col] = pd.to_numeric(df[_col], errors='coerce').fillna(0) if _col in df.columns else 0
df['Month_dt']            = pd.to_datetime(df['Month'], format='%Y-%m')
df['ACTION_DAYS']         = df['ACTION PLANS'].apply(extract_days)
df['total_mkt']           = df['MARKETING DM COUNT'] + df['MARKETING SMS COUNT']

# Marketing count: max of (total DM across all months, total SMS across all months)
# e.g. 7 SMS + 4 DM → 7, not 11 — measures strongest channel, not combined
ff_dm  = df.groupby('FOLIO')['MARKETING DM COUNT'].sum()
ff_sms = df.groupby('FOLIO')['MARKETING SMS COUNT'].sum()
ff_mkt = pd.concat([ff_dm, ff_sms], axis=1).max(axis=1).rename('ff_mkt')

ff_agg = df.groupby('FOLIO').agg(
    ff_likely = ('LIKELY DEAL SCORE', 'max'),
    ff_score  = ('SCORE',             'max'),
    ff_days   = ('ACTION_DAYS',       'min'),
).join(ff_mkt, how='left')

distress_main_cols = ['MAIN DISTRESS #1', 'MAIN DISTRESS #2', 'MAIN DISTRESS #3', 'MAIN DISTRESS #4']
no_dist_re = re.compile(r'^No distress', re.IGNORECASE)
dist_long = (
    df[['FOLIO'] + distress_main_cols]
    .melt(id_vars='FOLIO', value_vars=distress_main_cols, value_name='dtype')
    .dropna(subset=['dtype'])
)
dist_long = dist_long[~dist_long['dtype'].str.match(no_dist_re, na=True)].copy()
dist_long['dtype'] = dist_long['dtype'].str.strip()
vac_num  = pd.to_numeric(df['VACANT'], errors='coerce').fillna(0)
vac_rows = pd.DataFrame({'FOLIO': df[vac_num > 0]['FOLIO'].unique(), 'dtype': 'Vacant'})
dist_long = pd.concat([dist_long, vac_rows], ignore_index=True)
# Inject Default Risk for fulfillment FOLIOs that appear in the DR file
ff_folios_set  = set(df['FOLIO'].dropna().astype(str).str.strip().unique())
dr_in_ff_folios = ff_folios_set & dr_folios
dr_ff_rows = pd.DataFrame({'FOLIO': list(dr_in_ff_folios), 'dtype': 'Default Risk'})
dist_long = pd.concat([dist_long, dr_ff_rows], ignore_index=True)
ff_distress_map = dist_long.groupby('FOLIO')['dtype'].apply(set).to_dict()

C_dist   = {dtype: int(dist_long[dist_long['dtype'] == dtype]['FOLIO'].nunique())
            for dtype in distress_order}
C_likely = count_by_bin(ff_agg['ff_likely'], likely_bins, likely_labels)
C_score  = count_by_bin(ff_agg['ff_score'],  score_bins,  score_labels)
C_action = count_by_action(ff_agg['ff_days'], action_map)
# All fulfillment properties count in mkt — treat 0 contacts as 1 (recommended at least once)
ff_mkt_adj = ff_agg['ff_mkt'].fillna(0).clip(lower=1)
C_mkt    = count_by_bin(ff_mkt_adj, mkt_bins, mkt_labels)
print(f"  Column C ready: {len(ff_agg):,} properties")


# ══════════════════════════════════════════════════════════════════════════════
# STEP 2: Columns F and G — fulfillment-matched leads/deals only (345 + 4)
# ══════════════════════════════════════════════════════════════════════════════
print("\nLoading Deals and Leads Freedom...")
dl = pd.read_excel(DEALS_PATH)
dl['LIKELY DEAL SCORE']   = pd.to_numeric(dl['LIKELY DEAL SCORE'],   errors='coerce')
dl['SCORE']               = pd.to_numeric(dl['SCORE'],               errors='coerce')
for _col in ['MARKETING DM COUNT', 'MARKETING SMS COUNT']:
    dl[_col] = pd.to_numeric(dl[_col] if _col in dl.columns else 0, errors='coerce').fillna(0)
dl['dl_days'] = dl['ACTION PLANS'].apply(extract_days)
dl['dl_mkt']  = dl['MARKETING DM COUNT'] + dl['MARKETING SMS COUNT']
dl['FOLIO_key'] = dl['FOLIO'].astype(str).str.strip()
dl.loc[dl['FOLIO_key'] == 'nan', 'FOLIO_key'] = np.nan

# ── Filter leads/deals to on or after CLIENT_START_DATE ──────────────────────
# Date filter per status:
#   Lead        → LEAD DATE
#   Appointment → APPOINTMENT DATE  (counted as Lead in report)
#   Deal        → LAST SALE DATE
#   Contract    → no date column → excluded
print(f"  Rows before date filter: {len(dl):,}")
dl['_date_col'] = pd.NaT
dl.loc[dl['PROPERTY STATUS'] == 'Lead',        '_date_col'] = pd.to_datetime(dl.loc[dl['PROPERTY STATUS'] == 'Lead',        'LEAD DATE'],        errors='coerce')
dl.loc[dl['PROPERTY STATUS'] == 'Appointment', '_date_col'] = pd.to_datetime(dl.loc[dl['PROPERTY STATUS'] == 'Appointment', 'APPOINTMENT DATE'], errors='coerce')
dl.loc[dl['PROPERTY STATUS'] == 'Deal',        '_date_col'] = pd.to_datetime(dl.loc[dl['PROPERTY STATUS'] == 'Deal',        'LAST SALE DATE'],   errors='coerce')
# Contracts have no date → _date_col stays NaT → excluded by filter below

dl = dl[dl['_date_col'] >= CLIENT_START_DATE].drop(columns=['_date_col'])

# Appointments and Contracts counted as Leads in the final report
dl.loc[dl['PROPERTY STATUS'] == 'Appointment', 'PROPERTY STATUS'] = 'Lead'

print(f"  Rows after date filter (>= {CLIENT_START_DATE.date()}): {len(dl):,}")
print(f"    Leads (incl. appointments): {(dl['PROPERTY STATUS'] == 'Lead').sum():,} | Deals: {(dl['PROPERTY STATUS'] == 'Deal').sum():,}")

ff_for_join = ff_agg.reset_index().rename(columns={'FOLIO': 'FOLIO_key'})
dl = dl.merge(ff_for_join, on='FOLIO_key', how='left')

BINARY_DIST_MAP = {
    'PRE-FORECLOSURE': 'Pre-foreclosure', 'VACANT': 'Vacant',
    'TAXES': 'Taxes (Tax Delinquent)', 'ESTATE': 'Estate (Pre-Probate)',
    'INTER FAMILY TRANSFER': 'Inter family transfer', 'PROBATE': 'Probate',
    'JUDGEMENT': 'Judgement', 'LIENS CITY/COUNTY': 'Liens city/county',
    'LIENS OTHER': 'Liens other', 'LIENS MECHANIC': 'Liens Mechanic',
    'HIGH EQUITY': 'High Equity', '55+': '55+ (Senior)',
    'ABSENTEE': 'Absentee', 'DOWNSIZING': 'Downsizing',
}

def get_binary_distress(row):
    types = set()
    for col, label in BINARY_DIST_MAP.items():
        if col in row.index:
            val = pd.to_numeric(row[col], errors='coerce')
            if not pd.isna(val) and val > 0:
                types.add(label)
    return types

def resolve_row(row):
    in_ff = pd.notna(row.get('ff_likely'))
    if in_ff:
        return pd.Series({
            'Data_Source': 'Fulfillment',
            'r_likely':    row['ff_likely'],
            'r_score':     row['ff_score'],
            'r_days':      row['ff_days'],
            'r_mkt':       max(row['ff_mkt'], 1) if pd.notna(row.get('ff_mkt')) else 1,
            'r_distress':  ff_distress_map.get(str(row['FOLIO_key']).strip(), set()),
        })
    else:
        return pd.Series({
            'Data_Source': 'Deals and Leads Atlas',
            'r_likely':    row['LIKELY DEAL SCORE'],
            'r_score':     row['SCORE'],
            'r_days':      row['dl_days'],
            'r_mkt':       row['dl_mkt'],
            'r_distress':  get_binary_distress(row),
        })

resolved = dl.apply(resolve_row, axis=1)
dl = pd.concat([dl, resolved], axis=1)

def compute_col_counts(sub_df):
    likely_d = count_by_bin(sub_df['r_likely'], likely_bins, likely_labels)
    score_d  = count_by_bin(sub_df['r_score'],  score_bins,  score_labels)
    action_d = count_by_action(sub_df['r_days'], action_map)
    mkt_d    = count_by_bin(sub_df['r_mkt'],    mkt_bins,   mkt_labels)
    dist_d   = {dtype: int(
        sub_df['r_distress'].apply(lambda s: dtype in s if isinstance(s, set) else False).sum()
    ) for dtype in distress_order}
    return likely_d, score_d, action_d, mkt_d, dist_d

leads_ff = dl[(dl['PROPERTY STATUS'] == 'Lead') & (dl['Data_Source'] == 'Fulfillment')].copy()
deals_ff = dl[(dl['PROPERTY STATUS'] == 'Deal') & (dl['Data_Source'] == 'Fulfillment')].copy()
F_likely, F_score, F_action, F_mkt, F_dist = compute_col_counts(leads_ff)
G_likely, G_score, G_action, G_mkt, G_dist = compute_col_counts(deals_ff)
print(f"  Columns F/G ready — Leads: {len(leads_ff)} | Deals: {len(deals_ff)}")


# ══════════════════════════════════════════════════════════════════════════════
# STEP 3: Load domain → Column B + prep for D and E
# ══════════════════════════════════════════════════════════════════════════════
print("\nLoading domain files...")
_parquet = DOMAIN_DIR / "domain.parquet"
import os as _os
if _os.path.exists(_parquet):
    dom = pd.read_parquet(_parquet)
    print(f"  Loaded from parquet: {len(dom):,} rows")
else:
    print("  Parquet not found — reading xlsx and converting (first run is slower)...")
    _xlsx = sorted(f for f in DOMAIN_DIR.glob("*.xlsx") if not f.name.startswith("~$"))
    dom   = pd.concat([pd.read_excel(f) for f in _xlsx], ignore_index=True)
    for _c in dom.select_dtypes(include='object').columns:
        dom[_c] = dom[_c].astype(str)
    dom.to_parquet(_parquet, index=False)
    print(f"  Saved parquet cache: {_parquet.name}")
print(f"  {len(dom):,} rows, {dom['FOLIO'].nunique():,} unique FOLIOs")

dom['LIKELY DEAL SCORE']   = pd.to_numeric(dom['LIKELY DEAL SCORE'],   errors='coerce')
dom['SCORE']               = pd.to_numeric(dom['SCORE'],               errors='coerce')
for _col in ['MARKETING DM COUNT', 'MARKETING SMS COUNT']:
    dom[_col] = pd.to_numeric(dom[_col] if _col in dom.columns else 0, errors='coerce').fillna(0)
dom['total_mkt']           = dom[['MARKETING DM COUNT', 'MARKETING SMS COUNT']].max(axis=1)
dom['ACTION_DAYS']         = dom['ACTION PLANS'].apply(extract_days)
dom['LAST SALE DATE']      = pd.to_datetime(dom['LAST SALE DATE'], errors='coerce')
dom['buybox_int']          = pd.to_numeric(dom['PROPERTY ID (BUYBOX)'], errors='coerce')

# ── Owner Occupied cross-reference for Columns C / F / G ─────────────────────
# Fulfillment distress data has no Owner Occupied signal; derive it from
# domain ABSENTEE (0 = owner-occupied) by joining on FOLIO.
print("\nPatching Owner Occupied for Columns C/F/G via domain ABSENTEE...")
_dom_abs = (dom[['FOLIO', 'ABSENTEE']]
            .assign(FOLIO=dom['FOLIO'].astype(str).str.strip(),
                    ABSENTEE=pd.to_numeric(dom['ABSENTEE'], errors='coerce'))
            .drop_duplicates('FOLIO')
            .set_index('FOLIO')['ABSENTEE'])

ff_folio_idx = ff_agg.index.astype(str).str.strip()
C_dist['Owner Occupied'] = int((_dom_abs.reindex(ff_folio_idx) == 0).sum())

leads_ff_idx = leads_ff['FOLIO_key'].dropna().astype(str).str.strip().unique()
F_dist['Owner Occupied'] = int((_dom_abs.reindex(leads_ff_idx) == 0).sum())

deals_ff_idx = deals_ff['FOLIO_key'].dropna().astype(str).str.strip().unique()
G_dist['Owner Occupied'] = int((_dom_abs.reindex(deals_ff_idx) == 0).sum())
print(f"  Owner Occupied — C: {C_dist['Owner Occupied']:,} | F: {F_dist['Owner Occupied']:,} | G: {G_dist['Owner Occupied']:,}")

# Column B — filter to BUYBOX SCORE > 0
dom['BUYBOX SCORE'] = pd.to_numeric(dom['BUYBOX SCORE'], errors='coerce').fillna(0)
dom_b = dom[dom['BUYBOX SCORE'] > 0].copy()
print(f"  Domain with BUYBOX SCORE > 0: {len(dom_b):,}")

B_likely = count_by_bin(dom_b['LIKELY DEAL SCORE'], likely_bins, likely_labels)
B_score  = count_by_bin(dom_b['SCORE'],             score_bins,  score_labels)
B_action = count_by_action(dom_b['ACTION_DAYS'],    action_map)
B_mkt    = count_by_bin(dom_b[dom_b['total_mkt'] > 0]['total_mkt'], mkt_bins, mkt_labels)
B_dist   = domain_distress_counts(dom_b)
print(f"  Column B ready — total: {sum(B_likely.values()):,}")


# ══════════════════════════════════════════════════════════════════════════════
# STEP 4: Column D — Sold since Oct 1 2025
# ══════════════════════════════════════════════════════════════════════════════
print("\nComputing Column D (sold since Oct 1 2025, fulfillment only)...")
sold_all = dom[dom['LAST SALE DATE'] >= '2025-10-01'].copy()
ff_join  = ff_agg[['ff_likely', 'ff_score', 'ff_days', 'ff_mkt']].reset_index()
sold_all = sold_all.merge(ff_join, on='FOLIO', how='left')
# Restrict to fulfillment-recommended properties only
sold = sold_all[sold_all['ff_likely'].notna()].copy()
print(f"  Sold since Oct 1: {len(sold_all):,} total | In fulfillment: {len(sold):,}")

sold['d_likely'] = sold['ff_likely']
sold['d_score']  = sold['ff_score']
sold['d_days']   = sold['ff_days']
sold['d_mkt']    = sold['ff_mkt'].fillna(0).clip(lower=1)

D_likely = count_by_bin(sold['d_likely'], likely_bins, likely_labels)
D_score  = count_by_bin(sold['d_score'],  score_bins,  score_labels)
D_action = count_by_action(sold['d_days'], action_map)
D_mkt    = count_by_bin(sold['d_mkt'], mkt_bins, mkt_labels)
D_dist   = domain_distress_counts(sold)
print(f"  Column D ready — total: {sum(D_likely.values()):,}")


# ══════════════════════════════════════════════════════════════════════════════
# STEP 5: Column E — Market Deals (overlap: market deals ∩ domain sold, 1,367)
# Same score resolution as Column D: prefer fulfillment (pre-sale), else domain
# NOTE: Temporarily disabled — uncomment to re-enable Market Deals column
# ══════════════════════════════════════════════════════════════════════════════
# print("\nComputing Column E (market deals x fulfillment)...")
# md = pd.read_excel(MARKET_PATH)
# md_ids = set(md['PropertyID'].dropna().astype(int).unique())
#
# # Overlap: market deal BUYBOX IDs within the fulfillment-sold subset
# ff_sold_buybox_ids = set(sold['buybox_int'].dropna().astype(int).unique())
# overlap_ids = md_ids & ff_sold_buybox_ids
# print(f"  Market deals unique: {len(md_ids):,} | Fulfillment-sold: {len(ff_sold_buybox_ids):,} | Overlap: {len(overlap_ids):,}")
#
# # Filter fulfillment-sold to market deal properties
# mkt_sold = sold[sold['buybox_int'].isin(overlap_ids)].copy()
#
# E_likely = count_by_bin(mkt_sold['d_likely'], likely_bins, likely_labels)
# E_score  = count_by_bin(mkt_sold['d_score'],  score_bins,  score_labels)
# E_action = count_by_action(mkt_sold['d_days'], action_map)
# E_mkt    = count_by_bin(mkt_sold['d_mkt'], mkt_bins, mkt_labels)
# E_dist   = domain_distress_counts(mkt_sold)
# print(f"  Column E ready — total: {sum(E_likely.values()):,}")
#
# print(f"\n=== COLUMN E (market deals) ===")
# print("Likely Deal Score:", E_likely)
# print("Action Plan:", E_action)


# ══════════════════════════════════════════════════════════════════════════════
# STEP 6: Build report rows
# ══════════════════════════════════════════════════════════════════════════════
def build_rows(B_lik, B_sco, B_act, B_mkt, B_dis,
               C_lik, C_sco, C_act, C_mkt, C_dis,
               D_lik, D_sco, D_act, D_mkt, D_dis,
               # E_lik, E_sco, E_act, E_mkt, E_dis,  # Market Deals — disabled
               F_lik, F_sco, F_act, F_mkt, F_dis,
               G_lik, G_sco, G_act, G_mkt, G_dis):
    n = len(COLS)

    def section_header(name):
        r = list(COLS); r[0] = name; return r

    def data_row(label, b, c, d, f, g):
        r = [''] * n
        r[0] = label
        r[1] = b    # B — Total Properties
        r[2] = c    # C — Properties in the fulfillment
        r[3] = d    # D — Sold
        # r[4] = e  # E — Market Deals (disabled)
        r[4] = f    # E (Excel) — Clients Lead
        r[5] = g    # F (Excel) — Client Deals
        # G–I left empty here; formulas written during Excel rendering
        return r

    def blank():
        return [''] * n

    rows = []

    rows.append(section_header('Likely Deal Score'))
    for lbl in ['100-81', '80-61', '60-41', '40-21', '20-0']:
        rows.append(data_row(lbl, B_lik[lbl], C_lik[lbl], D_lik[lbl], F_lik[lbl], G_lik[lbl]))
    rows.append(data_row('Total', sum(B_lik.values()), sum(C_lik.values()), sum(D_lik.values()), sum(F_lik.values()), sum(G_lik.values())))
    rows.append(blank())

    rows.append(section_header('Total Score'))
    for lbl in ['1000-801', '800-601', '600-401', '400-201', '200-0']:
        rows.append(data_row(lbl, B_sco[lbl], C_sco[lbl], D_sco[lbl], F_sco[lbl], G_sco[lbl]))
    rows.append(data_row('Total', sum(B_sco.values()), sum(C_sco.values()), sum(D_sco.values()), sum(F_sco.values()), sum(G_sco.values())))
    rows.append(blank())

    rows.append(section_header('Action Plan'))
    for lbl in ['30 day', '60 day', '90 day']:
        rows.append(data_row(lbl, B_act[lbl], C_act[lbl], D_act[lbl], F_act[lbl], G_act[lbl]))
    rows.append(data_row('Total', sum(B_act.values()), sum(C_act.values()), sum(D_act.values()), sum(F_act.values()), sum(G_act.values())))
    rows.append(blank())

    rows.append(section_header('Mkt Count'))
    for lbl in ['1 to 5', '6 to 10', '11 to 19', '20+']:
        rows.append(data_row(lbl, B_mkt[lbl], C_mkt[lbl], D_mkt[lbl], F_mkt[lbl], G_mkt[lbl]))
    rows.append(data_row('Total', sum(B_mkt.values()), sum(C_mkt.values()), sum(D_mkt.values()), sum(F_mkt.values()), sum(G_mkt.values())))
    rows.append(blank())

    rows.append(section_header('Distress'))
    for dtype in distress_order:
        rows.append(data_row(dtype, B_dis[dtype], C_dis[dtype], D_dis[dtype], F_dis[dtype], G_dis[dtype]))
    rows.append(data_row('Total', sum(B_dis.values()), sum(C_dis.values()), sum(D_dis.values()), sum(F_dis.values()), sum(G_dis.values())))

    return rows


# ══════════════════════════════════════════════════════════════════════════════
# STEP 7: Write Excel with formulas for H–K
# ══════════════════════════════════════════════════════════════════════════════
def write_excel(path, rows, sheet_title):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title

    # Styles
    header_font   = Font(name='Helvetica', bold=True, size=10, color='FFFFFF')
    header_fill   = PatternFill('solid', fgColor='0B5394')
    header_align  = Alignment(horizontal='center', vertical='center', wrap_text=True)
    section_font  = Font(name='Georgia', bold=True, size=10)
    section_fill  = PatternFill('solid', fgColor='D9E2F3')
    section_align = Alignment(horizontal='left', vertical='center')
    total_font    = Font(name='Helvetica', bold=True, size=10)
    total_fill    = PatternFill('solid', fgColor='F2F2F2')
    data_font     = Font(name='Helvetica', size=10)
    count_font    = Font(name='Helvetica', bold=True, size=10, color='0B5394')
    pct_font      = Font(name='Helvetica', size=10, color='444444')
    pct_font_bold = Font(name='Helvetica', bold=True, size=10, color='444444')
    center_align  = Alignment(horizontal='center', vertical='center')
    left_align    = Alignment(horizontal='left',   vertical='center')
    thin_side     = Side(style='thin', color='CCCCCC')
    thin          = Border(left=thin_side, right=thin_side, bottom=thin_side, top=thin_side)

    # Header row
    ws.append(COLS)
    for cell in ws[1]:
        cell.font = header_font; cell.fill = header_fill
        cell.alignment = header_align; cell.border = thin
    ws.row_dimensions[1].height = 36

    # Header comments — one per column explaining the calculation
    header_comments = {
        'A': "Category\nSection headers (Likely Deal Score, Total Score, Action Plan, Mkt Count, Distress) and their range labels.",
        'B': "Total Properties\nAll domain properties with BUYBOX SCORE > 0.\nLikely/Total Score: all scored. Action Plan: only properties with an action plan. Mkt Count: only properties with mkt contacts > 0. Distress: binary columns including DEFAULT RISK (now in domain file).",
        'C': "Properties in the Fulfillment\n37,338 unique FOLIOs across 5 fulfillment files (Oct–Dec 2025).\nScores: MAX across all months. Action Plan: best (lowest days). Mkt Count: latest month DM+SMS; properties with 0 contacts counted as 1 (recommended at least once). Distress: MAIN DISTRESS #1–4 + binary VACANT.",
        'D': "Sold\nProperties sold since Oct 1 2025 that were also recommended in a fulfillment (390).\nScores come from fulfillment MAX (pre-sale picture). Distress from domain binary columns.",
        # 'E': "Market Deals — disabled (see STEP 5)",
        'E': "Clients Lead\n345 leads from Deals & Leads Atlas matched to the fulfillment via FOLIO.\nScores: fulfillment MAX. Mkt: 0 contacts treated as 1. 503 leads not in fulfillment are excluded.",
        'F': "Client Deals\n4 deals from Deals & Leads Atlas matched to the fulfillment via FOLIO.\nScores: fulfillment MAX. 147 deals not in fulfillment are excluded.",
        'G': "Sold Concentration %\nFormula: =D/C\nWhat % of fulfillment-recommended properties were sold since Oct 1.",
        # 'H': "Sold to Investors Concentration — disabled (depends on Market Deals)",
        'H': "Client Deals Concentration\nFormula: =F/C\nWhat % of fulfillment-recommended properties became a client deal.",
        'I': "Client Leads Concentration\nFormula: =E/C\nWhat % of fulfillment-recommended properties became a client lead.",
    }
    for col_letter, text in header_comments.items():
        cell = ws[f'{col_letter}1']
        cell.comment = Comment(text, "Atlas Report")


    # Count columns (0-based): B=1, C=2, D=3, E(Leads)=4, F(Deals)=5
    COUNT_COLS = [1, 2, 3, 4, 5]
    # Formula columns (0-based): G=6, H=7, I=8
    # Formulas reference 1-based Excel columns:
    #   G = D/C  →  =IFERROR(D{r}/C{r},"")   Sold Concentration
    #   H = F/C  →  =IFERROR(F{r}/C{r},"")   Client Deals Concentration
    #   I = E/C  →  =IFERROR(E{r}/C{r},"")   Client Leads Concentration
    #   [Market Deals col E and Sold to Investors col I disabled — see STEP 5]

    for row_data in rows:
        ws.append(row_data)
        r = ws.max_row
        label = row_data[0]

        if label in SECTION_NAMES:
            for cell in ws[r]:
                cell.font = section_font; cell.fill = section_fill
                cell.alignment = section_align; cell.border = thin

        elif label == '':
            pass  # blank separator

        else:
            # Data row or Total row
            is_total = (label == 'Total')
            base_font  = total_font if is_total else data_font
            base_fill  = total_fill if is_total else None
            c_font     = Font(name='Helvetica', bold=True, size=10, color='0B5394') if is_total else count_font
            p_font     = pct_font_bold if is_total else pct_font

            for cell in ws[r]:
                cell.font      = base_font
                cell.alignment = left_align
                cell.border    = thin
                if base_fill:
                    cell.fill = base_fill

            # Style count cells B–G
            for ci in COUNT_COLS:
                ws[r][ci].font      = c_font
                ws[r][ci].alignment = center_align

            # Write formulas for G–I and format as percentage
            ws[r][6].value      = f'=IFERROR(D{r}/C{r},"")'   # G: Sold / Fulfillment
            ws[r][7].value      = f'=IFERROR(F{r}/C{r},"")'   # H: Client Deals / Fulfillment
            ws[r][8].value      = f'=IFERROR(E{r}/C{r},"")'   # I: Client Leads / Fulfillment
            # ws[r][?].value = f'=IFERROR(E{r}/C{r},"")' — Market Deals disabled (see STEP 5)

            pct_fmt = '0.0%'
            for ci in [6, 7, 8]:
                ws[r][ci].number_format = pct_fmt
                ws[r][ci].font          = p_font
                ws[r][ci].alignment     = center_align
                ws[r][ci].border        = thin

    # Column widths
    ws.column_dimensions['A'].width = 32
    for col_letter in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']:
        ws.column_dimensions[col_letter].width = 22
    ws.freeze_panes = 'A2'

    wb.save(path)
    print(f"  Saved: {path}")


rows = build_rows(
    B_likely, B_score, B_action, B_mkt, B_dist,
    C_likely, C_score, C_action, C_mkt, C_dist,
    D_likely, D_score, D_action, D_mkt, D_dist,
    # E_likely, E_score, E_action, E_mkt, E_dist,  # Market Deals — disabled (see STEP 5)
    F_likely, F_score, F_action, F_mkt, F_dist,
    G_likely, G_score, G_action, G_mkt, G_dist,
)

print(f"\nWriting report...")
write_excel(OUTPUT, rows, "Max Logic - BCDEFG")
print("\nAll done.")
