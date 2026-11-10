"""
IFRS 17 Disclosure Analysis - IRA Kenya 2023
=====================================================
Reproduces the descriptive tables of:
"IFRS 17 as Supervisory Data Infrastructure: Evidence from Kenya and a
Proposed AI-Ready SupTech Architecture"
Author: Jeniffer Nasike Atetwe | ORCID: 0009-0009-0309-5045

All variables are computed from the visible, published sheets of:
  - IRA_Kenya_Annual_Statistics_2023.xlsx   (Appendix 1, Appendix 5 I-IV, Details)
  - IRA_Kenya_Annual_Statistics_2024.xlsx   (Appendix 1)
No hidden or protected sheets are read. No Capital Adequacy Ratio data are used.

Usage:
    pip install -r requirements.txt
    python ifrs17_analysis.py

Variable construction:
  IDCS  = IFRS 17 Disclosure Completeness Score (author-defined, 5 binary criteria)
          computed from Appendix 1 field presence and internal reconciliation.
  ISDI  = Income Statement Disclosure Depth Index
          0-4 ordinal: how far down the Appendix 1 income statement is populated.
  rev_growth = Insurance Revenue growth 2023 to 2024
  service_expense_ratio = Insurance Service Expense / Insurance Revenue
  reins_ratio = |Net Reinsurance Expense| / Insurance Revenue
  ln(assets) = log(Total Assets) from Appendix 5
"""

import os, sys, warnings
import numpy as np
import pandas as pd
import openpyxl
from datetime import date

warnings.filterwarnings('ignore')

FILE_2023 = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'IRA_Kenya_Annual_Statistics_2023.xlsx')
FILE_2024 = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'IRA_Kenya_Annual_Statistics_2024.xlsx')

for f in [FILE_2023, FILE_2024]:
    if not os.path.exists(f):
        sys.exit(f"ERROR: {f} not found in current folder.")

def sf(x):
    try: return float(x) if x is not None else 0.0
    except: return 0.0

def safe_log(x):
    return np.log(x) if x and x > 0 else np.nan

def clean(s):
    if not isinstance(s, str): return ''
    s = s.strip().upper()
    for w in ['LIMITED','LTD','COMPANY','COMPAN','INSURANE','INSURANCE',
              'GENERAL','ASSURANCE',' KENYA','(K)','CO.']:
        s = s.replace(w, '')
    return ' '.join(s.split())

print("=" * 65)
print("IFRS 17 Disclosure Analysis - IRA Kenya 2023")
print("=" * 65)

# ── LOAD 2023 DATA ────────────────────────────────────────────────
print(f"\nLoading {FILE_2023}...")
wb23 = openpyxl.load_workbook(FILE_2023, read_only=True, data_only=True)

# Appendix 1 — IFRS 17 P&L
rows_a1 = list(wb23['APPENDIX 1'].iter_rows(values_only=True))
for i, r in enumerate(rows_a1):
    if r[1] == 'Company': hdr = i; break

app1 = {}
for r in rows_a1[hdr+2:]:
    n = r[1]
    if not n or not isinstance(n, str): continue
    n = n.strip().upper()
    if n == 'REINSURERS': break
    if n in ('INSURERS','TOTAL','GRAND TOTAL',''): continue
    app1[n] = {
        'ins_rev': sf(r[2]),
        'ise':     sf(r[3]),
        'nre':     sf(r[4]),
        'isr':     sf(r[5]),
        'net_fin': sf(r[9]),
        'fin_exp': sf(r[7]),
        'fin_inc': sf(r[8]),
        'pbt':     sf(r[15]),
        'nii':     sf(r[6]),
        'pat':     sf(r[17]),
    }

# Total Assets (Appendix 5 I/II/III/IV — transposed layout)
def load_total_assets(wb):
    result = {}
    for sname in wb.sheetnames:
        if 'APPENDIX 5' not in sname.upper(): continue
        try:
            ws = wb[sname]
            rows = list(ws.iter_rows(values_only=True))
            cnames = None
            for r in rows:
                if r[1] and str(r[1]).strip() == 'Company':
                    cnames = [str(v).strip().upper() if v else '' for v in r[2:]]
                    break
            if not cnames: continue
            for r in rows:
                lbl = str(r[1]).strip().lower() if r[1] else ''
                if 'total assets' in lbl:
                    for c, v in zip(cnames, r[2:]):
                        if c:
                            val = sf(v)
                            if val > 0: result[c] = val
        except: continue
    return result

ta23 = load_total_assets(wb23)
wb23.close()

# ── LOAD 2024 INSURANCE REVENUE (for Insurance Revenue growth) ─────────────────
print(f"Loading {FILE_2024}...")
wb24 = openpyxl.load_workbook(FILE_2024, read_only=True, data_only=True)
rows_a1_24 = list(wb24['APPENDIX 1'].iter_rows(values_only=True))
for i, r in enumerate(rows_a1_24):
    if r[1] == 'Company': hdr24 = i; break

insrev_24 = {}
for r in rows_a1_24[hdr24+2:]:
    n = r[1]
    if not n or not isinstance(n, str): continue
    n = n.strip().upper()
    if n == 'REINSURERS': break
    if n in ('INSURERS','TOTAL','GRAND TOTAL',''): continue
    v = sf(r[2])
    if v != 0: insrev_24[n] = v
wb24.close()

# Name mapping: 2024 names → 2023 names
name_map_24_23 = {
    'AFRICA MERCHANT ASSURANCE': 'AFRICAN MERCHANT ASSURANCE',
    'APA INSURANCE LIMITED':     'APA INSURANCE COMPANY',
    'INTRA AFRICA ASSURANCE':    'INTRA-AFRICA ASSURANCE',
}

# ── ESTIMATED PUBLICATION-LAG PROXY (assumed date) ───────────────
# IRA Kenya 2023 Annual Statistics published August 2024
# Year-end: 31 December 2023
# All 40 insurers share this publication date (one aggregate report)
# Proxy = days from 31 Dec 2023 to an ASSUMED 31 Aug 2024 (source gives month only)
year_end   = date(2023, 12, 31)
pub_date   = date(2024, 8, 31)   # "August 2024" per Details sheet
lag_days   = (pub_date - year_end).days   # 244 days
lag_log    = np.log(lag_days)
print(f"  Estimated publication-lag proxy: {lag_days} days (31 Dec 2023 to assumed 31 Aug 2024)")

# ── BUILD DATASET ─────────────────────────────────────────────────
print("\nBuilding 2023 cross-section dataset...")
records = []
for a1_name, d in app1.items():
    cn = a1_name
    ins_rev = d['ins_rev']
    if ins_rev == 0: continue   # inactive insurer

    # ── IDCS (5 binary criteria) ──────────────────────────────────
    c1 = 1 if d['ins_rev'] != 0 else 0
    c2 = 1 if d['ise'] != 0 else 0
    c3 = 1 if d['nre'] != 0 else 0
    c4 = 1 if d['isr'] != 0 else 0
    if ins_rev != 0:
        computed_isr = ins_rev - d['ise'] - d['nre']
        c5 = 1 if abs(computed_isr - d['isr']) / abs(ins_rev) < 0.01 else 0
    else:
        c5 = 0
    IDCS = (c1 + c2 + c3 + c4 + c5) / 5

    # ── ISDI (0-4 ordinal, /4 normalised) ────────────────────────
    if ins_rev == 0:
        isdi_raw = 0
    elif d['pbt'] != 0:
        isdi_raw = 4
    elif d['fin_exp'] != 0 or d['fin_inc'] != 0:
        isdi_raw = 3
    elif d['net_fin'] != 0:
        isdi_raw = 2
    else:
        isdi_raw = 1
    ISDI = isdi_raw / 4

    # ── Insurance Revenue growth (InsRev 2023→2024) ────────────────────────────
    a1_clean = clean(a1_name)
    rev24 = insrev_24.get(a1_name)
    if rev24 is None:
        for n24, v24 in insrev_24.items():
            n23 = name_map_24_23.get(n24, n24)
            if n23 == a1_name:
                rev24 = v24; break
    if rev24 is None:
        for n24, v24 in insrev_24.items():
            if clean(n24) == a1_clean:
                rev24 = v24; break
    rev_growth_val = (rev24 - ins_rev) / ins_rev if rev24 and ins_rev > 0 else np.nan

    # ── Derived financial variables ───────────────────────────────
    service_expense_ratio = d['ise'] / ins_rev if ins_rev > 0 else np.nan
    reins_ratio    = abs(d['nre']) / ins_rev if ins_rev > 0 else np.nan

    # Total assets
    ta = ta23.get(a1_name)
    if ta is None:
        for k, v in ta23.items():
            if clean(k) == a1_clean:
                ta = v; break
    ln_assets = safe_log(ta)

    records.append({
        'insurer':        cn,
        'IDCS':           IDCS,
        'ISDI':           ISDI,
        'reporting_lag':  lag_log,
        'rev_growth':     rev_growth_val,
        'service_expense_ratio': service_expense_ratio,
        'reins_ratio':    reins_ratio,
        'ln_assets':      ln_assets,
        'total_assets':   ta,
        'ins_rev':        ins_rev,
        'ise': d['ise'], 'nre': d['nre'], 'isr': d['isr'], 'pbt': d['pbt'], 'nii': d['nii'], 'pat': d['pat'],
    })

df = pd.DataFrame(records)
print(f"  Total insurers matched: {len(df)}")
print(f"  Missing Insurance Revenue growth: {df['rev_growth'].isna().sum()}")
print(f"  Missing ln_assets: {df['ln_assets'].isna().sum()}")

df.to_csv('ifrs17_dataset_2023.csv', index=False)
print(f"  Dataset saved: ifrs17_dataset_2023.csv")

# -- DESCRIPTIVE STATISTICS ------------------------------------------
print(f"  Appendix 1 insurer entries: {len(app1)}  |  active (Insurance Revenue != 0): {len(df)}")
print("\n" + "=" * 65)
print("TABLE 4 - DESCRIPTIVE STATISTICS (n = all active insurers; distributions = Tables 5 and 6)")
print("=" * 65)
desc_cols = ['IDCS','ISDI','rev_growth','service_expense_ratio','reins_ratio','ln_assets']
print(df[desc_cols].describe().round(3).to_string())

print(f"\nIDCS distribution:")
from collections import Counter
print(f"  {sorted(Counter(df['IDCS'].round(2)).items())}")
print(f"\nISDI distribution (normalised):")
print(f"  {sorted(Counter(df['ISDI'].round(2)).items())}")
print(f"\nIDCS < 1 insurers: {list(df.loc[df['IDCS']<1,'insurer'])}")

# -- CORRELATION MATRIX (variables with variation only) ---------------
print("\n" + "=" * 65)
print("TABLE 12 - PAIRWISE CORRELATION MATRIX (financial variables with variation)")
print("=" * 65)
corr_cols = ['ln_assets','rev_growth','service_expense_ratio','reins_ratio']
corr_df = df[corr_cols].dropna()
print(corr_df.corr().round(3).to_string())
print(f"\n  n = {len(corr_df)} (complete cases)")
print("  IDCS and ISDI are (near-)constant in the sample and are excluded.")

# -- TABLE 11 / TABLE 13: group comparison (8 largest vs remaining by 2023 Insurance Revenue)
print("\n" + "=" * 65)
print("TABLE 11 - Income statement aggregates (KES thousand), by group")
print("=" * 65)
dfs = df.sort_values('ins_rev', ascending=False).reset_index(drop=True)
groups = {'8 largest': dfs.iloc[:8], f'{len(dfs)-8} remaining': dfs.iloc[8:], f'All {len(dfs)}': dfs}
for g, x in groups.items():
    rev, ise, nre, isr = x.ins_rev.sum(), x.ise.sum(), x.nre.sum(), x.isr.sum()
    print(f"  {g:<13} NII={x.nii.sum():,.0f} PBT={x.pbt.sum():,.0f} PAT={x.pat.sum():,.0f}")
    print(f"  {g:<13} Rev={rev:,.0f} ISE={ise:,.0f} NRE={nre:,.0f} ISR={isr:,.0f} "
          f"ISR margin={isr/rev*100:.1f}% residual(Rev-ISE-NRE-ISR)={rev-ise-nre-isr:,.3f}")
print("\nTABLE 13 - Group means")
for g, x in groups.items():
    print(f"  {g:<13} n={len(x)} IDCS={x.IDCS.mean():.3f} ISDI={x.ISDI.mean():.3f} "
          f"ln_assets={x.ln_assets.mean():.3f} (n_assets={x.ln_assets.count()}) "
          f"service_expense_ratio={x.service_expense_ratio.mean():.3f} reins_ratio={x.reins_ratio.mean():.3f} "
          f"rev_growth={x.rev_growth.mean():.3f} (n={x.rev_growth.count()})")

print("=" * 65)
print("SCRIPT COMPLETE")
print("=" * 65)
