"""
IFRS 17 / CAR Regression Analysis — IRA Kenya 2023
=====================================================
Replicates Table 14 of:
"IFRS 17 as Supervisory Data Infrastructure: Toward an AI-Ready
Insurance Supervision Architecture in Kenya"
Author: Jeniffer Nasike Atetwe | ORCID: 0009-0009-0309-5045

All variables computed automatically from:
  - IRA_Kenya_Annual_Statistics_2023.xlsx
  - IRA_Kenya_Annual_Statistics_2024.xlsx

Usage:
    pip install pandas openpyxl statsmodels numpy scipy
    python ifrs17_regression_final.py

Notes on variable construction:
  IDCS  = IFRS 17 Disclosure Completeness Score (author-defined, 5 binary criteria)
           Computed from Appendix 1 field presence and internal reconciliation.
  RAGI  = Risk Adjustment Granularity Index (author-defined, 0-4 ordinal)
           Computed from income statement disclosure depth in Appendix 1.
  GEP growth = Insurance Revenue growth 2023 to 2024 (proxy for insurer growth)
  Reporting lag = Days from 31 Dec 2023 to IRA annual report publication (Aug 2024)
                  All insurers share the same lag (one aggregate publication).
                  Variation comes from within-insurer filing timing differences.
  Combined ratio = ISE / Insurance Revenue
  Reinsurance ratio = |NRE| / Insurance Revenue
  ln(assets) = log(Total Assets) from Appendix 5
"""

import os, sys, warnings
import numpy as np
import pandas as pd
import openpyxl
import statsmodels.api as sm
from scipy import stats
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
print("IFRS 17 / CAR Regression — IRA Kenya 2023")
print("=" * 65)

# ── LOAD 2023 DATA ────────────────────────────────────────────────
print(f"\nLoading {FILE_2023}...")
wb23 = openpyxl.load_workbook(FILE_2023, read_only=True, data_only=True)

# CAR (Appendix CAR 2 = General Insurers)
car23 = {}
for row in wb23['CAR 2'].iter_rows(min_row=5, values_only=True):
    n, car = row[1], row[4]
    if not n or not isinstance(n, str): continue
    n = n.strip().upper()
    if n == 'REINSURERS': break
    if n == 'INSURERS': continue
    v = sf(car)
    if v != 0: car23[n] = v

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

# ── LOAD 2024 INSURANCE REVENUE (for GEP growth) ─────────────────
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

# Name mapping: CAR 2 names → Appendix 1 names
name_map_car = {
    'HERITAGE INSURANCE COMPANY':       'THE HERITAGE INSURANCE COMPANY',
    'JUBILEE GENERAL INSURANCE':        'JUBILEE ALLIANZ GENERAL INSURANCE',
    'MADISON GENERAL INSURANCE COMPANY':'MADISON INSURANCE COMPANY',
    'PIONEER GENERAL INSURANCE COMPANY':'PIONEER INSURANCE COMPANY',
    'SANLAM INSURANE COMPANY':          'SANLAM INSURANCE COMPANY',
    'UAP INSURANCE COMPANY':            'OLD MUTUAL GENERAL INSURANCE',
    'METROPOLITAN CANNON GENERAL':      'CANNON GENERAL INSURANCE (K) LIMITED',
}

# ── REPORTING LAG ─────────────────────────────────────────────────
# IRA Kenya 2023 Annual Statistics published August 2024
# Year-end: 31 December 2023
# All 40 insurers share this publication date (one aggregate report)
# Lag = days from 31 Dec 2023 to 31 Aug 2024
year_end   = date(2023, 12, 31)
pub_date   = date(2024, 8, 31)   # "August 2024" per Details sheet
lag_days   = (pub_date - year_end).days   # 244 days
lag_log    = np.log(lag_days)
print(f"  Reporting lag: {lag_days} days (31 Dec 2023 → Aug 2024)")

# ── BUILD DATASET ─────────────────────────────────────────────────
print("\nBuilding 2023 cross-section dataset...")
records = []
for car_name, car_val in car23.items():
    cn = car_name.strip()
    a1_name = cn if cn in app1 else name_map_car.get(cn)
    if not a1_name or a1_name not in app1:
        continue

    d = app1[a1_name]
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

    # ── RAGI (0-4 ordinal, /4 normalised) ────────────────────────
    if ins_rev == 0:
        ragi_raw = 0
    elif d['pbt'] != 0:
        ragi_raw = 4
    elif d['fin_exp'] != 0 or d['fin_inc'] != 0:
        ragi_raw = 3
    elif d['net_fin'] != 0:
        ragi_raw = 2
    else:
        ragi_raw = 1
    RAGI = ragi_raw / 4

    # ── GEP growth (InsRev 2023→2024) ────────────────────────────
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
    gep_growth = (rev24 - ins_rev) / ins_rev if rev24 and ins_rev > 0 else np.nan

    # ── Derived financial variables ───────────────────────────────
    combined_ratio = d['ise'] / ins_rev if ins_rev > 0 else np.nan
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
        'CAR':            car_val,
        'IDCS':           IDCS,
        'RAGI':           RAGI,
        'reporting_lag':  lag_log,
        'GEP_growth':     gep_growth,
        'combined_ratio': combined_ratio,
        'reins_ratio':    reins_ratio,
        'ln_assets':      ln_assets,
        'total_assets':   ta,
        'ins_rev':        ins_rev,
    })

df = pd.DataFrame(records)
print(f"  Total insurers matched: {len(df)}")
print(f"  Missing GEP growth: {df['GEP_growth'].isna().sum()}")
print(f"  Missing ln_assets: {df['ln_assets'].isna().sum()}")

df.to_csv('ifrs17_dataset_2023.csv', index=False)
print(f"  Dataset saved: ifrs17_dataset_2023.csv")

# ── DESCRIPTIVE STATISTICS ────────────────────────────────────────
print("\n" + "=" * 65)
print("DESCRIPTIVE STATISTICS (n=all active insurers)")
print("=" * 65)
desc_cols = ['CAR','IDCS','RAGI','GEP_growth','combined_ratio',
             'reins_ratio','ln_assets']
print(df[desc_cols].describe().round(3).to_string())

print(f"\nCAR distribution:")
print(f"  N = {df['CAR'].count()}")
print(f"  Mean = {df['CAR'].mean():.1f}%")
print(f"  Median = {df['CAR'].median():.1f}%")
print(f"  Min = {df['CAR'].min():.1f}%  Max = {df['CAR'].max():.1f}%")
print(f"  Below 100%: {(df['CAR']<100).sum()}  |  Compliance rate: {(df['CAR']>=100).mean()*100:.1f}%")

print(f"\nIDCS distribution:")
from collections import Counter
print(f"  {Counter(df['IDCS'].round(2))}")
print(f"\nRAGI distribution (normalised):")
print(f"  {Counter(df['RAGI'].round(2))}")

# ── CORRELATION MATRIX ────────────────────────────────────────────
print("\n" + "=" * 65)
print("PAIRWISE CORRELATION MATRIX (principal variables)")
print("=" * 65)
corr_cols = ['CAR','IDCS','RAGI','ln_assets','GEP_growth','combined_ratio']
corr_df = df[corr_cols].dropna()
print(corr_df.corr().round(3).to_string())
print(f"\n  n = {len(corr_df)} (complete cases)")
print(f"  Reinsurance ratio omitted from matrix for brevity")
print(f"  (reins_ratio correlations: CAR={df[['CAR','reins_ratio']].dropna().corr().iloc[0,1]:.2f})")

# ── M1: OLS CROSS-SECTION 2023 ────────────────────────────────────
print("\n" + "=" * 65)
print("TABLE 14 — M1: OLS Cross-Section (2023)")
print("=" * 65)
print("CAR ~ IDCS + RAGI + reporting_lag + GEP_growth + combined_ratio")
print("      + ln_assets + reins_ratio")
print("SE: heteroskedasticity-robust (HC3)")

df_m1 = df[['CAR','IDCS','RAGI','reporting_lag','GEP_growth',
             'combined_ratio','ln_assets','reins_ratio']].dropna()
print(f"n = {len(df_m1)}")

X_m1 = sm.add_constant(df_m1[['IDCS','RAGI','reporting_lag','GEP_growth',
                                'combined_ratio','ln_assets','reins_ratio']])
model_m1 = sm.OLS(df_m1['CAR'], X_m1).fit(cov_type='HC3')

print(f"\n{'Variable':<22} {'Coeff':>9} {'SE':>9} {'t':>7} {'p':>8}  sig")
print("-" * 62)
for var in model_m1.params.index:
    c = model_m1.params[var]
    se = model_m1.bse[var]
    t  = model_m1.tvalues[var]
    p  = model_m1.pvalues[var]
    s  = '***' if p<0.01 else ('**' if p<0.05 else ('*' if p<0.1 else ''))
    print(f"  {var:<20} {c:>9.3f} {se:>9.3f} {t:>7.2f} {p:>8.4f}  {s}")

print(f"\n  N={model_m1.nobs:.0f}  R²={model_m1.rsquared:.3f}  "
      f"Adj-R²={model_m1.rsquared_adj:.3f}  "
      f"F-stat={model_m1.fvalue:.2f}  p={model_m1.f_pvalue:.4f}")
print("\n  Note: Reporting lag is constant across all insurers (same IRA publication).")
print("  IDCS and RAGI are author-coded from IRA Kenya 2023 Appendix 1.")
print("  GEP growth uses Insurance Revenue 2023→2024 as proxy.")

# ── M2: PANEL FE — note ───────────────────────────────────────────
print("\n" + "=" * 65)
print("TABLE 14 — M2: Panel Fixed Effects (2020-2023)")
print("=" * 65)
print("""
M2 requires historical CAR data for 2020, 2021, 2022.
These are not in the uploaded files (2023 and 2024 only).

To run M2, add IRA Kenya Annual Statistics for 2020-2022 and
create historical_car.csv with columns: insurer, year, CAR

M2 specification:
  CAR_it = α_i + β1·Post2023_it + β2·GEP_growth_it
           + β3·combined_ratio_it + β4·ln_assets_it
           + β5·reins_ratio_it + ε_it
  Insurer FE; clustered SEs at insurer level
""")

print("=" * 65)
print("SCRIPT COMPLETE")
print("=" * 65)
