# ifrs17-kenya-suptech

Replication code for:

**"IFRS 17 as Supervisory Data Infrastructure: Toward an AI-Ready Insurance Supervision Architecture in Kenya"**

Author: Jeniffer Nasike Atetwe | ORCID: 0009-0009-0309-5045  
Journal: FinTech and Sustainable Innovation (FSI), Bon View Publishing

---

## How to run

```bash
pip install pandas openpyxl statsmodels numpy scipy
python ifrs17_regression.py
```

Place the two data files in the same folder as the script before running.

---

## Files

| File | Purpose |
|------|---------|
| `ifrs17_regression.py` | Main script — loads data, computes all variables, runs regression |
| `data/IRA_Kenya_Annual_Statistics_2023.xlsx` | Source: IRA Kenya (public) |
| `data/IRA_Kenya_Annual_Statistics_2024.xlsx` | Source: IRA Kenya (public) |

Data files are publicly available at: https://www.ira.go.ke/index.php/statistics

---

## What the script computes automatically

| Variable | Source | Method |
|----------|--------|--------|
| CAR | Appendix CAR 2 | As published |
| IDCS | Appendix 1 | 5 binary criteria: InsRev, ISE, NRE, ISR present + ISR reconciliation |
| RAGI | Appendix 1 | 0–4 ordinal: disclosure depth from ISR through PBT |
| GEP growth | Appendix 1 (2023 + 2024) | (InsRev_2024 − InsRev_2023) / InsRev_2023 |
| Reporting lag | IRA Details sheet | Days from 31 Dec 2023 to IRA annual report publication |
| Combined ratio | Appendix 1 | ISE / Insurance Revenue |
| Reinsurance ratio | Appendix 1 | |NRE| / Insurance Revenue |
| ln(assets) | Appendix 5 | log(Total Assets) |

---

## Key empirical finding

IFRS 17 achieved near-universal disclosure completeness in Year 1:
29 of 30 active general insurers scored IDCS = 1.0 in 2023.
All 30 disclosed through PBT (RAGI = 1.0).
Only ln(assets) is a significant predictor of CAR (coeff ≈ 57, p = 0.037).

---

## M2 panel regression

M2 requires IRA Kenya Annual Statistics for 2020–2022 (not uploaded here).
Download from the IRA Kenya website and place in the `data/` folder.
The script detects them automatically if named:
- `data/IRA_Kenya_Annual_Statistics_2023.xlsx`
- `data/IRA_Kenya_Annual_Statistics_2024.xlsx`

---

## Licence

MIT. Data sourced from IRA Kenya public statistics.
