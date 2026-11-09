# ifrs17-kenya-suptech

Replication code for the manuscript "IFRS 17 as Supervisory Data Infrastructure: Evidence from Kenya and a Proposed AI-Ready SupTech Architecture" (FinTech and Sustainable Innovation, Bon View Publishing).

Author: Jeniffer Nasike Atetwe

## How to reproduce

```bash
pip install -r requirements.txt
python ifrs17_regression.py
```

The script reads the two IRA Kenya workbooks in the `data/` folder (no other input is needed) and writes `ifrs17_dataset_2023.csv`. The same steps run automatically in GitHub Actions (`.github/workflows/run_analysis.yml`); the output log is the reference for the values in the paper. Library versions in `requirements.txt` are those used in the reference run (Python 3.11).

## Files

| File | Purpose |
|------|---------|
| `ifrs17_regression.py` | Computes all variables and prints Tables 4 to 6 and 11 to 15 |
| `requirements.txt` | Library versions used in the reference run |
| `data/IRA_Kenya_Annual_Statistics_2023.xlsx` | IRA Kenya public workbook (sheets: Appendix 1, CAR 2, Appendix 5, Details) |
| `data/IRA_Kenya_Annual_Statistics_2024.xlsx` | IRA Kenya public workbook (Appendix 1, used for revenue growth) |
| `ifrs17_dataset_2023.csv` | Analytical dataset written by the script |

Source of the workbooks: https://www.ira.go.ke/index.php/statistics

## Variables (as implemented)

| Variable in code | Meaning | Source |
|------------------|---------|--------|
| CAR | Reported Capital Adequacy Ratio (%). The reference date of the sheet could not be confirmed (header shows 31.12.2021) | 2023 workbook, CAR 2 |
| IDCS | Five rule-based checks on Appendix 1 fields (revenue, ISE, NRE, ISR non-zero; ISR reconciles within 1%), divided by five | 2023 Appendix 1 |
| RAGI (ISDI in the paper) | Income Statement Disclosure Depth, 0 to 4, divided by four. It does not measure risk-adjustment methodology | 2023 Appendix 1 |
| GEP_growth | Insurance Revenue growth 2023 to 2024 | 2023 and 2024 Appendix 1 |
| combined_ratio | Insurance Service Expense / Insurance Revenue | 2023 Appendix 1 |
| reins_ratio | absolute Net Reinsurance Expense / Insurance Revenue | 2023 Appendix 1 |
| ln_assets | Natural log of total assets (KES thousands) | 2023 Appendix 5 |
| reporting_lag | Constant (244 days, assumed publication date 31 August 2024); excluded from the regression | 2023 Details sheet |

## Expected output

30 active matched insurers; OLS on 24 complete cases with HC3 standard errors: ln(assets) coefficient 56.99 (SE 27.36, p=0.037 normal approximation; p=0.051 with t), R-squared 0.348, F=2.381 (p=0.088). Mean CAR 142.8%, 9 insurers below 100%. The regression is exploratory and associational.
