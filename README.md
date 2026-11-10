# ifrs17-kenya-suptech

Replication code for the manuscript "IFRS 17 as Supervisory Data Infrastructure: Evidence from Kenya and a Proposed AI-Ready SupTech Architecture" (FinTech and Sustainable Innovation, Bon View Publishing).

Author: Jeniffer Nasike Atetwe

## How to reproduce

```bash
pip install -r requirements.txt
python ifrs17_analysis.py
```

The script reads the two workbooks in the `data/` folder (no other input is needed) and writes `ifrs17_dataset_2023.csv`. The same steps run automatically in GitHub Actions (`.github/workflows/run_analysis.yml`); the output log is the reference for the values in the paper. Library versions in `requirements.txt` are those used in the reference run (Python 3.11).

## Scope

The analysis is descriptive. The script reads only the Appendix 1, Appendix 5 (I to IV) and Details sheets of the IRA Kenya Annual Insurance Industry Statistics workbooks. No other sheets, no capital adequacy data and no regression models are used.

## Files

| File | Purpose |
|------|---------|
| `ifrs17_analysis.py` | Computes all variables and prints Tables 4 to 6 and 11 to 13 |
| `requirements.txt` | Library versions used in the reference run |
| `data/IRA_Kenya_Annual_Statistics_2023.xlsx` | IRA Kenya 2023 workbook as downloaded from IRA (sheets read: Appendix 1, Appendix 5 I to IV, Details) |
| `data/IRA_Kenya_Annual_Statistics_2024.xlsx` | IRA Kenya 2024 workbook as downloaded from IRA (sheet read: Appendix 1, used for revenue growth) |
| `ifrs17_dataset_2023.csv` | Analytical dataset written by the script |

Source of the workbooks: https://www.ira.go.ke/index.php/statistics

## Variables (as implemented)

| Variable in code | Meaning | Source |
|------------------|---------|--------|
| IDCS | Five rule-based checks on Appendix 1 fields (revenue, ISE, NRE, ISR non-zero; ISR reconciles within 1%), divided by five. A reported zero scores 0 on its check | 2023 Appendix 1 |
| RAGI (ISDI in the paper) | Income Statement Disclosure Depth, 0 to 4, divided by four. It does not measure risk-adjustment methodology | 2023 Appendix 1 |
| GEP_growth | Insurance Revenue growth 2023 to 2024 | 2023 and 2024 Appendix 1 |
| combined_ratio | Insurance Service Expense / Insurance Revenue | 2023 Appendix 1 |
| reins_ratio | absolute Net Reinsurance Expense / Insurance Revenue | 2023 Appendix 1 |
| ln_assets | Natural log of total assets (KES thousands) | 2023 Appendix 5 |
| reporting_lag | Constant (244 days, assumed publication date 31 August 2024) | 2023 Details sheet |

## Expected output

37 Appendix 1 insurer entries; 32 active insurers (five report zero Insurance Revenue). IDCS mean 0.994 (31 of 32 insurers score 1.0; one scores 0.8); ISDI 1.000 for all 32. Correlations use the 25 complete cases among ln(assets), GEP growth, combined ratio and reinsurance ratio. The eight largest insurers by 2023 Insurance Revenue have a mean combined ratio of 0.832 against 0.876 for the other 24, and a mean reinsurance ratio of 0.127 against 0.107. All results are descriptive.
