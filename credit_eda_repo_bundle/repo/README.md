# Credit EDA Case Study — Payment Difficulty Analysis

Exploratory Data Analysis on the Home Credit loan dataset to identify applicant
patterns associated with payment difficulty (late payment on installments),
so the lender can make better approve / refuse / re-price decisions.

## Objective

Identify which applicant characteristics — at the time of loan application —
are associated with a client later having payment difficulty (`TARGET = 1`),
versus repaying on time (`TARGET = 0`). Insights are meant to support decisions
such as: approve, refuse, reduce loan amount, or price at a higher interest rate.

## Data

Two files (not included in this repo due to size — download from
[Kaggle: Credit EDA Case Study](https://www.kaggle.com/code/skagrawal/credit-eda-case-study/input)):

| File | Rows | Columns | Description |
|---|---|---|---|
| `application_data.csv` | 307,511 | 122 | One row per current loan application, includes `TARGET` |
| `previous_application.csv` | 1,670,214 | 37 | Every prior application each client made with this lender, including decision (`Approved` / `Canceled` / `Refused` / `Unused offer`) |

## Repo structure

```
notebooks/
  eda_analysis.py         # end-to-end analysis: cleaning, outliers, univariate/
                           # bivariate by TARGET, correlation shift, prior-history merge
  build_excel_report.py   # builds the formatted Excel deliverable from the results
outputs/
  credit_eda_case_study.xlsx   # final Excel report (8 sheets, see below)
  csv_results/                 # raw result tables produced by eda_analysis.py
```

## How to reproduce

```bash
pip install pandas numpy openpyxl

# 1. Run the analysis against the raw CSVs
python notebooks/eda_analysis.py --data-dir /path/to/csvs --out-dir outputs/csv_results

# 2. Build the Excel report (currently reads the computed constants baked in
#    from step 1 — see script header to point it at outputs/csv_results if you
#    want it to read the CSVs directly instead)
python notebooks/build_excel_report.py
```

## Excel deliverable — `outputs/credit_eda_case_study.xlsx`

| Sheet | Contents |
|---|---|
| Executive Summary | Headline findings and recommendations, one page |
| Data Quality | Missing-value ranking, outlier scan (IQR method) |
| Target Imbalance | Class balance (92% / 8%) and why accuracy is misleading here |
| Numeric Vars by Target | Median comparison, income-bucket and age-bucket default rates + charts |
| Categorical Vars by Target | Default rate by gender, education, occupation, housing, etc. + charts |
| Correlation Analysis | Correlation matrices for each TARGET segment, biggest shifts between them |
| Previous Application Analysis | Merge with prior-loan history; dose-response between prior refusals and current default |
| Recommendations | Underwriting actions derived from the findings |

## Key findings

- **Class imbalance:** 91.9% repay on time vs 8.1% have payment difficulty — a ~12:1 imbalance that makes plain accuracy a poor evaluation metric.
- **External bureau scores are the strongest signal:** `EXT_SOURCE_2`/`EXT_SOURCE_3` median 0.57/0.55 for repayers vs 0.44/0.38 for defaulters.
- **Age:** default rate falls monotonically from 11.4% (20–30 yrs) to 4.9% (60–70 yrs).
- **Prior refusal history is a strong, actionable, dose-response signal:** default rate rises from 6.98% (0 prior refusals) to 14.09% (5+ prior refusals) — roughly doubling.
- **Occupation/income type matters:** Low-skill Laborers (17.2%) and Drivers (11.3%) default well above Pensioners (5.4%).
- **Loan-size ratios (credit-to-income, annuity-to-income) are weak standalone predictors** — nearly identical medians between the two groups.
- **Correlation structure shifts:** among defaulters, `AMT_INCOME_TOTAL` becomes far less correlated with loan/annuity size than among repayers — loan sizing is less anchored to income for the risky segment.

## Recommendations

1. Weight `EXT_SOURCE_2` / `EXT_SOURCE_3` heavily in scoring.
2. Use prior-refusal count as an escalation trigger (manual review / repricing at 3+ refusals).
3. Apply extra income verification for young, short-tenure applicants.
4. Differentiate pricing for higher-risk occupations rather than blanket rejection.
5. Don't rely on credit-to-income ratio alone as a cutoff.
6. Cap/winsorize `AMT_INCOME_TOTAL` and `CNT_CHILDREN` outliers before any modeling.

## License / Data source

Dataset originally from the [Home Credit Default Risk Kaggle competition](https://www.kaggle.com/c/home-credit-default-risk),
packaged for this case study at the Kaggle link above. Used here for educational/case-study purposes.
