<div align="center">

<img src="assets/tata_logo.png" alt="Tata logo" width="180">

# Geldium Delinquency Prediction — Exploratory Data Analysis (EDA)

**Tata iQ Forage Simulation · GenAI-Powered Data Analytics Project**

![Python](https://img.shields.io/badge/Python-3.14-3776AB?logo=python&logoColor=white)
![pandas](https://img.shields.io/badge/pandas-3.0-150458?logo=pandas&logoColor=white)
![Status](https://img.shields.io/badge/Status-EDA%20Complete%20%F0%9F%86%95-brightgreen)
![Report](https://img.shields.io/badge/Report-DOCX%20%2B%20PDF-blue)
![Reproducible](https://img.shields.io/badge/Reproducible-Script%20%2B%20Notebook-informational)

**Objective:** Assess data quality and explore risk indicators for customer-account delinquency
**before** any predictive model is built.

</div>

---

> **Disclaimer:** This project is part of the Tata iQ *Forage* educational simulation. The Tata name
> and logo are the property of Tata Sons Pvt. Ltd. and are used here solely to identify the
> simulation scenario. This is **not** an official Tata deliverable, and no real customer data is
> involved — the dataset is a synthetic simulation file.

---

## 1. Business Context

Geldium is a consumer-credit lender within the Tata ecosystem. When a customer account becomes
**delinquent** (misses contracted payments), Geldium incurs credit losses and must act early —
collections outreach, limit reductions, or restructuring. A predictive model that flags
high-risk accounts *before* they go delinquent enables proactive, loss-mitigating action.

This repository documents the **Exploratory Data Analysis (EDA)** stage of that initiative:
profiling the delivered dataset, quantifying its quality, defining a justified missing-data
treatment, exploring relationships with the delinquency target, and shortlisting candidate risk
indicators for later validation.

## 2. Repository Structure

```
Tata---GenAI-Powered-Data-Analytics/
├── README.md                              ← repo overview (root)
└── Task_1_EDA/                            ← Task 1: this folder
    ├── README.md                          ← this file (full task documentation)
    ├── Delinquency_prediction_dataset.xlsx← SOURCE DATA (read-only, never modified)
    ├── EDA_SummaryReport_Template.docx    ← official report template (preserved)
    ├── geldium_eda.py                     ← reproducible end-to-end analysis script
    ├── assets/
    │   ├── tata_logo.png                  ← Tata logo (rendered, used in this README)
    │   └── tata_logo.svg                  ← Tata logo (original vector, Wikimedia Commons)
    └── Geldium_EDA_Outputs/               ← all generated deliverables
        ├── Geldium_EDA_Summary_Report.docx← completed official 6-section report
        ├── Geldium_EDA_Summary_Report.pdf ← PDF version (6 pages)
        ├── Geldium_EDA_Analysis.xlsx      ← 15-sheet analysis workbook
        ├── Geldium_Cleaned_Dataset.xlsx   ← cleaned working copy + cleaning log
        ├── Geldium_EDA_Code.ipynb         ← same analysis as a Jupyter notebook (36 cells)
        ├── Geldium_EDA_Charts/            ← 9 publication-quality PNG charts
        ├── verified_stats.json            ← machine-verified headline statistics (QA)
        └── run_log.txt                    ← full execution log of the final run
```

> All paths above are relative to this `Task_1_EDA/` folder.

## 3. Dataset Snapshot

| Attribute | Value |
|---|---|
| Records × variables | **500 × 19** (single worksheet) |
| Identifier | `Customer_ID` — 500 unique, **0 duplicates** |
| Target | `Delinquent_Account` (1 = delinquent, 0 = not) — **80/500 = 16.0% positive** |
| Predictors | Age, Income, Credit_Score, Credit_Utilization, Missed_Payments, Loan_Balance, Debt_to_Income_Ratio, Employment_Status, Account_Tenure, Credit_Card_Type, Location, Month_1..Month_6 |
| Data types | 8 numeric + 1 binary target + 9 categorical + 1 identifier |
| Missing data | Income 39 (7.8%), Loan_Balance 29 (5.8%), Credit_Score 2 (0.4%) — all other columns complete |
| Duplicates / outliers | 0 duplicate rows · 0 IQR (1.5×IQR) outliers · 0 sentinel or out-of-range values |

## 4. Methodology (executed phases)

| Phase | Activity | Status |
|---|---|---|
| 1 | Source inspection — worksheets, dtypes, sample records, target identification | ✅ |
| 2 | Data quality — structure, missing values, invalid/inconsistent values, outliers, dictionary | ✅ |
| 3 | Missing-data treatment — median imputation + indicators on a working copy, cleaning log | ✅ |
| 4 | EDA — univariate, bivariate (delinquency rates by groups/bands), correlations, charts | ✅ |
| 5 | Risk-indicator screening — ranked candidates with statistics and caveats | ✅ |
| 6 | AI & GenAI usage documentation — tools, prompts, human validation | ✅ |
| 7 | Official Word report populated from the template (6 sections) + PDF | ✅ |
| 8 | Supporting outputs — analysis workbook, cleaned dataset, charts, code | ✅ |
| 9 | Quality assurance — every figure re-verified against the untouched source | ✅ |

## 5. Key Findings (all statistics computed from the dataset)

| # | Finding | Evidence |
|---|---|---|
| 1 | **No predictor is significantly associated with the delinquency target** | All point-biserial \|r\| ≤ 0.05 (p > 0.33); all Mann-Whitney/chi² p > 0.10; every univariate AUC 0.47–0.54 (chance level) |
| 2 | **Predictors are mutually independent** | Every pairwise Pearson and Spearman \|r\| ≤ 0.10 — unusual for real credit data; suggests synthetic generation or a randomly assigned target |
| 3 | `Missed_Payments` contradicts the 6-month payment history | Agreement in only **72/500 rows** (corr = 0.007) — two conflicting behavioural feeds |
| 4 | `Debt_to_Income_Ratio` is not derivable from Loan_Balance ÷ Income | corr(DTI, Loan/Income) = 0.047 — business formula must be confirmed |
| 5 | Category-label inconsistency | `Employment_Status` has 6 raw variants ("EMP", "employed", "retired"…) → 4 standardised categories; **245 rows** relabelled |
| 6 | Logical impossibility | **81 rows (16.2%)** where Account_Tenure > Age − 18 (e.g., 18-year-olds with 12–18 years tenure) |
| 7 | Borderline utilisation values | 4 records with Credit_Utilization slightly > 1.0 (max 1.026) — retained and flagged, not deleted |
| 8 | Class imbalance | 16.0% delinquent — a naive "never delinquent" model already scores 84% accuracy |
| 9 | Missingness is treatable | 64 rows (12.8%) miss ≥1 value; delinquency 17.2% vs 15.8% in complete rows (Fisher p = 0.72) — consistent with MCAR |

### Candidate risk indicators (descriptive screening only — **none statistically validated**)

| Rank | Indicator | Observed pattern | Key statistics |
|---|---|---|---|
| 1 | Debt_to_Income_Ratio | 18.9% delinquent in >40% band vs 9.8% in <20% band | r = 0.034 (p = 0.44); z-test p = 0.10 |
| 2 | Employment_Status | Unemployed 19.4% vs Retired 11.5% | chi² p = 0.78 |
| 3 | Credit_Card_Type | Business 21.3% vs Platinum 11.8% | chi² p = 0.35 |
| 4 | Location *(fairness-sensitive)* | Los Angeles 19.6% vs New York 12.0% | chi² p = 0.66 |
| 5 | Age / Account_Tenure *(protected/proxy)* | 45–59 band highest (20.0%) | r = 0.023 / −0.040 (p = 0.62 / 0.37) |

> **Bottom line:** the dataset as delivered contains **no validated delinquency signal**. Confirm with
> the data owner how the target was generated before investing in predictive modelling.

## 6. Missing-Data Treatment (applied to the cleaned working copy only)

| Variable | Missing | Treatment | Rationale / limitation |
|---|---|---|---|
| Income | 39 (7.8%) | Median imputation + `Income_Missing` indicator | Near-symmetric (skew 0.05); understates variance; refit inside CV folds when modelling |
| Loan_Balance | 29 (5.8%) | Median imputation + `Loan_Balance_Missing` indicator | Near-symmetric (skew 0.11); same caveat |
| Credit_Score | 2 (0.4%) | Median imputation | Only 2 records; sensitivity check recommended |
| Target | 0 (0%) | None required | Complete — no label imputation was attempted |

**Guardrails honoured:** original file preserved · no synthetic customers generated · no rows deleted ·
missing payment history never treated as successful payment · flagged values retained, not overwritten.

## 7. AI & GenAI Usage Disclosure

A generative-AI coding assistant (LLM agent with Python execution) supported the analysis by drafting
and running the profiling, statistical-testing and reporting code and by suggesting narrative text.
Every number in the deliverables was **computed from the actual dataset by executed code** — none was
model-generated. Example prompts used:

- *"Summarise the dataset structure, missing values, anomalies and data quality risks; verify every figure against the file."*
- *"Recommend an appropriate missing-data treatment for each affected variable and justify it against the observed distribution."*
- *"Identify variables associated with delinquency, report the actual statistics, and flag anything that is not statistically significant."*
- *"Review the analysis for potential target leakage, selection bias and unsupported conclusions before finalising the report."*

Human validation consisted of re-running the statistics, cross-checking counts against the raw
workbook, and reviewing the analysis for leakage, bias and unsupported claims. AI did **not**
generate synthetic data and did **not** modify the source file.

## 8. How to Reproduce

**Requirements:** Python 3.10+ (tested on 3.14)

```bash
pip install pandas numpy scipy matplotlib seaborn openpyxl python-docx nbformat docx2pdf
python geldium_eda.py            # regenerates every deliverable in Geldium_EDA_Outputs/
```

Or open `Geldium_EDA_Outputs/Geldium_EDA_Code.ipynb` and run all cells.
Optional: `docx2pdf` (requires Microsoft Word) regenerates the PDF report; without it, the DOCX is
still fully generated.

> The script reads `Delinquency_prediction_dataset.xlsx` and `EDA_SummaryReport_Template.docx` from
> this `Task_1_EDA/` folder and **never writes to either file**.

## 9. Deliverables

| File | Contents |
|---|---|
| `Geldium_EDA_Summary_Report.docx/.pdf` | Official template completed: Introduction · Dataset Overview · Missing Data Analysis · Key Findings & Risk Indicators · AI & GenAI Usage · Conclusion & Next Steps (Tata-branded title, 3 tables, composite figure) |
| `Geldium_EDA_Analysis.xlsx` | 15 sheets: Data_Dictionary, Dataset_Structure, Data_Quality_Summary, Missing_Value_Report, Missing_Data_Treatment, Descriptive_Statistics, Category_Frequencies, Delinquency_By_Group, Statistical_Tests, Correlations (Pearson/Spearman), Top_Correlations, Unexpected_Findings, Risk_Indicator_Ranking, Cleaning_Log |
| `Geldium_Cleaned_Dataset.xlsx` | Cleaned data (500 × 29, zero remaining NaNs, target intact) + `Cleaning_Log` sheet (8 logged transformations) |
| `Geldium_EDA_Charts/` | 9 labelled charts (missing values, target distribution, correlation heatmap, delinquency by bands, box plots, missed-payment trend, label-quality comparison, monthly history, report composite) |
| `verified_stats.json` + `run_log.txt` | QA evidence: headline statistics and the full execution log |

## 10. Recommended Next Steps

1. **Resolve critical quality issues at source** — standardise `Employment_Status` labels, fix the 81 tenure-age inconsistencies, reconcile `Missed_Payments` with the monthly history, and obtain the official `Debt_to_Income_Ratio` formula.
2. **Confirm how the dataset and target were generated** — the absence of any predictor–target association must be explained before modelling effort is invested.
3. **Apply the documented imputation strategy strictly inside training folds** and keep the missingness indicators.
4. **Validate the five candidate risk indicators** with proper train/test validation rather than descriptive screens alone.
5. **Evaluate discrimination, calibration and fairness** in the modelling phase — PR-AUC (16% base rate), calibration curves, and performance monitored across Age, Location and Employment_Status segments.

## 11. Limitations

- `Updated_Dataset_Description_Guide.pdf` was not available in the project folder; variable definitions are **inferred** from column names and observed value ranges and should be confirmed with the data owner.
- No predictive model was trained in this phase; risk-indicator rankings are descriptive and explicitly not validated.
- The report is 6 pages (template target 3–5) to accommodate three tables and the composite figure.

## 12. Attribution & Licence

- **Tata logo:** sourced from [Wikimedia Commons — *Tata logo.svg*](https://commons.wikimedia.org/wiki/File:Tata_logo.svg) and rasterised locally for this README. © Tata Sons Pvt. Ltd.; used here only to identify the Forage simulation scenario.
- **Dataset:** Tata iQ Forage simulation file (`Delinquency_prediction_dataset.xlsx`) — synthetic data, no real customers.
- **Code:** original analysis code in this repository, free to reuse for educational purposes.

---

<div align="center">
<img src="assets/tata_logo.png" alt="Tata" width="70">
<br><br>
<sub>Tata iQ Forage Simulation · GenAI-Powered Data Analytics · EDA complete — dataset profiled,
quality issues quantified, missing-data treatment applied and documented, risk indicators screened
with full statistical transparency.</sub>
</div>
