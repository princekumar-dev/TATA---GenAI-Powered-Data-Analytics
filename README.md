<div align="center">

<img src="Task_1_EDA/assets/tata_logo.png" alt="Tata logo" width="170">

# Tata — GenAI-Powered Data Analytics

### End-to-End Delinquency Risk Analytics for Geldium
**Tata iQ Forage Simulation · Tasks 1–4 · All deliverables backed by executed, reproducible code**

![Tasks](https://img.shields.io/badge/Tasks-4%20Complete%20%F0%9F%86%95-brightgreen)
![Python](https://img.shields.io/badge/Python-3.14-3776AB?logo=python&logoColor=white)
![pandas](https://img.shields.io/badge/pandas-3.0-150458?logo=pandas&logoColor=white)
![Reproducible](https://img.shields.io/badge/Reproducible-Script%20%2B%20Notebook%20%2B%20PDF-blue)
![Honesty](https://img.shields.io/badge/Statistics-All%20executed%20%F0%9F%94%A7-orange)

</div>

---

> **Disclaimer:** This repository is part of the Tata iQ *Forage* educational simulation. The Tata
> name and logo are the property of Tata Sons Pvt. Ltd. and are used here solely to identify the
> simulation scenario. This is **not** an official Tata deliverable, and no real customer data is
> involved — the dataset is a synthetic simulation file.

---

## 📖 About This Project

Geldium — a consumer-credit lender — wants to predict which customer accounts will become
**delinquent** so collections outreach can be proactive instead of reactive. This repository walks
the complete analytics lifecycle across four tasks:

1. **EDA & data quality** — profile 500 customer records, quantify quality issues, treat missing
   data, and screen for risk indicators
2. **Predictive model plan** — model logic, justification, and an evaluation strategy with
   fairness guardrails — plus an **executed** cross-validated baseline check
3. **Business summary report** — a two-page stakeholder report with a SMART recommendation
4. **Strategy presentation** — a 6-slide agentic-AI collections deck built on the official template

**Headline honesty:** the executed evidence shows the delivered dataset carries **no validated
delinquency signal** (all models at the 0.16 no-skill PR-AUC floor). Every deliverable states this
plainly and recommends pilot-before-model instead of acting on noise — the professional call a
stakeholder needs to hear.

## 🗂️ Repository Layout

| Folder | Task | Status | Description |
|---|---|---|---|
| [`Task_1_EDA/`](Task_1_EDA/README.md) | **1 — Exploratory Data Analysis** | ✅ | Data-quality assessment, missing-data treatment, EDA, risk-indicator screening; official DOCX/PDF report, 15-sheet analysis workbook, cleaned dataset + cleaning log, 9 charts, notebook |
| [`Task_2_ModelPlan/`](Task_2_ModelPlan/README.md) | **2 — Predictive Model Plan** | ✅ | Model logic (GenAI-generated), logistic-regression champion + boosting challenger justification, evaluation/fairness strategy; **executed** CV baseline check (no-skill floor confirmed) |
| [`Task_3_Business_Report/`](Task_3_Business_Report/README.md) | **3 — Business Summary Report** | ✅ | Two-page report: Top 3 Risk Factors, SMART 90-day Collections pilot, ethics section — completed from the official template |
| [`Task_4_Presentation/`](Task_4_Presentation/README.md) | **4 — AI-Powered Collections Strategy** | ✅ | 6-slide deck on the official template: system workflow diagram, agentic-AI autonomy table, responsible-AI guardrails, business impact, evidence slide — with presenter notes |

## 🔍 Key Results at a Glance

| Metric | Value | Source |
|---|---|---|
| Records × variables | 500 × 19 (0 duplicates) | Task 1 |
| Delinquency base rate | **16.0%** (80/500) | Task 1 |
| Missing data | Income 7.8%, Loan_Balance 5.8%, Credit_Score 0.4% — median-imputed + indicators | Task 1 |
| Strongest observed gradient | DTI > 40%: **18.9%** vs < 20%: **9.8%** (~2×, not significant) | Task 1 |
| Executed model baseline | PR-AUC **0.15–0.17** ≈ **0.16 no-skill floor** → pilot-before-model | Task 2 |
| Data-quality issues fixed/logged | 245 label fixes, 81 tenure-age flags, 4 utilisation flags | Task 1 cleaning log |

## 🚀 Quick Start

```bash
# Task 1 — EDA (regenerates report, workbook, cleaned data, charts, notebook)
cd Task_1_EDA && pip install pandas numpy scipy matplotlib seaborn openpyxl python-docx nbformat docx2pdf
python geldium_eda.py

# Task 2 — Model plan (reruns CV baseline + rebuilds the DOCX)
cd ../Task_2_ModelPlan && pip install scikit-learn
python task2_model_plan.py

# Task 3 — Business report (rebuilds the 2-page DOCX from the template)
cd ../Task_3_Business_Report && python task3_business_report.py

# Task 4 — Presentation (rebuilds the 6-slide PPTX from the template)
cd ../Task_4_Presentation && pip install python-pptx pywin32
python task4_presentation.py
```

All scripts read their inputs **read-only** — official templates and the source dataset are never
modified.

## 📁 Reference Files

| File | Location | Purpose |
|---|---|---|
| `Updated_Business_Summary_Report_Template.docx` | repo root | Official Task 3 report template (preserved) |
| `Presentation_Template.pptx` | repo root | Official Task 4 slide template (preserved) |
| `Updated_Dataset_Description_Guide.pdf` | `Task_2_ModelPlan/` | Official variable definitions (re-validated against the Task 1 EDA) |
| `Delinquency_prediction_dataset.xlsx` | `Task_2_ModelPlan/` | Dataset copy re-uploaded for Task 2 — differs from the Task 1 source copy; analyses use the verified copy in `Task_1_EDA/` |

## 🧭 Repository Conventions

- One folder per task, each with its own README, deliverables and reproducible script
- Official templates and source data are **never overwritten** — completed versions are saved as new files
- Every statistic in every deliverable is computed by executed code and traceable to
  `verified_stats.json` / `baseline_check_results.json`
- AI-generated suggestions are documented, then **verified against the data** before inclusion

---

<div align="center">
<img src="Task_1_EDA/assets/tata_logo.png" alt="Tata" width="64">
<br><br>
<sub><b>Tata iQ Forage Simulation</b> · GenAI-Powered Data Analytics ·
Delinquency risk analytics done end-to-end — with the honesty to say when the data isn't ready.</sub>
</div>
