<div align="center">

<img src="../Task_1_EDA/assets/tata_logo.png" alt="Tata logo" width="140">

# Task 3 — Business Summary Report

**Tata iQ Forage Simulation · Predictive Insights for Collections Strategy**

![Status](https://img.shields.io/badge/Status-Complete-brightgreen)
![Length](https://img.shields.io/badge/Length-2%20pages-brightgreen)

</div>

> **Disclaimer:** Part of the Tata iQ *Forage* educational simulation. The Tata name and logo are
> the property of Tata Sons Pvt. Ltd., used here solely to identify the simulation scenario.

## Contents

| File | Description |
|---|---|
| `Task_3_Business_Report.docx` | Completed official template (`Updated_Business_Summary_Report_Template.docx`, preserved) — exactly 2 pages |
| `Task_3_Business_Report.pdf` | PDF version for submission |
| `task3_business_report.py` | Reproducible script that regenerates the DOCX from the template; pulls the Task 2 baseline numbers from `baseline_check_results.json` so no statistic is hard-typed |

## What the report covers

**1. Summary of Predictive Insights** — Task 2's executed baseline (all models at the 0.16
no-skill PR-AUC floor) framed honestly, then the three strongest *observed* segments: DTI > 40%
(18.9% vs 9.8%), unemployed customers (19.4%), Business card tier (21.3%) — with the guide-based
note that `Missed_Payments` (12-month window) and `Month_1..6` (6-month window) show no measurable
relationship to the target. Includes the **Top 3 Risk Factors** bullet list and a Key Insights
Summary Table.

**2. Recommendation Framework (SMART)** — a controlled 90-day Collections pilot offering supportive
restructuring outreach to the DTI > 40% cohort against a matched control: Specific, Measurable
(≥15% relative delinquency reduction at day 90), Actionable (existing team/CRM, rule-based — no
model deployment per the Task 2 gate), Relevant, Time-bound (launch 4 weeks, decision at day 91),
with business rationale.

**3. Ethical and Responsible AI Considerations** — two fairness risks with mitigations (proxy
discrimination via Location/Employment_Status; vulnerable-group impact on unemployed customers),
plain-language explainability, and responsible-AI practices (transparency, accountability, data
privacy, fairness re-audit before scale-up).

## Reproduce

```bash
cd Task_3_Business_Report
pip install python-docx docx2pdf
python task3_business_report.py     # rebuilds the completed DOCX from the official template
```
