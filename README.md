<div align="center">

<img src="Task_1_EDA/assets/tata_logo.png" alt="Tata logo" width="150">

# Tata — GenAI-Powered Data Analytics

**Tata iQ Forage Simulation · Multi-Task Project Repository**

![Tasks](https://img.shields.io/badge/Tasks-1%20Complete%20%F0%9F%86%95-brightgreen)
![Python](https://img.shields.io/badge/Python-3.14-3776AB?logo=python&logoColor=white)

</div>

> **Disclaimer:** This repository is part of the Tata iQ *Forage* educational simulation. The Tata
> name and logo are the property of Tata Sons Pvt. Ltd. and are used here solely to identify the
> simulation scenario. This is **not** an official Tata deliverable, and no real customer data is
> involved.

## Repository Layout

| Folder | Task | Status | Description |
|---|---|---|---|
| [`Task_1_EDA/`](Task_1_EDA/README.md) | **Task 1 — Exploratory Data Analysis** | ✅ Complete | Geldium delinquency dataset: data-quality assessment, missing-data treatment, EDA, risk-indicator screening, official DOCX/PDF report, analysis workbook, cleaned dataset, charts and reproducible code |
| [`Task_2_ModelPlan/`](Task_2_ModelPlan/README.md) | **Task 2 — Predictive Model Plan** | ✅ Complete | Model logic (GenAI-generated), justification for a transparent logistic-regression champion + boosting challenger, and an evaluation strategy with fairness/ethics; includes an **executed** cross-validated baseline check confirming the no-skill floor |
| _Task 3 — …_ | _Coming soon_ | ⏳ | _Will be added here when assigned_ |

**Task 2 headline result:** all baseline models (dummy, logistic regression, gradient boosting) sit
at the 0.16 no-skill PR-AUC floor on the delivered data — measured confirmation that the data
pipeline must be fixed before any model is deployed. The plan embeds an explicit go/no-go gate.

## Quick Start (Task 1)

```bash
cd Task_1_EDA
pip install pandas numpy scipy matplotlib seaborn openpyxl python-docx nbformat docx2pdf
python geldium_eda.py        # regenerates every deliverable in Task_1_EDA/Geldium_EDA_Outputs/
```

📄 **Full documentation, findings and methodology:**
see the [Task 1 README](Task_1_EDA/README.md).

---

<div align="center">
<img src="Task_1_EDA/assets/tata_logo.png" alt="Tata" width="60">
<br><br>
<sub>Tata iQ Forage Simulation · GenAI-Powered Data Analytics</sub>
</div>
