<div align="center">

<img src="../Task_1_EDA/assets/tata_logo.png" alt="Tata logo" width="140">

# Task 4 — AI-Powered Collections Strategy (Presentation)

**Tata iQ Forage Simulation · Agentic AI for Scalable, Fair and Effective Debt Management**

![Status](https://img.shields.io/badge/Status-Complete-brightgreen)
![Slides](https://img.shields.io/badge/Slides-6-brightgreen)

</div>

> **Disclaimer:** Part of the Tata iQ *Forage* educational simulation. The Tata name and logo are
> the property of Tata Sons Pvt. Ltd., used here solely to identify the simulation scenario.

## Contents

| File | Description |
|---|---|
| `AI_Powered_Collections_Strategy_Deck.pptx` | Completed deck built **on the official template** (`Presentation_Template.pptx`, preserved) — 6 slides, presenter notes on every slide |
| `AI_Powered_Collections_Strategy_Deck.pdf` | PDF version for submission |
| `task4_presentation.py` | Reproducible script: rebuilds the deck from the template (python-pptx) |

## Design approach

Built directly on the template's **marina** master so the deck stays on-brand: 16:9, dark-blue
(`#00517C`) background, Arial, template accent palette (`#0277BD`, `#009688`, `#039BE5`, `#558B2F`)
for chevrons and tables, white content cards, yellow lead-in accents, and a Tata logo badge on the
title slide. Every slide was rendered and visually QA'd before finalising.

## Slide map (per the Task 4 brief)

| Slide | Content | Design element |
|---|---|---|
| 1 | Title — AI-Powered Collections Strategy | Template title + Tata badge + byline |
| 2 | **How the System Works** — inputs → decision logic → actions → learning loop | 4-step chevron flow with detail cards (the required 4-part diagram) |
| 3 | **Role of Agentic AI** — autonomous vs human oversight | Two-column table, 5 paired activities + boundary principle |
| 4 | **Responsible AI Guardrails** — fairness, explainability, human-in-the-loop/compliance, monitoring | 4 bold-lead guardrails + build-requirement note |
| 5 | **Expected Business Impact** — KPIs vs customer outcomes | Two-column table (green/blue headers) tied to the Task 3 SMART pilot |
| 6 | **Evidence Base: Why Pilot-First** (extra slide) | Task 1 DTI chart in a card + key statistics + honesty caveat |

## Grounded in executed evidence

- Base rate 16.0% (80/500); strongest observed gradient DTI > 40% = 18.9% vs 9.8% (< 20%)
- Executed Task 2 baseline: logistic regression, boosting and dummy all at PR-AUC 0.15–0.17 ≈ the
  0.16 no-skill floor → **pilot before model** is the deck's spine
- Guardrails map to Task 2's evaluation strategy (fairness audits, reason codes, kill-switch) and
  Task 3's ethics section (proxy discrimination, vulnerable-group impact)
- Field definitions cross-checked against the official Dataset Description Guide

## Reproduce

```bash
cd Task_4_Presentation
pip install python-pptx pywin32
python task4_presentation.py        # rebuilds the PPTX from the official template
# PDF: File > Save As PDF from PowerPoint, or win32com SaveAs(format=32)
```
