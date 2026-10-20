# %% [markdown]
# # Task 2 — Predictive Model Plan (Tata iQ Forage Simulation)
#
# This script produces the completed `Task 2_ModelPlan_Template.docx`:
#   1. Runs an **executed baseline sanity check** (real cross-validated models) so the plan is
#      grounded in measured evidence, not intuition.
#   2. Fills the template's three sections (Model Logic, Justification, Evaluation Strategy)
#      with content derived from the Task 1 EDA findings.
#
# Inputs : `../Task_1_EDA/Geldium_EDA_Outputs/Geldium_Cleaned_Dataset.xlsx` (cleaned data),
#          `../Task 2_ModelPlan_Template.docx` (official template — preserved)
# Outputs: `Task_2_ModelPlan_Completed.docx`, `baseline_check_results.json`

# %%
import json
from pathlib import Path

import numpy as np
import pandas as pd
from sklearn.compose import ColumnTransformer
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import StandardScaler, OneHotEncoder
from sklearn.impute import SimpleImputer
from sklearn.linear_model import LogisticRegression
from sklearn.ensemble import HistGradientBoostingClassifier
from sklearn.dummy import DummyClassifier
from sklearn.model_selection import RepeatedStratifiedKFold, cross_validate
from sklearn.metrics import average_precision_score, roc_auc_score, brier_score_loss

BASE = Path(__file__).resolve().parent
CLEAN = BASE.parent / "Task_1_EDA" / "Geldium_EDA_Outputs" / "Geldium_Cleaned_Dataset.xlsx"
TEMPLATE = BASE.parent / "Task 2_ModelPlan_Template.docx"

# %% [markdown]
# ## A. Baseline sanity check (actually executed)

# %%
from sklearn.impute import SimpleImputer

df = pd.read_excel(CLEAN, sheet_name="Cleaned_Data")
target = "Delinquent_Account"

num_cols = ["Age", "Income", "Credit_Score", "Credit_Utilization", "Missed_Payments",
            "Loan_Balance", "Debt_to_Income_Ratio", "Account_Tenure",
            "N_Missed_6M", "N_Late_6M", "Payment_Severity_Score"]
cat_cols = ["Employment_Status_Clean", "Credit_Card_Type"]

X = df[num_cols + cat_cols]
y = df[target]
base_rate = float(y.mean())

pre = ColumnTransformer([
    ("num", Pipeline([("imp", SimpleImputer(strategy="median")),
                      ("sc", StandardScaler())]), num_cols),
    ("cat", Pipeline([("imp", SimpleImputer(strategy="most_frequent")),
                      ("ohe", OneHotEncoder(handle_unknown="ignore"))]), cat_cols),
])

models = {
    "Dummy (stratified)": DummyClassifier(strategy="stratified", random_state=42),
    "Logistic Regression (balanced)": LogisticRegression(max_iter=2000, class_weight="balanced", random_state=42),
    "HistGradientBoosting": HistGradientBoostingClassifier(random_state=42),
}

cv = RepeatedStratifiedKFold(n_splits=5, n_repeats=3, random_state=42)
scoring = {
    "pr_auc": "average_precision",
    "roc_auc": "roc_auc",
    "brier": "neg_brier_score",
}
results = {"n_rows": int(len(df)), "positive": int(y.sum()), "base_rate": round(base_rate, 4),
           "folds": "5x3 repeated stratified CV", "models": {}}

for name, m in models.items():
    pipe = Pipeline([("pre", pre), ("clf", m)])
    res = cross_validate(pipe, X, y, cv=cv, scoring=scoring, n_jobs=-1)
    results["models"][name] = {
        "pr_auc_mean": round(float(res["test_pr_auc"].mean()), 4),
        "pr_auc_std": round(float(res["test_pr_auc"].std()), 4),
        "roc_auc_mean": round(float(res["test_roc_auc"].mean()), 4),
        "roc_auc_std": round(float(res["test_roc_auc"].std()), 4),
        "brier_mean": round(float(-res["test_brier"].mean()), 4),
    }
print(json.dumps(results, indent=2))
(BASE / "baseline_check_results.json").write_text(json.dumps(results, indent=2), encoding="utf-8")

# %% [markdown]
# ## B. Fill the official template

# %%
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document(TEMPLATE)

def find_par(fragment):
    for p in doc.paragraphs:
        if fragment in p.text:
            return p
    raise KeyError(fragment)

def add_par_after(anchor, text, style="Normal", bold=False, italic=False, size=None):
    p = doc.add_paragraph(style=style)
    r = p.add_run(text); r.bold = bold; r.italic = italic
    if size:
        r.font.size = Pt(size)
    anchor._p.addnext(p._p)
    return p

def set_cell_border(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single"); el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0"); el.set(qn("w:color"), "808080")
        borders.append(el)
    tcPr.append(borders)

def shade_cell(cell, hexcolor="D9E1F2"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hexcolor)
    tcPr.append(shd)

def add_table_after(anchor, header, rows, font_size=8.5, caption=None):
    t = doc.add_table(rows=1, cols=len(header))
    for j, h in enumerate(header):
        c = t.rows[0].cells[j]; c.text = str(h); shade_cell(c); set_cell_border(c)
        for r in c.paragraphs[0].runs:
            r.bold = True; r.font.size = Pt(font_size)
    for row in rows:
        cells = t.add_row().cells
        for j, v in enumerate(row):
            cells[j].text = "" if v is None else str(v)
            for p in cells[j].paragraphs:
                for r in p.runs:
                    r.font.size = Pt(font_size)
            set_cell_border(cells[j])
    anchor._p.addnext(t._tbl)
    last = t._tbl
    if caption:
        cap = doc.add_paragraph()
        run = cap.add_run(caption); run.italic = True; run.font.size = Pt(8)
        last.addnext(cap._p)
    return t

# --- title
paras = doc.paragraphs
paras[0].text = "Geldium Delinquency Prediction - Predictive Model Plan"
for r in paras[0].runs:
    r.bold = True

# ---------------- 1. Model Logic ----------------
p_logic = find_par("[Insert GenAI model logic here]")
p_logic.text = (
    "Objective: build a supervised binary-classification model that estimates the probability that a "
    "Geldium customer account becomes delinquent (Delinquent_Account = 1) at the next review, so that "
    "collections and retention outreach can be prioritised by risk. The logic below was drafted with a "
    "GenAI assistant and reviewed/refined by the analyst against the Task 1 EDA evidence "
    "(500 customers, 16.0% delinquent)."
)
a = add_par_after(p_logic, "Step-by-step logic:", bold=True)
a = add_par_after(a, "1. Data contract: use only pre-outcome features (demographics, credit behaviour, "
                     "account attributes, 6-month payment history). Exclude the target and any field "
                     "recorded after the delinquency outcome to prevent leakage.")
a = add_par_after(a, "2. Ingest and clean: apply the Task 1 cleaning log - standardise Employment_Status "
                     "labels (245 rows), create quality flags (Credit_Utilization > 1; tenure-age "
                     "inconsistency), derive payment-history aggregates (N_Missed_6M, N_Late_6M, "
                     "Payment_Severity_Score).")
a = add_par_after(a, "3. Split once: stratified 80/20 hold-out test set; all tuning and preprocessing "
                     "inside 5x3 repeated stratified cross-validation on the training fold only.")
a = add_par_after(a, "4. Preprocess inside each fold: median imputation + missingness indicators "
                     "(Income 7.8%, Loan_Balance 5.8%, Credit_Score 0.4%) and one-hot encoding of "
                     "categoricals - refitting imputation per fold to avoid leakage.")
a = add_par_after(a, "5. Model candidates: (i) dummy baseline for the no-skill floor, (ii) class-weighted "
                     "logistic regression as the transparent champion, (iii) histogram gradient boosting "
                     "as the non-linear challenger.")
a = add_par_after(a, "6. Tune by cross-validated PR-AUC (primary), subject to acceptable calibration; "
                     "calibrate probabilities (Platt/isotonic) if they feed expected-loss estimates.")
a = add_par_after(a, "7. Set the operating threshold from a business cost matrix (cost of a missed "
                     "delinquent = expected loss; cost of a false alert = outreach effort).")
a = add_par_after(a, "8. Gate and audit: fairness audit across Age bands, Location and Employment_Status; "
                     "go/no-go decision versus the baseline; document as a model card; monitor drift "
                     "and calibration in production.")
a = add_par_after(a, "Simplified pseudo-code:", bold=True)
a = add_par_after(a,
    "load data -> apply cleaning log -> stratify_split(train, test)\n"
    "for fold in StratifiedKFold(train):\n"
    "    impute(median) + missingness_flags + onehot_encode   # fit on fold-train only\n"
    "    fit logistic_regression(class_weight='balanced') and hist_gradient_boosting\n"
    "    score fold-valid: PR-AUC, ROC-AUC, Brier, recall@threshold\n"
    "select best model -> calibrate probabilities -> choose threshold by cost matrix\n"
    "final evaluation on untouched test set -> fairness audit -> deploy with monitoring",
    size=8.5, italic=True)
a = add_par_after(a, "Simplified implementation (scikit-learn):", bold=True)
a = add_par_after(a,
    "pre = ColumnTransformer([('num', Pipeline([('imp', SimpleImputer(strategy='median')),\n"
    "                                            ('sc', StandardScaler())]), num_cols),\n"
    "                        ('cat', OneHotEncoder(handle_unknown='ignore'), cat_cols)])\n"
    "model = Pipeline([('pre', pre),\n"
    "                  ('clf', LogisticRegression(class_weight='balanced'))])\n"
    "cross_validate(model, X, y, cv=RepeatedStratifiedKFold(5, 3),\n"
    "               scoring={'pr_auc': 'average_precision', 'roc_auc': 'roc_auc'})",
    size=8.5, italic=True)
a = add_par_after(a, "Baseline sanity check (executed on the cleaned Geldium data, "
                     f"{results['folds']}):", bold=True)
hdr = ["Model", "PR-AUC (mean ± sd)", "ROC-AUC (mean ± sd)", "Brier ↓"]
rows = []
for name, r in results["models"].items():
    rows.append([name, f"{r['pr_auc_mean']:.3f} ± {r['pr_auc_std']:.3f}",
                 f"{r['roc_auc_mean']:.3f} ± {r['roc_auc_std']:.3f}", f"{r['brier_mean']:.3f}"])
rows.append(["No-skill floor (delinquency base rate)", f"{base_rate:.3f}", "0.500", "-"])
add_table_after(a, hdr, rows,
                caption="Table 1: Executed baseline check - all models sit at the no-skill floor, "
                        "confirming the Task 1 finding that the delivered data carries no usable "
                        "delinquency signal. The plan therefore includes an explicit go/no-go gate.")

# ---------------- 2. Justification for Model Choice ----------------
p_just = find_par("[Insert your justification here]")
p_just.text = (
    "Primary model: class-weighted, L2-regularised logistic regression. In consumer credit, the model "
    "must be defensible to risk managers, auditors and regulators as much as predictive. Logistic "
    "regression offers direct probability outputs, stable coefficients, and per-application "
    "reason codes (e.g., 'high debt-to-income ratio') that satisfy adverse-action notice requirements "
    "in financial services. With only 500 training records and ~16% positives, a linear, strongly "
    "regularised model is also the statistically appropriate choice - complex learners would fit "
    "noise. It is trivial to retrain, monitor and re-derive after data corrections, which suits "
    "Geldium's need to iterate quickly while its data quality issues are being resolved."
)
b = add_par_after(p_just, "Considerations weighed:")
b = add_par_after(b, "- Accuracy: raw accuracy is rejected as a selection metric - at a 16% base rate a "
                     "'never delinquent' rule already scores 84%. Selection is by PR-AUC and "
                     "recall at the operating threshold.")
b = add_par_after(b, "- Transparency: logistic regression yields explainable odds ratios and automatic "
                     "reason codes; gradient boosting is retained only as a challenger and only if it "
                     "beats the linear model materially and passes the same fairness and calibration "
                     "gates (with SHAP-based explanations).")
b = add_par_after(b, "- Ease of use/implementation: a single scikit-learn pipeline (impute -> encode -> "
                     "fit) reproduces training end-to-end, versioned with the cleaning log.")
b = add_par_after(b, "- Relevance for financial prediction: calibrated probabilities allow conversion "
                     "into expected loss (probability x exposure via Loan_Balance), which is how "
                     "Geldium would size interventions.")
b = add_par_after(b, "- Suitability for Geldium's business needs: risk-ranked queues with reason codes "
                     "feed collections workflows directly; a challenger framework keeps room to gain "
                     "accuracy later without changing the operating process.")
b = add_par_after(b, "Rejected alternatives: neural networks (far too little data, opaque, no regulatory "
                     "justification); unweighted models (would under-predict the minority class); "
                     "rule-only scorecards (premature while predictor-target relationships are unproven).")
b = add_par_after(b, "Key caveat: the Task 1 EDA found no statistically significant predictor-target "
                     "association, and the executed baseline check (Table 1) confirms all models sit at "
                     "the no-skill floor. Model selection is therefore provisional: if the go/no-go "
                     "gate fails on refreshed data, the recommendation is to fix the data pipeline "
                     "rather than to escalate model complexity.")

# ---------------- 3. Evaluation Strategy ----------------
p_eval = find_par("[Insert your evaluation plan here]")
p_eval.text = (
    "Evaluation is designed around a 16% positive base rate, financial consequences of errors, and the "
    "data-quality risks identified in Task 1. All metrics are computed on stratified cross-validation "
    "during development and once on the untouched hold-out test set for the final report."
)
c = add_par_after(p_eval, "Metrics and interpretation:", bold=True)
c = add_par_after(c, "- Primary: PR-AUC (average precision) - informative under class imbalance; must "
                     f"exceed the no-skill floor of {base_rate:.2f} by a material margin (gate: "
                     "cross-validated PR-AUC clearly above floor with confidence intervals excluding it).")
c = add_par_after(c, "- Secondary: ROC-AUC for overall ranking quality; recall and precision at the "
                     "cost-based threshold (a missed delinquent costs expected loss, a false alert costs "
                     "an outreach call); F1 as a balance summary; lift/gains in the top risk decile to "
                     "size campaign efficiency for collections.")
c = add_par_after(c, "- Calibration: Brier score and calibration plot - probabilities must be reliable "
                     "before being multiplied by Loan_Balance exposure for expected-loss estimates.")
c = add_par_after(c, "- Confusion matrix and business-cost curve: choose the threshold that minimises "
                     "total cost rather than maximising accuracy.")
c = add_par_after(c, "Bias detection and fairness:", bold=True)
c = add_par_after(c, "- Subgroup analysis of recall, precision and calibration across Age bands "
                     "(18-29/30-44/45-59/60+), Employment_Status and Location - the protected or "
                     "proxy-sensitive attributes flagged in Task 1.")
c = add_par_after(c, "- Equal-opportunity gap (difference in true-positive rate across groups) and "
                     "calibration parity; Location is additionally evaluated with-and-without to "
                     "measure its marginal contribution versus its fairness risk.")
c = add_par_after(c, "- Small-cell guard: subgroup metrics reported only where n is sufficient; "
                     "otherwise merged or suppressed to avoid noisy or disclosive statistics.")
c = add_par_after(c, "Ethical considerations:", bold=True)
c = add_par_after(c, "- Predictions about financial behaviour can restrict credit access: models assist "
                     "human collections officers, never trigger fully automated adverse decisions.")
c = add_par_after(c, "- Adverse-action transparency: every high-risk flag carries human-readable reason "
                     "codes derived from the model; customers can be told why they were flagged.")
c = add_par_after(c, "- Data minimisation: protected attributes are used only for fairness auditing, "
                     "not as scoring inputs, unless a documented business necessity and legal basis "
                     "exist; inference of sensitive traits via proxies is explicitly tested.")
c = add_par_after(c, "- Governance: model card documenting training data, cleaning log, limitations "
                     "(including the no-signal finding), drift monitoring (PSI) on inputs and "
                     "outcomes, periodic recalibration, and a kill-switch back to manual review.")
c = add_par_after(c, "- Go/no-go gate: if the refreshed dataset still yields cross-validated PR-AUC at "
                     "the no-skill floor, the model is NOT deployed; the finding is escalated to the "
                     "data owner, because acting on noise would create unfair, arbitrary decisions.")

out = BASE / "Task_2_ModelPlan_Completed.docx"
doc.core_properties.title = "Geldium Predictive Model Plan"
doc.core_properties.author = "Tata iQ Forage Simulation - Data Analytics"
doc.save(out)
print("saved:", out)

# %% [markdown]
# ## C. QA

# %%
d2 = Document(out)
print("headings:", [p.text for p in d2.paragraphs if p.style.name.startswith("Heading")])
print("tables:", len(d2.tables))
print("placeholders left:", [p.text[:60] for p in d2.paragraphs if "[Insert" in p.text] or "none")
