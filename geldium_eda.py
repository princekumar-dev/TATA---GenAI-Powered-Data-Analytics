# %% [markdown]
# # Geldium Delinquency EDA — Reproducible Analysis Script
#
# **Project:** Tata iQ Forage Simulation — Geldium Delinquency Prediction EDA
# **Source file:** `Delinquency_prediction_dataset.xlsx` (read-only; never modified)
# **Outputs:** `Geldium_EDA_Outputs/` (analysis workbook, cleaned dataset, charts, report)
#
# Phases: (1) source inspection, (2) data quality, (3) missing-data treatment,
# (4) exploratory analysis, (5) risk-indicator screening, (7) Word report, (8) outputs.
#
# Note: the dataset guide PDF (`Updated_Dataset_Description_Guide.pdf`) was not
# available in the project folder (user confirmed proceeding without it); variable
# definitions below are inferred from column names and observed value ranges and
# are flagged as such in the report.

# %%
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
from pathlib import Path
import json

BASE = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
SRC = BASE / "Delinquency_prediction_dataset.xlsx"
OUT = BASE / "Geldium_EDA_Outputs"
CHARTS = OUT / "Geldium_EDA_Charts"
for p in (OUT, CHARTS):
    p.mkdir(exist_ok=True)

sns.set_theme(style="whitegrid", context="notebook")
plt.rcParams.update({"figure.dpi": 150, "savefig.dpi": 150, "axes.titlesize": 12,
                     "axes.labelsize": 10})

STATS = {}  # key verified numbers reused in the report

# %% [markdown]
# ## Phase 1 — Inspect the source files

# %%
xl = pd.ExcelFile(SRC)
print("Worksheets:", xl.sheet_names)
df = xl.parse("Delinquency_prediction_dataset")   # main (only) data sheet
RAW = df.copy()                                    # raw working copy; source never overwritten
print("shape:", df.shape)
print(df.dtypes)
df.head()

# %% [markdown]
# ## Phase 2 — Data quality analysis
#
# ### A. Structure

# %%
month_cols = [f"Month_{i}" for i in range(1, 7)]
numeric_cols = ["Age", "Income", "Credit_Score", "Credit_Utilization", "Missed_Payments",
                "Loan_Balance", "Debt_to_Income_Ratio", "Account_Tenure"]
target = "Delinquent_Account"

structure = pd.DataFrame({
    "Metric": ["Rows", "Columns", "Numeric columns", "Categorical columns",
               "Unique Customer_IDs", "Duplicate rows (full)", "Duplicate rows (excluding ID)",
               "Worksheets in workbook"],
    "Value": [len(df), df.shape[1], len(numeric_cols) + 1,
              df.shape[1] - len(numeric_cols) - 2, df.Customer_ID.nunique(),
              int(df.duplicated().sum()),
              int(df.drop(columns=["Customer_ID"]).duplicated().sum()),
              len(xl.sheet_names)],
})
STATS.update({"n_rows": len(df), "n_cols": df.shape[1],
              "n_dup_full": int(df.duplicated().sum())})
print(structure.to_string(index=False))

# %% [markdown]
# ### B. Missing values (incl. blank/sentinel checks)

# %%
miss = df.isna().sum().rename("n_missing").to_frame()
miss["pct_missing"] = (miss.n_missing / len(df) * 100).round(2)

# blank / whitespace-only / sentinel-like string scan
blank_scan = {}
for c in df.columns:
    if df[c].dtype == object or str(df[c].dtype) == "str":
        s = df[c].astype("string").str.strip()
        n_blank = int((s == "").sum())
        n_sentinel = int(s.str.lower().isin(["nan", "none", "null", "na", "n/a", "-", "?"]).sum())
        if n_blank or n_sentinel:
            blank_scan[c] = (n_blank, n_sentinel)

sentinel_numeric = {c: int(df[c].isin([-1, 0, 99, 999, 9999, 99999, 999999]).sum())
                    for c in numeric_cols}

rows_any_miss = int(df.isna().any(axis=1).sum())
rows_multi_miss = int((df.isna().sum(axis=1) > 1).sum())
STATS.update({"miss_income": int(df.Income.isna().sum()),
              "miss_loan": int(df.Loan_Balance.isna().sum()),
              "miss_score": int(df.Credit_Score.isna().sum()),
              "rows_any_miss": rows_any_miss, "rows_multi_miss": rows_multi_miss})
print(miss[miss.n_missing > 0])
print("blank/sentinel-like strings:", blank_scan or "none")
print("numeric sentinel-like counts (0/-1/99...):", sentinel_numeric)
print(f"rows with >=1 missing: {rows_any_miss} ({rows_any_miss/len(df)*100:.1f}%); "
      f">1 missing: {rows_multi_miss}")

# is missingness related to the target? (two-proportion test)
has_miss = df.isna().any(axis=1)
r1 = df.loc[has_miss, target].mean(); r0 = df.loc[~has_miss, target].mean()
p_miss_target = stats.fisher_exact([
    [int(df.loc[has_miss, target].sum()), int(has_miss.sum() - df.loc[has_miss, target].sum())],
    [int(df.loc[~has_miss, target].sum()), int((~has_miss).sum() - df.loc[~has_miss, target].sum())]])[1]
STATS.update({"delinq_rate_miss_rows": round(float(r1), 4),
              "delinq_rate_complete_rows": round(float(r0), 4),
              "miss_vs_target_p": round(float(p_miss_target), 4)})
print(f"delinquency rate | rows with missing: {r1:.4f} | complete rows: {r0:.4f} "
      f"| Fisher p={p_miss_target:.3f}")

# %% [markdown]
# ### C. Invalid / inconsistent values

# %%
emp_raw = df.Employment_Status.value_counts()
emp_norm = (df.Employment_Status.str.strip().str.lower()
            .replace({"emp": "employed"}).replace(
                {"employed": "Employed", "unemployed": "Unemployed",
                 "retired": "Retired", "self-employed": "Self-employed"}))
n_emp_relabelled = int((df.Employment_Status != emp_norm).sum())

util_gt1 = df[df.Credit_Utilization > 1]
tenure_incons = df[df.Account_Tenure > df.Age - 18]
score_out = df[(df.Credit_Score < 300) | (df.Credit_Score > 850)]

# Missed_Payments vs Month_* internal consistency
n_missed_hist = (df[month_cols] == "Missed").sum(axis=1)
n_late_hist = (df[month_cols] == "Late").sum(axis=1)
n_ontime_hist = (df[month_cols] == "On-time").sum(axis=1)
hist_agrees = int((n_missed_hist == df.Missed_Payments).sum())
hist_corr = float(df.Missed_Payments.corr(n_missed_hist))

# Debt-to-income ratio identity check (is DTI = Loan_Balance / Income?)
sub = df.dropna(subset=["Income", "Loan_Balance"])
dti_identity_corr = float(sub.Debt_to_Income_Ratio.corr(sub.Loan_Balance / sub.Income))

STATS.update({"emp_relabelled": n_emp_relabelled, "util_gt1": len(util_gt1),
              "tenure_incons": len(tenure_incons), "score_out_of_range": len(score_out),
              "hist_agrees": hist_agrees, "hist_corr": round(hist_corr, 4),
              "dti_identity_corr": round(dti_identity_corr, 4),
              "emp_raw_labels": int(df.Employment_Status.nunique()),
              "emp_norm_labels": int(emp_norm.nunique())})
print("Employment_Status raw labels:\n", emp_raw)
print(f"rows relabelled by normalisation: {n_emp_relabelled}")
print(f"Credit_Utilization > 1.0: {len(util_gt1)} records (max {df.Credit_Utilization.max():.4f})")
print(f"Account_Tenure > Age-18 (implausible tenure): {len(tenure_incons)}")
print(f"Credit_Score outside 300-850: {len(score_out)}")
print(f"Missed_Payments == count of 'Missed' months: {hist_agrees}/{len(df)} rows "
      f"(corr={hist_corr:.4f})")
print(f"corr(DTI, Loan_Balance/Income): {dti_identity_corr:.4f}  -> DTI is NOT derivable "
      "from these two columns")

# %% [markdown]
# ### D. Outlier detection (1.5 x IQR rule)

# %%
rows = []
outlier_flags = {}
for c in numeric_cols:
    s = df[c].dropna()
    q1, q3 = s.quantile(0.25), s.quantile(0.75)
    iqr = q3 - q1
    lo, hi = q1 - 1.5 * iqr, q3 + 1.5 * iqr
    n_out = int(((s < lo) | (s > hi)).sum())
    outlier_flags[c] = n_out
    rows.append({"Variable": c, "count": int(s.count()), "missing": int(df[c].isna().sum()),
                 "min": round(s.min(), 4), "q1": round(q1, 4), "median": round(s.median(), 4),
                 "mean": round(s.mean(), 4), "q3": round(q3, 4), "max": round(s.max(), 4),
                 "std": round(s.std(), 4), "skew": round(float(s.skew()), 3),
                 "IQR_outliers": n_out, "lower_fence": round(lo, 4), "upper_fence": round(hi, 4)})
desc_num = pd.DataFrame(rows)
print(desc_num.to_string(index=False))
print("total IQR outliers:", sum(outlier_flags.values()))

# %% [markdown]
# ### E. Data dictionary (definitions inferred — guide PDF unavailable)

# %%
DEFS = {
    "Customer_ID": ("Unique customer account identifier (key)", "Identifier"),
    "Age": ("Customer age in years", "Predictor / sensitive attribute"),
    "Income": ("Annual income (currency units as provided)", "Predictor"),
    "Credit_Score": ("Credit bureau score, observed range 301-847", "Predictor"),
    "Credit_Utilization": ("Revolving credit usage ratio (balance / limit)", "Predictor"),
    "Missed_Payments": ("Count of missed payments (0-6)", "Predictor / behavioural"),
    "Delinquent_Account": ("Target: 1 = account currently delinquent, 0 = not", "Target"),
    "Loan_Balance": ("Outstanding loan balance", "Predictor"),
    "Debt_to_Income_Ratio": ("Debt-service-to-income ratio", "Predictor"),
    "Employment_Status": ("Employment category (label variants found)", "Predictor / proxy-sensitive"),
    "Account_Tenure": ("Years the customer has held the account", "Predictor"),
    "Credit_Card_Type": ("Card product tier", "Predictor"),
    "Location": ("Customer city (US metro)", "Predictor / sensitive proxy"),
}
for i, m in enumerate(month_cols, 1):
    DEFS[m] = (f"Payment status in month {i} (On-time / Late / Missed)", "Predictor / behavioural")

dict_rows = []
for c in df.columns:
    s = df[c]
    definition, role = DEFS[c]
    issues = []
    if s.isna().any():
        issues.append(f"{int(s.isna().sum())} missing")
    if c == "Employment_Status":
        issues.append("label variants (EMP/employed/retired)")
    if c == "Credit_Utilization" and (s > 1).any():
        issues.append(f"{int((s>1).sum())} values > 1.0")
    if c == "Account_Tenure":
        issues.append(f"{int((s > df.Age-18).sum())} rows exceed plausible tenure for age")
    if c == "Missed_Payments":
        issues.append("inconsistent with Month_1..6 history")
    if c in month_cols:
        issues.append("near-uniform frequencies; uncorrelated with Missed_Payments")
    if c == "Debt_to_Income_Ratio":
        issues.append("not derivable from Loan_Balance/Income")
    if c == target:
        issues.append("16% positive class (imbalanced)")
    dict_rows.append({
        "Column": c, "Business definition (inferred)": definition, "Role": role,
        "Data type": str(s.dtype), "Missing n": int(s.isna().sum()),
        "Missing %": round(s.isna().mean() * 100, 2), "Unique values": int(s.nunique()),
        "Data quality issues": "; ".join(issues) if issues else "None observed",
    })
data_dict = pd.DataFrame(dict_rows)

# %% [markdown]
# ### Data quality summary + missing-data treatment

# %%
quality_rows = [
    ("Q1", "Missing values: Income", f"{STATS['miss_income']} of {len(df)} ({STATS['miss_income']/len(df)*100:.1f}%)",
     "Reduces sample for income-based features; risk of biased model if MNAR",
     "Median imputation on working copy + income_missing indicator; confirm with source system"),
    ("Q2", "Missing values: Loan_Balance", f"{STATS['miss_loan']} of {len(df)} ({STATS['miss_loan']/len(df)*100:.1f}%)",
     "Weakens debt exposure features", "Median imputation + loan_balance_missing indicator"),
    ("Q3", "Missing values: Credit_Score", f"{STATS['miss_score']} of {len(df)} ({STATS['miss_score']/len(df)*100:.1f}%)",
     "Negligible impact", "Median imputation (only 2 records affected)"),
    ("Q4", "Inconsistent category labels: Employment_Status",
     f"{STATS['emp_raw_labels']} raw labels -> {STATS['emp_norm_labels']} categories; {n_emp_relabelled} rows relabelled (EMP, 'employed', 'retired')",
     "Splits categories, biases group rates and model coefficients",
     "Standardised casing and mapped 'EMP' to 'Employed' in cleaned copy"),
    ("Q5", "Credit_Utilization > 1.0", f"{len(util_gt1)} records (max {df.Credit_Utilization.max():.4f})",
     "Possible over-limit accounts or unit error", "Retained; capped at 1.0 NOT applied - flagged for business confirmation"),
    ("Q6", "Implausible Account_Tenure for age (tenure > age - 18)",
     f"{len(tenure_incons)} records ({len(tenure_incons)/len(df)*100:.1f}%)",
     "Logical inconsistency between Age and Account_Tenure",
     "Flagged as tenurable_inconsistency; retain rows, escalate to data owner"),
    ("Q7", "Missed_Payments inconsistent with Month_1..6 payment history",
     f"agreement in only {hist_agrees}/{len(df)} rows (corr={hist_corr:.4f})",
     "Two conflicting behavioural sources; combined features may add noise",
     "Keep both sources; do not derive one from the other; investigate source systems"),
    ("Q8", "Debt_to_Income_Ratio not reproducible from Loan_Balance / Income",
     f"corr(DTI, Loan/Income) = {dti_identity_corr:.3f}",
     "Definition ambiguity complicates interpretation",
     "Use as provided; request the business formula from the data owner"),
    ("Q9", "Target class imbalance", f"{int(df[target].sum())} delinquent of {len(df)} ({df[target].mean()*100:.1f}%)",
     "Accuracy misleading; models may ignore positives",
     "Use stratified splits, PR-based metrics, class weighting in modelling phase"),
    ("Q10", "No duplicate records", f"{int(df.duplicated().sum())} duplicates; {df.Customer_ID.nunique()} unique IDs",
     "None", "No action required"),
    ("Q11", "No IQR outliers by 1.5xIQR rule", "0 outliers across 8 numeric variables",
     "No trimming needed", "Retain all values; distributions are near-uniform by construction"),
    ("Q12", "Missing guide PDF (Updated_Dataset_Description_Guide.pdf)", "File absent from project folder",
     "Variable definitions inferred from names/ranges, not authoritative",
     "Obtain guide and re-validate definitions before modelling"),
]
quality = pd.DataFrame(quality_rows, columns=[
    "ID", "Issue", "Measured extent", "Potential impact", "Recommended action"])

treatment_rows = [
    ("Income", f"{STATS['miss_income']} ({STATS['miss_income']/len(df)*100:.1f}%)",
     "Median imputation + Income_Missing indicator",
     "Near-symmetric but uniform-like distribution (skew 0.05); median is robust and equals mean here (~107.7k vs ~108.4k)",
     "Understates variance; impute inside cross-validation folds when modelling; verify missingness mechanism with source system"),
    ("Loan_Balance", f"{STATS['miss_loan']} ({STATS['miss_loan']/len(df)*100:.1f}%)",
     "Median imputation + Loan_Balance_Missing indicator",
     "Similar near-symmetric distribution (skew 0.11); median robust",
     "Same variance-understatement caveat; 4 records missing both Income and Loan_Balance"),
    ("Credit_Score", f"{STATS['miss_score']} ({STATS['miss_score']/len(df)*100:.1f}%)",
     "Median imputation (alternative: drop the 2 rows in modelling)",
     "Only 2 records affected; imputation keeps them for descriptive work",
     "Both records are also missing other fields; sensitivity check recommended"),
    ("Delinquent_Account (target)", "0 (0%)",
     "No treatment - no missing labels",
     "Target is complete; no imputation of labels attempted",
     "None"),
    ("Missed_Payments vs Month_1..6", "Not missing, but mutually inconsistent",
     "No imputation; both retained and flagged",
     "Missing/inconsistent payment history must not be interpreted as on-time payment",
     "Source-system reconciliation required before using combined behavioural features"),
]
treatment = pd.DataFrame(treatment_rows, columns=[
    "Variable", "Missing n (%)", "Selected treatment", "Reason", "Limitations"])
print(quality.to_string(index=False))
print(treatment.to_string(index=False))

# %% [markdown]
# ## Phase 3 — Build the cleaned working dataset (original preserved)

# %%
clean = RAW.copy()
clean_log = []
clean_log.append(("L00", "Working copy created from source; source file never modified",
                  "All rows", 0))

# L1: employment label standardisation
clean["Employment_Status_Clean"] = emp_norm
clean_log.append(("L1", "Employment_Status standardised: EMP/employed->Employed, retired->Retired (casing)",
                  f"{n_emp_relabelled} rows", n_emp_relabelled))

# L2: median imputation (medians computed on the observed data; for modelling, refit per fold)
for c in ["Income", "Loan_Balance", "Credit_Score"]:
    med = clean[c].median()
    n = int(clean[c].isna().sum())
    clean[f"{c}_Missing"] = clean[c].isna().astype(int)
    clean[c] = clean[c].fillna(med)
    clean_log.append((f"L2-{c}", f"Median imputation ({med:,.1f}) + {c}_Missing indicator",
                      f"{n} rows", n))

# L3: quality flags (values retained, not overwritten)
clean["Credit_Util_Above_100pct"] = (RAW.Credit_Utilization > 1).astype(int)
clean_log.append(("L3", "Flag Credit_Utilization > 1.0 (values retained)",
                  f"{len(util_gt1)} rows", len(util_gt1)))
clean["Tenure_Age_Inconsistency"] = (RAW.Account_Tenure > RAW.Age - 18).astype(int)
clean_log.append(("L4", "Flag Account_Tenure > Age-18 (values retained)",
                  f"{len(tenure_incons)} rows", len(tenure_incons)))

# L4: derived behavioural aggregates from the 6-month history (independent of Missed_Payments)
clean["N_Missed_6M"] = n_missed_hist
clean["N_Late_6M"] = n_late_hist
clean["N_OnTime_6M"] = n_ontime_hist
clean["Payment_Severity_Score"] = n_missed_hist * 2 + n_late_hist   # Missed=2, Late=1, On-time=0
clean_log.append(("L5", "Derived N_Missed_6M / N_Late_6M / N_OnTime_6M / Payment_Severity_Score from Month_1..6",
                  "All rows", len(clean)))

clean_log = pd.DataFrame(clean_log, columns=["Log ID", "Transformation", "Scope", "Rows affected"])
cleaned_path = OUT / "Geldium_Cleaned_Dataset.xlsx"
with pd.ExcelWriter(cleaned_path, engine="openpyxl") as w:
    clean.to_excel(w, sheet_name="Cleaned_Data", index=False)
    clean_log.to_excel(w, sheet_name="Cleaning_Log", index=False)
print(clean_log.to_string(index=False))
print("cleaned dataset:", cleaned_path, clean.shape)

# %% [markdown]
# ## Phase 4 — Exploratory analysis
#
# ### A. Univariate

# %%
categ_summary = []
for c in ["Employment_Status", "Employment_Status_Clean", "Credit_Card_Type", "Location",
          "Delinquent_Account"] + month_cols:
    vc = clean[c].value_counts(dropna=False)
    for k, v in vc.items():
        categ_summary.append({"Variable": c, "Category": k, "Count": int(v),
                              "Percent": round(v / len(clean) * 100, 1)})
categ_summary = pd.DataFrame(categ_summary)
print(categ_summary[categ_summary.Variable.isin(
    ["Employment_Status", "Credit_Card_Type", "Location", "Delinquent_Account"])].to_string(index=False))

delinq_rate = float(df[target].mean())
STATS["delinq_rate"] = round(delinq_rate, 4)
STATS["delinq_n"] = int(df[target].sum())
print(f"target base rate: {STATS['delinq_n']}/{len(df)} = {delinq_rate*100:.1f}%")

# %% [markdown]
# ### B. Bivariate — delinquency rate by meaningful groups

# %%
def group_rates(col, bins=None, labels=None):
    if bins is None:
        g = clean.groupby(col)[target].agg(n="size", delinquent="sum", rate="mean")
    else:
        tmp = pd.DataFrame({col: clean[col], target: clean[target]})
        tmp["band"] = pd.cut(tmp[col], bins=bins, labels=labels)
        g = tmp.groupby("band", observed=True)[target].agg(n="size", delinquent="sum", rate="mean")
    g = g.reset_index().rename(columns={col: "Group", "band": "Group"})
    g["rate"] = (g["rate"] * 100).round(2)
    return g

bands = {
    "Credit_Score": ([-1, 500, 650, 750, 1000], ["Poor (<500)", "Fair (500-649)", "Good (650-749)", "Very good (750+)"]),
    "Credit_Utilization": ([-1, .3, .5, .7, 2], ["0-30%", "30-50%", "50-70%", ">70%"]),
    "Debt_to_Income_Ratio": ([-1, .2, .3, .4, 1], ["<20%", "20-30%", "30-40%", ">40%"]),
    "Income": ([-1, 60000, 120000, 1e9], ["<60k", "60-120k", ">120k"]),
    "Age": ([-1, 30, 45, 60, 100], ["18-29", "30-44", "45-59", "60+"]),
    "Loan_Balance": ([-1, 25000, 50000, 75000, 1e9], ["<25k", "25-50k", "50-75k", ">75k"]),
    "Account_Tenure": ([-1, 5, 10, 15, 20], ["0-4", "5-9", "10-14", "15+"]),
    "Missed_Payments": ([-1, 1, 3, 6], ["0-1", "2-3", "4-6"]),
}
bivar = []
for c, (b, lab) in bands.items():
    g = group_rates(c, b, lab); g.insert(0, "Variable", c); bivar.append(g)
bivar_num = pd.concat(bivar, ignore_index=True)

bivar_cat = []
for c in ["Employment_Status_Clean", "Credit_Card_Type", "Location"]:
    g = group_rates(c); g.insert(0, "Variable", c); bivar_cat.append(g)
bivar_cat = pd.concat(bivar_cat, ignore_index=True)
bivariate = pd.concat([bivar_num, bivar_cat], ignore_index=True)
print(bivariate.to_string(index=False))

# significance screening: t / Mann-Whitney / point-biserial for numerics, chi2 for categoricals
tests = []
for c in numeric_cols:
    s = clean[[c, target]].dropna()
    a = s.loc[s[target] == 1, c]; b = s.loc[s[target] == 0, c]
    t, p_t = stats.ttest_ind(a, b, equal_var=False)
    _, p_mw = stats.mannwhitneyu(a, b)
    r, p_pb = stats.pointbiserialr(s[target], s[c])
    ranks = stats.rankdata(s[c]); y = s[target].values
    auc = (ranks[y == 1].sum() - y.sum() * (y.sum() + 1) / 2) / (y.sum() * (y == 0).sum())
    tests.append({"Variable": c, "mean_delinquent": round(a.mean(), 4), "mean_non_delinquent": round(b.mean(), 4),
                  "t_test_p": round(p_t, 4), "mannwhitney_p": round(p_mw, 4),
                  "point_biserial_r": round(r, 4), "univariate_AUC": round(float(auc), 3)})
for c in ["Employment_Status_Clean", "Credit_Card_Type", "Location"] + month_cols:
    ct = pd.crosstab(clean[c], clean[target])
    chi, p, dof, _ = stats.chi2_contingency(ct)
    tests.append({"Variable": c, "mean_delinquent": np.nan, "mean_non_delinquent": np.nan,
                  "t_test_p": np.nan, "mannwhitney_p": np.nan, "point_biserial_r": np.nan,
                  "univariate_AUC": np.nan, "chi2_p": round(p, 4)})
tests_df = pd.DataFrame(tests)
if "chi2_p" not in tests_df:
    tests_df["chi2_p"] = np.nan
print(tests_df.to_string(index=False))

# two-proportion z-tests for the most visible band contrasts
def two_prop_z(x1, n1, x0, n0):
    p1, p0 = x1 / n1, x0 / n0
    p = (x1 + x0) / (n1 + n0)
    z = (p1 - p0) / np.sqrt(p * (1 - p) * (1 / n1 + 1 / n0))
    return round(float(z), 3), round(float(2 * (1 - stats.norm.cdf(abs(z)))), 4)

contrasts = {}
for c, lo_lab, hi_lab in [("Debt_to_Income_Ratio", "<20%", ">40%"),
                          ("Credit_Utilization", "0-30%", "30-50%")]:
    g = bivar_num[bivar_num.Variable == c].set_index("Group")
    contrasts[f"{c}: {hi_lab} vs {lo_lab}"] = two_prop_z(
        int(g.loc[hi_lab, "delinquent"]), int(g.loc[hi_lab, "n"]),
        int(g.loc[lo_lab, "delinquent"]), int(g.loc[lo_lab, "n"]))
for c, a_lab, b_lab in [("Employment_Status_Clean", "Unemployed", "Retired"),
                        ("Credit_Card_Type", "Business", "Platinum"),
                        ("Location", "Los Angeles", "New York")]:
    g = bivar_cat[bivar_cat.Variable == c].set_index("Group")
    contrasts[f"{c}: {a_lab} vs {b_lab}"] = two_prop_z(
        int(g.loc[a_lab, "delinquent"]), int(g.loc[a_lab, "n"]),
        int(g.loc[b_lab, "delinquent"]), int(g.loc[b_lab, "n"]))
print("two-proportion z-tests:", contrasts)

# %% [markdown]
# ### C. Correlation analysis

# %%
pearson = clean[numeric_cols].corr(method="pearson")
spearman = clean[numeric_cols].corr(method="spearman")
corr_long = pearson.round(4).where(np.triu(np.ones(pearson.shape), k=1).astype(bool)).stack().reset_index()
corr_long.columns = ["Variable 1", "Variable 2", "Pearson r"]
corr_long["abs_r"] = corr_long["Pearson r"].abs()
corr_long = corr_long.sort_values("abs_r", ascending=False)
print(corr_long.head(8).to_string(index=False))
print("max |r| among predictors:", corr_long.abs_r.max())

# %% [markdown]
# ### D. Charts

# %%
def savefig(fig, name):
    p = CHARTS / name
    fig.savefig(p, bbox_inches="tight")
    plt.close(fig)
    return p

# 1) missing values
fig, ax = plt.subplots(figsize=(7, 3.6))
mm = miss[miss.n_missing > 0].sort_values("n_missing", ascending=False)
bars = ax.bar(mm.index, mm.n_missing, color="#C0504D")
ax.bar_label(bars, labels=[f"{v} ({v/len(df)*100:.1f}%)" for v in mm.n_missing], fontsize=9)
ax.set_title("Missing values by variable (n = 500 records)")
ax.set_ylabel("Records with missing value"); ax.set_xlabel("")
savefig(fig, "fig1_missing_values.png")

# 2) target distribution
fig, ax = plt.subplots(figsize=(5.4, 3.6))
vc = df[target].value_counts().sort_index()
bars = ax.bar(["0 - Not delinquent", "1 - Delinquent"], vc.values, color=["#4F81BD", "#C0504D"])
ax.bar_label(bars, labels=[f"{v} ({v/len(df)*100:.1f}%)" for v in vc.values], fontsize=10)
ax.set_title(f"Target class distribution - base delinquency rate {delinq_rate*100:.1f}%")
ax.set_ylabel("Records")
savefig(fig, "fig2_target_distribution.png")

# 3) correlation heatmap
fig, ax = plt.subplots(figsize=(7.6, 6))
sns.heatmap(pearson, annot=True, fmt=".2f", cmap="RdBu_r", center=0, vmin=-1, vmax=1,
            square=True, cbar_kws={"label": "Pearson r"}, ax=ax)
ax.set_title("Pearson correlations among numeric variables\n(all |r| <= 0.10: no meaningful linear relationships)")
savefig(fig, "fig3_correlation_heatmap.png")

# 4) delinquency rate by DTI band
fig, ax = plt.subplots(figsize=(6.6, 3.8))
g = bivar_num[bivar_num.Variable == "Debt_to_Income_Ratio"]
bars = ax.bar(g.Group, g.rate, color="#4F81BD")
ax.bar_label(bars, labels=[f"{r:.1f}%\n(n={n})" for r, n in zip(g.rate, g.n)], fontsize=9)
ax.axhline(delinq_rate * 100, ls="--", color="grey", lw=1)
ax.text(3.4, delinq_rate * 100 + 0.6, f"overall {delinq_rate*100:.1f}%", fontsize=8, color="grey")
ax.set_title("Delinquency rate by debt-to-income band (descriptive only;\nband differences not statistically significant, z-test p > 0.05)")
ax.set_ylabel("Delinquency rate (%)"); ax.set_xlabel("Debt-to-Income ratio band")
savefig(fig, "fig4_delinq_by_dti.png")

# 5) key numeric distributions by target
fig, axes = plt.subplots(2, 2, figsize=(9, 6.4))
for ax, c in zip(axes.ravel(), ["Credit_Utilization", "Debt_to_Income_Ratio", "Credit_Score", "Missed_Payments"]):
    sns.boxplot(data=clean, x=target, y=c, hue=target, legend=False, ax=ax, palette=["#4F81BD", "#C0504D"])
    ax.set_xticklabels(["0 - Not delinquent", "1 - Delinquent"])
    ax.set_title(c.replace("_", " "))
fig.suptitle("Distribution of key risk variables by delinquency status (largely overlapping)", y=1.00, fontsize=12)
fig.tight_layout()
savefig(fig, "fig5_boxplots_by_target.png")

# 6) missed payments vs delinquency rate
fig, ax = plt.subplots(figsize=(7, 3.8))
g = clean.groupby("Missed_Payments")[target].agg(n="size", rate="mean").reset_index()
ax.plot(g.Missed_Payments, g.rate * 100, marker="o", color="#C0504D")
for _, r in g.iterrows():
    ax.annotate(f"{r.rate*100:.1f}%\n(n={r.n})", (r.Missed_Payments, r.rate * 100),
                textcoords="offset points", xytext=(0, 8), ha="center", fontsize=8)
ax.axhline(delinq_rate * 100, ls="--", color="grey", lw=1)
ax.set_title("Delinquency rate by missed-payment count - no monotonic trend\n(Mann-Whitney p = 0.58)")
ax.set_xlabel("Missed_Payments (0-6)"); ax.set_ylabel("Delinquency rate (%)")
ax.set_ylim(0, 30)
savefig(fig, "fig6_missed_payments_vs_delinq.png")

# 7) employment label inconsistency (data-quality chart)
fig, ax = plt.subplots(figsize=(7.2, 3.6))
order = ["Employed", "employed", "EMP", "Unemployed", "retired", "Retired", "Self-employed"]
raw_counts = df.Employment_Status.value_counts()
xs = np.arange(len(order))
b1 = ax.bar(xs - 0.2, [int(raw_counts.get(k, 0)) for k in order], width=0.4,
            label="Raw labels (6)", color="#C0504D")
norm_counts = emp_norm.value_counts()
order2 = ["Employed", "Unemployed", "Retired", "Self-employed"]
b2 = ax.bar(np.arange(len(order2)) + 0.6, [int(norm_counts[k]) for k in order2], width=0.4,
            label="Standardised (4)", color="#4F81BD")
ax.bar_label(b1, fontsize=8); ax.bar_label(b2, fontsize=8)
ax.set_xticks(list(xs - 0.2) + list(np.arange(len(order2)) + 0.6))
ax.set_xticklabels(order + order2, rotation=30, ha="right", fontsize=8)
ax.set_ylabel("Records"); ax.legend()
ax.set_title("Data quality: Employment_Status label inconsistency before/after standardisation")
savefig(fig, "fig7_employment_labels.png")

# 8) payment history monthly distribution
fig, ax = plt.subplots(figsize=(7.2, 3.6))
hist = pd.DataFrame({c: df[c].value_counts(normalize=True) * 100 for c in month_cols}).T
hist = hist[["On-time", "Late", "Missed"]]
hist.plot(kind="bar", stacked=True, color=["#4F81BD", "#F79646", "#C0504D"], ax=ax)
ax.set_title("6-month payment history: near-uniform status mix in every month")
ax.set_ylabel("% of records"); ax.set_xlabel("Month")
ax.legend(title="Status", bbox_to_anchor=(1.02, 1), loc="upper left")
ax.set_ylim(0, 100)
savefig(fig, "fig8_monthly_history.png")

# 9) single composite figure used in the Word report (2x2)
fig, axes = plt.subplots(2, 2, figsize=(11, 7.6))
ax = axes[0, 0]
vc = df[target].value_counts().sort_index()
bars = ax.bar(["0 - Not delinquent", "1 - Delinquent"], vc.values, color=["#4F81BD", "#C0504D"])
ax.bar_label(bars, labels=[f"{v} ({v/len(df)*100:.1f}%)" for v in vc.values], fontsize=9)
ax.set_title("(a) Target: 16.0% delinquent (80/500)", fontsize=10)
ax.set_ylabel("Records")
ax = axes[0, 1]
g = bivar_num[bivar_num.Variable == "Debt_to_Income_Ratio"]
bars = ax.bar(g.Group, g.rate, color="#4F81BD")
ax.bar_label(bars, labels=[f"{r:.1f}%\n(n={n})" for r, n in zip(g.rate, g.n)], fontsize=8)
ax.axhline(delinq_rate * 100, ls="--", color="grey", lw=1)
ax.set_title("(b) Delinquency by DTI band (descriptive only,\nnot statistically significant)", fontsize=10)
ax.set_ylabel("Delinquency rate (%)"); ax.set_xlabel("Debt-to-Income band")
ax = axes[1, 0]
sns.heatmap(pearson, annot=True, fmt=".2f", cmap="RdBu_r", center=0, vmin=-1, vmax=1,
            cbar_kws={"shrink": 0.8}, ax=ax, annot_kws={"fontsize": 6})
ax.set_title("(c) Predictor correlations are negligible (|r| <= 0.10)", fontsize=10)
ax = axes[1, 1]
g = clean.groupby("Missed_Payments")[target].agg(n="size", rate="mean").reset_index()
ax.plot(g.Missed_Payments, g.rate * 100, marker="o", color="#C0504D")
for _, r in g.iterrows():
    ax.annotate(f"{r.rate*100:.1f}%", (r.Missed_Payments, r.rate * 100),
                textcoords="offset points", xytext=(0, 8), ha="center", fontsize=8)
ax.axhline(delinq_rate * 100, ls="--", color="grey", lw=1)
ax.set_title("(d) No trend vs missed-payment count\n(Mann-Whitney p = 0.58)", fontsize=10)
ax.set_xlabel("Missed_Payments (0-6)"); ax.set_ylabel("Delinquency rate (%)")
ax.set_ylim(0, 30)
fig.suptitle("Key exploratory views of the Geldium delinquency dataset (n = 500)", fontsize=12)
fig.tight_layout(rect=[0, 0, 1, 0.97])
savefig(fig, "fig9_report_overview.png")
print("charts saved to", CHARTS)

# %% [markdown]
# ## Phase 5 — Candidate risk indicators (screened, NOT validated)
#
# **Critical caveat:** no predictor in this dataset shows a statistically significant
# association with `Delinquent_Account` (all p > 0.10; all univariate AUC 0.47-0.54).
# The ranking below is descriptive only.

# %%
risk_rows = [
    ("1", "Debt_to_Income_Ratio",
     "Highest delinquency in the >40% band (18.9%) vs lowest in <20% (9.8%) - ~2x gradient",
     "<20%: 8/82 = 9.8%; >40%: 14/74 = 18.9%; point-biserial r = 0.034 (p = 0.44); z-test p = 0.10",
     "Higher debt burden leaves less income buffer for EMI payments",
     "Overall association not significant; non-monotonic middle bands; DTI definition unverified",
     "Validate in a model with proper splits; confirm DTI formula with data owner"),
    ("2", "Employment_Status",
     "Unemployed 19.4% delinquent vs Retired 11.5%",
     "Unemployed 18/93; Retired 10/87; chi2 p = 0.78; z-test p = 0.15",
     "Income instability plausibly raises delinquency risk",
     "Not significant; label inconsistency (EMP/employed) had to be fixed first; proxy-sensitive attribute",
     "Re-test after label fix and with larger sample; run fairness checks"),
    ("3", "Credit_Card_Type",
     "Business card 21.3% vs Platinum 11.8%",
     "Business 23/108; Platinum 9/76; chi2 p = 0.35; z-test p = 0.10",
     "Business products may carry different underwriting and usage patterns",
     "Not significant; product mix confounded with employment and income",
     "Stratify by product in modelling; check small-cell stability"),
    ("4", "Location",
     "Los Angeles 19.6% vs New York 12.0%",
     "LA 21/107; NY 11/92; chi2 p = 0.66",
     "Regional economic conditions could differ",
     "Not significant; location is a fairness-sensitive proxy (ethnicity/socioeconomics)",
     "Treat cautiously; test model fairness with and without Location"),
    ("5", "Age / Account_Tenure",
     "45-59 age band highest (20.0%); tenure 10-14 lowest (10.8%)",
     "Age point-biserial r = 0.023 (p = 0.62); tenure r = -0.040 (p = 0.37)",
     "Relationships, if real, would reflect life-stage exposure and product maturity",
     "Not significant; 81 rows have tenure inconsistent with age",
     "Fix tenure-age inconsistency at source before use; age is a protected attribute"),
]
risk = pd.DataFrame(risk_rows, columns=[
    "Rank", "Candidate indicator", "Observed pattern", "Key statistics",
    "Possible business explanation", "Limitations / alternative explanations", "Recommended validation"])
print(risk.to_string(index=False))

# %% [markdown]
# ## Phase 8 (part 1) — Analysis workbook

# %%
missing_report = miss.reset_index().rename(columns={"index": "Variable", "n_missing": "Missing n",
                                                    "pct_missing": "Missing %"})
missing_report["Present n"] = len(df) - missing_report["Missing n"]
missing_report["Present %"] = (100 - missing_report["Missing %"]).round(2)
missing_report = missing_report[["Variable", "Missing n", "Missing %", "Present n", "Present %"]]

finding_rows = [
    ("F1", "No predictor is significantly associated with the delinquency target",
     "All point-biserial |r| <= 0.05; t/Mann-Whitney/chi2 p-values all > 0.10; univariate AUC 0.472-0.536",
     "The dataset as delivered offers no validated risk signal; predictive modelling may perform at chance level",
     "Obtain a dataset with real outcome linkage, or confirm whether the target was randomly generated"),
    ("F2", "Predictors are mutually independent (|r| <= 0.10 for all pairs)",
     "Pearson and Spearman correlation matrices both near zero; max |r| = 0.098 (Account_Tenure vs Missed_Payments)",
     "Unusual for real credit data (score, utilisation, DTI usually correlate); suggests synthetic/random generation",
     "Ask the data owner how the sample was generated"),
    ("F3", "Missed_Payments contradicts the 6-month payment history",
     "Agreement in only 72/500 rows (14.4%, close to chance given ~37% 'Missed' months); corr = 0.007",
     "Two conflicting behavioural sources; using both may add noise or double-count",
     "Reconcile source systems; establish which field is authoritative"),
    ("F4", "Debt_to_Income_Ratio cannot be derived from Loan_Balance / Income",
     "corr(DTI, Loan/Income) = 0.047; mean Loan/Income = 0.69 vs mean DTI = 0.30",
     "Definition ambiguity (likely monthly obligations vs monthly income)",
     "Request the official formula before interpretation"),
    ("F5", "Class imbalance",
     "80 delinquent vs 420 non-delinquent (16.0% positive)",
     "Accuracy would be 84% for a naive 'never delinquent' model",
     "Use stratified CV, PR-AUC, class weights; consider threshold tuning"),
    ("F6", "Logical inconsistency: Account_Tenure exceeds plausible age",
     "81 rows (16.2%) with tenure > age - 18, incl. 18-year-olds with 12-18 year tenures",
     "Undermines tenure-based risk interpretation",
     "Escalate to data owner; flag column in modelling"),
    ("F7", "Credit_Utilization slightly above 1.0",
     "4 records between 1.002 and 1.026",
     "Either genuine over-limit usage or a unit/cap error",
     "Confirm business rule; retain unless confirmed erroneous"),
    ("F8", "Missingness is not target-related but clusters mildly",
     "64 rows (12.8%) missing >=1 value; 6 rows missing >1; delinquency 17.2% vs 15.8% (Fisher p = 0.72)",
     "MCAR-like (Fisher p = 0.72); median imputation + indicators is reasonable",
     "Confirm with source system; keep indicators for the model"),
]
findings = pd.DataFrame(finding_rows, columns=[
    "ID", "Finding", "Evidence", "Implication", "Recommended follow-up"])

analysis_xlsx = OUT / "Geldium_EDA_Analysis.xlsx"
with pd.ExcelWriter(analysis_xlsx, engine="openpyxl") as w:
    data_dict.to_excel(w, sheet_name="Data_Dictionary", index=False)
    structure.to_excel(w, sheet_name="Dataset_Structure", index=False)
    quality.to_excel(w, sheet_name="Data_Quality_Summary", index=False)
    missing_report.to_excel(w, sheet_name="Missing_Value_Report", index=False)
    treatment.to_excel(w, sheet_name="Missing_Data_Treatment", index=False)
    desc_num.to_excel(w, sheet_name="Descriptive_Statistics", index=False)
    categ_summary.to_excel(w, sheet_name="Category_Frequencies", index=False)
    bivariate.to_excel(w, sheet_name="Delinquency_By_Group", index=False)
    tests_df.to_excel(w, sheet_name="Statistical_Tests", index=False)
    pearson.round(4).to_excel(w, sheet_name="Correlations_Pearson", index=True)
    spearman.round(4).to_excel(w, sheet_name="Correlations_Spearman", index=True)
    corr_long.head(10).to_excel(w, sheet_name="Top_Correlations", index=False)
    findings.to_excel(w, sheet_name="Unexpected_Findings", index=False)
    risk.to_excel(w, sheet_name="Risk_Indicator_Ranking", index=False)
    clean_log.to_excel(w, sheet_name="Cleaning_Log", index=False)
print("workbook saved:", analysis_xlsx)

with open(OUT / "verified_stats.json", "w") as f:
    json.dump(STATS, f, indent=2)
print(json.dumps(STATS, indent=2))

# %% [markdown]
# ## Phase 7 — Official Word report (template-based)

# %%
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

TPL = BASE / "EDA_SummaryReport_Template.docx"
docx_path = OUT / "Geldium_EDA_Summary_Report.docx"
doc = Document(TPL)


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


def add_table_after(anchor, header, rows, widths=None, font_size=8.5, caption=None):
    """Insert a bordered table (optionally with caption) after paragraph `anchor`."""
    t = doc.add_table(rows=1, cols=len(header))
    t.autofit = True
    for j, h in enumerate(header):
        cell = t.rows[0].cells[j]
        cell.text = str(h)
        for r in cell.paragraphs[0].runs:
            r.bold = True; r.font.size = Pt(font_size)
        shade_cell(cell)
        set_cell_border(cell)
    for row in rows:
        cells = t.add_row().cells
        for j, v in enumerate(row):
            cells[j].text = "" if v is None else str(v)
            for p in cells[j].paragraphs:
                for r in p.runs:
                    r.font.size = Pt(font_size)
            set_cell_border(cells[j])
    anchor._p.addnext(t._tbl)
    if caption:
        cap = doc.add_paragraph()
        run = cap.add_run(caption); run.italic = True; run.font.size = Pt(8)
        t._tbl.addnext(cap._p)
    return t


def add_par_after(anchor, text, style="Normal", bold=False, size=None, italic=False):
    p = doc.add_paragraph(style=style)
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    if size:
        run.font.size = Pt(size)
    anchor._p.addnext(p._p)
    return p


def add_image_after(anchor, path, width_in=6.0, caption=None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(path), width=Inches(width_in))
    anchor._p.addnext(p._p)
    last_p = p
    if caption:
        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = cap.add_run(caption); run.italic = True; run.font.size = Pt(8)
        p._p.addnext(cap._p)
        last_p = cap
    return last_p


paras = doc.paragraphs

def find_par(fragment):
    for p in doc.paragraphs:
        if fragment in p.text:
            return p
    raise KeyError(fragment)

# --- Tata logo + Title (branding added reproducibly; template content otherwise preserved)
logo_p = doc.add_paragraph()
logo_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
logo_p.add_run().add_picture(str(BASE / "assets" / "tata_logo.png"), width=Inches(0.95))
paras[0]._p.addprevious(logo_p._p)
paras[0].text = "Geldium Delinquency Prediction - Exploratory Data Analysis (EDA) Summary Report"
paras[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
for r in paras[0].runs:
    r.bold = True

# --- 1. Introduction
p_intro = find_par("[Insert purpose")
p_intro.text = (
    "This report presents the Exploratory Data Analysis (EDA) of the Geldium delinquency prediction "
    "dataset, prepared for the Tata iQ analytics team as part of the forage simulation. Geldium, a "
    "consumer-credit lender, aims to predict which customer accounts will become delinquent so that "
    "proactive, loss-mitigating outreach can be prioritised. Before any predictive model is built, the "
    "data must be understood, its quality quantified, and its usability for modelling judged. "
    "The analysis covers 500 customer records and 19 variables from the delivered workbook "
    "(single sheet, 'Delinquency_prediction_dataset'), including demographic attributes, credit "
    "behaviour, account characteristics and a six-month payment history. Objectives were to: (1) profile "
    "structure and content, (2) measure data quality issues, (3) define a justified missing-data "
    "treatment, (4) explore relationships with the delinquency target, and (5) shortlist candidate risk "
    "indicators for later validation. The source file was preserved unchanged; all cleaning was applied "
    "to a separate working copy with a full transformation log. The dataset description guide "
    "(Updated_Dataset_Description_Guide.pdf) was not available, so variable definitions are inferred from "
    "column names and observed value ranges and should be confirmed with the data owner."
)

# --- 2. Dataset Overview
p_kv = find_par("- Number of records:")
p_kv.text = f"- Number of records: {len(df)} (no duplicate rows; {df.Customer_ID.nunique()} unique Customer_IDs)"
p_key = find_par("- Key variables:")
p_key.text = (
    "- Key variables: Delinquent_Account (target, 1 = delinquent); credit behaviour: Credit_Score, "
    "Credit_Utilization, Missed_Payments, Month_1..Month_6 payment status; debt/exposure: Income, "
    "Loan_Balance, Debt_to_Income_Ratio; demographics/context: Age, Employment_Status, Location; "
    "account: Account_Tenure, Credit_Card_Type; Customer_ID (identifier)"
)
p_dt = find_par("- Data types:")
p_dt.text = (
    "- Data types: 8 numeric (Age, Income, Credit_Score, Credit_Utilization, Missed_Payments, "
    "Loan_Balance, Debt_to_Income_Ratio, Account_Tenure), 1 binary integer target "
    "(Delinquent_Account), 9 categorical (Employment_Status, Credit_Card_Type, Location, Month_1..6), "
    "1 identifier (Customer_ID)"
)
anchor = p_dt
anchor = add_par_after(anchor, "Headline metrics (Table 1):", bold=True)
hdr = ["Attribute", "Value", "Attribute", "Value"]
rows = [
    ["Records x variables", f"500 x 19", "Target base rate", f"{STATS['delinq_n']}/500 = 16.0% delinquent"],
    ["Unique IDs / duplicates", "500 / 0", "Rows with >=1 missing value", f"64 (12.8%)"],
    ["Numeric / categorical vars", "9 (incl. target) / 9", "Columns with missing data", "Income 7.8%, Loan_Balance 5.8%, Credit_Score 0.4%"],
    ["Age (mean, range)", "46.3 yrs (18-74)", "Income (median, range)", "107,658 (15,404-199,943)"],
    ["Credit_Score (mean, range)", "577.7 (301-847)", "Credit_Utilization (mean, range)", "0.49 (0.05-1.03)"],
    ["IQR outliers detected", "0 of 8 numeric vars", "Logical inconsistencies", "Tenure>age-18: 81 rows; Missed_Payments vs months: 428 rows disagree"],
]
add_table_after(anchor, hdr, rows, caption="Table 1: Dataset structure and headline quality metrics (n = 500).")

# add anomalies bullet list right after the Table-1 caption paragraph
cap1 = [p for p in doc.paragraphs if p.text.startswith("Table 1:")][0]
a = add_par_after(cap1, "Significant inconsistencies and anomalies:", bold=True)
a = add_par_after(a, "- Employment_Status contains 6 raw label variants ('EMP', 'employed', 'retired' etc.) that collapse to 4 categories; 245 rows required standardisation.", style="Normal")
a = add_par_after(a, "- Missed_Payments (0-6) contradicts the Month_1..6 payment history in 428 of 500 rows (agreement 14.4%, corr = 0.01) - two conflicting behavioural sources.", style="Normal")
a = add_par_after(a, "- Account_Tenure exceeds the plausible maximum for the customer's age in 81 rows (e.g., 18-year-olds with 12-18 years tenure).", style="Normal")
a = add_par_after(a, "- Debt_to_Income_Ratio cannot be reproduced from Loan_Balance / Income (corr = 0.05); its business formula must be confirmed.", style="Normal")
a = add_par_after(a, "- 4 records show Credit_Utilization slightly above 1.0 (max 1.026) - possible over-limit usage or a unit error; values were retained and flagged.", style="Normal")
a = add_par_after(a, "- No duplicate records, no negative or out-of-range values, no IQR outliers, and no sentinel values (e.g., 999/-1) were detected.", style="Normal")

# --- 3. Missing Data Analysis
p_var = find_par("- Variables with missing values:")
p_var.text = (
    "- Variables with missing values: Income - 39 records (7.8%); Loan_Balance - 29 records (5.8%); "
    "Credit_Score - 2 records (0.4%). All other 16 columns, including the target, are complete. "
    "In total 64 records (12.8%) miss at least one value and 6 records miss more than one. "
    "No blank or whitespace-only strings and no sentinel placeholders were found. Missingness is not "
    "related to the target (delinquency 17.2% in affected rows vs 15.8% in complete rows, Fisher p = 0.72), "
    "which is consistent with a missing-completely-at-random pattern."
)
p_tr = find_par("- Missing data treatment:")
p_tr.text = (
    "- Missing data treatment: median imputation for Income, Loan_Balance and Credit_Score on the "
    "cleaned working copy, each accompanied by a binary missingness indicator "
    "(Income_Missing, Loan_Balance_Missing, Credit_Score_Missing). All three distributions are close to "
    "symmetric (skew 0.05-0.11), so medians are stable and here coincide closely with the means. "
    "No rows were deleted, no synthetic customers were generated, no missing delinquency labels were "
    "imputed (none exist), and missing payment history was never treated as successful payment. "
    "When modelling begins, imputation parameters must be re-fitted inside each training fold to avoid "
    "leakage."
)
anchor = add_par_after(p_tr, "Missing-data treatment summary:", bold=True)
hdr = ["Variable", "Missing", "Treatment", "Rationale / limitation"]
rows = [
    ["Income", "39 (7.8%)", "Median imputation + indicator", "Near-symmetric (skew 0.05); understates variance; refit per fold"],
    ["Loan_Balance", "29 (5.8%)", "Median imputation + indicator", "Near-symmetric (skew 0.11); same caveat"],
    ["Credit_Score", "2 (0.4%)", "Median imputation", "Only 2 records; sensitivity check recommended"],
    ["Target (Delinquent_Account)", "0 (0%)", "None required", "Complete; no label imputation attempted"],
]
t = add_table_after(anchor, hdr, rows, caption="Table 2: Missing-data treatment plan (applied to the cleaned working copy).")
add_par_after([p for p in doc.paragraphs if p.text.startswith("Table 2:")][0],
              "Remaining concerns: the missingness mechanism could not be confirmed without the source systems, "
              "and imputed values must not be treated as observed facts in downstream modelling.", italic=True, size=9)

# --- 4. Key Findings and Risk Indicators
p_corr = find_par("- Correlations observed between key variables:")
p_corr.text = (
    "- Correlations observed between key variables: none of the predictors correlates meaningfully with "
    "the delinquency target - all point-biserial correlations are |r| <= 0.05 (p > 0.33) and all "
    "univariate AUCs fall between 0.47 and 0.54 (chance level). Predictors are also mutually "
    "independent: every pairwise Pearson and Spearman correlation is |r| <= 0.10, which is unusual for "
    "real credit data and suggests the sample may be synthetic or the target randomly assigned. "
    "Group-level rates vary (e.g., debt-to-income >40%: 18.9% vs <20%: 9.8%; unemployed 19.4% vs retired "
    "11.5%; Business cards 21.3% vs Platinum 11.8%), but no contrast is statistically significant "
    "(all two-proportion z-tests p > 0.05)."
)
p_unexp = find_par("- Unexpected anomalies:")
p_unexp.text = (
    "- Unexpected anomalies: (a) the target appears statistically independent of every recorded "
    "predictor, so a model built on this file would likely perform at chance - this must be clarified "
    "with the data owner before modelling; (b) Missed_Payments disagrees with the six monthly payment "
    "statuses in 428 of 500 rows, indicating two conflicting behavioural feeds; (c) the class "
    "imbalance (16.0% positive) means a naive 'never delinquent' classifier already achieves 84% "
    "accuracy; (d) 81 rows have an account tenure that is impossible given the customer's age."
)
anchor = add_par_after(p_unexp, "Candidate risk indicators (descriptive screening only):", bold=True)
hdr = ["Rank", "Indicator", "Observed pattern", "Key statistics"]
rows = [
    ["1", "Debt_to_Income_Ratio", "Delinquency 18.9% in >40% band vs 9.8% in <20% band",
     "r = 0.034 (p = 0.44); z-test p = 0.10"],
    ["2", "Employment_Status", "Unemployed 19.4% vs Retired 11.5%",
     "chi2 p = 0.78; z-test p = 0.15"],
    ["3", "Credit_Card_Type", "Business 21.3% vs Platinum 11.8%",
     "chi2 p = 0.35; z-test p = 0.10"],
    ["4", "Location (fairness-sensitive)", "Los Angeles 19.6% vs New York 12.0%",
     "chi2 p = 0.66"],
    ["5", "Age / Account_Tenure (protected/proxy)", "45-59 band highest (20.0%); tenure 10-14 lowest (10.8%)",
     "r = 0.023 / -0.040 (p = 0.62 / 0.37)"],
]
add_table_after(anchor, hdr, rows,
                caption="Table 3: Ranked candidate risk indicators - none is statistically significant; "
                        "definitive predictive importance requires model validation on a dataset with "
                        "genuine outcome linkage.")

# charts in section 4
p3cap = [p for p in doc.paragraphs if p.text.startswith("Table 3:")][0]
c1 = add_image_after(p3cap, CHARTS / "fig9_report_overview.png", width_in=6.3,
                     caption="Figure 1: Key exploratory views - (a) target class imbalance, (b) delinquency rate by "
                             "debt-to-income band (descriptive only, not significant), (c) negligible predictor "
                             "correlations, and (d) no trend between missed-payment count and delinquency.")

# --- 5. AI & GenAI Usage
p_ai = find_par("Generative AI tools were used")
p_ai.text = (
    "A generative-AI coding assistant (LLM agent with Python execution - pandas, numpy, scipy, "
    "matplotlib/seaborn, openpyxl and python-docx) supported this analysis by drafting and running the "
    "profiling, statistical-testing and reporting code and by suggesting narrative summaries. Every "
    "number in this report was computed from the actual dataset by executed code; none was generated by "
    "the model. AI suggested candidate imputation strategies (median vs mean vs mode, missingness "
    "indicators), which the analyst verified against the observed distributions (skewness 0.05-0.11, "
    "near-identical mean/median) before adopting. AI did not generate synthetic data and did not modify "
    "the source file. Human validation consisted of re-running the statistics, cross-checking counts "
    "against the raw workbook, and reviewing the analysis for leakage, bias and unsupported claims - "
    "which is why this report explicitly states that no risk indicator is statistically validated."
)
p_pr = find_par("- 'Summarize key patterns")
p_pr.text = "- 'Summarise the dataset structure, missing values, anomalies and data quality risks; verify every figure against the file.'"
p_pr2 = find_par("- 'Suggest an imputation strategy")
p_pr2.text = "- 'Recommend an appropriate missing-data treatment for each affected variable and justify it against the observed distribution.'"
p_pr3 = add_par_after(p_pr2, "- 'Identify variables associated with delinquency, report the actual statistics, and flag anything that is not statistically significant.'")
p_pr4 = add_par_after(p_pr3, "- 'Review the analysis for potential target leakage, selection bias and unsupported conclusions before finalising the report.'")
add_par_after(p_pr4,
              "AI-generated suggestions were used as hypotheses only; conclusions were retained only where "
              "the executed statistics supported them.", italic=True, size=9)

# --- 6. Conclusion & Next Steps
p_conc = find_par("[Summarize key findings")
p_conc.text = (
    "Overall, the Geldium dataset is structurally clean but not yet modelling-ready. Structure is sound: "
    "500 unique customers, no duplicates, no invalid ranges, and a complete delinquency target (16.0% "
    "positive). Missingness is limited and treatable (Income 7.8%, Loan_Balance 5.8%, Credit_Score 0.4%), "
    "and a documented median-imputation-plus-indicator strategy has been applied to a separate cleaned "
    "copy. The blocking issues are analytical rather than cosmetic: no predictor shows a statistically "
    "significant relationship with delinquency, predictors are mutually independent, and "
    "Missed_Payments conflicts with the monthly payment history - together these suggest the target may "
    "not be genuinely linked to the behavioural fields. Recommended next steps: (1) resolve the critical "
    "quality issues at source (employment-label standardisation, tenure-age inconsistencies, the "
    "Missed_Payments vs monthly-history conflict, and the Debt_to_Income_Ratio definition); (2) confirm "
    "how the dataset and target were generated before investing in modelling; (3) apply the documented "
    "imputation strategy strictly inside training folds and keep the missingness indicators; (4) validate "
    "the five candidate risk indicators with proper train/test validation rather than relying on these "
    "descriptive screens; (5) in the modelling phase evaluate discrimination (PR-AUC given the 16% base "
    "rate), calibration, and fairness across Age, Location and Employment_Status, and monitor performance "
    "by segment."
)

doc.core_properties.title = "Geldium EDA Summary Report"
doc.core_properties.author = "Tata iQ Forage Simulation - Data Analytics"
doc.save(docx_path)
print("report saved:", docx_path)

# %% [markdown]
# ## Phase 9 — QA cross-checks

# %%
chk = {
    "rows": len(df), "cols": df.shape[1], "dup": int(df.duplicated().sum()),
    "income_missing": int(df.Income.isna().sum()), "loan_missing": int(df.Loan_Balance.isna().sum()),
    "score_missing": int(df.Credit_Score.isna().sum()),
    "target_pos": int(df[target].sum()), "target_rate": round(float(df[target].mean()), 4),
    "util_gt1": int((df.Credit_Utilization > 1).sum()),
    "tenure_incons": int((df.Account_Tenure > df.Age - 18).sum()),
    "emp_relabelled": n_emp_relabelled, "hist_disagree": len(df) - hist_agrees,
    "rows_any_miss": rows_any_miss,
    "outputs": [str(p.name) for p in sorted(OUT.rglob("*")) if p.is_file()],
}
print(json.dumps(chk, indent=2))

# %% [markdown]
# ## Phase 8 (part 2) — Reproducible notebook export of this script

# %%
import re, nbformat as nbf

src_path = BASE / "geldium_eda.py"
nb_path = OUT / "Geldium_EDA_Code.ipynb"
cells, cur, kind = [], [], "code"
for line in src_path.read_text(encoding="utf-8").splitlines():
    if line.startswith("# %%"):
        if cur:
            cells.append((kind, "\n".join(cur).strip("\n")))
        cur, kind = [], ("markdown" if "[markdown]" in line else "code")
        if kind == "markdown":
            continue
        continue
    cur.append(line)
if cur:
    cells.append((kind, "\n".join(cur).strip("\n")))

nb = nbf.v4.new_notebook()
nb.cells = [nbf.v4.new_markdown_cell(text) if k == "markdown" else nbf.v4.new_code_cell(text)
            for k, text in cells if text.strip()]
nb.metadata["kernelspec"] = {"display_name": "Python 3", "language": "python", "name": "python3"}
nbf.write(nb, nb_path)
print("notebook exported:", nb_path, f"({len(nb.cells)} cells)")
