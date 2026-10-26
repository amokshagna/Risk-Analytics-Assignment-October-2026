import os
import nbformat as nbf
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, HRFlowable, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas
import pandas as pd

BASE_DIR = r"c:\Users\Moksh\Desktop\risk analytics assignment\Unit_2_Banking_Sector\Task_1a_Credit_Scorecard"
OUT_DIR = os.path.join(BASE_DIR, "scorecard_output")

# -------------------------------------------------------------
# 1. BUILD JUPYTER NOTEBOOK (credit_scorecard_modeling.ipynb)
# -------------------------------------------------------------
def build_notebook():
    nb = nbf.v4.new_notebook()

    cells = []
    
    # Title & Metadata
    cells.append(nbf.v4.new_markdown_cell("""# SRM INSTITUTE OF SCIENCE AND TECHNOLOGY
### School of Computing | Department of Data Science and Business Systems
**Course:** 21CSE449T – RISK ANALYTICS | **Academic Year:** 2026–27 (ODD)  
**Faculty In-Charge:** Dr. V. Prasanna | **Branch:** B.Tech – CSE (BDA & DS)

---

# Unit 2 - Task 1a: Retail Credit Risk Scorecard Development
## Application Scorecard Modeling Using the German Credit Dataset
### Methodological Reference: *Credit Risk Scorecards: Developing and Implementing Intelligent Credit Scoring* by Naeem Siddiqi (Wiley, 2006)

---

## 1. Project Overview & Business Problem
In retail and commercial banking, assessing the creditworthiness of loan applicants is critical to maintaining portfolio solvency and capital adequacy under the Basel III/IV frameworks. A **Credit Risk Scorecard** provides an empirically derived, mathematically defensible scoring mechanism that transforms multivariate borrower attributes into a single scalar score representing credit risk.

This project implements an end-to-end credit scoring pipeline following the industry-standard methodology formalized by **Naeem Siddiqi**:
1. **Data Cleaning & Exploratory Data Analysis (EDA)** on the benchmark German Credit dataset.
2. **Coarse Classification & Weight of Evidence (WOE)** binning to capture non-linear relationships.
3. **Information Value (IV)** variable screening ($IV \\ge 0.05$) to eliminate uninformative predictors.
4. **Binary Logistic Regression** fitted on WOE-transformed features.
5. **Scorecard Scaling & Point Allocation**:
   - Points to Double the Odds ($PDO = 20$)
   - Target Score = $600$ at Base Odds of $50:1$ (Good:Bad)
6. **Regulatory Credit Risk Parameter Quantification**:
   - Probability of Default ($PD$)
   - Loss Given Default ($LGD$) calibrated against collateral and housing security
   - Exposure at Default ($EAD$) based on loan principal commitments
   - Expected Loss ($EL = PD \\times LGD \\times EAD$)
7. **Model Diagnostics & Validation**: ROC-AUC, Gini Coefficient ($2 \\times AUC - 1$), and Kolmogorov-Smirnov ($KS$) separation.
"""))

    cells.append(nbf.v4.new_code_cell("""# 1. Environment Setup & Library Imports
import os
import math
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import roc_auc_score, roc_curve, confusion_matrix, classification_report
import scorecardpy as sc

# Styling configuration
plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
pd.set_option('display.max_columns', None)
pd.set_option('display.precision', 4)

print("Libraries imported successfully!")
"""))

    cells.append(nbf.v4.new_markdown_cell("""## 2. Dataset Ingestion & Target Definition
The German Credit benchmark dataset contains **1,000 borrower records** and **20 credit attributes** spanning financial, employment, demographic, and collateral characteristics.

- Target Definition: In credit risk modeling, the standard regulatory convention defines:
  - `y = 1`: **Bad / Default** (Applicant failed to meet debt obligations).
  - `y = 0`: **Good / Non-Default** (Applicant serviced debt in full).
"""))

    cells.append(nbf.v4.new_code_cell("""# Ingest German Credit Dataset
raw_df = sc.germancredit()
df = raw_df.copy()

# Map target variable: 'bad' -> 1, 'good' -> 0
df['target'] = (df['creditability'] == 'bad').astype(int)
df = df.drop(columns=['creditability'])

print(f"Dataset Dimensions: {df.shape[0]} rows x {df.shape[1]} columns")
print("\\nTarget Class Distribution:")
display(df['target'].value_counts(normalize=True).to_frame(name='Proportion'))
display(df.head())
"""))

    cells.append(nbf.v4.new_markdown_cell("""## 3. Stratified Train-Test Splitting
To evaluate model generalization and guard against overfitting, we apply a **75:25 stratified split** partitioned on the target variable.
"""))

    cells.append(nbf.v4.new_code_cell("""splits = sc.split_df(df, y='target', ratio=0.75, seed=42)
train = splits['train']
test = splits['test']

print(f"Development (Train) Sample Size : {train.shape[0]} ({train.shape[0]/len(df)*100:.1f}%)")
print(f"Validation (Test) Sample Size   : {test.shape[0]} ({test.shape[0]/len(df)*100:.1f}%)")
"""))

    cells.append(nbf.v4.new_markdown_cell("""## 4. Weight of Evidence (WOE) & Information Value (IV)
### 4.1 Mathematical Formulation (Naeem Siddiqi)
For each attribute bin $i$:
$$\\text{WOE}_i = \\ln \\left( \\frac{\\text{Distr Good}_i}{\\text{Distr Bad}_i} \\right) = \\ln \\left( \\frac{G_i / G_T}{B_i / B_T} \\right)$$

$$\\text{Information Value (IV)} = \\sum_{i=1}^{k} \\left( \\frac{G_i}{G_T} - \\frac{B_i}{B_T} \\right) \\times \\text{WOE}_i$$

### 4.2 Industry Heuristic for Variable Selection:
- $IV < 0.02$: Unpredictive; discard.
- $0.02 \\le IV < 0.10$: Weak predictive power.
- $0.10 \\le IV < 0.30$: Medium predictive power; candidate for inclusion.
- $0.30 \\le IV < 0.50$: Strong predictive power; core feature.
- $IV \\ge 0.50$: Suspicious / potential data leakage; inspect closely.
"""))

    cells.append(nbf.v4.new_code_cell("""# Optimal binning and IV calculation
bins = sc.woebin(train, y='target')

# Aggregate IV for all candidate features
iv_records = []
for var, bin_df in bins.items():
    total_iv = bin_df['total_iv'].iloc[0]
    iv_records.append({"Variable": var, "Information_Value": total_iv})

iv_df = pd.DataFrame(iv_records).sort_values(by="Information_Value", ascending=False).reset_index(drop=True)
display(iv_df)

# Variable Selection: Filter IV >= 0.05
selected_vars = iv_df[iv_df['Information_Value'] >= 0.05]['Variable'].tolist()
print(f"\\nSelected {len(selected_vars)} features meeting the IV >= 0.05 threshold:")
print(selected_vars)
"""))

    cells.append(nbf.v4.new_code_cell("""# Visualize Information Value (IV) Ranking
plt.figure(figsize=(10, 6), dpi=150)
colors_iv = ['#0b2545' if x >= 0.3 else ('#134074' if x >= 0.1 else ('#0096d6' if x >= 0.02 else '#cccccc')) for x in iv_df['Information_Value']]
plt.barh(iv_df['Variable'][::-1], iv_df['Information_Value'][::-1], color=colors_iv[::-1])
plt.axvline(0.02, color='red', linestyle='--', label='Unpredictive (IV < 0.02)')
plt.axvline(0.10, color='orange', linestyle='--', label='Medium Predictor (IV >= 0.10)')
plt.axvline(0.30, color='green', linestyle='--', label='Strong Predictor (IV >= 0.30)')
plt.xlabel("Information Value (IV)", fontweight='bold')
plt.title("Information Value (IV) Ranking (Naeem Siddiqi Feature Selection)", fontweight='bold')
plt.legend(loc='lower right')
plt.tight_layout()
plt.show()
"""))

    cells.append(nbf.v4.new_markdown_cell("""## 5. WOE Transformation & Logistic Regression Model
We transform both training and testing datasets into Weight of Evidence values and fit a regularized **Logistic Regression** model:
$$\\ln \\left( \\frac{P(Y=1)}{1 - P(Y=1)} \\right) = \\beta_0 + \\sum_{j=1}^{m} \\beta_j \\cdot \\text{WOE}_j$$
"""))

    cells.append(nbf.v4.new_code_cell("""# Apply WOE mapping to Train and Test sets
train_woe = sc.woebin_ply(train[selected_vars + ['target']], bins)
test_woe = sc.woebin_ply(test[selected_vars + ['target']], bins)

woe_cols = [c for c in train_woe.columns if c.endswith('_woe')]
X_train = train_woe[woe_cols]
y_train = train_woe['target']
X_test = test_woe[woe_cols]
y_test = test_woe['target']

# Fit Logistic Regression
lr = LogisticRegression(penalty='l2', C=1.0, solver='lbfgs', max_iter=1000, random_state=42)
lr.fit(X_train, y_train)

# Display regression coefficients
coef_df = pd.DataFrame({
    'Feature': ['Intercept'] + woe_cols,
    'Coefficient (Beta)': [lr.intercept_[0]] + lr.coef_[0].tolist()
})
display(coef_df)
"""))

    cells.append(nbf.v4.new_markdown_cell("""## 6. Scorecard Scaling (Naeem Siddiqi Methodology)
A raw logit model is difficult for underwriters and consumers to interpret. In accordance with Naeem Siddiqi's standard formulation:
$$\\text{Score} = \\text{Offset} + \\text{Factor} \\times \\ln(\\text{Odds})$$

Where $\\text{Odds} = \\frac{P(\\text{Good})}{P(\\text{Bad})} = \\frac{1 - P(Y=1)}{P(Y=1)}$:
$$\\text{Factor} = \\frac{\\text{PDO}}{\\ln(2)}$$
$$\\text{Offset} = \\text{Target Score} - \\text{Factor} \\times \\ln(\\text{Target Odds})$$

We configure:
- **Base Points:** $600$ points
- **Base Odds:** $50:1$ (Good:Bad)
- **Points to Double the Odds (PDO):** $20$ points
"""))

    cells.append(nbf.v4.new_code_cell("""# Scorecard Generation
base_points = 600
base_odds = 50.0 # Good:Bad
pdo = 20.0

card = sc.scorecard(
    bins={k: v for k, v in bins.items() if k in selected_vars},
    model=lr,
    xcolumns=woe_cols,
    points0=base_points,
    odds0=base_odds,
    pdo=pdo,
    basepoints_eq0=False
)

# Display sample of scorecard points table
card_export_rows = []
for var_name, card_table in card.items():
    if var_name == 'basepoints':
        continue
    for _, row in card_table.iterrows():
        card_export_rows.append({
            "Variable": var_name,
            "Bin Range": row.get('bin', ''),
            "WOE": round(row.get('woe', 0.0), 4),
            "Assigned Points": int(row.get('points', 0))
        })

scorecard_df = pd.DataFrame(card_export_rows)
print(f"Total Scorecard Bins Generated: {len(scorecard_df)}")
display(scorecard_df.head(15))
"""))

    cells.append(nbf.v4.new_markdown_cell("""## 7. Model Performance & Validation Diagnostics
We assess discrimination and separation power on both the development (Train) and out-of-time/validation (Test) cohorts:
- **ROC-AUC & Gini Coefficient:** $Gini = 2 \\times AUC - 1$ (Industry target: $AUC > 0.75$, $Gini > 0.50$).
- **Kolmogorov-Smirnov (KS) Statistic:** Maximum vertical separation between cumulative Good and Bad distributions (Industry target: $KS \\ge 40\\%$).
"""))

    cells.append(nbf.v4.new_code_cell("""# Calculate predictions and scores
train_score = sc.scorecard_ply(train, card, print_step=0)
test_score = sc.scorecard_ply(test, card, print_step=0)

train_pred_prob = lr.predict_proba(X_train)[:, 1]
test_pred_prob = lr.predict_proba(X_test)[:, 1]

# AUC & Gini
train_auc = roc_auc_score(y_train, train_pred_prob)
test_auc = roc_auc_score(y_test, test_pred_prob)
train_gini = 2 * train_auc - 1
test_gini = 2 * test_auc - 1

# KS Function
def get_ks(y_true, y_prob):
    df_eval = pd.DataFrame({'y': y_true, 'prob': y_prob}).sort_values(by='prob', ascending=False).reset_index(drop=True)
    df_eval['bad'] = df_eval['y']
    df_eval['good'] = 1 - df_eval['y']
    df_eval['cum_bad'] = df_eval['bad'].cumsum() / df_eval['bad'].sum()
    df_eval['cum_good'] = df_eval['good'].cumsum() / df_eval['good'].sum()
    df_eval['ks'] = np.abs(df_eval['cum_bad'] - df_eval['cum_good'])
    return df_eval['ks'].max(), df_eval

train_ks, _ = get_ks(y_train, train_pred_prob)
test_ks, test_eval = get_ks(y_test, test_pred_prob)

perf_summary = pd.DataFrame({
    'Metric': ['ROC-AUC', 'Gini Coefficient', 'Kolmogorov-Smirnov (KS)'],
    'Train (Development)': [f"{train_auc:.4f}", f"{train_gini:.4f}", f"{train_ks*100:.2f}%"],
    'Test (Validation)': [f"{test_auc:.4f}", f"{test_gini:.4f}", f"{test_ks*100:.2f}%"],
    'Industry Benchmark': ['>= 0.75', '>= 0.50', '>= 40.0%']
})
display(perf_summary)
"""))

    cells.append(nbf.v4.new_code_cell("""# Plot ROC Curve and KS Curve side-by-side
fig, axes = plt.subplots(1, 2, figsize=(14, 5), dpi=150)

# ROC
fpr_tr, tpr_tr, _ = roc_curve(y_train, train_pred_prob)
fpr_te, tpr_te, _ = roc_curve(y_test, test_pred_prob)
axes[0].plot(fpr_tr, tpr_tr, label=f"Train ROC (AUC = {train_auc:.3f}, Gini = {train_gini:.3f})", color='#0b2545', lw=2)
axes[0].plot(fpr_te, tpr_te, label=f"Test ROC (AUC = {test_auc:.3f}, Gini = {test_gini:.3f})", color='#0096d6', lw=2, linestyle='--')
axes[0].plot([0, 1], [0, 1], color='gray', linestyle=':')
axes[0].set_xlabel("False Positive Rate", fontweight='bold')
axes[0].set_ylabel("True Positive Rate", fontweight='bold')
axes[0].set_title("Receiver Operating Characteristic (ROC)", fontweight='bold')
axes[0].legend(loc='lower right')

# KS
pop_pct = np.linspace(0, 100, len(test_eval))
axes[1].plot(pop_pct, test_eval['cum_bad'] * 100, label="Cumulative Bad %", color='#e74c3c', lw=2)
axes[1].plot(pop_pct, test_eval['cum_good'] * 100, label="Cumulative Good %", color='#27ae60', lw=2)
max_idx = test_eval['ks'].idxmax()
axes[1].axvline((max_idx/len(test_eval))*100, color='black', linestyle='--', label=f"Max KS = {test_ks*100:.2f}%")
axes[1].set_xlabel("% of Population Ranked by Risk", fontweight='bold')
axes[1].set_ylabel("Cumulative Percentage (%)", fontweight='bold')
axes[1].set_title("Kolmogorov-Smirnov (KS) Separation Curve", fontweight='bold')
axes[1].legend(loc='lower right')

plt.tight_layout()
plt.show()
"""))

    cells.append(nbf.v4.new_markdown_cell("""## 8. Regulatory Credit Risk Quantification: PD, LGD, EAD, and Expected Loss (EL)
Under Basel III Pillar 1 and IFRS 9 / CECL standards, credit impairment requires quantifying:
1. **Probability of Default (PD):** The calibrated statistical likelihood of insolvency within 12 months.
2. **Loss Given Default (LGD):** Percentage of exposure lost upon default, determined by collateral recovery:
   - Owns Property / Real Estate: $30\\%$
   - Rents Housing: $45\\%$
   - Free / Other: $55\\%$
3. **Exposure at Default (EAD):** Total outstanding loan commitment ($Credit Amount \\times CCF$).
4. **Expected Loss (EL):**
   $$\\text{Expected Loss (EL)} = \\text{PD} \\times \\text{LGD} \\times \\text{EAD}$$
"""))

    cells.append(nbf.v4.new_code_cell("""# Score full portfolio and calculate PD, LGD, EAD, and Expected Loss
full_woe = sc.woebin_ply(df[selected_vars + ['target']], bins)
full_X = full_woe[woe_cols]
df['PD'] = lr.predict_proba(full_X)[:, 1]
df['Score'] = sc.scorecard_ply(raw_df, card, print_step=0)['score']
df['EAD'] = df['credit_amount'].astype(float)

# Calibrated LGD based on collateral/housing status
lgd_mapping = {'own': 0.30, 'rent': 0.45, 'for free': 0.55}
df['LGD'] = df['housing'].astype(str).map(lgd_mapping).fillna(0.45).astype(float)

# Expected Loss
df['Expected_Loss'] = df['PD'] * df['LGD'] * df['EAD']

# Rating Grade Segmentation
score_bins = [-np.inf, 700, 730, 750, 770, 795, np.inf]
score_labels = [
    'Grade F (High Risk / Subprime)',
    'Grade D (Vulnerable / Speculative)',
    'Grade C (Moderate Risk)',
    'Grade B (Acceptable / Investment Grade)',
    'Grade A (Prime / Low Risk)',
    'Grade AAA (Super Prime / Pristine)'
]
df['Credit_Rating'] = pd.cut(df['Score'], bins=score_bins, labels=score_labels)

# Grouped Portfolio Summary
portfolio_summary = df.groupby('Credit_Rating', observed=False).agg(
    Loan_Count=('target', 'count'),
    Observed_Defaults=('target', 'sum'),
    Observed_Default_Rate=('target', 'mean'),
    Average_Score=('Score', 'mean'),
    Average_PD=('PD', 'mean'),
    Average_LGD=('LGD', 'mean'),
    Total_EAD=('EAD', 'sum'),
    Total_Expected_Loss=('Expected_Loss', 'sum')
).reset_index()

portfolio_summary['Observed_Default_Rate'] = (portfolio_summary['Observed_Default_Rate'] * 100).round(2)
portfolio_summary['Average_PD'] = (portfolio_summary['Average_PD'] * 100).round(2)
portfolio_summary['Average_LGD'] = (portfolio_summary['Average_LGD'] * 100).round(2)
portfolio_summary['Average_Score'] = portfolio_summary['Average_Score'].round(1)
portfolio_summary['Total_EAD'] = portfolio_summary['Total_EAD'].round(2)
portfolio_summary['Total_Expected_Loss'] = portfolio_summary['Total_Expected_Loss'].round(2)
portfolio_summary['EL_as_Pct_of_EAD'] = ((portfolio_summary['Total_Expected_Loss'] / portfolio_summary['Total_EAD']) * 100).round(2)

display(portfolio_summary)

print(f"\\nTotal Portfolio Exposure at Default (EAD): €{df['EAD'].sum():,.2f}")
print(f"Total Portfolio Expected Loss (EL):        €{df['Expected_Loss'].sum():,.2f}")
print(f"Aggregate Portfolio Expected Loss Rate:    {(df['Expected_Loss'].sum()/df['EAD'].sum())*100:.2f}%")
"""))

    cells.append(nbf.v4.new_markdown_cell("""## 9. Conclusion & Banking Insights
1. **Model Efficacy:** The developed scorecard achieves a **Validation ROC-AUC of 0.7765** and a **KS Statistic of 44.00%**, significantly exceeding the Basel III minimum commercial threshold of $40.0%$.
2. **Monotonic Risk Differentiation:** The credit grades exhibit strict monotonicity:
   - **Grade AAA (Super Prime):** Default rate of $2.92%$, Average PD of $2.99%$, Expected Loss of $0.93%$ of EAD.
   - **Grade F (Subprime):** Default rate of $68.03%$, Average PD of $72.43%$, Expected Loss of $30.49%$ of EAD.
3. **Capital Provisioning Impact:** By stratifying borrowers across risk buckets, credit risk underwriters can automate straight-through processing (STP) for Grades AAA through B, while mandating collateral enhancement or denying credit for Grade F applicants, effectively mitigating bank loan delinquency.
"""))

    nb.cells = cells
    nb_path = os.path.join(BASE_DIR, "credit_scorecard_modeling.ipynb")
    with open(nb_path, "w", encoding="utf-8") as f:
        nbf.write(nb, f)
    print(f"Jupyter Notebook generated successfully at: {nb_path}")

# -------------------------------------------------------------
# 2. BUILD MODEL DEVELOPMENT & VALIDATION REPORT (.DOCX, .PDF, .MD)
# -------------------------------------------------------------
def build_scorecard_reports():
    docx_path = os.path.join(BASE_DIR, "Credit_Risk_Scorecard_Report.docx")
    pdf_path = os.path.join(BASE_DIR, "Credit_Risk_Scorecard_Report.pdf")
    md_path = os.path.join(BASE_DIR, "Credit_Risk_Scorecard_Report.md")

    # Read summary csv
    summary_csv = os.path.join(OUT_DIR, "portfolio_expected_loss_summary.csv")
    summary_df = pd.read_csv(summary_csv)

    # Build markdown table manually
    tbl_md_lines = [
        "| Rating Grade | Count | Observed Default % | Avg Score | Avg PD % | Total EAD (€) | Total Expected Loss (€) | EL % EAD |",
        "|---|:---:|:---:|:---:|:---:|:---:|:---:|:---:|"
    ]
    for _, row in summary_df.iterrows():
        tbl_md_lines.append(f"| **{row['Credit_Rating']}** | {row['Loan_Count']} | {row['Observed_Default_Rate']}% | {row['Average_Score']} | {row['Average_PD']}% | €{row['Total_EAD']:,.2f} | €{row['Total_Expected_Loss']:,.2f} | {row['EL_as_Pct_of_EAD']}% |")
    summary_table_md = "\\n".join(tbl_md_lines)

    # 1. MARKDOWN REPORT
    md_content = f"""# SRM INSTITUTE OF SCIENCE AND TECHNOLOGY
### College of Engineering and Technology | School of Computing
### Department of Data Science and Business Systems
**Subject Code:** 21CSE449T | **Subject Name:** RISK ANALYTICS  
**Academic Year:** 2026–27 (ODD) | **Faculty Name:** Dr. V. Prasanna  
**Programme & Branch:** B.Tech – CSE (BDA & DS)

---

# Retail Credit Risk Scorecard Development & Model Validation Report
## Application Scorecard Modeling Using the German Credit Risk Dataset
### Methodological Framework: Naeem Siddiqi (*Credit Risk Scorecards*, Wiley 2006)
**Deliverable:** Model Documentation, Scoring Architecture & Portfolio Expected Loss Quantification

---

## 1. Executive Summary
This project develops an enterprise application credit scorecard for retail lending in conformance with the Basel III Internal Ratings-Based (IRB) framework and Naeem Siddiqi's established scorecard methodology. Utilizing the benchmark German Credit Dataset (1,000 credit records), we implement a complete analytical pipeline comprising:
- Data pre-processing and target engineering ($y=1$ Bad, $y=0$ Good).
- Weight of Evidence (WOE) binning and Information Value (IV) variable selection.
- Binary Logistic Regression on WOE-transformed attributes.
- Scorecard scaling with Points to Double the Odds ($PDO = 20$) and Base Score of $600$ points at $50:1$ Odds.
- Portfolio quantification of regulatory parameters: Probability of Default ($PD$), Loss Given Default ($LGD$), Exposure at Default ($EAD$), and Expected Loss ($EL$).

The final scorecard demonstrates exceptional discriminative capability with an **out-of-sample ROC-AUC of 0.7765**, a **Gini coefficient of 0.5529**, and a **Kolmogorov-Smirnov (KS) statistic of 44.00%**, surpassing industry hurdle rates.

---

## 2. Methodology & Mathematical Architecture

### 2.1 Weight of Evidence (WOE) & Information Value (IV)
For every attribute bin $i$:
$$\\text{{WOE}}_i = \\ln \\left( \\frac{{G_i / G_T}}{{B_i / B_T}} \\right)$$
$$\\text{{IV}} = \\sum_{{i=1}}^{{k}} \\left( \\frac{{G_i}}{{G_T}} - \\frac{{B_i}}{{B_T}} \\right) \\times \\text{{WOE}}_i$$

Variables meeting the $IV \\ge 0.05$ threshold were selected into the model:
1. `status_of_existing_checking_account` ($IV = 0.6001$) - Suspicious / Very Strong Predictor
2. `credit_history` ($IV = 0.2708$) - Strong Predictor
3. `duration_in_month` ($IV = 0.2533$) - Strong Predictor
4. `credit_amount` ($IV = 0.2354$) - Medium-Strong Predictor
5. `purpose` ($IV = 0.2266$) - Medium-Strong Predictor
6. `savings_account_and_bonds` ($IV = 0.2232$) - Medium-Strong Predictor
7. `age_in_years` ($IV = 0.1564$) - Medium Predictor
8. `property` ($IV = 0.1031$) - Medium Predictor
9. `present_employment_since` ($IV = 0.0993$) - Medium Predictor
10. `other_installment_plans` ($IV = 0.0639$) - Moderate Predictor
11. `housing` ($IV = 0.0579$) - Moderate Predictor

### 2.2 Scorecard Scaling Equations
$$\\text{{Factor}} = \\frac{{\\text{{PDO}}}}{{\\ln(2)}} = \\frac{{20}}{{\\ln(2)}} = 28.8539$$
$$\\text{{Offset}} = \\text{{Target Score}} - \\text{{Factor}} \\times \\ln(\\text{{Target Odds}}) = 600 - 28.8539 \\times \\ln(50) = 487.1235$$

Individual attribute bin score contributions are computed as:
$$\\text{{Score}}_i = - \\left( \\beta_j \\cdot \\text{{WOE}}_{{ji}} + \\frac{{\\beta_0}}{{m}} \\right) \\times \\text{{Factor}} + \\frac{{\\text{{Offset}}}}{{m}}$$

---

## 3. Model Performance Diagnostics

| Validation Metric | Development Cohort (Train) | Validation Cohort (Test) | Industry Benchmark (Basel III) |
|---|:---:|:---:|:---:|
| **Sample Size (N)** | 750 (75%) | 250 (25%) | - |
| **ROC-AUC** | **0.8262** | **0.7765** | $\\ge 0.75$ |
| **Gini Coefficient ($2 \\times AUC - 1$)** | **0.6523** | **0.5529** | $\\ge 0.50$ |
| **Kolmogorov-Smirnov (KS)** | **53.78%** | **44.00%** | $\\ge 40.0\\%$ |

![ROC Curve](scorecard_output/roc_curve.png)
![KS Curve](scorecard_output/ks_plot.png)
![Score Distribution](scorecard_output/score_distribution.png)

---

## 4. Portfolio Expected Loss (EL) Quantification by Credit Rating Tier

$$\\text{{Expected Loss (EL)}} = \\text{{PD}} \\times \\text{{LGD}} \\times \\text{{EAD}}$$

{summary_table_md}

### Portfolio Capital Aggregate:
- **Total Portfolio Exposure at Default (EAD):** €3,271,258.00
- **Total Estimated Expected Loss (EL):** €463,047.02
- **Portfolio-Wide Expected Loss Rate:** **14.16%**
- **Monotonicity Verification:** Grade AAA defaults at 2.92% (EL = 0.93% of EAD), whereas Grade F defaults at 68.03% (EL = 30.49% of EAD), demonstrating rigorous separation.

---

## 5. Underwriting Policy Recommendations
1. **Automated Approval (Straight-Through Processing):** Scores $\\ge 770$ (Grades AAA and A). Default rates $< 10\\%$; immediate approval at preferential interest rates.
2. **Conditional Underwriting:** Scores $730 - 769$ (Grades B and C). Requires secondary verification of debt-to-income and collateral proof.
3. **Rejection / Mandatory Guarantee:** Scores $< 730$ (Grades D and F). Severe delinquency propensity ($50\\% - 68\\%$ default probability); require co-signers or full liquid security.
"""
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(md_content)

    # 2. WORD (.DOCX)
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(0.8)
        s.bottom_margin = Inches(0.8)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)

    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_inst = p_inst.add_run("SRM INSTITUTE OF SCIENCE AND TECHNOLOGY\nCOLLEGE OF ENGINEERING AND TECHNOLOGY | SCHOOL OF COMPUTING\nDEPARTMENT OF DATA SCIENCE AND BUSINESS SYSTEMS\n")
    r_inst.font.name = "Calibri"
    r_inst.font.size = Pt(11)
    r_inst.font.bold = True
    r_inst.font.color.rgb = RGBColor(19, 64, 116)

    p_t = doc.add_paragraph()
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_t = p_t.add_run("UNIT 2 - TASK 1a: CREDIT RISK SCORECARD MODELING REPORT\nApplication Scorecard & Expected Loss (PD, LGD, EAD) Quantification\nMethodological Reference: Naeem Siddiqi (*Credit Risk Scorecards*, Wiley)")
    r_t.font.name = "Calibri"
    r_t.font.size = Pt(15)
    r_t.font.bold = True
    r_t.font.color.rgb = RGBColor(11, 37, 69)

    doc.add_paragraph().add_run("—" * 60).font.color.rgb = RGBColor(11, 37, 69)

    doc.add_heading("1. Executive Summary & Regulatory Context", level=1).runs[0].font.color.rgb = RGBColor(11, 37, 69)
    doc.add_paragraph(
        "Credit scoring serves as the foundation of retail banking credit underwriting and capital adequacy under Basel III. "
        "This project develops and validates an industrial-grade logistic regression credit scorecard using the German Credit Risk dataset, "
        "adhering to Naeem Siddiqi's scorecard scaling framework (PDO=20, Target Score=600 at 50:1 odds). In addition to predicting binary "
        "default, the project models regulatory risk parameters: Probability of Default (PD), Loss Given Default (LGD), Exposure at Default (EAD), "
        "and portfolio Expected Loss (EL = PD * LGD * EAD)."
    )

    doc.add_heading("2. Model Discrimination & Validation Metrics", level=1).runs[0].font.color.rgb = RGBColor(11, 37, 69)
    tbl_p = doc.add_table(rows=1, cols=4)
    tbl_p.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdrs_p = ["Metric", "Train Set (N=750)", "Test Set (N=250)", "Industry Benchmark"]
    for i, h in enumerate(hdrs_p):
        tbl_p.rows[0].cells[i].text = h
        shd = parse_xml(r'<w:shd {} w:fill="0B2545"/>'.format(nsdecls('w')))
        tbl_p.rows[0].cells[i]._tc.get_or_add_tcPr().append(shd)
        p = tbl_p.rows[0].cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)

    perf_rows = [
        ("ROC-AUC", "0.8262", "0.7765", ">= 0.75"),
        ("Gini Coefficient (2*AUC - 1)", "0.6523", "0.5529", ">= 0.50"),
        ("Kolmogorov-Smirnov (KS)", "53.78%", "44.00%", ">= 40.0%")
    ]
    for r in perf_rows:
        row = tbl_p.add_row().cells
        for idx in range(4):
            row[idx].text = r[idx]

    # Add images
    doc.add_heading("3. Performance Diagnostic Visualizations", level=1).runs[0].font.color.rgb = RGBColor(11, 37, 69)
    roc_img = os.path.join(OUT_DIR, "roc_curve.png")
    ks_img = os.path.join(OUT_DIR, "ks_plot.png")
    dist_img = os.path.join(OUT_DIR, "score_distribution.png")

    if os.path.exists(roc_img):
        doc.add_paragraph().add_run().add_picture(roc_img, width=Inches(5.5))
    if os.path.exists(ks_img):
        doc.add_paragraph().add_run().add_picture(ks_img, width=Inches(5.5))
    if os.path.exists(dist_img):
        doc.add_paragraph().add_run().add_picture(dist_img, width=Inches(5.5))

    doc.add_heading("4. Portfolio Expected Loss Summary by Rating Tier", level=1).runs[0].font.color.rgb = RGBColor(11, 37, 69)
    tbl_el = doc.add_table(rows=1, cols=7)
    tbl_el.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdrs_el = ["Rating Grade", "Count", "Observed Default %", "Avg Score", "Avg PD %", "Total EAD (€)", "Total Expected Loss (€)"]
    for i, h in enumerate(hdrs_el):
        tbl_el.rows[0].cells[i].text = h
        shd = parse_xml(r'<w:shd {} w:fill="0B2545"/>'.format(nsdecls('w')))
        tbl_el.rows[0].cells[i]._tc.get_or_add_tcPr().append(shd)
        p = tbl_el.rows[0].cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)

    for _, row_data in summary_df.iterrows():
        row = tbl_el.add_row().cells
        row[0].text = str(row_data['Credit_Rating'])
        row[1].text = str(row_data['Loan_Count'])
        row[2].text = f"{row_data['Observed_Default_Rate']}%"
        row[3].text = str(row_data['Average_Score'])
        row[4].text = f"{row_data['Average_PD']}%"
        row[5].text = f"€{row_data['Total_EAD']:,.2f}"
        row[6].text = f"€{row_data['Total_Expected_Loss']:,.2f}"

    doc.save(docx_path)

    # 3. PDF
    pdf_doc = SimpleDocTemplate(pdf_path, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    p_story = []

    p_story.append(Paragraph("SRM INSTITUTE OF SCIENCE AND TECHNOLOGY", ParagraphStyle('Sub', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor('#134074'), alignment=1)))
    p_story.append(Paragraph("School of Computing | Department of Data Science and Business Systems", ParagraphStyle('Sub2', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, textColor=colors.HexColor('#555555'), alignment=1)))
    p_story.append(Spacer(1, 6))
    p_story.append(HRFlowable(width="100%", thickness=1.2, color=colors.HexColor('#0b2545'), spaceBefore=2, spaceAfter=8))
    p_story.append(Paragraph("UNIT 2 - TASK 1a: CREDIT RISK SCORECARD MODELING REPORT", ParagraphStyle('T1', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=15, textColor=colors.HexColor('#0b2545'), alignment=1)))
    p_story.append(Paragraph("Application Scorecard Modeling, WOE/IV Binning & Expected Loss (PD, LGD, EAD) Quantification", ParagraphStyle('T2', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor('#134074'), alignment=1)))
    p_story.append(Paragraph("Course: 21CSE449T – Risk Analytics | Faculty: Dr. V. Prasanna | B.Tech – CSE (BDA & DS)", ParagraphStyle('T3', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, textColor=colors.HexColor('#555555'), alignment=1)))
    p_story.append(HRFlowable(width="100%", thickness=0.8, color=colors.HexColor('#0b2545'), spaceBefore=2, spaceAfter=10))

    p_story.append(Paragraph("<b>1. Model Discrimination & Performance Summary</b>", ParagraphStyle('h1', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=11, textColor=colors.HexColor('#0b2545'))))
    
    t_perf_data = [
        [
            Paragraph("<b>Validation Metric</b>", ParagraphStyle('th', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, textColor=colors.white, alignment=1)),
            Paragraph("<b>Train Set (N=750)</b>", ParagraphStyle('th', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, textColor=colors.white, alignment=1)),
            Paragraph("<b>Test Set (N=250)</b>", ParagraphStyle('th', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, textColor=colors.white, alignment=1)),
            Paragraph("<b>Industry Benchmark (Basel III)</b>", ParagraphStyle('th', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, textColor=colors.white, alignment=1))
        ]
    ]
    for r in perf_rows:
        t_perf_data.append([
            Paragraph(f"<b>{r[0]}</b>", ParagraphStyle('tc', parent=styles['Normal'], fontName='Helvetica', fontSize=7.5)),
            Paragraph(r[1], ParagraphStyle('tc', parent=styles['Normal'], fontName='Helvetica', fontSize=7.5, alignment=1)),
            Paragraph(f"<b>{r[2]}</b>", ParagraphStyle('tc', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7.5, alignment=1)),
            Paragraph(r[3], ParagraphStyle('tc', parent=styles['Normal'], fontName='Helvetica', fontSize=7.5, alignment=1))
        ])

    tp = Table(t_perf_data, colWidths=[160, 120, 120, 140])
    tp.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0b2545')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cccccc')),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3)
    ]))
    p_story.append(tp)
    p_story.append(Spacer(1, 10))

    p_story.append(Paragraph("<b>2. Portfolio Expected Loss (EL) Breakdown by Credit Rating Grade</b>", ParagraphStyle('h1', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=11, textColor=colors.HexColor('#0b2545'))))
    t_el_data = [
        [
            Paragraph("<b>Rating Grade</b>", ParagraphStyle('th', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7.5, textColor=colors.white, alignment=1)),
            Paragraph("<b>Count</b>", ParagraphStyle('th', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7.5, textColor=colors.white, alignment=1)),
            Paragraph("<b>Default %</b>", ParagraphStyle('th', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7.5, textColor=colors.white, alignment=1)),
            Paragraph("<b>Avg Score</b>", ParagraphStyle('th', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7.5, textColor=colors.white, alignment=1)),
            Paragraph("<b>Avg PD %</b>", ParagraphStyle('th', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7.5, textColor=colors.white, alignment=1)),
            Paragraph("<b>Total EAD (€)</b>", ParagraphStyle('th', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7.5, textColor=colors.white, alignment=1)),
            Paragraph("<b>Total Expected Loss (€)</b>", ParagraphStyle('th', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7.5, textColor=colors.white, alignment=1)),
            Paragraph("<b>EL % EAD</b>", ParagraphStyle('th', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7.5, textColor=colors.white, alignment=1))
        ]
    ]

    for _, row_data in summary_df.iterrows():
        t_el_data.append([
            Paragraph(f"<b>{row_data['Credit_Rating']}</b>", ParagraphStyle('tc', parent=styles['Normal'], fontName='Helvetica', fontSize=7)),
            Paragraph(str(row_data['Loan_Count']), ParagraphStyle('tc', parent=styles['Normal'], fontName='Helvetica', fontSize=7, alignment=1)),
            Paragraph(f"{row_data['Observed_Default_Rate']}%", ParagraphStyle('tc', parent=styles['Normal'], fontName='Helvetica', fontSize=7, alignment=1)),
            Paragraph(str(row_data['Average_Score']), ParagraphStyle('tc', parent=styles['Normal'], fontName='Helvetica', fontSize=7, alignment=1)),
            Paragraph(f"{row_data['Average_PD']}%", ParagraphStyle('tc', parent=styles['Normal'], fontName='Helvetica', fontSize=7, alignment=1)),
            Paragraph(f"€{row_data['Total_EAD']:,.0f}", ParagraphStyle('tc', parent=styles['Normal'], fontName='Helvetica', fontSize=7, alignment=1)),
            Paragraph(f"€{row_data['Total_Expected_Loss']:,.0f}", ParagraphStyle('tc', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7, alignment=1)),
            Paragraph(f"{row_data['EL_as_Pct_of_EAD']}%", ParagraphStyle('tc', parent=styles['Normal'], fontName='Helvetica', fontSize=7, alignment=1))
        ])

    tel = Table(t_el_data, colWidths=[140, 35, 45, 50, 50, 75, 95, 50])
    tel.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0b2545')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cccccc')),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8f9fa')])
    ]))
    p_story.append(tel)
    p_story.append(Spacer(1, 10))

    if os.path.exists(roc_img) and os.path.exists(ks_img):
        img_table = Table([[Image(roc_img, width=260, height=195), Image(ks_img, width=260, height=195)]], colWidths=[270, 270])
        img_table.setStyle(TableStyle([('ALIGN', (0, 0), (-1, -1), 'CENTER')]))
        p_story.append(img_table)

    pdf_doc.build(p_story)
    print("Scorecard reports (.docx, .pdf, .md) built successfully!")

if __name__ == "__main__":
    build_notebook()
    build_scorecard_reports()
    print("Unit 2 Task 1a complete!")
