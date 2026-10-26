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

BASE_DIR = r"c:\Users\Moksh\Desktop\risk analytics assignment\Unit_2_Banking_Sector\Task_1a_Credit_Scorecard"
OUT_DIR = os.path.join(BASE_DIR, "scorecard_output")
os.makedirs(OUT_DIR, exist_ok=True)

print("--- STEP 1: LOADING & SAVING BENCHMARK GERMAN CREDIT DATASET ---")
raw_df = sc.germancredit()
# Save raw data to csv
data_path = os.path.join(BASE_DIR, "german_credit_data.csv")
raw_df.to_csv(data_path, index=False)
print(f"Dataset saved at: {data_path} (Shape: {raw_df.shape})")

# In credit risk modeling convention (Naeem Siddiqi):
# 1 = Bad (Default), 0 = Good (Non-default)
# scorecardpy default has 'good' and 'bad'
df = raw_df.copy()
df['target'] = (df['creditability'] == 'bad').astype(int)
df = df.drop(columns=['creditability'])

print("Target Distribution:")
print(df['target'].value_counts(normalize=True))

print("\n--- STEP 2: TRAIN-TEST SPLIT (STRATIFIED 75:25) ---")
splits = sc.split_df(df, y='target', ratio=0.75, seed=42)
train, test = splits['train'], splits['test']
print(f"Train Shape: {train.shape}, Test Shape: {test.shape}")

print("\n--- STEP 3: INFORMATION VALUE (IV) & WOE BINNING ---")
# Generate optimal bins using scorecardpy
bins = sc.woebin(train, y='target')

# Calculate Information Value for each variable
iv_records = []
for var, bin_df in bins.items():
    total_iv = bin_df['total_iv'].iloc[0]
    iv_records.append({"Variable": var, "Information_Value": total_iv})

iv_df = pd.DataFrame(iv_records).sort_values(by="Information_Value", ascending=False).reset_index(drop=True)
print("\nTop Predictors by Information Value (IV):")
print(iv_df.to_string())

# Plot IV ranking
plt.figure(figsize=(10, 6), dpi=300)
colors_iv = ['#0b2545' if x >= 0.3 else ('#134074' if x >= 0.1 else ('#0096d6' if x >= 0.02 else '#cccccc')) for x in iv_df['Information_Value']]
bars = plt.barh(iv_df['Variable'][::-1], iv_df['Information_Value'][::-1], color=colors_iv[::-1])
plt.axvline(0.02, color='red', linestyle='--', label='Unpredictive Cutoff (IV < 0.02)')
plt.axvline(0.1, color='orange', linestyle='--', label='Medium Predictor (IV >= 0.10)')
plt.axvline(0.3, color='green', linestyle='--', label='Strong Predictor (IV >= 0.30)')
plt.xlabel("Information Value (IV)", fontsize=11, fontweight='bold')
plt.title("Predictive Power of Variables (Information Value - Naeem Siddiqi Filter)", fontsize=12, fontweight='bold')
plt.legend(loc='lower right', fontsize=9)
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, "woe_iv_ranking.png"))
plt.close()

# Variable Selection: Keep variables with IV >= 0.05 for parsimonious model
selected_vars = iv_df[iv_df['Information_Value'] >= 0.05]['Variable'].tolist()
print(f"\nSelected {len(selected_vars)} variables with IV >= 0.05: {selected_vars}")

# Convert train and test to WOE values
train_woe = sc.woebin_ply(train[selected_vars + ['target']], bins)
test_woe = sc.woebin_ply(test[selected_vars + ['target']], bins)

woe_cols = [c for c in train_woe.columns if c.endswith('_woe')]
X_train = train_woe[woe_cols]
y_train = train_woe['target']
X_test = test_woe[woe_cols]
y_test = test_woe['target']

print("\n--- STEP 4: LOGISTIC REGRESSION MODELING ---")
lr = LogisticRegression(penalty='l2', C=1.0, solver='lbfgs', max_iter=1000, random_state=42)
lr.fit(X_train, y_train)

# Coefficients table
coef_df = pd.DataFrame({
    'Feature': ['Intercept'] + woe_cols,
    'Coefficient': [lr.intercept_[0]] + lr.coef_[0].tolist()
})
print("\nTrained Logistic Regression Parameters:")
print(coef_df.to_string())

print("\n--- STEP 5: SCORECARD SCALING (NAEEM SIDDIQI METHODOLOGY) ---")
# Reference: Naeem Siddiqi - Credit Risk Scorecards
# Target Score = 600 points at Odds 50:1 (Good:Bad)
# Points to Double the Odds (PDO) = 20
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

# Export scorecard points table
card_export_rows = []
for var_name, card_table in card.items():
    if var_name == 'basepoints':
        continue
    for _, row in card_table.iterrows():
        card_export_rows.append({
            "Variable": var_name,
            "Bin": row.get('bin', ''),
            "WOE": row.get('woe', 0.0),
            "Score_Points": row.get('points', 0)
        })

scorecard_df = pd.DataFrame(card_export_rows)
scorecard_df.to_csv(os.path.join(OUT_DIR, "scorecard_points_table.csv"), index=False)
print(f"Scorecard points table exported to {os.path.join(OUT_DIR, 'scorecard_points_table.csv')}")

# Calculate credit scores for train and test sets
train_score = sc.scorecard_ply(train, card, print_step=0)
test_score = sc.scorecard_ply(test, card, print_step=0)

train_pred_prob = lr.predict_proba(X_train)[:, 1]
test_pred_prob = lr.predict_proba(X_test)[:, 1]

# Model Evaluation
train_auc = roc_auc_score(y_train, train_pred_prob)
test_auc = roc_auc_score(y_test, test_pred_prob)
train_gini = 2 * train_auc - 1
test_gini = 2 * test_auc - 1

# KS Statistic Calculation
def calculate_ks(y_true, y_prob):
    df_eval = pd.DataFrame({'y': y_true, 'prob': y_prob})
    df_eval = df_eval.sort_values(by='prob', ascending=False).reset_index(drop=True)
    df_eval['bad'] = df_eval['y']
    df_eval['good'] = 1 - df_eval['y']
    df_eval['cum_bad'] = df_eval['bad'].cumsum() / df_eval['bad'].sum()
    df_eval['cum_good'] = df_eval['good'].cumsum() / df_eval['good'].sum()
    df_eval['ks'] = np.abs(df_eval['cum_bad'] - df_eval['cum_good'])
    ks_stat = df_eval['ks'].max()
    max_idx = df_eval['ks'].idxmax()
    return ks_stat, df_eval, max_idx

train_ks, _, _ = calculate_ks(y_train, train_pred_prob)
test_ks, test_eval, test_max_idx = calculate_ks(y_test, test_pred_prob)

print("\n--- MODEL PERFORMANCE METRICS ---")
print(f"Train ROC-AUC: {train_auc:.4f} | Train Gini: {train_gini:.4f} | Train KS: {train_ks*100:.2f}%")
print(f"Test  ROC-AUC: {test_auc:.4f} | Test  Gini: {test_gini:.4f} | Test  KS: {test_ks*100:.2f}%")

# Plot ROC Curves
plt.figure(figsize=(8, 6), dpi=300)
fpr_tr, tpr_tr, _ = roc_curve(y_train, train_pred_prob)
fpr_te, tpr_te, _ = roc_curve(y_test, test_pred_prob)
plt.plot(fpr_tr, tpr_tr, label=f"Train ROC (AUC = {train_auc:.3f}, Gini = {train_gini:.3f})", color='#0b2545', lw=2)
plt.plot(fpr_te, tpr_te, label=f"Test ROC (AUC = {test_auc:.3f}, Gini = {test_gini:.3f})", color='#0096d6', lw=2, linestyle='--')
plt.plot([0, 1], [0, 1], color='gray', linestyle=':')
plt.xlabel("False Positive Rate (1 - Specificity)", fontsize=11, fontweight='bold')
plt.ylabel("True Positive Rate (Sensitivity)", fontsize=11, fontweight='bold')
plt.title("Credit Scorecard ROC Curves (Train vs. Test Set)", fontsize=12, fontweight='bold')
plt.legend(loc='lower right', fontsize=10)
plt.grid(alpha=0.3)
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, "roc_curve.png"))
plt.close()

# Plot KS Chart
plt.figure(figsize=(8, 6), dpi=300)
pop_pct = np.linspace(0, 100, len(test_eval))
plt.plot(pop_pct, test_eval['cum_bad'] * 100, label="Cumulative Bad %", color='#e74c3c', lw=2)
plt.plot(pop_pct, test_eval['cum_good'] * 100, label="Cumulative Good %", color='#27ae60', lw=2)
max_ks_pop = (test_max_idx / len(test_eval)) * 100
plt.axvline(max_ks_pop, color='black', linestyle='--', label=f"Max KS = {test_ks*100:.2f}% at {max_ks_pop:.1f}% Pop")
plt.xlabel("% of Population Ranked by Predicted Risk", fontsize=11, fontweight='bold')
plt.ylabel("Cumulative Percentage (%)", fontsize=11, fontweight='bold')
plt.title(f"Kolmogorov-Smirnov (KS) Separation Curve (Test KS = {test_ks*100:.2f}%)", fontsize=12, fontweight='bold')
plt.legend(loc='lower right', fontsize=10)
plt.grid(alpha=0.3)
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, "ks_plot.png"))
plt.close()

# Plot Score Distribution
plt.figure(figsize=(9, 5), dpi=300)
test_results = test.copy()
test_results['score'] = test_score['score']
test_results['pred_pd'] = test_pred_prob
sns.histplot(data=test_results, x='score', hue='target', bins=25, kde=True, palette={0: '#27ae60', 1: '#e74c3c'}, alpha=0.6)
plt.axvline(test_results[test_results['target'] == 0]['score'].median(), color='#27ae60', linestyle='--', label='Median Good Score')
plt.axvline(test_results[test_results['target'] == 1]['score'].median(), color='#e74c3c', linestyle='--', label='Median Bad Score')
plt.xlabel("Credit Score (Points)", fontsize=11, fontweight='bold')
plt.ylabel("Customer Count", fontsize=11, fontweight='bold')
plt.title("Credit Score Distribution by Observed Default Status (0=Good, 1=Bad)", fontsize=12, fontweight='bold')
plt.legend(title="Default Status", labels=["Good", "Bad"], fontsize=10)
plt.grid(alpha=0.3)
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, "score_distribution.png"))
plt.close()

print("\n--- STEP 6: PD, LGD, EAD & EXPECTED LOSS (EL) QUANTIFICATION ---")
# Build complete full-portfolio prediction table
full_woe = sc.woebin_ply(df[selected_vars + ['target']], bins)
full_X = full_woe[woe_cols]
df['PD'] = lr.predict_proba(full_X)[:, 1]
df['Score'] = sc.scorecard_ply(raw_df, card, print_step=0)['score']

# Exposure at Default (EAD): Credit Amount * Credit Conversion Factor (CCF=1.0 for term loans)
df['EAD'] = df['credit_amount'].astype(float)

# Loss Given Default (LGD) Calibration based on collateral & housing status:
# Basel Foundation IRB standard conventions:
# - Housing = own (Collateralized/Secured): LGD = 30%
# - Housing = rent: LGD = 45%
# - Housing = for free / other: LGD = 55%
lgd_mapping = {
    'own': 0.30,
    'rent': 0.45,
    'for free': 0.55
}
df['LGD'] = df['housing'].astype(str).map(lgd_mapping).fillna(0.45).astype(float)

# Expected Loss (EL) = PD * LGD * EAD
df['Expected_Loss'] = df['PD'] * df['LGD'] * df['EAD']

# Credit Rating Tier Segmentation
# Calibrated Basel Risk Buckets (Super Prime to Subprime)
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

# Portfolio Summary by Credit Rating Grade
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

print("\n--- PORTFOLIO EXPECTED LOSS SUMMARY TABLE BY RATING GRADE ---")
print(portfolio_summary.to_string())

summary_csv_path = os.path.join(OUT_DIR, "portfolio_expected_loss_summary.csv")
portfolio_summary.to_csv(summary_csv_path, index=False)

loan_pred_path = os.path.join(OUT_DIR, "loan_level_predictions.csv")
df[['Score', 'Credit_Rating', 'PD', 'LGD', 'EAD', 'Expected_Loss', 'target']].to_csv(loan_pred_path, index=False)

total_ead = df['EAD'].sum()
total_el = df['Expected_Loss'].sum()
portfolio_el_rate = (total_el / total_ead) * 100
print(f"\n=======================================================")
print(f"TOTAL PORTFOLIO EXPOSURE AT DEFAULT (EAD): €{total_ead:,.2f}")
print(f"TOTAL ESTIMATED EXPECTED LOSS (EL):        €{total_el:,.2f}")
print(f"PORTFOLIO EXPECTED LOSS RATE:              {portfolio_el_rate:.2f}%")
print(f"=======================================================")
print("\nUnit 2 Task 1a Python calculations completed successfully!")
