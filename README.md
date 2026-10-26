# Risk Analytics Course Assignment (21CSE449T)
### SRM Institute of Science and Technology
**College of Engineering and Technology | School of Computing**  
**Department of Data Science and Business Systems**
**Submitted by Mokshagna Alaparthi**
**Programme & Branch:** B.Tech – CSE (Data Science)  
**Academic Year:** 2026–27 (ODD) | **Faculty In-Charge:** Dr. V. Prasanna

---

## Executive Project Directory Structure

```
risk analytics assignment/
│
├── README.md                                         <- Master repository index and documentation
├── submission_summary.html                           <- Interactive executive presentation & visual dashboard
├── .gitignore                                        <- Standard Git ignore rules
│
├── Unit_1_Risk_Introduction_and_Impacts/
│   ├── Task_1_Risk_Register/
│   │   ├── generate_risk_register_doc.py             <- Script generating PDF, DOCX, and Heatmap
│   │   ├── risk_heatmap.png                          <- 5x5 Matrix Risk Heat Map (Likelihood x Impact)
│   │   ├── Paytm_Risk_Register_and_Management_Plan.pdf <- Deliverable: STRICT 5-PAGE DOCUMENT
│   │   ├── Paytm_Risk_Register_and_Management_Plan.docx<- Word Document
│   │   └── Paytm_Risk_Register_and_Management_Plan.md  <- Markdown Report
│   │
│   └── Task_2_Tool_Exploration/
│       ├── generate_tool_exploration.py              <- Script generating 10-slide PPTX & docs
│       ├── Risk_Analytics_Tools_Comparison.pptx      <- Deliverable: 10-SLIDE PRESENTATION
│       ├── Tool_Exploration_Comparative_Table.docx   <- Comparative Table Document
│       ├── Tool_Exploration_Comparative_Table.pdf    <- Formatted PDF
│       └── Tool_Exploration_Comparative_Table.md     <- Markdown Report
│
└── Unit_2_Banking_Sector/
    ├── Task_1a_Credit_Scorecard/
    │   ├── credit_scorecard.py                       <- End-to-end Python model execution script
    │   ├── credit_scorecard_modeling.ipynb           <- Deliverable: JUPYTER NOTEBOOK (Naeem Siddiqi)
    │   ├── german_credit_data.csv                    <- Kaggle German Credit Risk Dataset
    │   ├── scorecard_output/
    │   │   ├── roc_curve.png                         <- ROC Curve (Train AUC: 0.826, Test AUC: 0.776)
    │   │   ├── ks_plot.png                           <- Kolmogorov-Smirnov Separation (KS: 44.00%)
    │   │   ├── score_distribution.png                <- Score Histogram by Good/Bad observed status
    │   │   ├── woe_iv_ranking.png                    <- Information Value (IV) predictor ranking
    │   │   ├── scorecard_points_table.csv            <- Scaled points per attribute bin
    │   │   ├── portfolio_expected_loss_summary.csv   <- Expected Loss by Rating Tier (AAA to F)
    │   │   └── loan_level_predictions.csv            <- Loan-level Score, PD, LGD, EAD, and EL
    │   ├── Credit_Risk_Scorecard_Report.docx         <- Model Validation Report
    │   ├── Credit_Risk_Scorecard_Report.pdf          <- Formatted PDF Report
    │   └── Credit_Risk_Scorecard_Report.md           <- Markdown Documentation
    │
    └── Task_2_Regulatory_Gap_Analysis/
        ├── generate_icaap_paper.py                   <- Script generating 2-page research paper review
        ├── RBI_ICAAP_vs_Basel_III_Gap_Analysis.pdf   <- Deliverable: STRICT 2-PAGE RESEARCH PAPER REVIEW
        ├── RBI_ICAAP_vs_Basel_III_Gap_Analysis.docx  <- Academic Word Document
        └── RBI_ICAAP_vs_Basel_III_Gap_Analysis.md    <- Markdown Report
```

---

## Detailed Task Summaries

### Unit 1: Risk Introduction and Impacts

#### Task 1: Enterprise Risk Register & Management Plan (Paytm / One97 Communications Ltd)
- **Company Examined:** One97 Communications Limited (Paytm) [NSE: PAYTM | BSE: 543396] post-RBI January 31, 2024 Section 35A supervisory action against Paytm Payments Bank Limited (PPBL).
- **Core Requirements Met:**
  - Evaluated **15 comprehensive corporate risks** spanning Regulatory & Compliance, Operational Continuity, Strategic, IT & Cybersecurity, Fraud, Human Capital, Liquidity, Legal, and Technological Obsolescence.
  - Multi-dimensional scoring: Likelihood (1–5), Impact (1–5), and Risk Score ($L \times I$).
  - High-resolution **5x5 Risk Heat Map** (`risk_heatmap.png`) with color-coded severity bands (Critical, High, Moderate, Low).
  - Four-Pillar Strategic Mitigation Architecture across 0–3M, 3–6M, 6–12M, and 12–24M horizons.
- **Deliverables:**
  - `Paytm_Risk_Register_and_Management_Plan.pdf` (Strictly **5 Pages**, formatted with running headers/footers and audit sign-off).
  - `Paytm_Risk_Register_and_Management_Plan.docx` and `.md`.

#### Task 2: Risk Analytics Tool Exploration (OpenFAIR vs. RiskLens vs. LogicManager)
- **Core Requirements Met:**
  - Evaluated on **5 mandatory parameters:** Cost, Features, Scalability, Reporting, Compliance.
  - Deep pros and cons evaluation comparing open-source mathematical simulation (`pyfair` / OpenFAIR) with commercial cyber risk quantification (RiskLens SaaS) and enterprise GRC platforms (LogicManager).
- **Deliverables:**
  - `Risk_Analytics_Tools_Comparison.pptx`: **10 Widescreen Slides** designed with corporate dark navy and electric blue styling, structured cards, data tables, and strategic recommendations.
  - `Tool_Exploration_Comparative_Table.docx`, `.pdf`, and `.md`.

---

### Unit 2: Banking Sector

#### Task 1a: Credit Scorecard Modeling (Reference: Naeem Siddiqi)
- **Methodology Reference:** *Credit Risk Scorecards: Developing and Implementing Intelligent Credit Scoring* by Naeem Siddiqi (John Wiley & Sons, 2006).
- **Dataset:** Kaggle German Credit Risk benchmark dataset (1,000 credit applications, 20 predictors).
- **Pipeline Executed:**
  1. *Target Definition:* 1 = Bad (Default), 0 = Good.
  2. *Data Partitioning:* Stratified 75% Training ($N=750$), 25% Testing ($N=250$).
  3. *Weight of Evidence (WOE) & Information Value (IV):* Screened features with $IV \ge 0.05$ (Status of checking account $IV=0.60$, Credit history $IV=0.27$, Duration $IV=0.25$, Credit amount $IV=0.24$, Purpose $IV=0.23$, Savings $IV=0.22$, etc.).
  4. *Logistic Regression:* Fit on WOE variables with monotonic coefficients.
  5. *Scorecard Scaling:*
     $$\text{Factor} = \frac{\text{PDO}}{\ln(2)} = \frac{20}{\ln(2)} = 28.8539$$
     $$\text{Offset} = \text{Target Score} - \text{Factor} \times \ln(\text{Target Odds}) = 600 - 28.8539 \times \ln(50) = 487.1235$$
  6. *Model Validation Metrics:*
     - **Train ROC-AUC:** `0.8262` | **Train Gini:** `0.6523` | **Train KS:** `53.78%`
     - **Test ROC-AUC:** `0.7765` | **Test Gini:** `0.5529` | **Test KS:** `44.00%` (Exceeds industry standard $\ge 40\%$)
  7. *Regulatory Parameters & Expected Loss:*
     $$\text{Expected Loss (EL)} = \text{PD} \times \text{LGD} \times \text{EAD}$$
     - Total Portfolio Exposure at Default (EAD): **€3,271,258.00**
     - Total Estimated Expected Loss (EL): **€463,047.02**
     - Portfolio Expected Loss Rate: **14.16%**
     - Monotonic risk gradient verified from Grade AAA (Default rate: 2.92%, EL: 0.93% of EAD) to Grade F (Default rate: 68.03%, EL: 30.49% of EAD).
- **Deliverables:**
  - `credit_scorecard_modeling.ipynb`: Complete runnable Jupyter Notebook with markdown narratives and inline plots.
  - `credit_scorecard.py`: Standalone script for automated pipeline reproduction.
  - `Credit_Risk_Scorecard_Report.docx`, `.pdf`, and `.md`.

#### Task 2: Regulatory Gap Analysis (RBI ICAAP Guidelines vs. Basel III)
- **Paper Critique:** Examined the Reserve Bank of India's (RBI) Master Circular on Basel III Capital Regulations and ICAAP guidelines versus the Basel Committee on Banking Supervision (BCBS) global standards.
- **Identified 10 Key Structural Differences:**
  1. Minimum Common Equity Tier 1 (CET1): 5.5% (RBI) vs. 4.5% (Basel III) [+100 bps gold-plating].
  2. Minimum Tier 1 Capital: 7.0% (RBI) vs. 6.0% (Basel III).
  3. Minimum Total Capital (CRAR): 9.0% (RBI) vs. 8.0% (Basel III) [11.5% with CCB].
  4. Capital Conservation Buffer (CCB): Multiple COVID-19 countercyclical deferrals.
  5. Countercyclical Capital Buffer (CCCB): Maintained at 0% with reliance on targeted risk-weight hikes.
  6. Sovereign Debt Risk Weights: 0% risk weight for all Central and State government debt (including SDLs).
  7. Liquidity Framework: Integration of SLR (18%) via FALLCR and MSF carve-outs into Level 1 HQLA.
  8. Asset Classification: Incurred loss model with 90-day NPA provisioning vs. IFRS 9 ECL.
  9. Interest Rate Risk in Banking Book (IRRBB): Explicit Pillar 2 capital add-on under ICAAP.
  10. D-SIB Surcharges & Board Oversight: CET1 surcharges for SBI, HDFC Bank, ICICI Bank, with formal board sign-off.
- **Deliverables:**
  - `RBI_ICAAP_vs_Basel_III_Gap_Analysis.pdf`: Strict **2-Page Academic Research Paper Review** with abstract, comparative matrix, deep critique, and references.
  - `RBI_ICAAP_vs_Basel_III_Gap_Analysis.docx` and `.md`.
\
