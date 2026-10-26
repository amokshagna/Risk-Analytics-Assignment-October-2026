# SRM INSTITUTE OF SCIENCE AND TECHNOLOGY
### College of Engineering and Technology | School of Computing
### Department of Data Science and Business Systems
**Subject Code:** 21CSE449T | **Subject Name:** RISK ANALYTICS  
**Academic Year:** 2026–27 (ODD) | **Faculty Name:** Dr. V. Prasanna  
**Programme & Branch:** B.Tech – CSE (BDA & DS)

---

# Unit 1 - Task 2: Risk Analytics Tool Exploration
## Comparative Evaluation: OpenFAIR vs. RiskLens vs. LogicManager
### Systematic Benchmark Across 5 Key Enterprise Dimensions: Cost, Features, Scalability, Reporting, Compliance

---

## 1. Executive Summary
Modern enterprise risk management (ERM) is experiencing a historic structural transition from subjective ordinal scoring models (such as traditional "High/Medium/Low" 5x5 heat maps) toward defensible, mathematically validated **Quantitative Cyber Risk Analytics (CRQ)**. Executive leadership, statutory auditors, and regulatory bodies (including the Reserve Bank of India, US SEC, and European Central Bank) increasingly demand that operational and cybersecurity exposures be expressed in standardized financial metrics: Value at Risk (VaR), Annualized Loss Exposure (ALE), and Loss Exceedance Probabilities.

This study conducts an in-depth technical and financial exploration comparing two leading open-source and commercial risk analytics engines based on the **Factor Analysis of Information Risk (FAIR)** standard—**OpenFAIR** and **RiskLens**—while incorporating **LogicManager**, an industry-benchmark Enterprise Governance, Risk, and Compliance (GRC) platform, to evaluate organizational integration breadth.

---

## 2. Comparative Evaluation Matrix (5 Core Parameters)

| Evaluation Parameter | OpenFAIR (Open-Source Standard) | RiskLens (Commercial FAIR SaaS) | LogicManager (Enterprise GRC) |
|---|---|---|---|
| **1. Total Cost of Ownership (Cost & Licensing)** | • **License:** $0 (Free open-source / GNU GPL / Apache).<br/>• **Infrastructure:** Local workstations or internal JupyterHub ($0 - $5k).<br/>• **Human Capital:** High internal quantitative engineering costs ($120k-$180k/yr).<br/>• **Rating:** ★★★★★ (Zero software cost) | • **License:** Premium Enterprise SaaS ($50k - $250k+ annually).<br/>• **Implementation:** $20k - $60k professional onboarding & FAIR training.<br/>• **Infrastructure:** Managed cloud (Zero hosting overhead).<br/>• **Rating:** ★★☆☆☆ (High software expense) | • **License:** Tiered user subscription ($25k - $120k annually).<br/>• **Implementation:** $15k - $40k workflow configuration.<br/>• **Multi-Unit Sharing:** Amortized across Legal, Compliance, Audit, and IT.<br/>• **Rating:** ★★★☆☆ (Moderate corporate cost) |
| **2. Features & Quantitative Modeling Depth** | • **Math Core:** Probabilistic Monte Carlo simulation (Beta-PERT, log-normal).<br/>• **Parameter Customization:** Complete control over Loss Event Frequency (LEF) and Loss Magnitude (LM).<br/>• **Limitations:** Lacks built-in asset catalog, automated questionnaires, or threat feeds.<br/>• **Rating:** ★★★★☆ (High mathematical rigor) | • **Math Core:** Enhanced FAIR Enterprise simulation engine.<br/>• **Empirical Database:** Proprietary loss cost tables (ransomware, business interruption, records).<br/>• **Decision Support:** Control investment ROI simulator (e.g., EDR vs. MFA impact on ALE).<br/>• **Rating:** ★★★★★ (Gold standard in CRQ) | • **Math Core:** Qualitative & semi-quantitative risk scoring matrices.<br/>• **Workflow Engines:** Automated risk assessment surveys, remediation tracking.<br/>• **Limitations:** Lacks native Monte Carlo / probabilistic Loss Exceedance curves.<br/>• **Rating:** ★★★☆☆ (Qualitative GRC breadth) |
| **3. Architecture, Scalability & Integrations** | • **Architecture:** Standalone Python (`pyfair`) / R (`Evaluate-Core`) scripts.<br/>• **Scalability:** High throughput on local/HPC cores; poor multi-user enterprise tenancy.<br/>• **Integrations:** Requires custom REST API and ETL scripting for SIEM/CMDB.<br/>• **Rating:** ★★☆☆☆ (Single-user / developer-centric) | • **Architecture:** Multi-tenant elastic cloud microservices.<br/>• **Scalability:** Distributed parallel simulations processing millions of iterations.<br/>• **Integrations:** Turnkey API connectors for ServiceNow CMDB, Splunk, Qualys.<br/>• **Rating:** ★★★★☆ (Cloud-native enterprise grade) | • **Architecture:** Enterprise relational cloud GRC platform.<br/>• **Scalability:** Scales seamlessly to tens of thousands of corporate employees.<br/>• **Integrations:** Pre-built connectors with Jira, Workday, Azure AD, MS Teams.<br/>• **Rating:** ★★★★★ (Universal organizational scale) |
| **4. Reporting & Decision Support** | • **Visualizations:** Matplotlib / Seaborn histograms, probability density curves.<br/>• **Loss Metrics:** Raw statistical tables (10th, 50th, 90th percentile loss).<br/>• **Executive Translation:** Requires heavy manual interpretation for Board members.<br/>• **Rating:** ★★☆☆☆ (Technical / statistical focus) | • **Visualizations:** C-suite financial loss exceedance curves & exposure rankings.<br/>• **Loss Metrics:** Dollarized ALE, Value at Risk, Probable Maximum Loss.<br/>• **Executive Translation:** 1-click board presentation exports & SEC Item 106 reports.<br/>• **Rating:** ★★★★★ (Premier executive clarity) | • **Visualizations:** Multi-dimensional heat maps, compliance radar charts.<br/>• **Loss Metrics:** Residual risk scores, audit findings, control deficiency aging.<br/>• **Executive Translation:** Tailored dashboards for audit committees and front-line managers.<br/>• **Rating:** ★★★★☆ (Comprehensive audit oversight) |
| **5. Regulatory Compliance & Industry Standards** | • **Standards:** 100% compliant with The Open Group OpenFAIR (O-RA and O-RT).<br/>• **Frameworks:** Methodological backing for ISO 27005 and NIST SP 800-30.<br/>• **Audit Trail:** Manual version control via Git repository.<br/>• **Rating:** ★★★☆☆ (Methodologically sound) | • **Standards:** Designed for SEC Cybersecurity Disclosures (Item 106 & Form 8-K).<br/>• **Frameworks:** Basel Operational Risk SMA, DORA, and NYDFS 500.<br/>• **Audit Trail:** Immutable enterprise audit logs for model parameters and SME interviews.<br/>• **Rating:** ★★★★★ (Built for regulatory defense) | • **Standards:** Native mapping to SOX 404, ISO 31000, ISO 27001, COSO, HIPAA, GDPR.<br/>• **Frameworks:** Universal GRC regulatory change tracking.<br/>• **Audit Trail:** Comprehensive workpaper documentation, sample testing sign-offs.<br/>• **Rating:** ★★★★★ (Broadest regulatory library) |

---

## 3. Deep-Dive Pros and Cons Analysis

### 3.1 OpenFAIR (Open-Source Standard / pyfair)
- **Key Pros:**
  1. *Zero Licensing Overhead:* Unlimited deployment across unlimited nodes without recurring SaaS costs.
  2. *Algorithmic Transparency:* Complete visibility into underlying distributions, random number seeds, and mathematical equations.
  3. *Unconstrained Customization:* Data scientists can alter parameter copulas, tail distributions, and correlation matrices to model novel threats.
- **Key Cons:**
  1. *High Talent Dependency:* Demands skilled data scientists proficient in probability and scripting; inaccessible to traditional compliance officers.
  2. *No Enterprise Collaboration:* Lacks built-in authentication, role-based access control (RBAC), and multi-analyst consensus workflows.
  3. *Manual Reporting Pipeline:* Generating executive presentations requires manual data wrangling and chart formatting.

### 3.2 RiskLens (Commercial FAIR Enterprise SaaS)
- **Key Pros:**
  1. *Defensible Dollar-Based Metrics:* Directly translates technical cyber vulnerabilities into annualized financial exposure ($ ALE) understood by CFOs.
  2. *Empirical Industry Loss Tables:* Provides calibrated benchmark data for ransomware, business interruption, and data records, drastically accelerating triage.
  3. *Board & SEC Regulatory Alignment:* Turnkey reporting templates directly aligned with SEC cybersecurity disclosure rules and cyber insurance renewals.
- **Key Cons:**
  1. *Substantial Capital Expenditure:* High annual SaaS subscription costs ($50k–$250k+) limit adoption to large corporations and banks.
  2. *Proprietary IP Enclosure:* Core simulation logic and loss calibration databases are proprietary, preventing underlying code alterations.

### 3.3 LogicManager (Enterprise GRC Hub)
- **Key Pros:**
  1. *Comprehensive Cross-Taxonomy Mapping:* Connects an operational risk to underlying policies, IT controls, statutory regulations, and audit workpapers.
  2. *Organizational Inclusivity:* Unlimited user models allow decentralized front-line employees to submit incident tickets and compliance surveys.
- **Key Cons:**
  1. *Lack of Probabilistic Simulation:* Cannot execute Monte Carlo Loss Exceedance modeling to compute Value at Risk in currency terms.

---

## 4. Strategic Implementation Recommendations

1. **For Systemically Important Financial Institutions & Listed Enterprises (e.g., Paytm, Yes Bank):**
   - **Recommended Stack:** Implement **RiskLens** integrated into enterprise SIEM (Splunk) and CMDB (ServiceNow). The substantial software cost is easily justified by regulatory compliance with SEC/RBI mandates and optimized cyber insurance underwriting.
2. **For Fintech Startups, Academic Research Labs & Boutique Consultancies:**
   - **Recommended Stack:** Deploy **OpenFAIR (`pyfair`)** inside a managed Python/JupyterHub environment. This delivers maximum quantitative rigor at zero software licensing cost.
3. **For Diversified Enterprises Requiring Broad Regulatory Oversight (SOX, ISO, HIPAA):**
   - **Recommended Stack:** Deploy a **Hybrid Architecture** utilizing **LogicManager** as the central GRC operational hub, supplemented by OpenFAIR / RiskLens plugins for high-risk quantitative cyber assets.
