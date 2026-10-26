import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from docx import Document
from docx.shared import Inches as DocInches, Pt as DocPt, RGBColor as DocRGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas

BASE_DIR = r"c:\Users\Moksh\Desktop\risk analytics assignment\Unit_1_Risk_Introduction_and_Impacts\Task_2_Tool_Exploration"
os.makedirs(BASE_DIR, exist_ok=True)

# -------------------------------------------------------------
# 1. GENERATE 10-SLIDE PRESENTATION (.PPTX)
# -------------------------------------------------------------
def build_pptx():
    pptx_path = os.path.join(BASE_DIR, "Risk_Analytics_Tools_Comparison.pptx")
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5) # 16:9 Widescreen
    blank_layout = prs.slide_layouts[6]

    # Corporate Colors
    NAVY = RGBColor(11, 37, 69)      # #0B2545 - Primary Dark
    BLUE = RGBColor(19, 64, 116)     # #134074 - Secondary Blue
    ACCENT = RGBColor(0, 150, 214)   # #0096D6 - Vibrant Cyan/Accent
    LIGHT_BG = RGBColor(245, 247, 250) # Card BG
    WHITE = RGBColor(255, 255, 255)
    GRAY_TEXT = RGBColor(100, 110, 125)
    DARK_TEXT = RGBColor(30, 41, 59)
    GREEN = RGBColor(39, 174, 96)
    RED = RGBColor(231, 76, 60)
    GOLD = RGBColor(243, 156, 18)

    def add_header(slide, title_text, category_text="UNIT 1 - TASK 2: TOOL EXPLORATION"):
        # Header banner shape
        header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(1.1))
        tf = header_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        
        p_cat = tf.paragraphs[0]
        p_cat.text = category_text.upper()
        p_cat.font.name = "Calibri"
        p_cat.font.size = Pt(11)
        p_cat.font.bold = True
        p_cat.font.color.rgb = ACCENT

        p_title = tf.add_paragraph()
        p_title.text = title_text
        p_title.font.name = "Calibri"
        p_title.font.size = Pt(24)
        p_title.font.bold = True
        p_title.font.color.rgb = NAVY

    def add_card(slide, left, top, width, height, title, subtitle=None, bg_color=LIGHT_BG, border_color=None):
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = bg_color
        if border_color:
            shape.line.color.rgb = border_color
            shape.line.width = Pt(1.5)
        else:
            shape.line.fill.background()
        
        # Add textbox inside
        tb = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.2), width - Inches(0.4), height - Inches(0.4))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p = tf.paragraphs[0]
        p.text = title
        p.font.name = "Calibri"
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = NAVY

        if subtitle:
            p2 = tf.add_paragraph()
            p2.text = subtitle
            p2.font.name = "Calibri"
            p2.font.size = Pt(11)
            p2.font.color.rgb = ACCENT
            p2.space_after = Pt(8)

        return tf

    # ---------------- SLIDE 1: TITLE SLIDE ----------------
    slide1 = prs.slides.add_slide(blank_layout)
    bg1 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = NAVY
    bg1.line.fill.background()

    # Title box
    tb_t = slide1.shapes.add_textbox(Inches(1.2), Inches(1.5), Inches(10.9), Inches(4.5))
    tf_t = tb_t.text_frame
    tf_t.word_wrap = True

    p_inst = tf_t.paragraphs[0]
    p_inst.text = "SRM INSTITUTE OF SCIENCE AND TECHNOLOGY\nCOLLEGE OF ENGINEERING AND TECHNOLOGY | SCHOOL OF COMPUTING\nDEPARTMENT OF DATA SCIENCE AND BUSINESS SYSTEMS"
    p_inst.font.name = "Calibri"
    p_inst.font.size = Pt(13)
    p_inst.font.bold = True
    p_inst.font.color.rgb = ACCENT
    p_inst.space_after = Pt(20)

    p_main = tf_t.add_paragraph()
    p_main.text = "QUANTITATIVE RISK ANALYTICS TOOL EXPLORATION"
    p_main.font.name = "Calibri"
    p_main.font.size = Pt(36)
    p_main.font.bold = True
    p_main.font.color.rgb = WHITE
    p_main.space_after = Pt(10)

    p_sub = tf_t.add_paragraph()
    p_sub.text = "Comparative Benchmark: OpenFAIR vs. RiskLens vs. LogicManager Across 5 Core Enterprise Dimensions"
    p_sub.font.name = "Calibri"
    p_sub.font.size = Pt(18)
    p_sub.font.color.rgb = RGBColor(200, 215, 235)
    p_sub.space_after = Pt(35)

    p_meta = tf_t.add_paragraph()
    p_meta.text = "Course: 21CSE449T – Risk Analytics | Academic Year: 2026–27 (ODD)\nFaculty In-Charge: Dr. V. Prasanna | Branch: B.Tech – CSE (Big Data Analytics & Data Science)"
    p_meta.font.name = "Calibri"
    p_meta.font.size = Pt(13)
    p_meta.font.color.rgb = WHITE

    # ---------------- SLIDE 2: INDUSTRY CONTEXT ----------------
    slide2 = prs.slides.add_slide(blank_layout)
    add_header(slide2, "Executive Context: The Paradigm Shift in Risk Quantification")
    
    tf_c1 = add_card(slide2, Inches(0.8), Inches(1.8), Inches(3.6), Inches(5.0), "Traditional Qualitative Risk", "Heatmaps & High/Med/Low Scales")
    p = tf_c1.add_paragraph()
    p.text = "• Subjective Scoring: Relies on opinion-based ordinal scales (e.g., 'Likelihood 4 × Impact 4').\n• The 'Flaw of Averages': Obscures tail risks and compound exposure scenarios.\n• Inactionable for CFOs: Boards cannot allocate millions in cyber capital based on 'yellow' or 'amber' colors.\n• Incomparable Portfolios: Fails to correlate disparate threat communities across operational divisions."
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.color.rgb = DARK_TEXT

    tf_c2 = add_card(slide2, Inches(4.8), Inches(1.8), Inches(3.6), Inches(5.0), "Quantitative Cyber Risk (CRQ)", "FAIR Standard & Probabilistic Models", border_color=ACCENT)
    p = tf_c2.add_paragraph()
    p.text = "• Factor Analysis of Information Risk: The only international standard (Open Group) for value-at-risk modeling.\n• Probabilistic Monte Carlo: Simulates tens of thousands of loss events per year to derive Loss Exceedance Curves (LEC).\n• Financial Loss Metrics: Outputs Expected Annual Loss (ALE), Value at Risk (VaR), and probable maximum loss (PML) in currency.\n• Defensible ROI: Enables precise cost-benefit calculation for cybersecurity and insurance controls."
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.color.rgb = DARK_TEXT

    tf_c3 = add_card(slide2, Inches(8.8), Inches(1.8), Inches(3.6), Inches(5.0), "Enterprise Scope: Tool Spectrum", "From Open Scripting to Full GRC")
    p = tf_c3.add_paragraph()
    p.text = "• Open-Source FAIR (pyfair, OpenFAIR): Pure quantitative algorithms, high practitioner control, zero software licensing fees.\n• Commercial FAIR (RiskLens / RSA Archer CRQ): Enterprise SaaS, automated data ingestion, pre-calibrated loss tables, board dashboards.\n• Integrated ERM (LogicManager): Full-spectrum GRC, qualitative controls, regulatory mapping, ERM operational taxonomy."
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.color.rgb = DARK_TEXT

    # ---------------- SLIDE 3: CONTENDER OVERVIEW ----------------
    slide3 = prs.slides.add_slide(blank_layout)
    add_header(slide3, "Contender Profiles: The Evaluated Risk Analytics Platforms")

    tf_p1 = add_card(slide3, Inches(0.8), Inches(1.8), Inches(3.6), Inches(5.0), "OpenFAIR (Open-Source)", "Standard by The Open Group / pyfair")
    p = tf_p1.add_paragraph()
    p.text = "• Platform Type: Open-source computational framework and taxonomy standard.\n• Underlying Tech: Python (`pyfair`), R (`Evaluate-Core`), Excel/Jupyter.\n• Primary Focus: Cyber & operational risk quantification via Monte Carlo simulation.\n• Target Audience: Data scientists, risk modelers, academic researchers, and boutique consultancies.\n• Key Strengths: 100% transparent code, no vendor lock-in, infinite customization."
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.color.rgb = DARK_TEXT

    tf_p2 = add_card(slide3, Inches(4.8), Inches(1.8), Inches(3.6), Inches(5.0), "RiskLens (Commercial CRQ)", "Industry Leader in FAIR SaaS")
    p = tf_p2.add_paragraph()
    p.text = "• Platform Type: Enterprise SaaS platform built by the creators of FAIR.\n• Underlying Tech: Proprietary cloud simulation engine with API connectors.\n• Primary Focus: Financial cyber risk quantification and executive decision support.\n• Target Audience: CISOs, CROs, Fortune 500 enterprises, and financial institutions.\n• Key Strengths: Pre-packaged threat tables, benchmark loss data, C-suite reporting, automated triage."
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.color.rgb = DARK_TEXT

    tf_p3 = add_card(slide3, Inches(8.8), Inches(1.8), Inches(3.6), Inches(5.0), "LogicManager (Integrated ERM)", "Leading Enterprise GRC & Risk Software")
    p = tf_p3.add_paragraph()
    p.text = "• Platform Type: Enterprise Governance, Risk, and Compliance (GRC) platform.\n• Underlying Tech: Relational SaaS platform with risk-taxonomy cross-mapping.\n• Primary Focus: Holistic ERM, operational risk registers, audit, and regulatory compliance.\n• Target Audience: Chief Compliance Officers, ERM directors, and internal audit teams.\n• Key Strengths: Pre-built compliance frameworks (SOX, HIPAA, ISO), drag-and-drop workflows."
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.color.rgb = DARK_TEXT

    # ---------------- SLIDE 4: PARAMETER 1 - COST ----------------
    slide4 = prs.slides.add_slide(blank_layout)
    add_header(slide4, "Evaluation Parameter 1: Total Cost of Ownership (TCO) & Licensing")

    tf_cost1 = add_card(slide4, Inches(0.8), Inches(1.8), Inches(3.6), Inches(5.0), "OpenFAIR (Open-Source)", "TCO: Low Software / High Human Capital")
    p = tf_cost1.add_paragraph()
    p.text = "• License Fee: $0 (GNU GPL / Apache 2.0 / Open Group Open Access).\n• Infrastructure Cost: Negligible ($0-$5k for internal servers or Jupyterhub).\n• Hidden Implementation Costs: Requires senior data scientists / quantitative analysts ($120k-$180k/yr).\n• Maintenance: Manual code patching, pipeline debugging, and internal script upkeep.\n• Verdict: Ideal for budget-constrained teams with strong quantitative Python/R engineering skills."
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.color.rgb = DARK_TEXT

    tf_cost2 = add_card(slide4, Inches(4.8), Inches(1.8), Inches(3.6), Inches(5.0), "RiskLens (Commercial CRQ)", "TCO: High SaaS / Rapid Value Delivery")
    p = tf_cost2.add_paragraph()
    p.text = "• License Fee: Premium Enterprise SaaS ($50k - $250k+ annually depending on modules/assets).\n• Implementation: $20k - $60k onboarding & FAIR academy professional services.\n• Infrastructure: 0 (Managed Cloud SaaS with high availability).\n• Operational Efficiency: Pre-calibrated industry data reduces analysis time from weeks to hours.\n• Verdict: Justified for large enterprises where a single cyber capital decision exceeds $2M."
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.color.rgb = DARK_TEXT

    tf_cost3 = add_card(slide4, Inches(8.8), Inches(1.8), Inches(3.6), Inches(5.0), "LogicManager (ERM Platform)", "TCO: Moderate to High SaaS / Broad Utility")
    p = tf_cost3.add_paragraph()
    p.text = "• License Fee: Subscription based on user tiers ($25k - $120k annually).\n• Unlimited User Model: Many plans offer unlimited end-user survey and incident loggers.\n• Implementation: $15k - $40k for workflow configuration and taxonomy setup.\n• Multi-Department Amortization: Cost is shared across Audit, Compliance, Legal, and IT.\n• Verdict: High cost-efficiency if replacing multiple standalone GRC point tools."
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.color.rgb = DARK_TEXT

    # ---------------- SLIDE 5: PARAMETER 2 - FEATURES ----------------
    slide5 = prs.slides.add_slide(blank_layout)
    add_header(slide5, "Evaluation Parameter 2: Features & Quantitative Modeling Depth")

    tf_f1 = add_card(slide5, Inches(0.8), Inches(1.8), Inches(3.6), Inches(5.0), "OpenFAIR (Open-Source)", "Mathematical Purity & Flexibility")
    p = tf_f1.add_paragraph()
    p.text = "• Core Math Engine: Monte Carlo simulations (Beta-PERT, log-normal, triangular distributions).\n• Modeling Granularity: Full access to Loss Event Frequency (LEF) and Loss Magnitude (LM) formulas.\n• Custom Distributions: Practitioners can implement any statistical copula or custom fat-tailed loss model.\n• Weaknesses: No out-of-the-box asset inventory, no automated questionnaire engine, command-line only."
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.color.rgb = DARK_TEXT

    tf_f2 = add_card(slide5, Inches(4.8), Inches(1.8), Inches(3.6), Inches(5.0), "RiskLens (Commercial CRQ)", "Specialized CRQ & Loss Calibration")
    p = tf_f2.add_paragraph()
    p.text = "• FAIR Enterprise Model: Proprietary algorithms expanding the standard FAIR ontology.\n• Cyber Risk Triage: Rapid screening to identify high-exposure assets in minutes.\n• Benchmark Loss Library: Built-in empirical cost tables for ransomware payouts, records breaches, and business interruption.\n• Scenario Comparison: Side-by-side simulation of control investments (e.g., EDR vs. MFA impact on ALE)."
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.color.rgb = DARK_TEXT

    tf_f3 = add_card(slide5, Inches(8.8), Inches(1.8), Inches(3.6), Inches(5.0), "LogicManager (ERM Platform)", "Qualitative Breadth & Workflow Automation")
    p = tf_f3.add_paragraph()
    p.text = "• Taxonomy Interconnectedness: Connects risks across vendors, regulations, business processes, and incidents.\n• Automated Assessment: Survey workflows, automated reminders, and root cause analysis.\n• Quantitative Limitations: Primarily relies on qualitative / semi-quantitative matrices; lacks native Monte Carlo engines.\n• Incident Management: Real-time logging of operational loss events and remediation ticketing."
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.color.rgb = DARK_TEXT

    # ---------------- SLIDE 6: PARAMETER 3 - SCALABILITY ----------------
    slide6 = prs.slides.add_slide(blank_layout)
    add_header(slide6, "Evaluation Parameter 3: Architecture, Scalability & Integrations")

    tf_s1 = add_card(slide6, Inches(0.8), Inches(1.8), Inches(3.6), Inches(5.0), "OpenFAIR (Open-Source)", "Local Workstation to HPC Scripts")
    p = tf_s1.add_paragraph()
    p.text = "• Execution Architecture: Client-side Python/R script or Jupyter Notebook runtime.\n• Computational Scale: High simulation throughput when executed on multicore CPUs or GPU clusters.\n• Enterprise Scalability: Poor; lacks multi-tenant user access controls, role-based security, and database concurrency.\n• System Integrations: Requires custom-built REST APIs and ETL pipelines to connect with CMDB or SIEM."
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.color.rgb = DARK_TEXT

    tf_s2 = add_card(slide6, Inches(4.8), Inches(1.8), Inches(3.6), Inches(5.0), "RiskLens (Commercial CRQ)", "Cloud-Native Enterprise CRQ Pipeline")
    p = tf_s2.add_paragraph()
    p.text = "• Execution Architecture: Multi-tenant cloud SaaS with elastic simulation microservices.\n• Computational Scale: Runs millions of Monte Carlo iterations asynchronously across distributed cloud nodes.\n• Enterprise Connectors: Pre-built API integrations with ServiceNow CMDB, Splunk, Qualys, and RSA Archer.\n• Continuous Quantification: Ingests vulnerability and asset telemetry to update risk scores dynamically."
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.color.rgb = DARK_TEXT

    tf_s3 = add_card(slide6, Inches(8.8), Inches(1.8), Inches(3.6), Inches(5.0), "LogicManager (ERM Platform)", "Enterprise-Wide Collaborative Scaling")
    p = tf_s3.add_paragraph()
    p.text = "• Execution Architecture: Enterprise cloud GRC database with configurable role hierarchies.\n• Organizational Scale: Easily scales to tens of thousands of business users submitting risk surveys.\n• Ecosystem Integrations: Connects with Slack, Microsoft Teams, Active Directory, Jira, and Workday.\n• Data Architecture: Graph-like relational linking between regulatory mandates, risks, controls, and audits."
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.color.rgb = DARK_TEXT

    # ---------------- SLIDE 7: PARAMETER 4 - REPORTING ----------------
    slide7 = prs.slides.add_slide(blank_layout)
    add_header(slide7, "Evaluation Parameter 4: Executive Reporting & Decision Support")

    tf_r1 = add_card(slide7, Inches(0.8), Inches(1.8), Inches(3.6), Inches(5.0), "OpenFAIR (Open-Source)", "Technical Statistical Charts")
    p = tf_r1.add_paragraph()
    p.text = "• Visual Outputs: Matplotlib/Seaborn probability density functions (PDF) and cumulative distribution functions (CDF).\n• Loss Exceedance Curves: Generates raw percentile curves ($1M, $5M, $10M loss thresholds).\n• Stakeholder Usability: Requires technical translation for non-technical executives and board members.\n• Report Generation: Manual export of charts to Word, PowerPoint, or LaTeX."
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.color.rgb = DARK_TEXT

    tf_r2 = add_card(slide7, Inches(4.8), Inches(1.8), Inches(3.6), Inches(5.0), "RiskLens (Commercial CRQ)", "Board-Ready Financial Dashboards")
    p = tf_r2.add_paragraph()
    p.text = "• Financial Risk Framing: Presents risk strictly in currency terms (e.g., '$14.2M Annualized Loss Exposure').\n• Executive Action Dashboards: Top 10 Cyber Risk Scenarios, Capital at Risk, and Insurance Adequacy metrics.\n• Cost-Benefit ROI Modeling: Visualizes risk reduction per dollar spent across competing security investments.\n• Board Presentation Ready: One-click export of executive-ready summary decks and SEC-compliant disclosures."
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.color.rgb = DARK_TEXT

    tf_r3 = add_card(slide7, Inches(8.8), Inches(1.8), Inches(3.6), Inches(5.0), "LogicManager (ERM Platform)", "Comprehensive Audit & Governance Heatmaps")
    p = tf_r3.add_paragraph()
    p.text = "• Visual Formats: Enterprise heatmaps, risk matrices, compliance radar charts, and audit dashboards.\n• Role-Based Views: Granular dashboards for front-line managers, audit committees, and executive directors.\n• Deficiency Tracking: Real-time remediation tracking reports with automated aging alerts.\n• Regulatory Package Generation: Pre-configured audit packages for external regulators and statutory auditors."
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.color.rgb = DARK_TEXT

    # ---------------- SLIDE 8: PARAMETER 5 - COMPLIANCE ----------------
    slide8 = prs.slides.add_slide(blank_layout)
    add_header(slide8, "Evaluation Parameter 5: Regulatory Compliance & Industry Standards")

    tf_cp1 = add_card(slide8, Inches(0.8), Inches(1.8), Inches(3.6), Inches(5.0), "OpenFAIR (Open-Source)", "Pure Methodological Compliance")
    p = tf_cp1.add_paragraph()
    p.text = "• Standards Alignment: Fully conforms to The Open Group OpenFAIR (O-RA and O-RT standards).\n• ISO/NIST Alignment: Provides quantitative defensibility for ISO/IEC 27005 and NIST SP 800-30.\n• Audit Trail: Low out-of-the-box auditability; requires strict manual version control (Git) of scripts and input files.\n• Regulatory Status: Widely respected by quantitative regulators (OCC, Fed, PRA) when documented rigorously."
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.color.rgb = DARK_TEXT

    tf_cp2 = add_card(slide8, Inches(4.8), Inches(1.8), Inches(3.6), Inches(5.0), "RiskLens (Commercial CRQ)", "SEC Cyber & Global Financial Mandates")
    p = tf_cp2.add_paragraph()
    p.text = "• Regulatory Mandates: Specifically designed to satisfy SEC Cybersecurity Disclosure Rules (Item 106 & Form 8-K).\n• Financial Standards: Directly supports Basel Committee Operational Risk AMA/SMA and DORA (Digital Operational Resilience Act).\n• Evidence & Traceability: Immutable audit logs of model assumptions, SME interviews, and calibration rationale.\n• Industry Benchmarks: Standardized risk profiles recognized by cyber insurance underwriters."
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.color.rgb = DARK_TEXT

    tf_cp3 = add_card(slide8, Inches(8.8), Inches(1.8), Inches(3.6), Inches(5.0), "LogicManager (ERM Platform)", "Universal Regulatory Governance Hub")
    p = tf_cp3.add_paragraph()
    p.text = "• Framework Libraries: Built-in mapping for SOX 404, ISO 31000, ISO 27001, COSO, HIPAA, GDPR, and PCI-DSS.\n• Regulatory Change Management: Automatically flags control gaps when central bank or statutory regulations update.\n• Audit Workflows: Complete workpaper documentation, sample testing, and independent audit sign-offs.\n• Policy Management: Built-in policy lifecycle management, employee attestations, and exception logging."
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.color.rgb = DARK_TEXT

    # ---------------- SLIDE 9: SYNTHESIS MATRIX ----------------
    slide9 = prs.slides.add_slide(blank_layout)
    add_header(slide9, "Comprehensive Evaluation Matrix: 5-Parameter Comparative Scorecard")

    # Add a styled table
    rows, cols = 6, 4
    left = Inches(0.8)
    top = Inches(1.8)
    width = Inches(11.733)
    height = Inches(4.9)

    tbl_shape = slide9.shapes.add_table(rows, cols, left, top, width, height)
    tbl = tbl_shape.table
    tbl.columns[0].width = Inches(2.333)
    tbl.columns[1].width = Inches(3.133)
    tbl.columns[2].width = Inches(3.133)
    tbl.columns[3].width = Inches(3.133)

    matrix_data = [
        ["Parameter", "OpenFAIR (Open-Source)", "RiskLens (Commercial CRQ)", "LogicManager (Enterprise GRC)"],
        [
            "1. Cost & TCO",
            "★★★★★ (Zero License)\n$0 license; requires skilled in-house quantitative analysts.",
            "★★☆☆☆ (High SaaS)\n$50k–$250k/yr; high initial investment, rapid ROI for cyber.",
            "★★★☆☆ (Moderate SaaS)\n$25k–$120k/yr; cost amortized across enterprise departments."
        ],
        [
            "2. Features & Modeling",
            "★★★★☆ (High Math Rigor)\nPure Monte Carlo, Beta-PERT, complete parameter control.",
            "★★★★★ (Benchmark CRQ)\nFAIR Enterprise, empirical loss data, control ROI simulation.",
            "★★★☆☆ (Broad GRC)\nQualitative scoring, risk matrices, incident workflows."
        ],
        [
            "3. Scalability & APIs",
            "★★☆☆☆ (Limited Multi-user)\nLocal Python/R runtime; requires custom API pipelines.",
            "★★★★☆ (Enterprise SaaS)\nElastic cloud simulations, ServiceNow/Splunk integrations.",
            "★★★★★ (Org-wide Scale)\nUnlimited user surveys, graph-relational database structure."
        ],
        [
            "4. Reporting & Dashboards",
            "★★☆☆☆ (Raw Statistical)\nMatplotlib curves, histograms; requires manual presentation.",
            "★★★★★ (CFO/Board Ready)\nDollarized ALE, Loss Exceedance Curves, Board decks.",
            "★★★★☆ (Audit & Compliance)\nHeatmaps, compliance dashboards, deficiency trackers."
        ],
        [
            "5. Regulatory & Audit",
            "★★★☆☆ (Standard Compliant)\nO-RA standard compliant; manual audit documentation.",
            "★★★★★ (SEC / DORA / Basel)\nSEC Item 106, DORA, Basel Pillar 2 operational risk.",
            "★★★★★ (Universal GRC)\nSOX 404, ISO 31000, ISO 27001, COSO, HIPAA, GDPR."
        ]
    ]

    for r_idx, row in enumerate(matrix_data):
        for c_idx, val in enumerate(row):
            cell = tbl.cell(r_idx, c_idx)
            cell.text = val
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if c_idx > 0 else PP_ALIGN.LEFT
            if r_idx == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = NAVY
                p.font.name = "Calibri"
                p.font.size = Pt(13)
                p.font.bold = True
                p.font.color.rgb = WHITE
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = WHITE if r_idx % 2 == 1 else RGBColor(245, 248, 252)
                p.font.name = "Calibri"
                p.font.size = Pt(10.5)
                p.font.color.rgb = DARK_TEXT

    # ---------------- SLIDE 10: STRATEGIC RECOMMENDATIONS ----------------
    slide10 = prs.slides.add_slide(blank_layout)
    add_header(slide10, "Strategic Recommendations: Institutional Tool Selection Blueprint")

    tf_rec1 = add_card(slide10, Inches(0.8), Inches(1.8), Inches(3.6), Inches(5.0), "Tier 1: Research & Startups", "Choose OpenFAIR (pyfair)", border_color=GREEN)
    p = tf_rec1.add_paragraph()
    p.text = "• Target Profile: Universities, academic risk labs, fintech startups, and internal quantitative R&D teams.\n• Decision Rationale: Maximizes mathematical transparency without capital expenditure.\n• Implementation Path: Deploy containerized JupyterHub with `pyfair` and `Evaluate-Core`.\n• Key Prerequisite: In-house staff proficient in probability distributions and Python/R."
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.color.rgb = DARK_TEXT

    tf_rec2 = add_card(slide10, Inches(4.8), Inches(1.8), Inches(3.6), Inches(5.0), "Tier 2: Systemic Banks & Large Firms", "Choose RiskLens (Commercial FAIR)", border_color=BLUE)
    p = tf_rec2.add_paragraph()
    p.text = "• Target Profile: Tier-1 Banks, listed enterprises, critical infrastructure, and high-scrutiny fintechs.\n• Decision Rationale: Satisfies board fiduciary mandates and SEC/central bank cyber disclosure requirements.\n• Implementation Path: Integrate RiskLens SaaS with enterprise CMDB and SIEM for automated data ingestion.\n• Key Prerequisite: Dedicated enterprise budget and formal C-suite mandate for financial risk quantification."
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.color.rgb = DARK_TEXT

    tf_rec3 = add_card(slide10, Inches(8.8), Inches(1.8), Inches(3.6), Inches(5.0), "Tier 3: Mid-Market & Broad ERM", "Choose LogicManager (GRC Hub)", border_color=GOLD)
    p = tf_rec3.add_paragraph()
    p.text = "• Target Profile: Regulated mid-market firms, healthcare systems, insurance, and multi-regulatory organizations.\n• Decision Rationale: Replaces fragmented spreadsheets with unified audit, compliance, and policy workflows.\n• Hybrid Architecture Recommendation: Combine LogicManager as the central GRC registry with OpenFAIR/RiskLens for quantitative cyber modules."
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.color.rgb = DARK_TEXT

    prs.save(pptx_path)
    print(f"10-Slide Presentation (.pptx) generated successfully at: {pptx_path}")

# -------------------------------------------------------------
# 2. GENERATE COMPARATIVE TABLE & REPORT (.DOCX, .PDF, .MD)
# -------------------------------------------------------------
def build_comparative_table_docs():
    docx_path = os.path.join(BASE_DIR, "Tool_Exploration_Comparative_Table.docx")
    pdf_path = os.path.join(BASE_DIR, "Tool_Exploration_Comparative_Table.pdf")
    md_path = os.path.join(BASE_DIR, "Tool_Exploration_Comparative_Table.md")

    # 1. MARKDOWN
    md_content = """# SRM INSTITUTE OF SCIENCE AND TECHNOLOGY
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
"""
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(md_content)

    # 2. WORD (.DOCX)
    doc = Document()
    for s in doc.sections:
        s.top_margin = DocInches(0.8)
        s.bottom_margin = DocInches(0.8)
        s.left_margin = DocInches(0.8)
        s.right_margin = DocInches(0.8)

    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_inst = p_inst.add_run("SRM INSTITUTE OF SCIENCE AND TECHNOLOGY\nCOLLEGE OF ENGINEERING AND TECHNOLOGY | SCHOOL OF COMPUTING\nDEPARTMENT OF DATA SCIENCE AND BUSINESS SYSTEMS\n")
    r_inst.font.name = "Calibri"
    r_inst.font.size = DocPt(11)
    r_inst.font.bold = True
    r_inst.font.color.rgb = DocRGBColor(19, 64, 116)

    p_t = doc.add_paragraph()
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_t = p_t.add_run("UNIT 1 - TASK 2: RISK ANALYTICS TOOL EXPLORATION\nComparative Evaluation of OpenFAIR, RiskLens, and LogicManager\n5-Parameter Benchmark: Cost, Features, Scalability, Reporting, Compliance")
    r_t.font.name = "Calibri"
    r_t.font.size = DocPt(15)
    r_t.font.bold = True
    r_t.font.color.rgb = DocRGBColor(11, 37, 69)

    doc.add_paragraph().add_run("—" * 60).font.color.rgb = DocRGBColor(11, 37, 69)

    doc.add_heading("1. Executive Summary", level=1).runs[0].font.color.rgb = DocRGBColor(11, 37, 69)
    doc.add_paragraph(
        "Modern enterprise risk management is transitioning from qualitative heat maps toward quantitative cyber risk analytics (CRQ). "
        "This evaluation systematically compares OpenFAIR (open source standard), RiskLens (commercial SaaS FAIR), and LogicManager (integrated GRC) "
        "across the 5 mandatory dimensions: Cost, Features, Scalability, Reporting, and Compliance."
    )

    doc.add_heading("2. Five-Parameter Comparative Benchmark Matrix", level=1).runs[0].font.color.rgb = DocRGBColor(11, 37, 69)
    tbl = doc.add_table(rows=1, cols=4)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = tbl.rows[0].cells
    hdrs = ["Parameter", "OpenFAIR (Open-Source)", "RiskLens (Commercial FAIR)", "LogicManager (Enterprise GRC)"]
    for i, h in enumerate(hdrs):
        hdr[i].text = h
        shd = parse_xml(r'<w:shd {} w:fill="0B2545"/>'.format(nsdecls('w')))
        hdr[i]._tc.get_or_add_tcPr().append(shd)
        p = hdr[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = DocPt(9)
            run.font.bold = True
            run.font.color.rgb = DocRGBColor(255, 255, 255)

    matrix_rows = [
        ("1. Total Cost of Ownership", "License: $0.\nRequires in-house quant analysts.\nRating: 5/5 (Free)", "License: $50k-$250k/yr.\nHigh SaaS cost, rapid cyber ROI.\nRating: 2/5 (Expensive)", "License: $25k-$120k/yr.\nAmortized across departments.\nRating: 3/5 (Moderate)"),
        ("2. Features & Modeling", "Pure Monte Carlo, Beta-PERT, full parameter control.\nRating: 4/5 (Rigor)", "Empirical loss tables, control ROI simulator, FAIR Enterprise.\nRating: 5/5 (Gold Standard)", "Qualitative scoring, survey workflows, incident logging.\nRating: 3/5 (Qualitative)"),
        ("3. Scalability & APIs", "Local workstation runtime, manual API scripting.\nRating: 2/5 (Limited)", "Elastic cloud microservices, ServiceNow/Splunk APIs.\nRating: 4/5 (Enterprise SaaS)", "Unlimited user surveys, graph relational GRC database.\nRating: 5/5 (Org-wide)"),
        ("4. Reporting & Output", "Statistical histograms, CDF/PDF curves.\nRating: 2/5 (Technical)", "Dollarized ALE, Loss Exceedance Curves, Board decks.\nRating: 5/5 (Board Ready)", "Heatmaps, compliance radar, deficiency tracking.\nRating: 4/5 (Audit Focus)"),
        ("5. Regulatory Compliance", "Open Group OpenFAIR (O-RA / O-RT) standard compliant.\nRating: 3/5 (Sound)", "SEC Item 106, DORA, Basel Pillar 2 operational risk.\nRating: 5/5 (Reg Defense)", "SOX 404, ISO 31000, ISO 27001, COSO, HIPAA, GDPR.\nRating: 5/5 (Universal)")
    ]

    for m in matrix_rows:
        row = tbl.add_row().cells
        for idx in range(4):
            row[idx].text = m[idx]
            for run in row[idx].paragraphs[0].runs:
                run.font.name = "Calibri"
                run.font.size = DocPt(8)

    doc.add_heading("3. Pros and Cons Detailed Analysis", level=1).runs[0].font.color.rgb = DocRGBColor(11, 37, 69)
    doc.add_paragraph(
        "• OpenFAIR: Pros include zero license fees and full algorithmic transparency; Cons include high reliance on senior data scientists and lack of native reporting dashboards.\n"
        "• RiskLens: Pros include defensible dollar-denominated risk metrics and empirical benchmark loss databases; Cons include premium subscription costs and proprietary model enclosure.\n"
        "• LogicManager: Pros include universal regulatory mapping (SOX, ISO, HIPAA) and broad organizational workflow scaling; Cons include absence of probabilistic Monte Carlo risk engines."
    )

    doc.add_heading("4. Strategic Selection Roadmap", level=1).runs[0].font.color.rgb = DocRGBColor(11, 37, 69)
    doc.add_paragraph(
        "1. Academic Labs & Boutique Consultancies: Adopt OpenFAIR (pyfair) for mathematical flexibility at zero software cost.\n"
        "2. Listed Banks & Systemic Enterprises: Adopt RiskLens for defensible C-suite and SEC-mandated financial cyber quantification.\n"
        "3. Regulated Mid-Market Organizations: Adopt LogicManager for integrated ERM governance, supplemented with OpenFAIR scripts for targeted cyber assets."
    )

    doc.save(docx_path)

    # 3. PDF
    pdf_doc = SimpleDocTemplate(pdf_path, pagesize=letter, leftMargin=40, rightMargin=40, topMargin=40, bottomMargin=40)
    styles = getSampleStyleSheet()
    p_story = []

    p_story.append(Paragraph("SRM INSTITUTE OF SCIENCE AND TECHNOLOGY", ParagraphStyle('Sub', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor('#134074'), alignment=1)))
    p_story.append(Paragraph("School of Computing | Department of Data Science and Business Systems", ParagraphStyle('Sub2', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, textColor=colors.HexColor('#555555'), alignment=1)))
    p_story.append(Spacer(1, 6))
    p_story.append(HRFlowable(width="100%", thickness=1.2, color=colors.HexColor('#0b2545'), spaceBefore=2, spaceAfter=8))
    p_story.append(Paragraph("UNIT 1 - TASK 2: RISK ANALYTICS TOOL EXPLORATION", ParagraphStyle('T1', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=16, textColor=colors.HexColor('#0b2545'), alignment=1)))
    p_story.append(Paragraph("Comparative Evaluation: OpenFAIR vs. RiskLens vs. LogicManager Across 5 Parameters", ParagraphStyle('T2', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor('#134074'), alignment=1)))
    p_story.append(Paragraph("Course: 21CSE449T – Risk Analytics | Faculty: Dr. V. Prasanna | B.Tech – CSE (BDA & DS)", ParagraphStyle('T3', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, textColor=colors.HexColor('#555555'), alignment=1)))
    p_story.append(HRFlowable(width="100%", thickness=0.8, color=colors.HexColor('#0b2545'), spaceBefore=2, spaceAfter=10))

    pdf_tbl_data = [
        [
            Paragraph("<b>Parameter</b>", ParagraphStyle('th', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, textColor=colors.white, alignment=1)),
            Paragraph("<b>OpenFAIR (Open-Source)</b>", ParagraphStyle('th', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, textColor=colors.white, alignment=1)),
            Paragraph("<b>RiskLens (Commercial CRQ)</b>", ParagraphStyle('th', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, textColor=colors.white, alignment=1)),
            Paragraph("<b>LogicManager (Enterprise GRC)</b>", ParagraphStyle('th', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, textColor=colors.white, alignment=1))
        ]
    ]

    cell_style = ParagraphStyle('tc', parent=styles['Normal'], fontName='Helvetica', fontSize=7.2, leading=9.2, textColor=colors.HexColor('#222222'))
    cell_bold = ParagraphStyle('tcb', parent=cell_style, fontName='Helvetica-Bold')

    for m in matrix_rows:
        pdf_tbl_data.append([
            Paragraph(f"<b>{m[0]}</b>", cell_bold),
            Paragraph(m[1].replace("\n", "<br/>"), cell_style),
            Paragraph(m[2].replace("\n", "<br/>"), cell_style),
            Paragraph(m[3].replace("\n", "<br/>"), cell_style)
        ])

    pdf_tbl = Table(pdf_tbl_data, colWidths=[90, 150, 150, 142], repeatRows=1)
    pdf_tbl.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0b2545')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cccccc')),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8f9fa')])
    ]))

    p_story.append(pdf_tbl)
    p_story.append(Spacer(1, 10))

    p_story.append(Paragraph("<b>Strategic Implementation Summary:</b>", ParagraphStyle('h', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor('#0b2545'))))
    summary_txt = (
        "1. <b>OpenFAIR</b> delivers maximum statistical freedom and zero licensing overhead for academic and data-science practitioners.<br/>"
        "2. <b>RiskLens</b> represents the corporate benchmark for listed banks and enterprises needing dollar-based Value-at-Risk outputs for Board oversight and SEC disclosures.<br/>"
        "3. <b>LogicManager</b> provides universal GRC workflow integration across multiple compliance mandates (SOX, ISO, HIPAA, GDPR)."
    )
    p_story.append(Paragraph(summary_txt, ParagraphStyle('b', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=11, textColor=colors.HexColor('#333333'))))

    pdf_doc.build(p_story)
    print("Comparative Table documents (.docx, .pdf, .md) built successfully!")

if __name__ == "__main__":
    build_pptx()
    build_comparative_table_docs()
    print("Unit 1 - Task 2 complete!")
