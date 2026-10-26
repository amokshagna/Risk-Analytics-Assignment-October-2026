import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, KeepTogether, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas

BASE_DIR = r"c:\Users\Moksh\Desktop\risk analytics assignment\Unit_2_Banking_Sector\Task_2_Regulatory_Gap_Analysis"
os.makedirs(BASE_DIR, exist_ok=True)

# -------------------------------------------------------------
# 1. 10 KEY REGULATORY DIFFERENCES DATASET
# -------------------------------------------------------------
diff_data = [
    {
        "num": "1",
        "dimension": "Common Equity Tier 1 (CET1) Minimum",
        "basel": "Minimum CET1 ratio of 4.5% of Risk-Weighted Assets (RWAs).",
        "rbi": "Minimum CET1 ratio of 5.5% of RWAs (100 bps super-equivalent gold-plating).",
        "rationale": "Buffers against higher emerging-market asset volatility and historical public-sector bank capital erosion."
    },
    {
        "num": "2",
        "dimension": "Minimum Tier 1 Capital Ratio",
        "basel": "Minimum Tier 1 capital ratio of 6.0% of RWAs.",
        "rbi": "Minimum Tier 1 capital ratio of 7.0% of RWAs (100 bps higher than BCBS standard).",
        "rationale": "Ensures that loss-absorption capacity remains anchored in permanent high-quality going-concern equity."
    },
    {
        "num": "3",
        "dimension": "Minimum Total Capital (CRAR) Requirement",
        "basel": "Minimum Capital to Risk-Weighted Assets Ratio (CRAR) of 8.0%.",
        "rbi": "Minimum CRAR of 9.0% for commercial banks (11.5% including 2.5% CCB).",
        "rationale": "Stricter Pillar 1 baseline to absorb elevated structural credit and concentration risks in the Indian financial system."
    },
    {
        "num": "4",
        "dimension": "Capital Conservation Buffer (CCB) Enforcement",
        "basel": "2.5% CCB phased in from 2016-2019; uniform distribution restriction tiers.",
        "rbi": "2.5% CCB enforced; transition deferred multiple times during COVID-19 shock to support bank lending.",
        "rationale": "Calibrated counter-cyclical forbearance balancing macroeconomic credit flow with systemic solvency."
    },
    {
        "num": "5",
        "dimension": "Countercyclical Capital Buffer (CCCB) Activation",
        "basel": "0% to 2.5% activated mechanistically based on the Credit-to-GDP gap.",
        "rbi": "Maintained at 0% since inception; RBI relies instead on targeted sectoral risk-weight hikes (unsecured consumer credit).",
        "rationale": "Credit-to-GDP gap in developing economies is skewed by structural financial deepening; targeted micro-prudential hikes are more effective."
    },
    {
        "num": "6",
        "dimension": "Sovereign Debt Risk Weights (Credit Risk)",
        "basel": "0% risk weight tied to external credit rating (AAA to AA-); non-zero for lower ratings.",
        "rbi": "Mandatory 0% risk weight for all Central and State Government securities (including State Development Loans - SDLs).",
        "rationale": "Facilitates seamless domestic sovereign borrowing and statutory pre-emption without capital penalty."
    },
    {
        "num": "7",
        "dimension": "Liquidity Coverage: SLR Carve-Out vs HQLA",
        "basel": "High-Quality Liquid Assets (HQLA) strictly unencumbered cash and central bank reserves.",
        "rbi": "Permits banks to carve out up to 15-18% of Statutory Liquidity Ratio (SLR) via FALLCR and MSF as Level 1 HQLA.",
        "rationale": "Harmonizes India's legacy statutory reserve requirement (SLR ~18%) with modern Basel LCR compliance."
    },
    {
        "num": "8",
        "dimension": "Asset Classification & Loan Loss Provisioning",
        "basel": "Expected Credit Loss (ECL) 3-stage forward-looking impairment model (IFRS 9).",
        "rbi": "Incurred Loss Model with formulaic 90-day NPA delinquency provisioning (Sub-standard 15%, Doubtful 25-100%); ECL in phased transition.",
        "rationale": "Rule-based supervisory provisioning prevents subjective bank management under-provisioning during credit cycles."
    },
    {
        "num": "9",
        "dimension": "Interest Rate Risk in Banking Book (IRRBB)",
        "basel": "Pillar 2 supervisory outlier test (+/-200 bps parallel or 6 standardized shock scenarios).",
        "rbi": "Explicit Pillar 2 capital add-on under ICAAP combining Earnings at Risk (EaR) and Economic Value of Equity (EVE) simulation.",
        "rationale": "Indian commercial banks hold heavy fixed-rate government bond portfolios, creating acute duration and yield curve vulnerabilities."
    },
    {
        "num": "10",
        "dimension": "D-SIB Surcharges & Board Governance",
        "basel": "National discretion for D-SIB assessment; general principle-based governance.",
        "rbi": "Designates SBI (0.80%), HDFC Bank (0.40%), ICICI Bank (0.20%) with tiered CET1 surcharges; mandatory Board-approved ICAAP stress testing.",
        "rationale": "Mitigates moral hazard of 'too-big-to-fail' lenders and enforces strict board accountability and proportional compliance."
    }
]

# -------------------------------------------------------------
# 2. GENERATE 2-PAGE PDF DOCUMENT (ReportLab)
# -------------------------------------------------------------
class AcademicCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super(AcademicCanvas, self).__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super(AcademicCanvas, self).showPage()
        super(AcademicCanvas, self).save()

    def draw_page_decorations(self, page_count):
        self.saveState()
        self.setFont("Times-Roman", 8)
        self.setFillColor(colors.HexColor("#444444"))
        
        # Header (page 2)
        if self._pageNumber > 1:
            self.drawString(36, 756, "SRMIST | 21CSE449T: Risk Analytics | Regulatory Gap Analysis: RBI ICAAP vs Basel III")
            self.drawRightString(576, 756, "Academic Research Paper Review")
            self.setStrokeColor(colors.HexColor("#0b2545"))
            self.setLineWidth(0.6)
            self.line(36, 748, 576, 748)

        # Footer (all pages)
        self.setStrokeColor(colors.HexColor("#cccccc"))
        self.setLineWidth(0.6)
        self.line(36, 36, 576, 36)
        self.drawString(36, 26, "Department of Data Science and Business Systems | School of Computing | Academic Year 2026-27")
        page_str = f"Page {self._pageNumber} of {page_count}"
        self.drawRightString(576, 26, page_str)
        self.restoreState()

def build_pdf_paper():
    pdf_path = os.path.join(BASE_DIR, "RBI_ICAAP_vs_Basel_III_Gap_Analysis.pdf")
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=letter,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36
    )

    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'PaperTitle',
        parent=styles['Normal'],
        fontName='Times-Bold',
        fontSize=13,
        leading=16,
        textColor=colors.HexColor('#0b2545'),
        alignment=1, # Center
        spaceAfter=4
    )
    author_style = ParagraphStyle(
        'PaperAuthor',
        parent=styles['Normal'],
        fontName='Times-Italic',
        fontSize=8,
        leading=10.5,
        textColor=colors.HexColor('#333333'),
        alignment=1,
        spaceAfter=6
    )
    abstract_heading = ParagraphStyle(
        'AbsHead',
        parent=styles['Normal'],
        fontName='Times-Bold',
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor('#0b2545')
    )
    abstract_body = ParagraphStyle(
        'AbsBody',
        parent=styles['Normal'],
        fontName='Times-Roman',
        fontSize=7.5,
        leading=9.5,
        textColor=colors.HexColor('#222222'),
        spaceAfter=6
    )
    sec_heading = ParagraphStyle(
        'SecHead',
        parent=styles['Normal'],
        fontName='Times-Bold',
        fontSize=9,
        leading=11.5,
        textColor=colors.HexColor('#0b2545'),
        spaceBefore=4,
        spaceAfter=2,
        keepWithNext=True
    )
    body_text = ParagraphStyle(
        'Body',
        parent=styles['Normal'],
        fontName='Times-Roman',
        fontSize=7.2,
        leading=9.2,
        textColor=colors.HexColor('#222222'),
        spaceAfter=4
    )
    tbl_th = ParagraphStyle(
        'TblH',
        parent=styles['Normal'],
        fontName='Times-Bold',
        fontSize=6.8,
        leading=8.5,
        textColor=colors.white,
        alignment=1
    )
    tbl_td = ParagraphStyle(
        'TblD',
        parent=styles['Normal'],
        fontName='Times-Roman',
        fontSize=6.2,
        leading=7.8,
        textColor=colors.HexColor('#222222')
    )
    tbl_td_bold = ParagraphStyle(
        'TblDBold',
        parent=tbl_td,
        fontName='Times-Bold'
    )

    story = []

    # ---------------- PAGE 1 ----------------
    story.append(Paragraph("A CRITICAL GAP ANALYSIS OF RBI'S ICAAP GUIDELINES VERSUS BASEL III STANDARDS: SUPER-EQUIVALENCE, STRUCTURAL DIVERGENCES, AND MACROPRUDENTIAL IMPLICATIONS", title_style))
    story.append(Paragraph("<b>Course:</b> 21CSE449T – Risk Analytics | <b>Faculty Evaluator:</b> Dr. V. Prasanna<br/>Department of Data Science and Business Systems, School of Computing, SRM Institute of Science and Technology", author_style))
    story.append(HRFlowable(width="100%", thickness=0.8, color=colors.HexColor('#0b2545'), spaceBefore=1, spaceAfter=4))

    # Abstract Box
    abs_text = (
        "<b>Abstract—</b>Following the 2007–2009 Global Financial Crisis, the Basel Committee on Banking Supervision (BCBS) formulated the Basel III regulatory framework to fortify microprudential regulation and mitigate systemic banking crises. While global jurisdictions adopted Basel III as a harmonized benchmark, the <b>Reserve Bank of India (RBI)</b> operationalized its <b>Internal Capital Adequacy Assessment Process (ICAAP)</b> under the Master Circular on Basel III Capital Regulations through a deliberate posture of <i>super-equivalence</i> ('gold-plating'). This research paper critique conducts a systematic comparative gap analysis between RBI's ICAAP architecture and BCBS Basel III global standards. We isolate and dissect <b>ten pivotal structural divergences</b> spanning capital thresholds, sovereign risk-weighting, liquidity interaction with statutory pre-emptions (SLR), credit impairment models, and Interest Rate Risk in the Banking Book (IRRBB). We argue that while RBI's conservative calibration reinforces systemic solvency against emerging market cyclicality, it introduces capital allocation frictions and structural trade-offs between regulatory stringency and credit growth."
    )
    story.append(Paragraph(abs_text, abstract_body))
    story.append(HRFlowable(width="100%", thickness=0.4, color=colors.HexColor('#aaaaaa'), spaceBefore=1, spaceAfter=4))

    # Section 1
    story.append(Paragraph("1. Introduction & The Political Economy of RBI's Super-Equivalence", sec_heading))
    sec1_text = (
        "The Basel III accord represents an international soft-law standard establishing minimum regulatory thresholds for capital adequacy, leverage, "
        "and funding liquidity. In emerging market economies such as India, banking stability is uniquely challenged by shallow corporate bond markets, "
        "elevated state-directed lending priorities, and cyclical asset quality deteriorations in public-sector banks (PSBs). Consequently, the RBI "
        "rejected a verbatim transposition of Basel III, designing an ICAAP framework anchored in regulatory conservatism. Pillar 2 (Supervisory Review "
        "Process) in India is not merely an auxiliary review; it functions as a binding institutional covenant requiring commercial banks to hold "
        "statutory capital buffers well in excess of international minimums."
    )
    story.append(Paragraph(sec1_text, body_text))

    # Section 2 & Table (Part 1: Differences 1 to 5)
    story.append(Paragraph("2. Comparative Structural Matrix: Ten Foundational Gaps", sec_heading))
    story.append(Paragraph("The comparative matrix below delineates the precise technical divergences between the BCBS Basel III standards and RBI's ICAAP guidelines:", body_text))
    
    t_data_1 = [
        [
            Paragraph("<b>#</b>", tbl_th),
            Paragraph("<b>Regulatory Dimension</b>", tbl_th),
            Paragraph("<b>Basel III International Standard (BCBS)</b>", tbl_th),
            Paragraph("<b>RBI ICAAP Guidelines (Indian Banking)</b>", tbl_th),
            Paragraph("<b>Macroprudential & Economic Critique</b>", tbl_th)
        ]
    ]

    for d in diff_data[:5]:
        t_data_1.append([
            Paragraph(f"<b>{d['num']}</b>", tbl_td_bold),
            Paragraph(f"<b>{d['dimension']}</b>", tbl_td_bold),
            Paragraph(d['basel'], tbl_td),
            Paragraph(d['rbi'], tbl_td),
            Paragraph(d['rationale'], tbl_td)
        ])

    tbl_1 = Table(t_data_1, colWidths=[14, 100, 135, 145, 146])
    tbl_1.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0b2545')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cccccc')),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8f9fa')])
    ]))
    story.append(tbl_1)
    story.append(PageBreak())

    # ---------------- PAGE 2 ----------------
    story.append(Paragraph("2. Comparative Structural Matrix (Contd: Differences 6 to 10)", sec_heading))
    
    t_data_2 = [
        [
            Paragraph("<b>#</b>", tbl_th),
            Paragraph("<b>Regulatory Dimension</b>", tbl_th),
            Paragraph("<b>Basel III International Standard (BCBS)</b>", tbl_th),
            Paragraph("<b>RBI ICAAP Guidelines (Indian Banking)</b>", tbl_th),
            Paragraph("<b>Macroprudential & Economic Critique</b>", tbl_th)
        ]
    ]

    for d in diff_data[5:]:
        t_data_2.append([
            Paragraph(f"<b>{d['num']}</b>", tbl_td_bold),
            Paragraph(f"<b>{d['dimension']}</b>", tbl_td_bold),
            Paragraph(d['basel'], tbl_td),
            Paragraph(d['rbi'], tbl_td),
            Paragraph(d['rationale'], tbl_td)
        ])

    tbl_2 = Table(t_data_2, colWidths=[14, 100, 135, 145, 146])
    tbl_2.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0b2545')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cccccc')),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8f9fa')])
    ]))
    story.append(tbl_2)
    story.append(Spacer(1, 4))

    # Section 3: In-Depth Critique
    story.append(Paragraph("3. Deep-Dive Regulatory Critique & Systemic Implications", sec_heading))
    critique_text = (
        "<b>A. Super-Equivalence as Counter-Cyclical Insulation:</b> The RBI's enforcement of a 9.0% minimum CRAR and 5.5% CET1 (versus Basel's 8.0% and 4.5%) "
        "reflects structural wisdom. During the twin-balance-sheet crisis (2015–2020), Indian banks absorbed massive NPA provisioning without breaching insolvency "
        "thresholds, validating the 100 bps equity buffer as an indispensable systemic shock absorber.<br/>"
        "<b>B. The Sovereign Nexus & Statutory Liquidity Distortions:</b> While Basel III penalizes sub-investment sovereign exposures, the RBI assigns a 0% risk weight "
        "to all Central and State government debt, while simultaneously mandating an 18% Statutory Liquidity Ratio (SLR). By allowing the Facility to Avail "
        "Liquidity for Liquidity Coverage Ratio (FALLCR), the RBI synthesized a hybrid liquidity model that satisfies Basel LCR requirements while preserving captive demand for sovereign debt.<br/>"
        "<b>C. The ECL Transition Impasse:</b> The global banking sector operates under IFRS 9 forward-looking Expected Credit Loss models. In contrast, Indian banks "
        "remain predominantly bound to the incurred loss framework with formulaic provisioning rules. While the RBI's prescriptive approach prevents subjective "
        "under-provisioning, it creates earnings volatility and regulatory divergence for internationally active Indian banks."
    )
    story.append(Paragraph(critique_text, body_text))

    # Section 4: Policy Recommendations & Conclusion
    story.append(Paragraph("4. Policy Roadmap & Concluding Assessment", sec_heading))
    concl_text = (
        "<b>Conclusion:</b> The comparative analysis confirms that the RBI's ICAAP architecture is fundamentally more conservative and macroprudentially intrusive "
        "than the BCBS baseline. Moving toward Basel IV, the RBI should: (1) Expedite the phased rollout of the Expected Credit Loss (ECL) framework; "
        "(2) Calibrate differentiated risk weights for State Development Loans (SDLs) to foster sub-national fiscal discipline; and (3) Refine the Countercyclical Capital "
        "Buffer (CCCB) framework by transitioning from credit-to-GDP gap metrics toward high-frequency systemic leverage indices."
    )
    story.append(Paragraph(concl_text, body_text))

    # References
    story.append(Paragraph("5. References & Regulatory Authorities", sec_heading))
    refs_text = (
        "[1] Basel Committee on Banking Supervision (BCBS), <i>Basel III: A global regulatory framework for more resilient banks and banking systems</i>, BIS, 2011.<br/>"
        "[2] Reserve Bank of India (RBI), <i>Master Circular - Basel III Capital Regulations</i>, RBI/2023-24/108, updated 2024.<br/>"
        "[3] Reserve Bank of India (RBI), <i>Guidelines on Internal Capital Adequacy Assessment Process (ICAAP) and Pillar 2 Framework</i>, 2018.<br/>"
        "[4] Acharya, V. V., et al., <i>Quest for Restoring Financial Stability in India</i>, SAGE Publications, 2020."
    )
    story.append(Paragraph(refs_text, ParagraphStyle('Refs', parent=body_text, fontSize=5.8, leading=7.2, textColor=colors.HexColor('#555555'))))

    doc.build(story, canvasmaker=AcademicCanvas)
    print(f"2-Page Regulatory Gap Analysis PDF generated successfully at: {pdf_path}")

# -------------------------------------------------------------
# 3. GENERATE WORD (.DOCX) & MARKDOWN (.MD)
# -------------------------------------------------------------
def build_docx_and_md():
    docx_path = os.path.join(BASE_DIR, "RBI_ICAAP_vs_Basel_III_Gap_Analysis.docx")
    md_path = os.path.join(BASE_DIR, "RBI_ICAAP_vs_Basel_III_Gap_Analysis.md")

    # 1. MARKDOWN
    md_content = """# SRM INSTITUTE OF SCIENCE AND TECHNOLOGY
### School of Computing | Department of Data Science and Business Systems
**Subject Code:** 21CSE449T | **Subject Name:** RISK ANALYTICS  
**Academic Year:** 2026–27 (ODD) | **Faculty Name:** Dr. V. Prasanna  
**Programme & Branch:** B.Tech – CSE (BDA & DS)

---

# Unit 2 - Task 2: Regulatory Gap Analysis (Research Paper Review)
## A Critical Gap Analysis of RBI's ICAAP Guidelines versus Basel III Standards: Super-Equivalence, Structural Divergences, and Macroprudential Implications
**Deliverable:** 2-Page Research Paper Review

---

## Abstract
Following the 2007–2009 Global Financial Crisis, the Basel Committee on Banking Supervision (BCBS) formulated the Basel III regulatory framework to fortify microprudential regulation and mitigate systemic banking crises. While global jurisdictions adopted Basel III as a harmonized benchmark, the **Reserve Bank of India (RBI)** operationalized its **Internal Capital Adequacy Assessment Process (ICAAP)** under the Master Circular on Basel III Capital Regulations through a deliberate posture of *super-equivalence* ("gold-plating"). This research paper critique conducts a systematic comparative gap analysis between RBI's ICAAP architecture and BCBS Basel III global standards. We isolate and dissect **ten pivotal structural divergences** spanning capital thresholds, sovereign risk-weighting, liquidity interaction with statutory pre-emptions (SLR), credit impairment models, and Interest Rate Risk in the Banking Book (IRRBB). We argue that while RBI's conservative calibration reinforces systemic solvency against emerging market cyclicality, it introduces capital allocation frictions and structural trade-offs between regulatory stringency and credit growth.

---

## 1. Introduction & Regulatory Philosophy
The Basel III accord represents an international soft-law standard establishing minimum regulatory thresholds for capital adequacy, leverage, and funding liquidity. In emerging market economies such as India, banking stability is uniquely challenged by shallow corporate bond markets, elevated state-directed lending priorities, and cyclical asset quality deteriorations in public-sector banks (PSBs). Consequently, the RBI rejected a verbatim transposition of Basel III, designing an ICAAP framework anchored in regulatory conservatism. Pillar 2 (Supervisory Review Process) in India is not merely an auxiliary review; it functions as a binding institutional covenant requiring commercial banks to hold statutory capital buffers well in excess of international minimums.

---

## 2. The 10 Key Differences: Comparative Synthesis Matrix

| # | Regulatory Dimension | Basel III International Standard (BCBS) | RBI ICAAP Guidelines (Indian Banking) | Macroprudential & Economic Critique |
|---|---|---|---|---|
| **1** | **Common Equity Tier 1 (CET1) Minimum** | Minimum CET1 ratio of **4.5%** of Risk-Weighted Assets (RWAs). | Minimum CET1 ratio of **5.5%** of RWAs (100 bps super-equivalent gold-plating). | Buffers against higher emerging-market asset volatility and historical public-sector bank capital erosion. |
| **2** | **Minimum Tier 1 Capital Ratio** | Minimum Tier 1 capital ratio of **6.0%** of RWAs. | Minimum Tier 1 capital ratio of **7.0%** of RWAs (100 bps higher than BCBS standard). | Ensures that loss-absorption capacity remains anchored in permanent high-quality going-concern equity. |
| **3** | **Minimum Total Capital (CRAR)** | Minimum CRAR of **8.0%** of RWAs. | Minimum CRAR of **9.0%** for commercial banks (11.5% including 2.5% CCB). | Stricter Pillar 1 baseline to absorb elevated structural credit and concentration risks in the Indian financial system. |
| **4** | **Capital Conservation Buffer (CCB)** | 2.5% CCB phased in from 2016-2019; uniform distribution restriction tiers. | 2.5% CCB enforced; transition deferred multiple times during COVID-19 shock to support bank lending. | Calibrated counter-cyclical forbearance balancing macroeconomic credit flow with systemic solvency. |
| **5** | **Countercyclical Capital Buffer (CCCB)** | 0% to 2.5% activated mechanistically based on the Credit-to-GDP gap. | Maintained at 0% since inception; RBI relies instead on targeted sectoral risk-weight hikes (unsecured consumer credit). | Credit-to-GDP gap in developing economies is skewed by structural financial deepening; targeted micro-prudential hikes are more effective. |
| **6** | **Sovereign Debt Risk Weights** | 0% risk weight tied to external credit rating (AAA to AA-); non-zero for lower ratings. | Mandatory **0% risk weight** for all Central and State Government securities (including State Development Loans - SDLs). | Facilitates seamless domestic sovereign borrowing and statutory pre-emption without capital penalty. |
| **7** | **Liquidity Coverage: SLR vs HQLA** | High-Quality Liquid Assets (HQLA) strictly unencumbered cash and central bank reserves. | Permits banks to carve out up to 15-18% of Statutory Liquidity Ratio (SLR) via **FALLCR** and **MSF** as Level 1 HQLA. | Harmonizes India's legacy statutory reserve requirement (SLR ~18%) with modern Basel LCR compliance. |
| **8** | **Asset Classification & Provisioning** | Expected Credit Loss (ECL) 3-stage forward-looking impairment model (IFRS 9). | Incurred Loss Model with formulaic 90-day NPA delinquency provisioning (Sub-standard 15%, Doubtful 25-100%); ECL in phased transition. | Rule-based supervisory provisioning prevents subjective bank management under-provisioning during credit cycles. |
| **9** | **Interest Rate Risk in Banking Book (IRRBB)** | Pillar 2 supervisory outlier test (+/-200 bps parallel or 6 standardized shock scenarios). | Explicit Pillar 2 capital add-on under ICAAP combining Earnings at Risk (EaR) and Economic Value of Equity (EVE) simulation. | Indian commercial banks hold heavy fixed-rate government bond portfolios, creating acute duration and yield curve vulnerabilities. |
| **10** | **D-SIB Surcharges & Board Governance** | National discretion for D-SIB assessment; general principle-based governance. | Designates **SBI (0.80%)**, **HDFC Bank (0.40%)**, **ICICI Bank (0.20%)** with tiered CET1 surcharges; mandatory Board-approved ICAAP stress testing. | Mitigates moral hazard of 'too-big-to-fail' lenders and enforces strict board accountability and proportional compliance. |

---

## 3. Deep-Dive Regulatory Critique & Systemic Implications

### 3.1 Super-Equivalence as Counter-Cyclical Insulation
The RBI's enforcement of a 9.0% minimum CRAR and 5.5% CET1 (versus Basel's 8.0% and 4.5%) reflects structural wisdom. During the twin-balance-sheet crisis (2015–2020), Indian banks absorbed massive NPA provisioning without breaching insolvency thresholds, validating the 100 bps equity buffer as an indispensable systemic shock absorber.

### 3.2 The Sovereign Nexus & Statutory Liquidity Distortions
While Basel III penalizes sub-investment sovereign exposures, the RBI assigns a 0% risk weight to all Central and State government debt, while simultaneously mandating an 18% Statutory Liquidity Ratio (SLR). By allowing the Facility to Avail Liquidity for Liquidity Coverage Ratio (FALLCR), the RBI synthesized a hybrid liquidity model that satisfies Basel LCR requirements while preserving captive demand for sovereign debt.

### 3.3 The ECL Transition Impasse
The global banking sector operates under IFRS 9 forward-looking Expected Credit Loss models. In contrast, Indian banks remain predominantly bound to the incurred loss framework with formulaic provisioning rules. While the RBI's prescriptive approach prevents subjective under-provisioning, it creates earnings volatility and regulatory divergence for internationally active Indian banks.

---

## 4. Policy Roadmap & Concluding Assessment
The comparative analysis confirms that the RBI's ICAAP architecture is fundamentally more conservative and macroprudentially intrusive than the BCBS baseline. Moving toward Basel IV, the RBI should:
1. Expedite the phased rollout of the Expected Credit Loss (ECL) framework.
2. Calibrate differentiated risk weights for State Development Loans (SDLs) to foster sub-national fiscal discipline.
3. Refine the Countercyclical Capital Buffer (CCCB) framework by transitioning from credit-to-GDP gap metrics toward high-frequency systemic leverage indices.

---

## 5. References
1. Basel Committee on Banking Supervision (BCBS), *Basel III: A global regulatory framework for more resilient banks and banking systems*, BIS, 2011.
2. Reserve Bank of India (RBI), *Master Circular - Basel III Capital Regulations*, RBI/2023-24/108, updated 2024.
3. Reserve Bank of India (RBI), *Guidelines on Internal Capital Adequacy Assessment Process (ICAAP) and Pillar 2 Framework*, 2018.
4. Acharya, V. V., et al., *Quest for Restoring Financial Stability in India*, SAGE Publications, 2020.
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
    r_t = p_t.add_run("UNIT 2 - TASK 2: REGULATORY GAP ANALYSIS (RESEARCH PAPER REVIEW)\nA Critical Analysis of RBI's ICAAP Guidelines vs. Basel III Standards\nSuper-Equivalence, Structural Divergences, and Macroprudential Implications")
    r_t.font.name = "Calibri"
    r_t.font.size = Pt(14)
    r_t.font.bold = True
    r_t.font.color.rgb = RGBColor(11, 37, 69)

    doc.add_paragraph().add_run("—" * 60).font.color.rgb = RGBColor(11, 37, 69)

    doc.add_heading("Abstract", level=1).runs[0].font.color.rgb = RGBColor(11, 37, 69)
    doc.add_paragraph(
        "This research paper critique evaluates the Reserve Bank of India's (RBI) Internal Capital Adequacy Assessment Process (ICAAP) "
        "guidelines against the Basel Committee on Banking Supervision (BCBS) Basel III global standard. We isolate and dissect 10 critical "
        "structural divergences, including capital ratios, sovereign risk-weighting, SLR liquidity carve-outs, credit provisioning, and IRRBB."
    )

    doc.add_heading("Ten Key Structural Differences Matrix", level=1).runs[0].font.color.rgb = RGBColor(11, 37, 69)
    tbl = doc.add_table(rows=1, cols=5)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdrs = ["#", "Regulatory Dimension", "Basel III Global Standard", "RBI ICAAP Guidelines", "Macroprudential Critique"]
    for i, h in enumerate(hdrs):
        tbl.rows[0].cells[i].text = h
        shd = parse_xml(r'<w:shd {} w:fill="0B2545"/>'.format(nsdecls('w')))
        tbl.rows[0].cells[i]._tc.get_or_add_tcPr().append(shd)
        p = tbl.rows[0].cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)

    for d in diff_data:
        row = tbl.add_row().cells
        row[0].text = d['num']
        row[1].text = d['dimension']
        row[2].text = d['basel']
        row[3].text = d['rbi']
        row[4].text = d['rationale']
        for idx in range(5):
            for run in row[idx].paragraphs[0].runs:
                run.font.size = Pt(8)

    doc.add_heading("Critical Analysis & Policy Recommendations", level=1).runs[0].font.color.rgb = RGBColor(11, 37, 69)
    doc.add_paragraph(
        "1. Super-Equivalence Benefits: The +100 bps equity buffer provided critical absorption during India's twin-balance-sheet crisis.\n"
        "2. Sovereign Risk-Weighting: The 0% risk-weight on state debt (SDLs) creates implicit subsidies for state deficits.\n"
        "3. Future Roadmap: Accelerated transition to IFRS 9 ECL provisioning and high-frequency cyclical indicators."
    )

    doc.save(docx_path)
    print("Word (.docx) and Markdown (.md) generated successfully!")

if __name__ == "__main__":
    build_pdf_paper()
    build_docx_and_md()
    print("Unit 2 Task 2 complete!")
