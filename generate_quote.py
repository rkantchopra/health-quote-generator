#!/usr/bin/env python3
"""
generate_quote_pdf.py — increMint Health Insurance Quote (PDF)
Mirrors generate_quote.py exactly: Cover page (portrait A4) + Comparison (landscape A4)
Uses ReportLab platypus for reliable, device-independent rendering.
"""

import sys, os, json, glob
from datetime import datetime
from typing import Dict, List

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import cm, mm, inch
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable, Image, KeepTogether
)
from reportlab.platypus.flowables import Flowable
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT

# ── COLORS ────────────────────────────────────────────────────────────────────
GREEN       = colors.HexColor("#00C853")
GREEN_DARK  = colors.HexColor("#00A846")
GREEN_LIGHT = colors.HexColor("#E8FFF3")
GREEN_PALE  = colors.HexColor("#F5FFFB")
YELLOW_HL   = colors.HexColor("#FFFDE7")
YELLOW_BG   = colors.HexColor("#FFF9C4")
WHITE       = colors.white
OFFWHITE    = colors.HexColor("#FAFAFA")
LGRAY       = colors.HexColor("#F5F5F5")
BORDER_C    = colors.HexColor("#E0E0E0")
DARK        = colors.HexColor("#1A1A1A")
GRAY        = colors.HexColor("#666666")
RED         = colors.HexColor("#D32F2F")
BLACK       = colors.black

# ── MASTER DATA ───────────────────────────────────────────────────────────────
MASTER: Dict[str, Dict[str, str]] = {
    "ICICI Lombard - Elevate": {
        "Room Rent":                  "Single Private A/C Room",
        "ICU Limit":                  "No limit on ICU",
        "Pre-Hospitalization":        "90 days",
        "Post-Hospitalization":       "180 days",
        "Day Care Procedures":        "All procedures up to S.I.",
        "Copay":                      "0%",
        "Restoration Benefit":        "Unlimited (incl. same illness)",
        "NCB Benefit":                "Unlimited 100%/yr — No Cap",
        "Non-Consumables":            "Covered",
        "Hospitalization @ Home":     "Up to Sum Assured",
        "Ambulance Cover":            "Up to Sum Assured",
        "Air Ambulance":              "Up to Sum Assured",
        "AYUSH Treatment":            "Up to Sum Assured",
        "Organ Donor":                "Up to Sum Assured",
        "Modern Treatments":          "Up to Sum Assured",
        "2-Hour Hospitalization":     "Covered",
        "E-Consultation":             "Unlimited",
        "Preventive Health Check-up": "Covered (Optional)",
        "Maternity Cover":            "Optional Rider (10% SA, Max Rs.1L)",
        "Newborn Cover":              "Day 1 (10% SA)",
        "Worldwide Cover":            "Not Available",
        "OPD Cover":                  "Optional Rider",
        "Lock-the-Clock":             "Not Applicable",
        "Priority Claim Desk":        "Not Available",
        "Cash+ Wallet":               "—",
        "Unique Features":            "Unlimited NCB | 2-hr hospitalization | Child till 30 yrs | Newborn Day-1"
    },
    "Niva Bupa - ReAssure 3.0": {
        "Room Rent":                  "Any Room incl. Suite",
        "ICU Limit":                  "No limit on ICU",
        "Pre-Hospitalization":        "60 days",
        "Post-Hospitalization":       "180 days",
        "Day Care Procedures":        "All procedures up to S.I.",
        "Copay":                      "0%",
        "Restoration Benefit":        "Unlimited — Same illness, never ends",
        "NCB Benefit":                "N/A — Unlimited Cover Plan",
        "Non-Consumables":            "Covered",
        "Hospitalization @ Home":     "Up to Sum Insured",
        "Ambulance Cover":            "Up to Sum Insured",
        "Air Ambulance":              "Up to Rs. 5 Lakh",
        "AYUSH Treatment":            "Up to Sum Assured",
        "Organ Donor":                "Up to Sum Assured",
        "Modern Treatments":          "Up to Sum Assured",
        "2-Hour Hospitalization":     "Covered",
        "E-Consultation":             "Unlimited",
        "Preventive Health Check-up": "Annual check-up covered",
        "Maternity Cover":            "Not Available",
        "Newborn Cover":              "Not Available",
        "Worldwide Cover":            "Available (with rider)",
        "OPD Cover":                  "Optional (Rs.1L annual)",
        "Lock-the-Clock":             "Premium stays locked",
        "Priority Claim Desk":        "Optional (HNI/Prime)",
        "Cash+ Wallet":               "Cashback every claim-free year",
        "Unique Features":            "Unlimited Sum Insured | Worldwide Cover | Lock-the-Clock | Cash+ Wallet"
    },
    "Niva Bupa - Aspire Platinum": {
        "Room Rent":                  "Single Private A/C Room",
        "ICU Limit":                  "No limit on ICU",
        "Pre-Hospitalization":        "60 days",
        "Post-Hospitalization":       "90 days",
        "Day Care Procedures":        "All procedures up to S.I.",
        "Copay":                      "0%",
        "Restoration Benefit":        "Unlimited (incl. same illness)",
        "NCB Benefit":                "Up to 5X (300%)",
        "Non-Consumables":            "Covered",
        "Hospitalization @ Home":     "Up to Sum Assured",
        "Ambulance Cover":            "Up to Sum Assured",
        "Air Ambulance":              "Up to Sum Assured",
        "AYUSH Treatment":            "Up to Sum Assured",
        "Organ Donor":                "Up to Sum Assured",
        "Modern Treatments":          "Up to Sum Assured",
        "2-Hour Hospitalization":     "—",
        "E-Consultation":             "Unlimited",
        "Preventive Health Check-up": "Standard",
        "Maternity Cover":            "Standard Rs.12k/yr",
        "Newborn Cover":              "Available",
        "Worldwide Cover":            "Not Available",
        "OPD Cover":                  "Optional Rider",
        "Lock-the-Clock":             "Premium stays locked",
        "Priority Claim Desk":        "—",
        "Cash+ Wallet":               "—",
        "Unique Features":            "Premium lock | Child cover till 60 yrs | NCB up to 5X"
    },
    "Tata AIG - Medicare Select": {
        "Room Rent":                  "Single Private A/C Room",
        "ICU Limit":                  "No limit on ICU",
        "Pre-Hospitalization":        "60 days",
        "Post-Hospitalization":       "90 days",
        "Day Care Procedures":        "All procedures up to S.I.",
        "Copay":                      "0%",
        "Restoration Benefit":        "Unlimited (incl. same illness)",
        "NCB Benefit":                "100%/yr up to 500% (Super NCB)",
        "Non-Consumables":            "Covered",
        "Hospitalization @ Home":     "Up to Sum Assured",
        "Ambulance Cover":            "Up to Sum Assured",
        "Air Ambulance":              "Up to Sum Assured",
        "AYUSH Treatment":            "Up to Sum Assured",
        "Organ Donor":                "Up to Sum Assured",
        "Modern Treatments":          "Up to Sum Assured",
        "2-Hour Hospitalization":     "—",
        "E-Consultation":             "Unlimited",
        "Preventive Health Check-up": "All covered",
        "Maternity Cover":            "Optional Rider (10% SA, Max Rs.1L)",
        "Newborn Cover":              "Available with rider",
        "Worldwide Cover":            "Not Available",
        "OPD Cover":                  "Optional Rider",
        "Lock-the-Clock":             "Not Applicable",
        "Priority Claim Desk":        "—",
        "Cash+ Wallet":               "—",
        "Unique Features":            "Salary discount 7.5% | Super NCB up to 500%"
    },
    "HDFC ERGO - Optima Secure": {
        "Room Rent":                  "Single Private A/C Room",
        "ICU Limit":                  "No limit on ICU",
        "Pre-Hospitalization":        "60 days",
        "Post-Hospitalization":       "180 days",
        "Day Care Procedures":        "All procedures up to S.I.",
        "Copay":                      "0%",
        "Restoration Benefit":        "Unlimited (incl. same illness)",
        "NCB Benefit":                "2X from Day 1",
        "Non-Consumables":            "Covered",
        "Hospitalization @ Home":     "Up to Sum Assured",
        "Ambulance Cover":            "Up to Sum Assured",
        "Air Ambulance":              "Up to Sum Assured",
        "AYUSH Treatment":            "Up to Sum Assured",
        "Organ Donor":                "Up to Sum Assured",
        "Modern Treatments":          "Up to Sum Assured",
        "2-Hour Hospitalization":     "—",
        "E-Consultation":             "Unlimited",
        "Preventive Health Check-up": "Standard",
        "Maternity Cover":            "Not Available",
        "Newborn Cover":              "Not Available",
        "Worldwide Cover":            "Not Available",
        "OPD Cover":                  "Optional Rider",
        "Lock-the-Clock":             "Not Applicable",
        "Priority Claim Desk":        "—",
        "Cash+ Wallet":               "—",
        "Unique Features":            "2X Cover from Day 1 | Health check-up included"
    },
    "Care Health - Supreme": {
        "Room Rent":                  "Single Private A/C Room",
        "ICU Limit":                  "No limit on ICU",
        "Pre-Hospitalization":        "60 days",
        "Post-Hospitalization":       "180 days",
        "Day Care Procedures":        "All procedures up to S.I.",
        "Copay":                      "0%",
        "Restoration Benefit":        "Unlimited (incl. same illness)",
        "NCB Benefit":                "100%/yr up to 600% (6X)",
        "Non-Consumables":            "Covered",
        "Hospitalization @ Home":     "Up to Sum Assured",
        "Ambulance Cover":            "Up to Sum Assured",
        "Air Ambulance":              "Up to Sum Assured",
        "AYUSH Treatment":            "Up to Sum Assured",
        "Organ Donor":                "Up to Sum Assured",
        "Modern Treatments":          "Up to Sum Assured",
        "2-Hour Hospitalization":     "—",
        "E-Consultation":             "Unlimited",
        "Preventive Health Check-up": "All covered",
        "Maternity Cover":            "Not Available",
        "Newborn Cover":              "Not Available",
        "Worldwide Cover":            "Not Available",
        "OPD Cover":                  "Optional Rider",
        "Lock-the-Clock":             "Not Applicable",
        "Priority Claim Desk":        "—",
        "Cash+ Wallet":               "—",
        "Unique Features":            "NCB up to 600% (6X) | All non-consumables covered"
    }
}

FEATURE_GROUPS = [
    ("MAIN BENEFITS", [
        ("Room Rent",            "Maximum room type allowed during hospitalization"),
        ("ICU Limit",            "Coverage limit for Intensive Care Unit charges"),
        ("Pre-Hospitalization",  "Medical costs covered before hospital admission"),
        ("Post-Hospitalization", "Medical costs covered after discharge"),
        ("Day Care Procedures",  "Treatments needing less than 24-hr hospitalization"),
        ("Copay",                "Your share of the bill — lower is better"),
    ]),
    ("COVERAGE BENEFITS", [
        ("Restoration Benefit",    "Coverage restored after a claim is made"),
        ("NCB Benefit",            "Reward for claim-free years — increases your cover"),
        ("Non-Consumables",        "Covers items like gloves, syringes, PPE kits"),
        ("Hospitalization @ Home", "Treatment at home if hospitalization is not possible"),
        ("Ambulance Cover",        "Road ambulance charges during emergency"),
        ("Air Ambulance",          "Air transport charges during critical emergency"),
        ("AYUSH Treatment",        "Ayurveda, Yoga, Unani, Siddha & Homeopathy"),
        ("Organ Donor",            "Covers organ donor's medical expenses"),
        ("Modern Treatments",      "Robotic surgery, stem cell therapy, etc."),
        ("2-Hour Hospitalization", "Short-stay admission of 2+ hours qualifies"),
    ]),
    ("WELLNESS & ADD-ONS", [
        ("E-Consultation",             "Online doctor consultations covered"),
        ("Preventive Health Check-up", "Annual health check-up covered"),
        ("Maternity Cover",            "Delivery and related expenses"),
        ("Newborn Cover",              "Coverage for newborn from Day 1"),
        ("Worldwide Cover",            "Coverage outside India for emergencies"),
        ("OPD Cover",                  "Out-patient doctor visits, tests & medicines"),
        ("Lock-the-Clock",             "Premium does not increase until a claim is made"),
        ("Priority Claim Desk",        "Dedicated claims support for faster settlement"),
        ("Cash+ Wallet",               "Cashback rewards for claim-free policy years"),
    ]),
    ("STANDOUT FEATURES", [
        ("Unique Features", "Exclusive benefits that set this plan apart"),
    ]),
]

LOGO_MAP = {
    "ICICI Lombard - Elevate":     "icici_lombard",
    "Niva Bupa - ReAssure 3.0":    "niva_reassure3",
    "Niva Bupa - Aspire Platinum": "niva_aspire",
    "Tata AIG - Medicare Select":  "tata_aig",
    "HDFC ERGO - Optima Secure":   "hdfc_ergo",
    "Care Health - Supreme":       "care_health",
}

HIGHLIGHTS = {
    "ICICI Lombard - Elevate":     ["Unlimited NCB (no cap)", "Newborn Day-1 cover (10% SA)", "2-hr hospitalization covered"],
    "Niva Bupa - ReAssure 3.0":    ["Truly Unlimited Sum Insured", "Worldwide cover with rider", "Lock-the-Clock premium freeze"],
    "Niva Bupa - Aspire Platinum": ["Premium lock feature", "Child cover till 60 yrs", "NCB up to 5X (300%)"],
    "Tata AIG - Medicare Select":  ["Salary-linked discount 7.5%", "Super NCB up to 500%", "Optional maternity/newborn rider"],
    "HDFC ERGO - Optima Secure":   ["2X cover from Day 1", "Deductible on 1st claim", "Health check-up included"],
    "Care Health - Supreme":       ["NCB up to 600% (6X)", "All non-consumables covered", "Health check-up rider"],
}

# ── HELPERS ───────────────────────────────────────────────────────────────────
def map_master(name):
    if not name: return None
    n = name.lower().strip()
    if "reassure" in n or "3.0" in n:   return "Niva Bupa - ReAssure 3.0"
    if "aspire" in n:                    return "Niva Bupa - Aspire Platinum"
    if "icici" in n or "elevate" in n:   return "ICICI Lombard - Elevate"
    if "tata" in n or "medicare" in n:   return "Tata AIG - Medicare Select"
    if "hdfc" in n or "optima" in n:     return "HDFC ERGO - Optima Secure"
    if "care" in n or "supreme" in n:    return "Care Health - Supreme"
    for k in MASTER:
        if k.lower() == n: return k
    return None

def find_logo(mk, logo_folder):
    if not mk or not logo_folder: return None
    base = LOGO_MAP.get(mk, "")
    if not base: base = "".join(c if c.isalnum() else "_" for c in mk).lower()
    for ext in (".png", ".jpg", ".jpeg"):
        p = os.path.join(logo_folder, base + ext)
        if os.path.exists(p): return p
    return None

def find_incremint_logo(logo_folder):
    if not logo_folder or not os.path.isdir(logo_folder): return None
    for pat in ("*.png", "*.jpg", "*.jpeg"):
        for p in glob.glob(os.path.join(logo_folder, pat)):
            if "incremint" in os.path.basename(p).lower(): return p
    return None

def fmt_prem(v):
    if not v or str(v).strip() in ("", "0"): return "—"
    try: return f"Rs. {int(str(v).replace(',','').replace(' ','')):,}"
    except: return str(v)

def val_color(val):
    if not val or val in ("—", "-", ""): return GRAY
    lv = val.lower()
    if "not available" in lv or "not applicable" in lv: return RED
    return DARK

def sp(pts): return Spacer(1, pts * mm * 0.4)

# ── STYLES ────────────────────────────────────────────────────────────────────
def make_styles():
    s = {}
    base = dict(fontName="Helvetica", fontSize=10, leading=14, textColor=DARK)

    s['normal']      = ParagraphStyle('normal',      **base)
    s['bold']        = ParagraphStyle('bold',        fontName="Helvetica-Bold", fontSize=10, leading=14, textColor=DARK)
    s['center']      = ParagraphStyle('center',      alignment=TA_CENTER, **base)
    s['right']       = ParagraphStyle('right',       alignment=TA_RIGHT,  **base)
    s['small']       = ParagraphStyle('small',       fontName="Helvetica", fontSize=8.5, leading=12, textColor=GRAY)
    s['small_c']     = ParagraphStyle('small_c',     fontName="Helvetica", fontSize=8.5, leading=12, textColor=GRAY, alignment=TA_CENTER)
    s['gray']        = ParagraphStyle('gray',        fontName="Helvetica", fontSize=9,   leading=13, textColor=GRAY)
    s['gray_c']      = ParagraphStyle('gray_c',      fontName="Helvetica", fontSize=9,   leading=13, textColor=GRAY, alignment=TA_CENTER)
    s['red_c']       = ParagraphStyle('red_c',       fontName="Helvetica", fontSize=9,   leading=13, textColor=RED,  alignment=TA_CENTER)
    s['dark_c']      = ParagraphStyle('dark_c',      fontName="Helvetica", fontSize=9,   leading=13, textColor=DARK, alignment=TA_CENTER)
    s['green_bold']  = ParagraphStyle('green_bold',  fontName="Helvetica-Bold", fontSize=10, leading=14, textColor=GREEN_DARK)
    s['green_c']     = ParagraphStyle('green_c',     fontName="Helvetica-Bold", fontSize=10, leading=14, textColor=GREEN_DARK, alignment=TA_CENTER)
    s['white_bold']  = ParagraphStyle('white_bold',  fontName="Helvetica-Bold", fontSize=10, leading=14, textColor=WHITE)
    s['white_bold_c']= ParagraphStyle('white_bold_c',fontName="Helvetica-Bold", fontSize=10, leading=14, textColor=WHITE, alignment=TA_CENTER)
    s['white_sm']    = ParagraphStyle('white_sm',    fontName="Helvetica", fontSize=8.5, leading=12, textColor=colors.HexColor("#CCFFE8"))
    s['title']       = ParagraphStyle('title',       fontName="Helvetica-Bold", fontSize=16, leading=22, textColor=DARK, alignment=TA_CENTER)
    s['subtitle']    = ParagraphStyle('subtitle',    fontName="Helvetica", fontSize=10,   leading=14, textColor=GRAY, alignment=TA_CENTER)
    s['prem_big']    = ParagraphStyle('prem_big',    fontName="Helvetica-Bold", fontSize=15, leading=20, textColor=GREEN_DARK, alignment=TA_CENTER)
    s['prem_yr']     = ParagraphStyle('prem_yr',     fontName="Helvetica", fontSize=9, leading=13, textColor=GRAY, alignment=TA_CENTER)
    s['sa_big']      = ParagraphStyle('sa_big',      fontName="Helvetica-Bold", fontSize=13, leading=18, textColor=DARK, alignment=TA_CENTER)
    s['feat_name']   = ParagraphStyle('feat_name',   fontName="Helvetica-Bold", fontSize=10, leading=14, textColor=DARK)
    s['feat_desc']   = ParagraphStyle('feat_desc',   fontName="Helvetica",      fontSize=8,  leading=11, textColor=GRAY)
    s['uniq_name']   = ParagraphStyle('uniq_name',   fontName="Helvetica-Bold", fontSize=10, leading=14, textColor=colors.HexColor("#5D4037"))
    s['section_bar'] = ParagraphStyle('section_bar', fontName="Helvetica-Bold", fontSize=11, leading=16, textColor=DARK)
    s['footer']      = ParagraphStyle('footer',      fontName="Helvetica", fontSize=8.5, leading=12, textColor=GRAY)
    s['footer_r']    = ParagraphStyle('footer_r',    fontName="Helvetica", fontSize=8, leading=12, textColor=GRAY, alignment=TA_RIGHT)
    s['advisor_lbl'] = ParagraphStyle('advisor_lbl', fontName="Helvetica-Bold", fontSize=11, leading=15, textColor=GREEN_DARK)
    s['advisor_txt'] = ParagraphStyle('advisor_txt', fontName="Helvetica-Oblique", fontSize=10.5, leading=15, textColor=GRAY)
    s['bullet']      = ParagraphStyle('bullet',      fontName="Helvetica", fontSize=11, leading=16, textColor=DARK)
    s['med_label']   = ParagraphStyle('med_label',   fontName="Helvetica-Bold", fontSize=10, leading=14, textColor=GRAY)
    s['med_val']     = ParagraphStyle('med_val',     fontName="Helvetica", fontSize=10, leading=14, textColor=DARK)
    s['mem_hdr']     = ParagraphStyle('mem_hdr',     fontName="Helvetica-Bold", fontSize=11, leading=15, textColor=GREEN_DARK)
    return s

ST = make_styles()

# ── BORDER / STYLE HELPERS ────────────────────────────────────────────────────
def tbl_style(cmds): return TableStyle(cmds)

GRID = lambda c=BORDER_C: [
    ('GRID',       (0,0), (-1,-1), 0.5, c),
    ('ROWBACKGROUNDS', (0,0), (-1,-1), [WHITE, WHITE]),
]

def header_row_style(ncols, bg=GREEN):
    return [
        ('BACKGROUND', (0,0), (ncols-1, 0), bg),
        ('TEXTCOLOR',  (0,0), (ncols-1, 0), WHITE),
        ('GRID',       (0,0), (-1,-1), 0.5, BORDER_C),
        ('FONTNAME',   (0,0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0), (-1, 0), 10),
        ('ALIGN',      (0,0), (-1, 0), 'CENTER'),
        ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING',(0,0),(-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING',(0,0), (-1,-1), 6),
    ]

def section_bar_table(text, width):
    """Grey section divider bar like Coverfox MAIN BENEFITS."""
    t = Table([[Paragraph(f"  {text}", ST['section_bar'])]], colWidths=[width])
    t.setStyle(TableStyle([
        ('BACKGROUND',    (0,0), (-1,-1), LGRAY),
        ('TOPPADDING',    (0,0), (-1,-1), 7),
        ('BOTTOMPADDING', (0,0), (-1,-1), 7),
        ('LEFTPADDING',   (0,0), (-1,-1), 4),
        ('RIGHTPADDING',  (0,0), (-1,-1), 4),
        ('GRID',          (0,0), (-1,-1), 0, WHITE),
    ]))
    return t

def green_divider(width):
    return HRFlowable(width=width, thickness=3, color=GREEN, spaceAfter=4, spaceBefore=2)

def load_logo_img(path, max_w, max_h=1.0*cm):
    if not path: return None
    try:
        img = Image(path)
        iw, ih = img.imageWidth, img.imageHeight
        ratio = iw / ih
        w = min(max_w, max_h * ratio)
        h = w / ratio
        if h > max_h:
            h = max_h
            w = h * ratio
        img.drawWidth  = w
        img.drawHeight = h
        return img
    except:
        return None

# ── GENERATOR ─────────────────────────────────────────────────────────────────
class PDFQuoteGenerator:
    def __init__(self, data, logo_folder):
        self.data     = data
        self.lf       = logo_folder
        self.members  = data.get("members", [])
        self.insurers = data.get("insurers", [])
        self.advisor  = data.get("advisor_name", "your trusted advisor")
        self.date     = data.get("date", datetime.now().strftime("%d-%m-%Y"))
        self.qid      = data.get("quote_id", "")
        self.company  = data.get("company_name", "increMint")
        self.note     = data.get("advisor_note", "")

        self.included = []
        self.prem_map = {}
        for ins in self.insurers:
            mk  = map_master(ins.get("name", ""))
            key = mk if mk else ins.get("name", "")
            if key and key not in self.included:
                self.included.append(key)
            self.prem_map[key] = ins

    # ── header block ─────────────────────────────────────────────────────────
    def _header(self, W):
        elems = []
        logo_path = find_incremint_logo(self.lf)
        logo_img  = load_logo_img(logo_path, 3.5*cm, 1.0*cm) if logo_path else None

        left_cell  = logo_img if logo_img else Paragraph(
            f'<font color="#00C853"><b>{self.company}</b></font>', ST['bold'])
        right_cell = [
            Paragraph(f'<b>Quote ID: {self.qid}</b>', ST['right']),
            Paragraph(f'Prepared by {self.advisor}   —   {self.date}', ST['right']),
        ]
        hdr_t = Table(
            [[left_cell, right_cell]],
            colWidths=[W*0.42, W*0.58]
        )
        hdr_t.setStyle(TableStyle([
            ('VALIGN',      (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING',  (0,0), (-1,-1), 4),
            ('BOTTOMPADDING',(0,0),(-1,-1), 4),
            ('GRID',        (0,0), (-1,-1), 0, WHITE),
        ]))
        elems.append(hdr_t)
        elems.append(green_divider(W))
        return elems

    # ── footer block ─────────────────────────────────────────────────────────
    def _footer_elems(self, W):
        elems = []
        elems.append(green_divider(W))
        ft = Table(
            [[
                Paragraph(f'  {self.company}  •  Quote: {self.qid}  •  {self.date}', ST['footer']),
                Paragraph('Prepared by a licensed insurance advisor. For information purposes only.  ', ST['footer_r']),
            ]],
            colWidths=[W*0.55, W*0.45]
        )
        ft.setStyle(TableStyle([
            ('GRID',         (0,0),(-1,-1), 0, WHITE),
            ('TOPPADDING',   (0,0),(-1,-1), 3),
            ('BOTTOMPADDING',(0,0),(-1,-1), 3),
        ]))
        elems.append(ft)
        return elems

    # ── Cover page ────────────────────────────────────────────────────────────
    def build_cover(self):
        W_pt = A4[0] - 2*1.4*cm   # usable width in points (margins 0.55in each)
        W    = W_pt               # just use same variable

        story = []
        story += self._header(W)
        story.append(sp(6))

        # Title
        story.append(Paragraph("COMPARE PLANS", ST['title']))
        name0 = self.members[0].get('name','your family') if self.members else 'your family'
        story.append(Paragraph(f"Health Insurance Comparison for {name0}", ST['subtitle']))
        story.append(sp(8))

        # ── Insured Members table ─────────────────────────────────────────────
        story.append(section_bar_table("INSURED MEMBERS", W))
        story.append(sp(1))

        ct_hdrs = ["#","Member Name","Relation","Date of Birth","Age","City / Pincode","Sum Assured"]
        col_ratios = [0.04, 0.24, 0.12, 0.15, 0.05, 0.21, 0.14]
        col_ws = [W * r for r in col_ratios]

        ct_data = [[Paragraph(h, ST['white_bold_c']) for h in ct_hdrs]]
        for idx, m in enumerate(self.members):
            dob = str(m.get("dob",""))
            try:
                from datetime import datetime as _dt
                dob = _dt.strptime(dob, "%Y-%m-%d").strftime("%d-%m-%Y")
            except: pass
            bg = GREEN_LIGHT if idx % 2 == 0 else WHITE
            row = [
                Paragraph(str(idx+1), ST['center']),
                Paragraph(m.get("name",""), ST['normal']),
                Paragraph(m.get("relation",""), ST['center']),
                Paragraph(dob, ST['center']),
                Paragraph(str(m.get("age","")), ST['center']),
                Paragraph(m.get("city",""), ST['normal']),
                Paragraph(m.get("sum_assured",""), ST['center']),
            ]
            ct_data.append(row)

        ct = Table(ct_data, colWidths=col_ws, repeatRows=1)
        ct_style_cmds = header_row_style(7)
        for idx in range(1, len(ct_data)):
            bg = GREEN_LIGHT if (idx-1) % 2 == 0 else WHITE
            ct_style_cmds.append(('BACKGROUND', (0,idx), (-1,idx), bg))
        ct.setStyle(TableStyle(ct_style_cmds))
        story.append(ct)
        story.append(sp(8))

        # ── Medical History ───────────────────────────────────────────────────
        has_med = any(
            (m.get("diseases") and len(m["diseases"])>0) or
            m.get("policy_type")=="port" or
            (m.get("surgeries") and len(m["surgeries"])>0)
            for m in self.members
        )
        if has_med:
            story.append(section_bar_table("MEDICAL HISTORY", W))
            story.append(sp(1))
            for m in self.members:
                has_dis = m.get("diseases") and len(m["diseases"])>0
                has_sur = m.get("surgeries") and len(m["surgeries"])>0
                is_port = m.get("policy_type")=="port"
                if not has_dis and not has_sur and not is_port: continue

                med_data = [[Paragraph(
                    f"{m.get('member_no','?')}.  {m.get('name','')}  ({m.get('relation','')})",
                    ST['mem_hdr']), ""]]
                pol = (f"Port from {m.get('port_details',{}).get('current_insurer','Unknown')}"
                       if is_port else "Fresh Policy")
                med_data.append([Paragraph("Policy Type", ST['med_label']),
                                  Paragraph(pol, ST['med_val'])])
                if is_port and m.get("port_details"):
                    p = m["port_details"]; pts=[]
                    if p.get("expiry_date"):     pts.append(f"Expiry: {p['expiry_date']}")
                    if p.get("current_sa"):      pts.append(f"SA: {p['current_sa']}")
                    if p.get("claim_last_year"): pts.append(f"Claim: {p['claim_last_year']}")
                    if p.get("claim_amount"):    pts.append(f"Amount: Rs.{p['claim_amount']}")
                    if pts:
                        med_data.append([Paragraph("Port Details", ST['med_label']),
                                         Paragraph("  |  ".join(pts), ST['med_val'])])
                conds = (", ".join(d["name"]+(f" ({d['description']})" if d.get("description") else "")
                                   for d in m.get("diseases",[])) if has_dis else "None declared")
                med_data.append([Paragraph("Pre-existing Conditions", ST['med_label']),
                                  Paragraph(conds, ST['med_val'])])
                surgs = (", ".join(f"{s['name']} ({s['year']})" for s in m.get("surgeries",[]))
                         if has_sur else "None")
                med_data.append([Paragraph("Surgeries", ST['med_label']),
                                  Paragraph(surgs, ST['med_val'])])

                med_t = Table(med_data, colWidths=[W*0.25, W*0.75])
                med_cmds = [
                    ('GRID',        (0,0), (-1,-1), 0.5, BORDER_C),
                    ('SPAN',        (0,0), (1,0)),
                    ('BACKGROUND',  (0,0), (-1,0), GREEN_LIGHT),
                    ('TOPPADDING',  (0,0), (-1,-1), 5),
                    ('BOTTOMPADDING',(0,0),(-1,-1), 5),
                    ('LEFTPADDING', (0,0), (-1,-1), 7),
                    ('RIGHTPADDING',(0,0), (-1,-1), 7),
                    ('VALIGN',      (0,0), (-1,-1), 'MIDDLE'),
                ]
                for ri in range(1, len(med_data)):
                    bg = LGRAY if ri % 2 == 1 else WHITE
                    med_cmds.append(('BACKGROUND', (0,ri), (-1,ri), bg))
                med_t.setStyle(TableStyle(med_cmds))
                story.append(med_t)
                story.append(sp(4))
            story.append(sp(6))

        # ── Advisor Note ──────────────────────────────────────────────────────
        if self.note:
            note_t = Table(
                [[Paragraph(f'<b><font color="#00A846">Advisor Note:   </font></b>'
                            f'<i><font color="#666666">{self.note}</font></i>',
                            ST['normal'])]],
                colWidths=[W]
            )
            note_t.setStyle(TableStyle([
                ('BACKGROUND',    (0,0), (-1,-1), GREEN_LIGHT),
                ('GRID',          (0,0), (-1,-1), 1.5, GREEN),
                ('TOPPADDING',    (0,0), (-1,-1), 8),
                ('BOTTOMPADDING', (0,0), (-1,-1), 8),
                ('LEFTPADDING',   (0,0), (-1,-1), 10),
                ('RIGHTPADDING',  (0,0), (-1,-1), 10),
            ]))
            story.append(note_t)
            story.append(sp(8))

        story += self._footer_elems(W)
        return story

    # ── Comparison page (landscape) ───────────────────────────────────────────
    def build_comparison(self):
        if not self.included: return []

        W_pt = A4[1] - 2*1.4*cm   # landscape: A4[1] is the longer side
        W    = W_pt

        feat_w  = W * 0.24
        plan_w  = (W - feat_w) / len(self.included)

        story = []
        story += self._header(W)
        story.append(sp(4))

        # ── Plan cards: logo + sum assured + premium ──────────────────────────
        logo_row = []
        sa_row   = []
        prem_row = []

        logo_row.append(Paragraph(""))
        sa_row.append(Paragraph("Sum Assured", ST['med_label']))
        prem_row.append([
            Paragraph("<b>Premium</b>", ST['bold']),
            Paragraph("Annual (1 year)", ST['small']),
        ])

        for name in self.included:
            ins_d    = self.prem_map.get(name, {})
            logo_path = find_logo(name if name in MASTER else map_master(name), self.lf)
            logo_img  = load_logo_img(logo_path, plan_w * 0.65, 1.1*cm)
            name_p    = Paragraph(f'<b><font color="#1A1A1A">{name}</font></b>', ST['center'])
            logo_cell = [logo_img, name_p] if logo_img else [name_p]
            logo_row.append(logo_cell)

            sa = ins_d.get("sum_assured","") or "As selected"
            sa_row.append(Paragraph(sa, ST['sa_big']))

            prem1 = fmt_prem(ins_d.get("premium_1yr",""))
            prem_row.append([
                Paragraph(prem1, ST['prem_big']),
                Paragraph("/ year", ST['prem_yr']),
            ])

        col_ws = [feat_w] + [plan_w] * len(self.included)

        plan_t = Table(
            [logo_row, sa_row, prem_row],
            colWidths=col_ws,
            rowHeights=[2.2*cm, 1.0*cm, 1.3*cm]
        )
        plan_cmds = [
            ('GRID',         (0,0), (-1,-1), 0.5, BORDER_C),
            ('VALIGN',       (0,0), (-1,-1), 'MIDDLE'),
            ('ALIGN',        (0,0), (-1,-1), 'CENTER'),
            ('ALIGN',        (0,0), (0,-1),  'LEFT'),
            ('LEFTPADDING',  (0,0), (-1,-1), 6),
            ('RIGHTPADDING', (0,0), (-1,-1), 6),
            ('TOPPADDING',   (0,0), (-1,-1), 5),
            ('BOTTOMPADDING',(0,0), (-1,-1), 5),
            ('BACKGROUND',   (0,1), (-1,1),  LGRAY),
            ('BACKGROUND',   (0,0), (-1,0),  WHITE),
            ('BACKGROUND',   (0,2), (-1,2),  WHITE),
        ]
        plan_t.setStyle(TableStyle(plan_cmds))
        story.append(plan_t)

        # Multi-year premiums
        has2 = any(str(self.prem_map.get(n,{}).get("premium_2yr","")).strip() not in ("","0") for n in self.included)
        has3 = any(str(self.prem_map.get(n,{}).get("premium_3yr","")).strip() not in ("","0") for n in self.included)
        if has2 or has3:
            multi_row = [Paragraph("Multi-year Premium", ST['med_label'])]
            for name in self.included:
                ins_d = self.prem_map.get(name,{})
                parts = []
                if has2:
                    v2 = fmt_prem(ins_d.get("premium_2yr",""))
                    if v2 != "—": parts.append(f"2yr: {v2}")
                if has3:
                    v3 = fmt_prem(ins_d.get("premium_3yr",""))
                    if v3 != "—": parts.append(f"3yr: {v3}")
                txt = "  |  ".join(parts) if parts else "—"
                col = GREEN_DARK if parts else GRAY
                multi_row.append(Paragraph(txt, ParagraphStyle('_mc', fontName="Helvetica",
                    fontSize=10, leading=14, textColor=col, alignment=TA_CENTER)))
            mt = Table([multi_row], colWidths=col_ws)
            mt.setStyle(TableStyle([
                ('GRID',         (0,0),(-1,-1), 0.5, BORDER_C),
                ('BACKGROUND',   (0,0),(-1,-1), LGRAY),
                ('VALIGN',       (0,0),(-1,-1), 'MIDDLE'),
                ('ALIGN',        (1,0),(-1,-1), 'CENTER'),
                ('TOPPADDING',   (0,0),(-1,-1), 5),
                ('BOTTOMPADDING',(0,0),(-1,-1), 5),
                ('LEFTPADDING',  (0,0),(-1,-1), 6),
                ('RIGHTPADDING', (0,0),(-1,-1), 6),
            ]))
            story.append(mt)

        story.append(sp(5))

        # ── Feature comparison groups ─────────────────────────────────────────
        for group_name, features in FEATURE_GROUPS:
            story.append(section_bar_table(group_name, W))
            story.append(sp(0.5))

            feat_rows = []
            for fi, (feat, desc) in enumerate(features):
                is_unique = (feat == "Unique Features")
                bg_feat = YELLOW_HL if is_unique else (LGRAY   if fi%2==0 else WHITE)
                bg_val  = YELLOW_BG if is_unique else (OFFWHITE if fi%2==0 else WHITE)

                feat_cell = [
                    Paragraph(feat, ST['uniq_name'] if is_unique else ST['feat_name']),
                    Paragraph(desc, ST['feat_desc']),
                ]
                row = [feat_cell]
                for name in self.included:
                    mk2 = name if name in MASTER else map_master(name)
                    val = MASTER[mk2].get(feat, "—") if mk2 and mk2 in MASTER else "—"
                    if val in ("-",): val = "—"
                    col = val_color(val)
                    sty = ParagraphStyle('_v', fontName="Helvetica",
                                         fontSize=9 if not is_unique else 8.5,
                                         leading=13, textColor=col, alignment=TA_CENTER)
                    row.append(Paragraph(val, sty))
                feat_rows.append((row, bg_feat, bg_val))

            tdata  = [r for r, _, _ in feat_rows]
            ft_tbl = Table(tdata, colWidths=col_ws)

            ft_cmds = [
                ('GRID',        (0,0), (-1,-1), 0.5, BORDER_C),
                ('VALIGN',      (0,0), (-1,-1), 'MIDDLE'),
                ('ALIGN',       (1,0), (-1,-1), 'CENTER'),
                ('TOPPADDING',  (0,0), (-1,-1), 5),
                ('BOTTOMPADDING',(0,0),(-1,-1), 5),
                ('LEFTPADDING', (0,0), (-1,-1), 6),
                ('RIGHTPADDING',(0,0), (-1,-1), 6),
            ]
            for ri, (_, bg_feat, bg_val) in enumerate(feat_rows):
                ft_cmds.append(('BACKGROUND', (0,ri), (0,ri), bg_feat))
                for ci in range(1, len(self.included)+1):
                    ft_cmds.append(('BACKGROUND', (ci,ri), (ci,ri), bg_val))
            ft_tbl.setStyle(TableStyle(ft_cmds))
            story.append(ft_tbl)
            story.append(sp(3))

        story.append(sp(5))

        # ── Advisor Recommendation ────────────────────────────────────────────
        story.append(section_bar_table("ADVISOR'S RECOMMENDATION", W))
        story.append(sp(1))

        rec_col_w = W * 0.38
        val_col_w = W * 0.62
        rec_data  = [
            [Paragraph("Plan", ST['white_bold_c']),
             Paragraph("Why choose this plan", ST['white_bold'])],
        ]
        for idx, name in enumerate(self.included):
            logo_path = find_logo(name if name in MASTER else map_master(name), self.lf)
            logo_img  = load_logo_img(logo_path, rec_col_w * 0.55, 1.0*cm)
            name_p    = Paragraph(f'<b><font color="#00A846">{name}</font></b>', ST['green_c'])
            left_cell = [logo_img, name_p] if logo_img else [name_p]

            bullets = HIGHLIGHTS.get(name, HIGHLIGHTS.get(map_master(name) or "", []))
            bullet_p = [Paragraph(f"  \u2022  {b}", ST['bullet']) for b in bullets]
            rec_data.append([left_cell, bullet_p])

        rec_t = Table(rec_data, colWidths=[rec_col_w, val_col_w])
        rec_cmds = [
            ('GRID',         (0,0), (-1,-1), 0.5, BORDER_C),
            ('BACKGROUND',   (0,0), (-1,0),  GREEN),
            ('VALIGN',       (0,0), (-1,-1), 'MIDDLE'),
            ('ALIGN',        (0,0), (0,-1),  'CENTER'),
            ('TOPPADDING',   (0,0), (-1,-1), 6),
            ('BOTTOMPADDING',(0,0), (-1,-1), 6),
            ('LEFTPADDING',  (0,0), (-1,-1), 8),
            ('RIGHTPADDING', (0,0), (-1,-1), 8),
        ]
        for idx in range(1, len(rec_data)):
            bg = GREEN_LIGHT if (idx-1) % 2 == 0 else WHITE
            rec_cmds.append(('BACKGROUND', (0,idx), (-1,idx), bg))
        rec_t.setStyle(TableStyle(rec_cmds))
        story.append(rec_t)
        story.append(sp(8))
        story += self._footer_elems(W)
        return story

    # ── Build both sections and save ─────────────────────────────────────────
    def generate(self, output_path):
        os.makedirs(os.path.dirname(output_path) if os.path.dirname(output_path) else ".", exist_ok=True)

        cover_story = self.build_cover()
        comp_story  = self.build_comparison()

        # Cover = portrait, Comparison = landscape
        # We write them as separate docs then merge
        import tempfile
        from pypdf import PdfWriter, PdfReader

        cover_path = output_path.replace(".pdf", "_cover_tmp.pdf")
        comp_path  = output_path.replace(".pdf", "_comp_tmp.pdf")

        # Portrait cover
        doc_cover = SimpleDocTemplate(
            cover_path,
            pagesize=A4,
            leftMargin=1.4*cm, rightMargin=1.4*cm,
            topMargin=1.2*cm, bottomMargin=1.2*cm,
        )
        doc_cover.build(cover_story)

        # Landscape comparison
        if comp_story:
            doc_comp = SimpleDocTemplate(
                comp_path,
                pagesize=landscape(A4),
                leftMargin=1.4*cm, rightMargin=1.4*cm,
                topMargin=1.0*cm, bottomMargin=1.0*cm,
            )
            doc_comp.build(comp_story)

            # Merge
            writer = PdfWriter()
            for path in [cover_path, comp_path]:
                reader = PdfReader(path)
                for page in reader.pages:
                    writer.add_page(page)
            with open(output_path, "wb") as f:
                writer.write(f)
            os.remove(cover_path)
            os.remove(comp_path)
        else:
            import shutil
            shutil.move(cover_path, output_path)

        return output_path


# ── ENTRY POINT ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python generate_quote_pdf.py '<json>' '<output_path>' '<logo_folder>'",
              file=sys.stderr)
        sys.exit(1)
    data        = json.loads(sys.argv[1])
    output_path = sys.argv[2]
    logo_folder = sys.argv[3] if len(sys.argv) > 3 else "logos"
    gen = PDFQuoteGenerator(data, logo_folder)
    print(gen.generate(output_path))
