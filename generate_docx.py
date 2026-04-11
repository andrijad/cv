#!/usr/bin/env python3
"""Generate DOCX CV using python-docx."""

from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── Palette ───────────────────────────────────────────────────────────────────
ACCENT    = RGBColor(0xB8, 0x54, 0x1E)
INK       = RGBColor(0x1C, 0x1A, 0x18)
INK_MID   = RGBColor(0x4A, 0x46, 0x43)
INK_FAINT = RGBColor(0x8A, 0x84, 0x80)
STONE200  = "EDE9E3"
STONE300  = "DDD8D1"
ACCENT_LT = "F2EAE3"
GREEN     = RGBColor(0x2D, 0x5E, 0x1E)
BLUE      = RGBColor(0x1D, 0x4E, 0x87)
WHITE     = "FFFFFF"

FONT = "Calibri"

# ── Helpers ───────────────────────────────────────────────────────────────────

def cell_bg(cell, hex_color: str):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)

def cell_borders(cell, top=None, bottom=None, left=None, right=None, color="DDD8D1"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcB = OxmlElement("w:tcBorders")
    for side, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        if val:
            tag = OxmlElement(f"w:{side}")
            tag.set(qn("w:val"),   val)
            tag.set(qn("w:sz"),    "4")
            tag.set(qn("w:color"), color)
            tcB.append(tag)
    tcPr.append(tcB)

def no_borders(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcB = OxmlElement("w:tcBorders")
    for side in ("top", "bottom", "left", "right", "insideH", "insideV"):
        tag = OxmlElement(f"w:{side}")
        tag.set(qn("w:val"), "none")
        tcB.append(tag)
    tcPr.append(tcB)

def cell_padding(cell, top=60, bottom=60, left=80, right=80):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for side, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        tag = OxmlElement(f"w:{side}")
        tag.set(qn("w:w"),    str(val))
        tag.set(qn("w:type"), "dxa")
        tcMar.append(tag)
    tcPr.append(tcMar)

def run(para, text, bold=False, size=10, color=None, italic=False):
    r = para.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size)
    r.font.name = FONT
    r.font.color.rgb = color if color else INK
    return r

def para_spacing(para, before=0, after=2, line=None):
    pf = para.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after  = Pt(after)
    if line:
        from docx.shared import Pt as _Pt
        pf.line_spacing = _Pt(line)

def section_heading(cell, text):
    """Uppercase accent-colored section label."""
    p = cell.add_paragraph()
    para_spacing(p, before=6, after=3)
    r = p.add_run(text.upper())
    r.bold = True
    r.font.size = Pt(7.5)
    r.font.name = FONT
    r.font.color.rgb = ACCENT
    # Bottom border via XML
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot = OxmlElement("w:bottom")
    bot.set(qn("w:val"),   "single")
    bot.set(qn("w:sz"),    "4")
    bot.set(qn("w:color"), "DDD8D1")
    pBdr.append(bot)
    pPr.append(pBdr)
    return p

def sidebar_item(cell, strong_text, normal_text="", small=False):
    sz = 8.5 if small else 9
    p = cell.add_paragraph()
    para_spacing(p, after=1)
    if strong_text:
        r = p.add_run(strong_text)
        r.bold = True
        r.font.size = Pt(sz)
        r.font.name = FONT
        r.font.color.rgb = INK
    if normal_text:
        r2 = p.add_run(("  " if strong_text else "") + normal_text)
        r2.font.size = Pt(sz)
        r2.font.name = FONT
        r2.font.color.rgb = INK_MID
    return p

def bullet(cell, text, size=9):
    p = cell.add_paragraph(style="List Bullet")
    para_spacing(p, after=1)
    pf = p.paragraph_format
    pf.left_indent = Cm(0.3)
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.name = FONT
    r.font.color.rgb = INK_MID
    return p

def job_header(cell, title, period):
    p = cell.add_paragraph()
    para_spacing(p, before=8, after=1)
    run(p, title, bold=True, size=10, color=INK)
    # right-aligned period via tab stop
    tab = OxmlElement("w:tab")
    p._p.find(qn("w:r")).append(tab)
    r2 = p.add_run(f"\t{period}")
    r2.font.size = Pt(8.5)
    r2.font.name = FONT
    r2.font.color.rgb = INK_FAINT
    r2.italic = True
    # Add right tab stop
    pPr = p._p.get_or_add_pPr()
    tabs = OxmlElement("w:tabs")
    t = OxmlElement("w:tab")
    t.set(qn("w:val"), "right")
    t.set(qn("w:pos"), "8800")
    tabs.append(t)
    pPr.append(tabs)
    return p

def job_meta(cell, text):
    p = cell.add_paragraph()
    para_spacing(p, after=3)
    r = p.add_run(text)
    r.font.size = Pt(8.5)
    r.font.name = FONT
    r.font.color.rgb = INK_FAINT
    r.italic = True
    return p

# ── Document setup ────────────────────────────────────────────────────────────
doc = Document()

sec = doc.sections[0]
sec.page_width   = Cm(21)
sec.page_height  = Cm(29.7)
sec.left_margin  = Cm(0)
sec.right_margin = Cm(0)
sec.top_margin   = Cm(0)
sec.bottom_margin = Cm(0)

# Remove default paragraph styles spacing
style = doc.styles["Normal"]
style.font.name = FONT
style.font.size = Pt(10)

# ─────────────────────────────────────────────────────────────────────────────
# HEADER TABLE  (full width, white bg, name | spacer | photo)
# ─────────────────────────────────────────────────────────────────────────────
hdr = doc.add_table(rows=1, cols=2)
hdr.style = "Table Grid"
hdr.autofit = False

# Column widths (in EMU: 1 cm = 360000 EMU)
hdr.columns[0].width = Cm(14.5)   # text side
hdr.columns[1].width = Cm(6.5)    # photo side

cText = hdr.cell(0, 0)
cPhoto = hdr.cell(0, 1)

cell_bg(cText,  WHITE)
cell_bg(cPhoto, WHITE)
no_borders(cText)
no_borders(cPhoto)
cell_padding(cText,  top=280, bottom=200, left=400, right=200)
cell_padding(cPhoto, top=280, bottom=200, left=200, right=400)

# Name
p_name = cText.paragraphs[0]
para_spacing(p_name, after=3)
r_name = p_name.add_run("Andrija Despotović")
r_name.bold = True
r_name.font.size = Pt(26)
r_name.font.name = FONT
r_name.font.color.rgb = INK

# Title
p_title = cText.add_paragraph()
para_spacing(p_title, after=10)
r_title = p_title.add_run("Senior Product Manager  ·  B2B Platforms, AI Automation & Agent Orchestration")
r_title.font.size = Pt(10)
r_title.font.name = FONT
r_title.font.color.rgb = ACCENT

# Contact info (one line)
contacts = [
    "Belgrade, Serbia",
    "andrija@devsmachina.app",
    "linkedin.com/in/andrija-despotovic",
    "github.com/andrijad",
    "devsmachina.app",
    "+381 (63) 274-702",
]
p_contact = cText.add_paragraph()
para_spacing(p_contact, after=0)
first = True
for c in contacts:
    if not first:
        sep = p_contact.add_run("  ·  ")
        sep.font.size = Pt(8.5)
        sep.font.color.rgb = INK_FAINT
        sep.font.name = FONT
    r = p_contact.add_run(c)
    r.font.size = Pt(8.5)
    r.font.color.rgb = INK_MID
    r.font.name = FONT
    first = False

# Photo placeholder
cPhoto.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

p_photo = cPhoto.paragraphs[0]
p_photo.alignment = WD_ALIGN_PARAGRAPH.CENTER
para_spacing(p_photo, after=2)
r_ph = p_photo.add_run("[ Photo ]")
r_ph.font.size = Pt(9)
r_ph.font.name = FONT
r_ph.font.color.rgb = INK_FAINT
r_ph.italic = True

p_ph2 = cPhoto.add_paragraph()
p_ph2.alignment = WD_ALIGN_PARAGRAPH.CENTER
para_spacing(p_ph2, after=2)
r_ph2 = p_ph2.add_run("3 × 3.5 cm")
r_ph2.font.size = Pt(7.5)
r_ph2.font.name = FONT
r_ph2.font.color.rgb = INK_FAINT

# Draw dashed border around photo cell using paragraph border trick
# (DOCX doesn't natively do dashed cell borders easily — use a box border on paragraphs)
tc = cPhoto._tc
tcPr = tc.get_or_add_tcPr()
tcB = OxmlElement("w:tcBorders")
for side in ("top", "bottom", "left", "right"):
    tag = OxmlElement(f"w:{side}")
    tag.set(qn("w:val"),   "dashed")
    tag.set(qn("w:sz"),    "8")
    tag.set(qn("w:color"), "DDD8D1")
    tcB.append(tag)
tcPr.append(tcB)

# Separator line below header
doc.add_paragraph()   # small spacer

# ─────────────────────────────────────────────────────────────────────────────
# BODY TABLE  (sidebar 5.5cm | main 15.5cm)
# ─────────────────────────────────────────────────────────────────────────────
body = doc.add_table(rows=1, cols=2)
body.style = "Table Grid"
body.autofit = False

body.columns[0].width = Cm(5.5)   # sidebar
body.columns[1].width = Cm(15.5)  # main

SB  = body.cell(0, 0)   # sidebar cell
MN  = body.cell(0, 1)   # main cell

cell_bg(SB, STONE200)
cell_bg(MN, WHITE)
no_borders(SB)
no_borders(MN)
cell_padding(SB, top=180, bottom=180, left=400, right=200)
cell_padding(MN, top=180, bottom=180, left=320, right=400)

SB.vertical_alignment = WD_ALIGN_VERTICAL.TOP
MN.vertical_alignment = WD_ALIGN_VERTICAL.TOP

# Clear default empty paragraph
SB.paragraphs[0].clear()
MN.paragraphs[0].clear()

# ── SIDEBAR ──────────────────────────────────────────────────────────────────

# Languages
section_heading(SB, "Languages")
sidebar_item(SB, "Serbian", "native")
sidebar_item(SB, "English", "professional")
sidebar_item(SB, "Spanish", "A2")

# Education
section_heading(SB, "Education")
sidebar_item(SB, "MA Sociology", "")
p = SB.add_paragraph()
para_spacing(p, after=4)
r = p.add_run("Univ. of Belgrade  ·  9.8 GPA  ·  2010")
r.font.size = Pt(8.5); r.font.name = FONT; r.font.color.rgb = INK_MID

sidebar_item(SB, "MA Anthropology", "")
p = SB.add_paragraph()
para_spacing(p, after=4)
r = p.add_run("Univ. of Belgrade  ·  9.0 GPA  ·  2007")
r.font.size = Pt(8.5); r.font.name = FONT; r.font.color.rgb = INK_MID

# Clients
section_heading(SB, "Clients")
for client in ["Keestrack", "MDConnect", "Canvas", "BASF",
               "Schneider Electric", "National Museum of Serbia", "GIFT", "Vellis"]:
    p = SB.add_paragraph()
    para_spacing(p, after=1)
    r = p.add_run(client)
    r.font.size = Pt(8.5); r.font.name = FONT; r.font.color.rgb = INK_MID

# Traits
section_heading(SB, "Traits")
traits = ["builder", "tinkerer", "ships without direction", "early adopter",
          "pragmatic", "goal oriented", "quick learner", "communicator",
          "organizer", "comfortable with ambiguity"]
p = SB.add_paragraph()
para_spacing(p, after=0)
r = p.add_run(" · ".join(traits))
r.font.size = Pt(8); r.font.name = FONT; r.font.color.rgb = INK_MID

# ── MAIN CONTENT ─────────────────────────────────────────────────────────────

# ── SUMMARY ──────────────────────────────────────────────────────────────────
section_heading(MN, "Summary")

# Summary box (shaded paragraph)
p_sum = MN.add_paragraph()
para_spacing(p_sum, before=2, after=12)
r_sum = p_sum.add_run(
    "Senior PM with 9+ years shipping B2B platforms and AI-integrated products. "
    "I build and operate AI automation infrastructure — multi-agent workflows, LLM API integrations, "
    "and Cloudflare-based platform services — alongside managing enterprise ecommerce platforms "
    "for multi-region clients. Full ownership from product strategy to hands-on deployment, "
    "without big budgets or large teams. Comfortable with ambiguity, shipping fast, and iterating "
    "on real partner feedback."
)
r_sum.font.size = Pt(9.5)
r_sum.font.name = FONT
r_sum.font.color.rgb = INK

# Left border effect on summary via paragraph shading
pPr = p_sum._p.get_or_add_pPr()
shd = OxmlElement("w:shd")
shd.set(qn("w:val"),   "clear")
shd.set(qn("w:color"), "auto")
shd.set(qn("w:fill"),  "F2EAE3")
pPr.append(shd)

pBdr = OxmlElement("w:pBdr")
left = OxmlElement("w:left")
left.set(qn("w:val"),   "single")
left.set(qn("w:sz"),    "18")
left.set(qn("w:color"), "B8541E")
pBdr.append(left)
pPr.append(pBdr)

# ── SKILLS TABLE ─────────────────────────────────────────────────────────────
section_heading(MN, "Skills & Technical Stack")

skills = [
    ("AI / LLM Integration",
     "Anthropic Claude API, OpenAI API, RAG architecture, agent orchestration, prompt engineering, "
     "LLM cost/latency tradeoffs, hallucination mitigation strategies"),
    ("Automation & Infra",
     "n8n (self-hosted), Docker, Cloudflare Workers/Pages/Zero Trust, GitHub Actions CI/CD, "
     "Linux VPS, webhook pipeline design"),
    ("Platform & API",
     "API integration architecture, partner API discovery & requirements, Cloudflare Workers routing, "
     "third-party API orchestration"),
    ("Product Management",
     "Backlog management, user story writing, stakeholder alignment, Agile/Scrum, roadmap ownership, "
     "A/B testing, zero-to-one delivery"),
    ("eCommerce & SaaS",
     "WooCommerce, multi-tenant B2B platforms, ERP/BI integration, SEO, UX/UI collaboration, "
     "CMS ecosystems"),
    ("Data & Analytics",
     "GA4, GSC, GTM, BI tools, data-driven prioritization; SQL fundamentals"),
    ("Tools",
     "Jira, Confluence, Trello, Asana, GitHub, Notion, Figma"),
]

sk_table = MN.add_table(rows=len(skills), cols=2)
sk_table.style = "Table Grid"
sk_table.autofit = False
sk_table.columns[0].width = Cm(3.8)
sk_table.columns[1].width = Cm(11.2)

for i, (label, value) in enumerate(skills):
    bg = "F7F4F0" if i % 2 == 0 else WHITE
    c0 = sk_table.cell(i, 0)
    c1 = sk_table.cell(i, 1)
    cell_bg(c0, bg)
    cell_bg(c1, bg)
    cell_padding(c0, top=60, bottom=60, left=80, right=80)
    cell_padding(c1, top=60, bottom=60, left=80, right=80)
    cell_borders(c0, top="single", bottom="single", left="single", right="single")
    cell_borders(c1, top="single", bottom="single", left="single", right="single")

    p0 = c0.paragraphs[0]
    r0 = p0.add_run(label)
    r0.bold = True; r0.font.size = Pt(9); r0.font.name = FONT; r0.font.color.rgb = INK

    p1 = c1.paragraphs[0]
    r1 = p1.add_run(value)
    r1.font.size = Pt(9); r1.font.name = FONT; r1.font.color.rgb = INK_MID

p_gap = MN.add_paragraph()
para_spacing(p_gap, after=8)

# ── WORK EXPERIENCE ───────────────────────────────────────────────────────────
section_heading(MN, "Work Experience")

jobs = [
    {
        "title":  "Founder & Product Builder — DevsMachina",
        "period": "2025 – present",
        "meta":   "devsmachina.app · Belgrade, Serbia",
        "bullets": [
            "Founded and operate an AI studio focused on agent orchestration, automation workflows, and web platform development for B2B clients.",
            "Shipped products from zero without design teams or external budget: CutTool (free B2B browser app for steel coil cutting), Tasklog (vanilla HTML/JS/CSS tracker via Cloudflare Workers + GitHub), and a WooCommerce AI description generator (Supabase + Claude API + Lemon Squeezy + Next.js, credits-based pricing).",
            "Designed and deployed multi-agent LLM workflows using Claude API and n8n — RAG pipelines, prompt chaining, and cost-optimized API orchestration.",
            "Self-managed full infrastructure: Cloudflare Workers/Pages, Zero Trust access control, Docker, Linux VPS, GitHub Actions CI/CD.",
            "Maintained and optimized client websites across medical ecommerce, B2B heavy equipment, and SaaS — billing in EUR and USD.",
        ],
    },
    {
        "title":  "Product Owner / Account Manager — Keystone Studio",
        "period": "2023 – present",
        "meta":   "Belgrade, Serbia",
        "bullets": [
            "Owning B2B platform product strategy for multi-tenant ecommerce systems — defining API integration requirements, managing partner onboarding flows, and shipping AI-assisted features for enterprise clients.",
            "Aligned business, UX/UI, and engineering teams; managed backlog, delivery planning, and UAT cycles across concurrent client projects.",
            "Led integration initiatives connecting third-party logistics, ERP, and BI systems into the core platform.",
            "Direct stakeholder management across enterprise clients billed in EUR and USD.",
        ],
    },
    {
        "title":  "Senior Project Manager / Product Owner — Devokado / Avokado",
        "period": "2019 – present",
        "meta":   "devokado.rs · Belgrade, Serbia",
        "bullets": [
            "Managed web and custom SaaS development across multiple parallel client projects — platform performance, feature delivery, and cross-functional coordination.",
            "Delivered tailored tools and integrations for Keestrack, Canvas, BASF, Schneider Electric, MDConnect, and others.",
        ],
    },
    {
        "title":  "Account Manager — Degordian",
        "period": "2020 – 2022",
        "meta":   "degordian.com · Belgrade, Serbia; Zagreb, Croatia",
        "bullets": [
            "Led digital projects for Štark, BASF, Jelen, Schneider Electric, and the Innovation Fund — web development, SEO, PPC, social campaigns.",
            "Main point of contact across parallel client workstreams; shaped requirements and aligned stakeholders through delivery.",
        ],
    },
    {
        "title":  "Account Director / Manager — Next Game",
        "period": "2016 – 2019",
        "meta":   "nextgame.rs · Belgrade, Serbia",
        "bullets": [
            "Led digital and creative accounts for B2B and consumer clients; managed cross-functional teams across design, development, and media.",
        ],
    },
    {
        "title":  "Curator & Project Manager — National Museum of Serbia",
        "period": "2009 – 2016",
        "meta":   "narodnimuzej.rs · Belgrade, Serbia",
        "bullets": [
            "Managed digital and cultural projects including the EU-funded GIFT Horizon 2020 interactive gamification initiative for European museums.",
            "Coordinated cross-border teams across UX, content, and academic domains.",
            "Museum website was chosen as the best website in CH sector in 2018 by the PC magazine.",
        ],
    },
    {
        "title":  "Misc — Early Career",
        "period": "2007 – 2009",
        "meta":   "",
        "bullets": [
            "Reporter at Večernje novosti (daily newspaper) and Svet kompjutera (tech magazine) — covered technology, culture, and current affairs.",
        ],
    },
]

for job in jobs:
    job_header(MN, job["title"], job["period"])
    if job["meta"]:
        job_meta(MN, job["meta"])
    for b in job["bullets"]:
        bullet(MN, b)

p_gap2 = MN.add_paragraph()
para_spacing(p_gap2, after=8)

# ── PROJECTS ─────────────────────────────────────────────────────────────────
section_heading(MN, "Selected Products (Zero-to-One)")

projects = [
    ("DevsMachina",
     "AI studio platform — hub for agent orchestration, automation, and web development services. "
     "Built on Astro + Cloudflare Pages + GitHub Actions CI/CD.",
     "devsmachina.app", "MVP"),
    ("CutTool",
     "Free B2B browser app for steel coil cutting optimization. "
     "Built, shipped, and maintained as sole PM/developer. Zero external budget.",
     "cuttool.app", "live"),
    ("Tasklog",
     "Single-file vanilla HTML/JS/CSS project tracker with localStorage persistence. "
     "Deployed via Cloudflare Worker serving directly from GitHub repo.",
     "tasklog.devsmachina.app", "live"),
    ("WooCommerce AI Desc. Gen.",
     "Micro-SaaS AI product description generator for WooCommerce stores. "
     "Stack: Supabase, Claude API, Lemon Squeezy, n8n Cloud, Next.js. "
     "4 vertical Expert Modes, credits-based pricing.",
     "", "in dev"),
    ("GIFT — Horizon 2020",
     "EU-funded interactive gamification project for European museums. "
     "Digital engagement experiences across cultural institutions.",
     "cordis.europa.eu/project/id/727040", "shipped"),
]

pr_table = MN.add_table(rows=len(projects), cols=3)
pr_table.style = "Table Grid"
pr_table.autofit = False
pr_table.columns[0].width = Cm(3.2)
pr_table.columns[1].width = Cm(9.3)
pr_table.columns[2].width = Cm(2.5)

for i, (name, desc, url, status) in enumerate(projects):
    bg = "E8F2E2" if i % 2 == 0 else WHITE
    for col in range(3):
        c = pr_table.cell(i, col)
        cell_bg(c, bg)
        cell_padding(c, top=60, bottom=60, left=80, right=80)
        cell_borders(c, top="single", bottom="single", left="single", right="single")

    c0 = pr_table.cell(i, 0)
    p0 = c0.paragraphs[0]
    r = p0.add_run(name)
    r.bold = True; r.font.size = Pt(9); r.font.name = FONT; r.font.color.rgb = INK

    c1 = pr_table.cell(i, 1)
    p1 = c1.paragraphs[0]
    r1 = p1.add_run(desc)
    r1.font.size = Pt(9); r1.font.name = FONT; r1.font.color.rgb = INK_MID
    if url:
        p_url = c1.add_paragraph()
        para_spacing(p_url, after=0)
        r_url = p_url.add_run(url)
        r_url.font.size = Pt(8); r_url.font.name = FONT; r_url.font.color.rgb = BLUE

    c2 = pr_table.cell(i, 2)
    c2.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p2 = c2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    status_colors = {"live": GREEN, "MVP": ACCENT, "in dev": INK_FAINT, "shipped": BLUE}
    r2 = p2.add_run(status.upper())
    r2.bold = True; r2.font.size = Pt(7.5); r2.font.name = FONT
    r2.font.color.rgb = status_colors.get(status, INK_MID)

# ── Save ──────────────────────────────────────────────────────────────────────
out = "Andrija_Despotovic_CV.docx"
doc.save(out)
from pathlib import Path
sz = Path(out).stat().st_size // 1024
print(f"Done: {out} ({sz} KB)")
