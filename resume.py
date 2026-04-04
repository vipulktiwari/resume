from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

OUTPUT_PATH = "VipulKumarTiwari_Resume.docx"

DARK_GRAY = RGBColor(0x2C, 0x2C, 0x2C)
MUTED = RGBColor(0x55, 0x55, 0x55)

doc = Document()

# ── Page margins ──────────────────────────────────────────────────────────────
for section in doc.sections:
    section.page_width  = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin    = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin   = Inches(0.5)
    section.right_margin  = Inches(0.5)

# ── Helpers ───────────────────────────────────────────────────────────────────

def clear_paragraph_spacing(p):
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)


def set_run(run, text, bold=False, size=9, color=DARK_GRAY, font="Arial"):
    run.text = text
    run.bold = bold
    run.font.name = font
    run.font.size = Pt(size)
    run.font.color.rgb = color


def add_paragraph(cell_or_doc, text, bold=False, size=9, color=DARK_GRAY,
                  align=WD_ALIGN_PARAGRAPH.LEFT, space_before=2, space_after=2):
    p = cell_or_doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    run = p.add_run(text)
    set_run(run, text, bold=bold, size=size, color=color)
    return p


def add_section_header(cell, text):
    p = cell.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(1)
    run = p.add_run(text.upper())
    run.bold = True
    run.font.name = "Arial"
    run.font.size = Pt(9)
    run.font.color.rgb = DARK_GRAY
    # bottom border as section line
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '2C2C2C')
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p


def add_bullet(cell, text, size=9):
    p = cell.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(2)
    p.paragraph_format.left_indent  = Inches(0.2)
    run = p.add_run(text)
    run.font.name = "Arial"
    run.font.size = Pt(size)
    run.font.color.rgb = DARK_GRAY


def add_job_header(cell, title, company, date, location):
    p = cell.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(1)
    r = p.add_run(title)
    r.bold = True; r.font.name = "Arial"; r.font.size = Pt(9); r.font.color.rgb = DARK_GRAY

    p2 = cell.add_paragraph()
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after  = Pt(3)
    for part in [company + "   ", date + "   ", location]:
        r2 = p2.add_run(part)
        r2.font.name = "Arial"; r2.font.size = Pt(8); r2.font.color.rgb = MUTED


def set_cell_border_none(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side in ('top', 'left', 'bottom', 'right'):
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'), 'none')
        tcBorders.append(el)
    tcPr.append(tcBorders)


def set_col_width(cell, width_inches):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), str(int(width_inches * 1440)))
    tcW.set(qn('w:type'), 'dxa')
    tcPr.append(tcW)


# ── Header (name / title / contact) ───────────────────────────────────────────

add_paragraph(doc, "VIPUL KUMAR TIWARI", bold=True, size=18, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=2)
add_paragraph(doc, "Principal Member of Technical Staff", size=11, color=MUTED, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=2)
add_paragraph(doc, "+91-8839814859  |  er.vktcs@gmail.com  |  linkedin.com/in/vipulkumartiwari",
              size=9, color=MUTED, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=6)

# ── Two-column table ───────────────────────────────────────────────────────────
table = doc.add_table(rows=1, cols=2)
table.style = 'Table Grid'

left  = table.cell(0, 0)
right = table.cell(0, 1)

set_cell_border_none(left)
set_cell_border_none(right)
set_col_width(left,  5.2)
set_col_width(right, 2.3)

# Remove default empty paragraph in cells
for cell in [left, right]:
    for p in cell.paragraphs:
        p._element.getparent().remove(p._element)

# ════════════════════════════════════════════════════════════
# LEFT COLUMN
# ════════════════════════════════════════════════════════════

add_section_header(left, "Summary")
add_paragraph(left,
    "Software Developer with 10+ years of experience in designing and building scalable, "
    "high-performance products. Expert in system design, problem-solving, and developing "
    "efficient, maintainable code. Extensive experience in databases and storage systems, "
    "with a strong focus on performance optimization, scalability, and reliability of "
    "data-driven applications.",
    size=9, space_before=3, space_after=6)

add_section_header(left, "Experience")

# Oracle
add_job_header(left, "Principal Member of Technical Staff", "Oracle", "02/2026 - Present", "Pune, India")
add_bullet(left, "Responsible for enhancing the LogMiner component in Oracle RDBMS, used in High Availability (HA) solutions and Oracle GoldenGate for database replication and migration.")

# SAP Labs
add_job_header(left, "Senior Software Developer", "SAP Labs", "07/2020 - 02/2026", "Pune, India")
add_bullet(left, "Contributed to the development of SAP Adaptive Server Enterprise a high-performance OLTP database system, focusing on kernel-layer enhancements and improving the Job Scheduler subsystem for better performance and reliability.")
add_bullet(left, "Handled high-priority, business-critical escalations for enterprise SAP ASE customers, leveraging strong expertise in database internals.")
add_bullet(left, "Designed and developed the Hanaservice Sync Agent, a Python-based microservice deployed on a Kubernetes cloud architecture, enabling reliable cross-cluster synchronization of Hanaservice Custom Resources (CRs).")
add_bullet(left, "Owned end-to-end delivery of cloud-plane services, including component testing frameworks, integration and E2E test suites, observability (alerts & monitoring), and comprehensive customer-facing and internal documentation.")

# Druva
add_job_header(left, "Staff Software Engineer", "Druva Data Solutions", "03/2018 - 07/2020", "Pune, India")
add_bullet(left, "Involved in design and development of Quaere, a metadata search microservice, from scratch using AWS services (DynamoDB and S3) for scalable storage, enabling efficient and reliable metadata search functionality.")
add_bullet(left, "Enhanced mstore, a mail storage service leveraging Quaere as the underlying storage namespace provider, improving system efficiency and integration.")
add_bullet(left, "Optimized search performance and reduced COGS by improving underlying architecture and resource utilization.")
add_bullet(left, "Designed and developed a CI system in Python using Docker, enabling automated testing and seamless code integration/merge for Quaere.")

# Amdocs
add_job_header(left, "Software Developer", "Amdocs DVCI", "09/2015 - 03/2018", "Pune, India")
add_bullet(left, "Involved in full lifecycle development of Change Requests for Accounts Receivable and Billing modules for AT&T Telecom, ensuring timely and high-quality delivery.")
add_bullet(left, "Provided on-site production support in Mexico for a Revenue Recognition change request, resolving critical issues and ensuring system stability.")
add_bullet(left, "Contributed to the Ensemble System MPS component, developing solutions to process Call Detail Records (CDRs) into usage data, which was further used for billing and charge generation.")
add_bullet(left, "Successfully delivered multiple Change Requests for the T-Mobile US client as an MAF/MPS developer, owning DCUT activities to ensure smooth deployment and high-quality releases.")

# ════════════════════════════════════════════════════════════
# RIGHT COLUMN
# ════════════════════════════════════════════════════════════

add_section_header(right, "Key Achievements")
add_paragraph(right, "Achievements and Awards", bold=True, size=9, space_before=3, space_after=2)
add_bullet(right, "Received several on-the-spot awards from SAP ASE customer for exceptional fix delivery, responsiveness, and professionalism.")
add_bullet(right, "Received outstanding achievement award at Druva for developing Quaere")

add_section_header(right, "Skills")
for line in ["System Design  Databases  Storage",
             "C/C++  Python  Golang  AWS",
             "DynamoDB  S3  Docker",
             "Data Structures  Algorithms"]:
    add_paragraph(right, line, size=9, space_before=2, space_after=2)

add_section_header(right, "Education")
add_paragraph(right, "Bachelor's degree in Computer Science & Engineering", bold=True, size=9, space_before=4, space_after=1)
add_paragraph(right, "Rajiv Gandhi Technical University", size=8, color=MUTED, space_before=0, space_after=1)
add_paragraph(right, "07/2011 - 06/2015  Bhopal", size=8, color=MUTED, space_before=0, space_after=1)
add_paragraph(right, "Grade: 8.07 / 10", size=9, space_before=0, space_after=5)

add_paragraph(right, "AISSCE", bold=True, size=9, space_before=3, space_after=1)
add_paragraph(right, "Jawahar Navodaya Vidhyalaya Narsinghpur", size=8, color=MUTED, space_before=0, space_after=1)
add_paragraph(right, "07/2008 - 04/2010", size=8, color=MUTED, space_before=0, space_after=1)
add_paragraph(right, "Grade: 79.6 / 100", size=9, space_before=0, space_after=5)

add_section_header(right, "Languages")
add_paragraph(right, "Hindi — Native", size=9, space_before=3, space_after=2)
add_paragraph(right, "English — Proficient", size=9, space_before=0, space_after=2)

# ── Save ───────────────────────────────────────────────────────────────────────
doc.save(OUTPUT_PATH)
print(f"Saved: {OUTPUT_PATH}")
