from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from textwrap import wrap
import json
import os
import tempfile

from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from report_diagrams import (
    create_sequence_diagram,
    create_state_diagram,
    create_ui_component_contract,
    create_use_case_flow,
    create_vibe_layout,
)

ROOT = Path(os.environ.get('VIBE_CODE_ROOT', Path(__file__).resolve().parents[2]))
REPORTS = ROOT / 'wiki/product/reports'
PRD_OUT = REPORTS / 'vibe-code-no-code-poc-prd-and-delivery-plan.docx'
TECH_OUT = REPORTS / 'vibe-code-poc-technical-design-and-runtime-architecture.docx'
DIAGRAM_DIR = Path(os.environ.get('VIBE_REPORT_DIAGRAM_DIR', Path(tempfile.gettempdir()) / 'vibe-doc-v2-diagrams'))
STATE_CONTRACT_PATH = Path(__file__).with_name('state_contract.json')
STATE_CONTRACT = json.loads(STATE_CONTRACT_PATH.read_text(encoding='utf-8'))
DIAGRAM_DIR.mkdir(parents=True, exist_ok=True)
REPORTS.mkdir(parents=True, exist_ok=True)

NOW = datetime.now(timezone.utc)

# ---------------- shared visual helpers ----------------
NAVY = '0F172A'
TEAL = '0F766E'
BLUE = '2563EB'
PALE_BLUE = 'DBEAFE'
PALE_TEAL = 'CCFBF1'
PALE_AMBER = 'FEF3C7'
PALE_RED = 'FEE2E2'
LIGHT = 'F8FAFC'
WHITE = 'FFFFFF'
MUTED = '475569'
BORDER = 'CBD5E1'


def font(size: int, bold: bool = False):
    choices = [
        '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf' if bold else '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
        '/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf' if bold else '/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf',
    ]
    for path in choices:
        if Path(path).exists():
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def draw_box(d, xy, title, lines, fill=WHITE, accent=BLUE):
    x1, y1, x2, y2 = xy
    d.rounded_rectangle(xy, radius=20, fill=fill, outline='#94A3B8', width=3)
    d.rectangle((x1, y1, x1 + 9, y2), fill=accent)
    d.text((x1 + 28, y1 + 20), title, font=font(25, True), fill='#0F172A')
    y = y1 + 64
    for line in lines:
        d.text((x1 + 28, y), line, font=font(18), fill='#475569')
        y += 29


def arrow(d, start, end, label=''):
    import math
    x1, y1 = start
    x2, y2 = end
    d.line((x1, y1, x2, y2), fill='#64748B', width=4)
    ang = math.atan2(y2-y1, x2-x1)
    p1 = (x2 - 18 * math.cos(ang - .55), y2 - 18 * math.sin(ang - .55))
    p2 = (x2 - 18 * math.cos(ang + .55), y2 - 18 * math.sin(ang + .55))
    d.polygon([(x2, y2), p1, p2], fill='#64748B')
    if label:
        mx, my = (x1+x2)//2, (y1+y2)//2
        bbox = d.textbbox((0, 0), label, font=font(15, True))
        w = bbox[2]-bbox[0]
        d.rounded_rectangle((mx-w//2-8, my-25, mx+w//2+8, my+2), radius=6, fill='#F8FAFC')
        d.text((mx-w//2, my-22), label, font=font(15, True), fill='#475569')


def create_product_flow(path: Path):
    im = Image.new('RGB', (1900, 760), '#F8FAFC')
    d = ImageDraw.Draw(im)
    d.text((60, 35), 'Stage 1 Product Journey', font=font(38, True), fill='#0F172A')
    d.text((60, 88), 'The product promise, independent of low-level runtime implementation', font=font(21), fill='#475569')
    boxes = [
        (70, 190, 360, 405, '1  Describe', ['Ask for a business', 'change in chat'], '#EFF6FF', '#2563EB'),
        (440, 190, 730, 405, '2  Build', ['Agent proposes and', 'checks file changes'], '#ECFEFF', '#0F766E'),
        (810, 190, 1100, 405, '3  Review', ['Inspect diff and', 'accept, reject or iterate'], '#FFF7ED', '#D97706'),
        (1180, 190, 1470, 405, '4  Preview', ['See eligible changes', 'update automatically'], '#F0FDF4', '#16A34A'),
        (1550, 190, 1840, 405, '5  Recover', ['Checkpoint, undo,', 'resume and retry'], '#F5F3FF', '#7C3AED'),
    ]
    for i, (x1,y1,x2,y2,title,lines,fill,accent) in enumerate(boxes):
        draw_box(d, (x1,y1,x2,y2), title, lines, fill, accent)
        if i < len(boxes)-1:
            arrow(d, (x2, 298), (boxes[i+1][0], 298))
    d.rounded_rectangle((185, 530, 1715, 680), radius=22, fill='#FFFFFF', outline='#CBD5E1', width=3)
    d.text((230, 560), 'Stage 1 boundary', font=font(25, True), fill='#0F172A')
    d.text((230, 606), 'Mutable owner-only preview + checkpoints. No Deploy button, published release, registry, promotion or rollback.', font=font(20), fill='#475569')
    im.save(path)


def create_architecture(path: Path):
    im = Image.new('RGB', (1900, 1120), '#F8FAFC')
    d = ImageDraw.Draw(im)
    d.text((55, 28), 'Vibe Code Stage 1 — Technical Runtime Context', font=font(38, True), fill='#0F172A')
    d.text((55, 78), 'Single-node K3s validates the future workload model without claiming node-level high availability.', font=font(20), fill='#475569')
    draw_box(d, (65, 145, 420, 330), 'Browser', ['React product UI', 'Chat + authenticated preview'], '#EFF6FF', '#2563EB')
    draw_box(d, (520, 145, 875, 330), 'Traefik', ['TLS + wildcard ingress', 'HTTP / SSE / WSS'], '#FFF7ED', '#D97706')
    draw_box(d, (980, 145, 1450, 330), 'Platform Services', ['NestJS + Prisma', 'AdminJS + workers', 'PostgreSQL + Redis'], '#ECFEFF', '#0F766E')
    draw_box(d, (1535, 145, 1840, 330), 'Model Gateway', ['Scoped model route', 'No provider key in project'], '#F5F3FF', '#7C3AED')
    arrow(d, (420, 236), (520, 236), 'HTTPS/WSS')
    arrow(d, (875, 236), (980, 236), 'app/API')
    arrow(d, (1450, 236), (1535, 236), 'approved')

    d.rounded_rectangle((60, 405, 1840, 1055), radius=24, fill='#FFFFFF', outline='#94A3B8', width=3)
    d.text((95, 430), 'Single-node K3s runtime', font=font(29, True), fill='#0F172A')
    d.text((95, 472), 'containerd + runc; exact namespace, storage class and network-policy implementation confirmed by spike', font=font(18), fill='#475569')

    draw_box(d, (110, 555, 470, 755), 'Preview Gateway', ['Owner authorization', 'Project route resolution', 'Same-origin /api routing'], '#FFF7ED', '#D97706')
    draw_box(d, (580, 555, 970, 755), 'Preview Pod  0 or 1', ['Vite :5173', 'FastAPI/Uvicorn :8000', 'Mutable development runtime'], '#EFF6FF', '#2563EB')
    draw_box(d, (1090, 555, 1450, 755), 'Workspace PVC', ['Generated source', 'SQLite app data', 'Durable across Pod restart'], '#F0FDF4', '#16A34A')
    draw_box(d, (1530, 555, 1790, 755), 'Agent Job', ['Mounts same PVC', 'One mutation/project'], '#F5F3FF', '#7C3AED')
    arrow(d, (470, 650), (580, 650), 'Service')
    arrow(d, (970, 650), (1090, 650), 'mount')
    arrow(d, (1530, 650), (1450, 650), 'writes')
    arrow(d, (700, 330), (280, 555), '*.preview')
    arrow(d, (1215, 330), (1660, 555), 'creates Job')

    d.rounded_rectangle((130, 845, 1770, 990), radius=18, fill='#F8FAFC', outline='#CBD5E1', width=2)
    d.text((175, 870), 'Load-balancing boundary', font=font(23, True), fill='#0F172A')
    d.text((175, 914), 'Stateless platform services may have multiple Pods on the same node. Each mutable preview remains one replica because of Vite state, shared workspace and SQLite.', font=font(18), fill='#475569')
    d.text((175, 947), 'Pod/process recovery: yes. Host/node high availability: no.', font=font(18, True), fill='#B45309')
    im.save(path)


def create_reload_flow(path: Path):
    im = Image.new('RGB', (1900, 980), '#F8FAFC')
    d = ImageDraw.Draw(im)
    d.text((55, 28), 'Natural-language Request to Live Preview', font=font(38, True), fill='#0F172A')
    d.text((55, 80), 'The live-preview requirement is satisfied by explicit frontend and backend update paths.', font=font(20), fill='#475569')
    actors = [('Browser',150),('NestJS',480),('Agent Job',810),('Workspace PVC',1140),('Preview Pod',1470),('Preview Gateway',1780)]
    for name,x in actors:
        d.text((x-60, 135), name, font=font(20, True), fill='#0F172A')
        d.line((x,175,x,910),fill='#CBD5E1',width=3)
    steps = [
        (150,480,220,'1  Submit request'),
        (480,810,300,'2  Create admitted Agent Job'),
        (810,1140,390,'3  Write validated source files'),
        (1140,1470,480,'4  Filesystem watcher detects change'),
        (1470,1780,570,'5a  Vite HMR update via WSS'),
        (1470,1470,660,'5b  Uvicorn reloads Python process'),
        (1780,150,750,'6  Preview updates/reconnects'),
        (480,150,840,'7  Persist result, evidence and checkpoint'),
    ]
    for x1,x2,y,label in steps:
        if x1 == x2:
            d.arc((x1-25,y-15,x1+90,y+35),0,300,fill='#64748B',width=4)
            d.text((x1+95,y-5),label,font=font(17,True),fill='#475569')
        else:
            arrow(d,(x1,y),(x2,y),label)
    d.rounded_rectangle((80, 915, 1820, 965), radius=10, fill='#FEF3C7', outline='#F59E0B', width=2)
    d.text((110, 928), 'Dependency, environment, migration or runtime-contract changes are classified as restart, rebuild or approval required — not falsely reported as HMR.', font=font(17, True), fill='#92400E')
    im.save(path)

PRODUCT_FLOW = DIAGRAM_DIR / 'product-flow.png'
ARCHITECTURE = DIAGRAM_DIR / 'technical-architecture.png'
RELOAD_FLOW = DIAGRAM_DIR / 'live-reload-flow.png'
create_product_flow(PRODUCT_FLOW)
create_architecture(ARCHITECTURE)
create_reload_flow(RELOAD_FLOW)


class DocBuilder:
    def __init__(self, title: str, header: str, subject: str):
        self.doc = Document()
        self.title = title
        self.header_text = header
        self.subject = subject
        self._setup()

    def _setup(self):
        doc = self.doc
        for section in doc.sections:
            section.top_margin = Inches(.62)
            section.bottom_margin = Inches(.62)
            section.left_margin = Inches(.68)
            section.right_margin = Inches(.68)
        normal = doc.styles['Normal']
        normal.font.name = 'Aptos'
        normal.font.size = Pt(9.4)
        normal._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')
        for name, size, color in [('Title',30,NAVY),('Heading 1',20,NAVY),('Heading 2',14,TEAL),('Heading 3',11,MUTED)]:
            st = doc.styles[name]
            st.font.name = 'Aptos Display'
            st.font.size = Pt(size)
            st.font.bold = True
            st.font.color.rgb = RGBColor.from_string(color)
            st._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')
            if name != 'Title':
                st.paragraph_format.keep_with_next = True
        doc.styles['Heading 1'].paragraph_format.space_before = Pt(11)
        doc.styles['Heading 1'].paragraph_format.space_after = Pt(5)
        doc.styles['Heading 2'].paragraph_format.space_before = Pt(8)
        doc.styles['Heading 2'].paragraph_format.space_after = Pt(3)
        for section in doc.sections:
            h = section.header.paragraphs[0]
            h.text = self.header_text
            for r in h.runs:
                r.font.size = Pt(8)
                r.font.color.rgb = RGBColor.from_string(MUTED)
            f = section.footer.paragraphs[0]
            f.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r = f.add_run('DRAFT  |  26 August 2026  |  ')
            r.font.size = Pt(8)
            r.font.color.rgb = RGBColor.from_string(MUTED)
            fld = OxmlElement('w:fldSimple')
            fld.set(qn('w:instr'), 'PAGE')
            f._p.append(fld)

    @staticmethod
    def shade(cell, fill):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = tcPr.find(qn('w:shd'))
        if shd is None:
            shd = OxmlElement('w:shd')
            tcPr.append(shd)
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:fill'), fill)

    @staticmethod
    def set_cell(cell, text, bold=False, color='1E293B', size=7.6):
        cell.text = ''
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.keep_together = True
        r = p.add_run(str(text))
        r.bold = bold
        r.font.name = 'Aptos'
        r.font.size = Pt(size)
        r.font.color.rgb = RGBColor.from_string(color)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    def para(self, text='', style=None, after=4):
        p = self.doc.add_paragraph(text, style=style)
        p.paragraph_format.space_after = Pt(after)
        return p

    def bullet(self, text, level=0):
        style = 'List Bullet' if level == 0 else 'List Bullet 2'
        p = self.doc.add_paragraph(style=style)
        p.paragraph_format.space_after = Pt(1.8)
        p.add_run(text)
        return p

    def number(self, text):
        p = self.doc.add_paragraph(style='List Number')
        p.paragraph_format.space_after = Pt(2)
        p.add_run(text)
        return p

    def callout(self, text, fill=PALE_TEAL, color=TEAL):
        t = self.doc.add_table(rows=1, cols=1)
        c = t.cell(0,0)
        self.shade(c, fill)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(text)
        r.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor.from_string(color)
        self.doc.add_paragraph().paragraph_format.space_after = Pt(0)

    def table(self, headers, rows, widths=None, size=7.5, header_fill=PALE_BLUE):
        t = self.doc.add_table(rows=1, cols=len(headers))
        t.style = 'Table Grid'
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i,h in enumerate(headers):
            self.shade(t.rows[0].cells[i], header_fill)
            self.set_cell(t.rows[0].cells[i], h, True, '1E3A8A', size)
        trPr = t.rows[0]._tr.get_or_add_trPr()
        rep = OxmlElement('w:tblHeader')
        rep.set(qn('w:val'), 'true')
        trPr.append(rep)
        for ridx,row in enumerate(rows):
            cells = t.add_row().cells
            trPr = t.rows[-1]._tr.get_or_add_trPr()
            no_split = OxmlElement('w:cantSplit')
            trPr.append(no_split)
            for i,v in enumerate(row):
                if ridx % 2:
                    self.shade(cells[i], LIGHT)
                self.set_cell(cells[i], v, False, '1E293B', size)
        if widths:
            for row in t.rows:
                for i,w in enumerate(widths):
                    row.cells[i].width = Inches(w)
        self.doc.add_paragraph().paragraph_format.space_after = Pt(0)
        return t

    def image(self, path, width, caption):
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(2)
        p.add_run().add_picture(str(path), width=Inches(width))
        cap = self.doc.add_paragraph(caption)
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap.paragraph_format.space_after = Pt(4)
        for r in cap.runs:
            r.font.size = Pt(8)
            r.font.italic = True
            r.font.color.rgb = RGBColor.from_string(MUTED)

    def page(self):
        self.doc.add_page_break()

    def landscape(self):
        s = self.doc.add_section(WD_SECTION.NEW_PAGE)
        s.orientation = WD_ORIENT.LANDSCAPE
        s.page_width = Inches(11)
        s.page_height = Inches(8.5)
        s.top_margin = Inches(.48)
        s.bottom_margin = Inches(.48)
        s.left_margin = Inches(.55)
        s.right_margin = Inches(.55)
        return s

    def portrait(self):
        s = self.doc.add_section(WD_SECTION.NEW_PAGE)
        s.orientation = WD_ORIENT.PORTRAIT
        s.page_width = Inches(8.5)
        s.page_height = Inches(11)
        s.top_margin = Inches(.62)
        s.bottom_margin = Inches(.62)
        s.left_margin = Inches(.68)
        s.right_margin = Inches(.68)
        return s

    def toc(self, entries):
        self.para('Contents', 'Heading 1')
        self.para('Static page index for the issued document. Section headings remain available in the Word navigation pane.')
        self.table(['Section','Page'], entries, [6.2,.8], 9.0)

    def cover(self, label, subtitle, fields, toc_entries):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(38)
        r = p.add_run(label)
        r.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = RGBColor.from_string(TEAL)
        self.doc.add_paragraph(self.title, style='Title')
        p = self.para(subtitle)
        for r in p.runs:
            r.font.size = Pt(13)
            r.font.color.rgb = RGBColor.from_string(MUTED)
        self.table(['Field','Value'], fields, [1.65,5.4], 8.2)
        self.page()
        self.toc(toc_entries)
        self.page()

    def save(self, path: Path, keywords: str):
        cp = self.doc.core_properties
        cp.title = self.title
        cp.subject = self.subject
        cp.author = 'Vibe-Code Planning Team'
        cp.keywords = keywords
        cp.created = datetime(2026,8,26,tzinfo=timezone.utc)
        cp.modified = NOW
        self.doc.save(path)


# ---------------- PRD data ----------------
actors = [
    ('ACT-01','Vibe user','Creates and continues owned projects through conversation, review and preview.','Own projects only; no infrastructure or privileged administration.'),
    ('ACT-02','Platform administrator','Supports users, reviews safe operational evidence and performs bounded recovery.','Restricted AdminJS; explicit permissions; break-glass actions are audited.'),
]

features = [
    ('F-01','Authentication and project ownership','Sign in and access only owned projects.'),
    ('F-02','Project creation from fixed template','Start from one approved generated-application template.'),
    ('F-03','Chatbot sessions','Create, continue, archive and branch durable conversations.'),
    ('F-04','Streamed progress and Stop','See plain-language progress and stop an active run.'),
    ('F-05','Generated file changes','Agent proposes bounded project-file changes.'),
    ('F-06','Diff, accept and reject','Review changes before accepting, rejecting or iterating.'),
    ('F-07','Whole-app versions and recovery','Initialize each project with Version 1, autosave a rolling recovery state, save numbered code-and-data versions and restore either as one verified unit.'),
    ('F-08','Live frontend preview','Eligible frontend edits update automatically in the authenticated preview.'),
    ('F-09','Automatic backend refresh','Supported backend source edits reload automatically; the UI reconnects safely.'),
    ('F-10','Preview lifecycle','Start, stop, resume and idle-stop previews without losing durable work.'),
    ('F-11','Queue, limits and recovery','Show Waiting, Recovering and Action required rather than generic failure.'),
    ('F-12','Restricted operations','Permit only audited, bounded administrative recovery and configuration.'),
    ('F-13','Cline session continuity and recovery','Refreshing the web page reconnects to the same run; a stalled Cline session can be terminated and recovered safely.'),
    ('F-14','Future-ready organization placement','Assign each POC user to the default Organization, OU and Department without changing owner-only project access.'),
]

epics = [
    ('EP-01','Identity and owned projects','Users can sign in, receive a default organization placement, create projects and see only their own work.','F-01, F-02, F-14'),
    ('EP-02','Conversational project sessions','Users can describe outcomes, retain context and control or recover an active run.','F-03, F-04, F-13'),
    ('EP-03','Generate, review and accept changes','Users can inspect a proposal before accepting, rejecting or iterating.','F-05, F-06'),
    ('EP-04','Mutable live preview','Accepted or proposed eligible source changes appear in an authenticated preview.','F-08, F-09'),
    ('EP-05','Checkpoints and recovery','Accepted work is recoverable and failed/rejected work cannot silently replace it.','F-07'),
    ('EP-06','Preview lifecycle and capacity','Preview resources start when needed, stop when idle and queue when capacity is full.','F-10, F-11'),
    ('EP-07','Security and governance','Ownership, isolation, least privilege, secrets and audit controls fail closed.','F-12'),
    ('EP-08','Guidance and operational clarity','Users understand status and recovery; administrators receive safe evidence.','F-11, F-12, F-13'),
]

stories = [
    ('US-01','EP-01','Must','As a Vibe user, I can sign in and receive a secure session.','AC-ID-01'),
    ('US-02','EP-01','Must','As a Vibe user, I can see and open only projects I own.','AC-ID-02'),
    ('US-03','EP-01','Must','As a Vibe user, I can create a project from the approved fixed template.','AC-PRJ-01'),
    ('US-04','EP-01','Must','As a Vibe user, I can rename, stop, reopen and soft-delete my project.','AC-PRJ-02'),
    ('US-05','EP-02','Must','As a Vibe user, I can create, continue and archive a chatbot session.','AC-CHAT-01'),
    ('US-06','EP-02','Must','As a Vibe user, I see streamed plain-language progress and can stop the run.','AC-CHAT-02'),
    ('US-07','EP-02','Must','As a Vibe user, I can branch from an earlier message without overwriting history.','AC-CHAT-03'),
    ('US-08','EP-02','Must','As a Vibe user, I am warned before context compaction and can continue without silent loss.','AC-CHAT-04'),
    ('US-09','EP-03','Must','As a Vibe user, I can inspect a safe diff and preview before accepting a change.','AC-CHG-01'),
    ('US-10','EP-03','Must','As a Vibe user, I can accept, reject or request another iteration.','AC-CHG-02'),
    ('US-11','EP-03','Must','As a Vibe user, destructive changes require explicit approval.','AC-CHG-03'),
    ('US-12','EP-04','Must','As a Vibe user, eligible frontend edits appear automatically in the current preview.','AC-PRV-01'),
    ('US-13','EP-04','Must','As a Vibe user, supported backend source edits refresh automatically and reconnect safely.','AC-PRV-02'),
    ('US-14','EP-04','Must','As a Vibe user, I am told when a change requires restart, rebuild or approval.','AC-PRV-03'),
    ('US-15','EP-05','Must','As a Vibe user, a new project starts at Version 1 and Save creates the next numbered whole-app version.','AC-VER-01, AC-VER-02'),
    ('US-16','EP-05','Must','As a Vibe user, I can restore code and project data together from a healthy saved version without erasing history.','AC-VER-04'),
    ('US-17','EP-06','Must','As a Vibe user, opening a stopped project resumes its preview without losing source.','AC-LFC-01'),
    ('US-18','EP-06','Must','As a Vibe user, capacity exhaustion shows Waiting instead of an unexplained error.','AC-CAP-01'),
    ('US-19','EP-06','Should','As a Vibe user, an inactive preview stops automatically and resumes explicitly.','AC-LFC-02'),
    ('US-20','EP-07','Must','As a Vibe user, my preview and live-update channel cannot be opened by another owner.','AC-SEC-01'),
    ('US-21','EP-07','Must','As a Platform administrator, I can inspect audited lifecycle and recovery evidence.','AC-OPS-01'),
    ('US-22','EP-07','Must','As a Platform administrator, I can change only allowlisted non-secret configuration.','AC-OPS-02'),
    ('US-23','EP-08','Should','As a Vibe user, I see plain-language statuses and actionable recovery choices.','AC-UX-01'),
    ('US-24','EP-08','Should','As a first-time user, I receive concise guidance and prompt starters.','AC-UX-02'),
    ('US-25','EP-02','Must','As a Vibe user, refreshing the web page reconnects to the current Cline run without starting a duplicate session.','AC-CHAT-05'),
    ('US-26','EP-02','Must','As a Vibe user, I can recover an unresponsive Cline session while preserving the last healthy project state.','AC-CHAT-06'),
    ('US-27','EP-01','Must','As a provisioned POC user, I am assigned to the default Organization, OU and Department while my project list remains owner-only.','AC-ID-03'),
    ('US-28','EP-05','Must','As a Vibe user, my active project has a rolling whole-app autosave for crash recovery without creating a new visible version every interval.','AC-VER-03'),
]

UC_BUSINESS_RULES = {
    'UC-01': ('BR-01', 'BR-02', 'BR-17'),
    'UC-02': ('BR-01', 'BR-03', 'BR-13', 'BR-17'),
    'UC-03': ('BR-01', 'BR-04', 'BR-05', 'BR-06', 'BR-13', 'BR-17'),
    'UC-04': ('BR-01', 'BR-04', 'BR-07', 'BR-09', 'BR-13', 'BR-17'),
    'UC-05': ('BR-01', 'BR-04', 'BR-08', 'BR-09', 'BR-13', 'BR-17'),
    'UC-06': ('BR-01', 'BR-04', 'BR-10', 'BR-11', 'BR-12', 'BR-13', 'BR-17'),
    'UC-07': ('BR-01', 'BR-04', 'BR-13', 'BR-14', 'BR-15', 'BR-17'),
    'UC-08': ('BR-01', 'BR-04', 'BR-06', 'BR-12', 'BR-13', 'BR-16', 'BR-17'),
    'UC-09': ('BR-01', 'BR-13', 'BR-16', 'BR-17'),
}

business_rules = [
    ('BR-01','Every project, conversation, run, snapshot, route and command is owner-scoped; direct-ID and cross-owner access fail closed.','UC-01–UC-09','Object authorization on every command/query plus cross-owner negative tests.'),
    ('BR-02','Organization, OU and Department placement is administrative metadata and never grants Stage 1 project access.','UC-01','Provision one current primary membership; same-unit users still fail project access tests.'),
    ('BR-03','Project creation is idempotent and pins exactly one approved immutable template version.','UC-02','Unique idempotency key, atomic project/workspace initialization and retry evidence.'),
    ('BR-04','Only one mutating operation may hold the project mutation lease; terminal attempts are immutable.','UC-03–UC-08','Lease/CAS checks and linked retry rather than in-place history rewrite.'),
    ('BR-05','Browser refresh reconnects to the existing run and persisted event stream; it never creates a duplicate run, Job or lease.','UC-03','Run identity and event cursor survive reconnect; duplicate-creation probes fail.'),
    ('BR-06','Stalled-agent recovery starts a linked retry only after prior Job/Pod deletion and failed write-probe evidence proves the old attempt cannot write.','UC-03, UC-08','UID-bound runtime-resource evidence; metadata lease expiry alone is rejected.'),
    ('BR-07','A proposal becomes accepted only through one explicit durable decision; destructive changes require explicit approval.','UC-04','Decision idempotency, approval gate and pre-change checkpoint evidence.'),
    ('BR-08','Mutable preview is owner-only and uses a one-time short-lived asymmetric JWT with exact-host audience, exchanged for a short-lived host-only cookie.','UC-05','Active kid/signature, exact claims, hashed-jti CAS consume and replay/revocation tests.'),
    ('BR-09','Changes are classified honestly as hot update, reload, restart, rebuild, approval required or unsupported.','UC-04, UC-05','Change matrix tests; unsupported changes never report Ready/HMR.'),
    ('BR-10','Project initialization creates Version 1; autosave is rotating recovery only; User Save and successful Restore create immutable numbered whole-app history.','UC-06','Kind/number/idempotency checks across Git, SQLite and snapshot records.'),
    ('BR-11','A whole-app snapshot is AVAILABLE only when Git, SQLite online backup, migration/lock/template manifests and hashes verify together.','UC-06','No raw live SQLite copy; partial or mismatched capture remains unavailable.'),
    ('BR-12','Restore is owner-only, serialized, idempotent and generation-staged; current pointers switch atomically only after code, DB, lock and health verification.','UC-06, UC-08','Before-images, provisional result, journal decision and pointer CAS; failure protects prior current state.'),
    ('BR-13','Only transitions listed in the authoritative state contract are legal; unknown or unlisted transitions fail closed and are audited.','UC-02–UC-09','State/edge equality against Data Specification plus illegal-edge mutation tests.'),
    ('BR-14','Capacity admission is durable, bounded and fair: owner round-robin across owners, FIFO within owner, explicit slot charges and deterministic timeout/cancel.','UC-07','CAPACITY_POOL row lock, queue sequence, durable cursor and slot-sum tests.'),
    ('BR-15','Preview idle-stop never interrupts an active mutation and never deletes durable source, SQLite data or checkpoints.','UC-07','Lease check, route withdrawal, scale-to-zero and resume evidence.'),
    ('BR-16','Administrative changes are restricted to named commands and typed allowlisted non-secret configuration; no shell or direct state forgery.','UC-08, UC-09','Permission, type/range, stale-write, secret-canary and immutable audit tests.'),
    ('BR-17','Stage 1 has no Deploy button, published release, stable application URL, registry promotion or production rollback.','UC-01–UC-09','UI/content scan and route/runtime acceptance tests reject hidden release behavior.'),
]
business_rules = [
    (rule_id, rule, ', '.join(uc for uc, rules in UC_BUSINESS_RULES.items() if rule_id in rules), evidence)
    for rule_id, rule, _legacy_scope, evidence in business_rules
]

use_cases = [
    ('UC-01','Sign in and view projects','ACT-01','M01 — Identity, organization placement and owned projects','Presentation + Application + Domain','Active account with one current primary default organization membership.','Sign in → validate organization placement → create secure session → show owner-scoped project list.','BR-01, BR-02, BR-17','Secure session, CSRF, session rotation and object-level owner authorization.','Invalid membership fails provisioning closed; same OU/Department never reveals another owner’s project.','Authenticated session; default Organization/OU/Department exists; no cross-owner project visibility.','US-01, US-02, US-27'),
    ('UC-02','Create and open project','ACT-01','M02 — Project and immutable template registry','Presentation + Application + Domain + Runtime','Authenticated owner; approved template available; idempotency key allocated.','Enter project name and goal → create once → pin template → initialize Git/SQLite/workspace → open project.','BR-01, BR-03, BR-13, BR-17','Owner command, idempotency uniqueness, template allowlist and path-safe initialization.','Initialization remains recoverable and retry returns the same project rather than creating a duplicate.','Project owns one template lock, generated repository, workspace and preview identity.','US-03, US-04'),
    ('UC-03','Use and recover Vibe Code chatbot','ACT-01','M03 — Conversation, Agent adapter and Agent Job','Presentation + Application + Runtime','Owned project; durable chat session exists or may be created; mutation slot is available or queueable.','Open prompt guidance → submit outcome → persist message → admit one run → stream progress → reconnect or recover → show result and usage.','BR-01, BR-04, BR-05, BR-06, BR-13, BR-17','Owner/session checks, one mutation lease, admitted Job identity, heartbeat and UID-bound old-writer revocation.','Refresh reconnects without duplication. A stale heartbeat exposes recovery; uncertain old writers keep mutation blocked.','Ordered messages/events survive reconnect; linked attempts and runtime evidence remain immutable and attributable.','US-05–US-08, US-24–US-26'),
    ('UC-04','Review and decide a change','ACT-01','M04 — Change proposal, diff, approval and decision','Presentation + Application + Domain','Agent produced a validated proposal tied to the owned project and run.','Open safe diff and preview → inspect impact → accept, reject or iterate → persist one decision → checkpoint accepted result.','BR-01, BR-04, BR-07, BR-09, BR-13, BR-17','Owner authorization, safe path/diff redaction, explicit destructive approval and decision idempotency.','Rejected or failed work never replaces the last healthy accepted version; approval denial is terminal evidence.','Decision, audit evidence, run and accepted checkpoint are linked exactly once.','US-09–US-11'),
    ('UC-05','Observe live preview update','ACT-01','M05 — Preview Gateway, route and mutable Preview Pod','Presentation + Gateway + Runtime','Owned preview is admitted and healthy; proposed edit is eligible for the classified update path.','Request owner preview → exchange the one-time exact-audience JWT → load preview → apply eligible source change → HMR/reload or bounded restart → confirm Ready for review.','BR-01, BR-08, BR-09, BR-13, BR-17','Active kid/signature, exact JWT claims, hashed-jti CAS consume, host-only cookie and fixed-target routing.','Unsupported change is classified as restart, rebuild, approval required or unsupported; stale/replayed/cross-owner route is denied.','Frontend/backend update and WSS paths use the authenticated owner route with safe readiness evidence.','US-12–US-14, US-20'),
    ('UC-06','Save and restore a whole-app version','ACT-01','M06 — Git workspace, snapshot and Restore transaction','Presentation + Application + Persistence + Runtime','Owned project; no conflicting mutation; target snapshot is AVAILABLE and same-project for Restore.','Capture rotating autosave when changed → user selects Save → verify next numbered whole-app version → select earlier version → stage and verify Restore → atomically switch current pointers.','BR-01, BR-04, BR-10, BR-11, BR-12, BR-13, BR-17','Owner-only idempotent command, project lease, path-safe Git, SQLite online backup, hashes, before-images and pointer CAS.','No-change autosave is skipped. Any code, DB, lock, hash, health or cutover failure protects the previous current generation and enters verified rollback/manual intervention.','Manifest binds Git, SQLite, migration/lock/template metadata and actor; successful Restore creates new history and healthy preview.','US-15, US-16, US-28'),
    ('UC-07','Stop, idle-stop and resume','ACT-01','M07 — Workspace lifecycle and capacity admission','Presentation + Application + Runtime','Owned active, queued or stopped project; current mutation and capacity state are known.','Warn on idle → verify no active mutation → withdraw route and stop → retain durable workspace → owner opens project → fair admission → readiness confirmed.','BR-01, BR-04, BR-13, BR-14, BR-15, BR-17','Owner command, pool row lock, slot/reserve checks, durable fairness cursor, lease check and generation-safe route.','Active mutation blocks idle-stop; full queue shows bounded Waiting/cancel/timeout; failed start/stop remains recoverable.','Source, SQLite and checkpoints persist while disposable runtime is stopped or recreated.','US-17–US-19'),
    ('UC-08','Recover a failed operation','ACT-01\nACT-02','M08 — Recovery controller, errors, evidence and audit','Presentation + Application + Runtime + Operations','Stable safe error/reference exists; actor owns the project or has explicit bounded administrator permission.','Inspect safe status → choose retry, revert, recover agent or bounded reconcile → validate prior terminal state and proof → create linked attempt → verify healthy result.','BR-01, BR-04, BR-06, BR-12, BR-13, BR-16, BR-17','Owner or explicit admin permission, idempotency, prior-attempt UID proof, mutation serialization and redacted audit.','No automatic replay of uncertain mutation. Missing or mismatched proof fails closed and keeps route withdrawn/mutation blocked.','New linked attempt and correlation evidence exist; prior terminal history and last healthy version remain unchanged.','US-21, US-23, US-26'),
    ('UC-09','Perform restricted administration','ACT-02','M09 — Restricted administration and typed configuration','Presentation + Application + Operations','Active administrator session; explicit role/permission; expected aggregate row version loaded.','Inspect safe state → choose named allowlisted command or config key → validate permission/type/range/version → commit CAS and audit → show safe result.','BR-01, BR-13, BR-16, BR-17','Restricted AdminJS, explicit permission, last-admin protection, typed non-secret allowlist and atomic CAS plus command dedup.','Unknown command/key, secret value, wrong type/range, stale row or direct-state attempt is rejected atomically and audited.','Accepted and denied privileged actions are attributable, redacted and reviewable.','US-21, US-22'),
]
# Business Rule references are derived from one canonical UC mapping; tuple literals above
# remain descriptive source data only and cannot override the released mapping.
use_cases = [
    use_case[:7] + (', '.join(UC_BUSINESS_RULES[use_case[0]]),) + use_case[8:]
    for use_case in use_cases
]

security_baseline = [
    ('SEC-01','Identity, organization and ownership','Every user has one current primary organization placement, while every project, chat, run, checkpoint and preview remains owner-scoped in Stage 1.','Default membership validation plus same/cross-Department negative tests for list, direct URL, API, preview and live update.'),
    ('SEC-02','Authenticated preview','Mutable previews are private and require the owner’s active session.','Anonymous, expired-session and wrong-owner requests fail closed.'),
    ('SEC-03','Project isolation','One project cannot read or change another project’s files, data or runtime.','Cross-project filesystem, service and route tests are denied, including absolute/parent paths, symlink ancestors/leaves, broken links, special nodes and concurrent path-swap probes before Git/snapshot/restore reads.'),
    ('SEC-04','Least privilege','Generated workloads run without privileged host access or unrestricted administration.','Isolation and privilege-escape negative tests pass.'),
    ('SEC-05','Resource protection','Each project has bounded CPU, memory, process and storage use; capacity is admitted.','Limit and exhaustion tests show Waiting or controlled failure.'),
    ('SEC-06','Network control','Preview routes expose only intended application traffic; internal services remain private.','No database/debug host exposure; blocked private/metadata destinations.'),
    ('SEC-07','Secret handling','Credentials never enter generated source, chat display, logs, exports or non-secret configuration.','Secret scans and redaction tests pass with synthetic canaries.'),
    ('SEC-08','Change and recovery safety','Destructive/unsupported changes require approval; stalled-session recovery cannot overlap two mutating Cline attempts; snapshot capture and restore run under one project mutation lock.','Approval, proven old-Pod write revocation, generation staging, journaled pointer cutover and crash recovery preserve the last healthy checkpoint; metadata lease fencing alone is insufficient.'),
    ('SEC-09','Audit and evidence','Privileged actions, security decisions and lifecycle changes have actor, action, result and correlation reference.','Accepted, denied and failed actions appear in immutable audit evidence.'),
    ('SEC-10','Retention and recovery','Project source, rolling autosave and numbered whole-app versions survive ordinary stop/restart; SQLite is captured through a consistent backup operation, not by copying a live database file.','Code+database restore drill passes; autosave/version retention remains an approved configurable policy.'),
]

acceptance = [
    ('AC-ID-01','Login/logout, secure session, CSRF and session rotation pass.'),
    ('AC-ID-02','User A cannot list, open, run, preview or delete User B’s project.'),
    ('AC-PRJ-01','A duplicate create request returns one project pinned to the approved template.'),
    ('AC-PRJ-02','Rename, stop, reopen and soft-delete show correct durable state.'),
    ('AC-CHAT-01','Chat sessions can be created, continued and archived without deleting history.'),
    ('AC-CHAT-02','Progress streams in plain language; Stop prevents new agent actions and reaches a safe terminal state.'),
    ('AC-CHAT-03','Editing a prior message creates traceable branch lineage.'),
    ('AC-CHAT-04','Pre-compaction notice and persisted summary prevent silent context loss.'),
    ('AC-CHG-01','Diff and preview are safe, owner-scoped and tied to the proposed run.'),
    ('AC-CHG-02','Accept, reject and iterate produce one explicit durable decision.'),
    ('AC-CHG-03','A destructive change cannot proceed without explicit user approval.'),
    ('AC-PRV-01','A React/TypeScript/CSS edit updates the current preview through live frontend hot reload without deployment.'),
    ('AC-PRV-02','A supported Python source edit reloads the backend preview and the product reconnects safely.'),
    ('AC-PRV-03','Dependency, environment, migration and runtime-contract changes show the correct restart/rebuild/approval state.'),
    ('AC-VER-01','Project creation initializes Git on main, creates an initial commit and exposes Version 1 only after a matching SQLite backup and manifest are verified.'),
    ('AC-VER-02','Each successful Save creates exactly one next numbered immutable version linked to a Git commit, private SQLite backup, hashes, locks and actor; an idempotent retry creates no duplicate version.'),
    ('AC-VER-03','While a project is active, the configured interval captures a rotating whole-app recovery autosave only when state changed; autosaves do not increment or clutter the user-visible version number.'),
    ('AC-VER-04','Before any code/DB mutation, idempotent Restore atomically captures PRE_RESTORE, prior generation/snapshot/route generation/exact route status, target generation and one provisional RESTORE_RESULT. Early or late rollback marks that result INVALID; cutover or verified rollback uses one PostgreSQL pointer transaction. Failure keeps route withdrawn and mutation blocked.'),
    ('AC-LFC-01','Stopping/recreating a preview does not lose Git source, checkpoints or approved persistent data.'),
    ('AC-CAP-01','Capacity exhaustion creates a durable bounded Waiting request with retry/cancel/timeout, owner-round-robin then FIFO order and separate Preview/Agent reservations; it never becomes login failure, generic 500 or silent oversubscription.'),
    ('AC-LFC-02','Idle stop does not interrupt an active mutation and requires explicit resume.'),
    ('AC-SEC-01','Owner, anonymous and cross-owner preview/live-update authorization tests pass.'),
    ('AC-OPS-01','Injected startup, agent, preview and restore failures produce safe recovery and correlation evidence.'),
    ('AC-OPS-02','Only typed allowlisted non-secret configuration can be changed; stale and invalid writes fail atomically.'),
    ('AC-UX-01','Normal users never see raw commands, ports, hostnames, stack traces, infrastructure manifests or credentials.'),
    ('AC-UX-02','Prompt starters and a concise first-run guide are available.'),
    ('AC-CHAT-05','Refreshing the browser during an active run reconnects to persisted messages/events and does not create a second AGENT_RUN, Agent Job or mutation lease.'),
    ('AC-CHAT-06','After the approved heartbeat threshold, Recover agent session persists one unique Agent Job resource reference whose agent_run_id equals the stopped/prior attempt, plus old Job/Pod UID, deletion observation and failed-closed write probe. STOPPED/retry reject missing, duplicate or other-attempt proof; lease expiry never authorizes overlap.'),
    ('AC-ID-03','Account provisioning creates one current primary Default Organization → Default OU → Default Department membership; users in the same unit still see only projects they own.'),
]

stage_rows = [
    ('Stage 1 — POC','Chatbot, generated edits, diff/accept/reject, rolling autosave, numbered whole-app Save/Restore, owner-only live preview, automatic frontend/backend refresh, preview lifecycle, queue and idle stop.','Deploy button; immutable production release; stable published application URL; promotion/rollback; multi-node HA.'),
    ('Stage 2 — release lifecycle','Deploy button, controlled immutable release, stable application route, health gate, promotion and rollback.','Only starts after separate scope, security, operating-model and capacity approval.'),
]

delivery = [
    ('W1','24–28 Aug','Runtime feasibility, Cline integration/recovery gate, authenticated preview route and security spike evidence.'),
    ('W2','31 Aug–4 Sep','Authentication, project ownership, fixed template, durable project/workspace metadata and restricted administration foundation.'),
    ('W3','7–11 Sep','Chat sessions, streamed events, mutation control, Agent Job and workspace persistence.'),
    ('W4','14–18 Sep','Owner preview gateway, frontend hot reload, backend source reload and lifecycle reconciliation.'),
    ('W5','21–25 Sep','Diff/review, accept/reject, Git workspace autosave, whole-app Save/Restore and failure recovery.'),
    ('W6','28 Sep–2 Oct','Integrated beta by 30 Sep; security baseline, load evidence and remediation continue through 2 Oct.'),
    ('W7','5–9 Oct','Contingency close: operational guidance, final evidence, traceability and POC acceptance.'),
]

risks = [
    ('Runtime details not proven','High / High','Time-box runtime, storage, live update, routing and isolation spikes before locking the implementation.'),
    ('Cline SDK compatibility','High / High','Pin 0.0.79 and prove events, cancel, reconnect, isolation and file mutation early.'),
    ('Cline session stalls','Medium / High','Persist events, monitor heartbeat/lease, expose bounded Recover agent session and prove no overlapping mutation.'),
    ('Live update overstated','Medium / High','Use explicit classification; dependency, data and runtime changes must show restart/rebuild/approval honestly.'),
    ('Single-node outage','Medium / High','State clearly: process recovery only; durable backup and restore evidence required.'),
    ('Runtime isolation','Medium / Critical','Apply the Technical Design isolation baseline and verify cross-project/privilege-escape tests.'),
    ('Mutable preview conflict','High / High','Keep each mutable preview at zero or one active instance.'),
    ('Code/data version skew','High / High','Bind Git commit and SQLite online backup in one verified manifest; expose no version until all parts validate.'),
    ('Deadline pressure','High / High','Prioritize the 30 Sep integrated beta; use 9 Oct only as contingency.'),
]

decisions = [
    ('DEC-01','Stage 1 orchestration','Single-node K3s','CONFIRMED'),
    ('DEC-02','Stage 1 runtime detail','Single-node K3s; container/runtime mechanics are owned by the Technical Design','CONFIRMED'),
    ('DEC-03','Preview routing','Authenticated preview gateway; exact routing technology is owned by the Technical Design','CONFIRMED'),
    ('DEC-04','Generated-app template','One approved fixed generated-application template; exact stack is owned by the Technical Design','CONFIRMED'),
    ('DEC-05','Published deployment','Deploy button and release lifecycle start in Stage 2','CONFIRMED'),
    ('DEC-06','Live update','Automatic eligible frontend/backend refresh and honest restart/rebuild classification are mandatory','CONFIRMED'),
    ('DEC-07','Cline SDK target','Pin @cline/sdk@0.0.79; compatibility evidence required','CONFIRMED TARGET'),
    ('DEC-08','Namespace/storage/network implementation','Defined by Technical Design spikes','DECISION REQUIRED'),
    ('DEC-09','Retention and backup policy','Per data class; no invented duration','DECISION REQUIRED'),
    ('DEC-10','Launch admission limits','Start conservatively; raise only after load evidence','PROVISIONAL'),
    ('DEC-11','Cline stuck-session recovery','Page refresh reconnects only; watchdog/owner cancels and proves the old Job/Pod can no longer write before an in-place linked retry. Lease fencing metadata alone is insufficient; generation-staged retries require controller-only CAS promotion.','CONFIRMED DIRECTION'),
    ('DEC-12','Organization hierarchy','Seed Default Organization → Default OU → Default Department; Stage 1 project access remains owner-only and organization placement grants no project access','CONFIRMED DIRECTION'),
    ('DEC-13','Whole-app version model','main mirrors the latest saved version; workspace/<project-id> holds active work; interval autosave is rotating recovery state; Save creates a numbered immutable code+SQLite version; Restore creates new history','CONFIRMED DIRECTION'),
]

VIBE_LAYOUT = DIAGRAM_DIR / 'vibe-page-layout.png'
UI_COMPONENT_CONTRACT = DIAGRAM_DIR / 'vibe-page-component-contract.png'
create_vibe_layout(VIBE_LAYOUT)
create_ui_component_contract(UI_COMPONENT_CONTRACT)

UC_FLOW_PATHS = {}
for use_case in use_cases:
    uid = use_case[0]
    path = DIAGRAM_DIR / f'{uid.lower()}-flow.png'
    create_use_case_flow(path, use_case)
    UC_FLOW_PATHS[uid] = path

STATE_DIAGRAM_PATHS = []
for index, machine in enumerate(STATE_CONTRACT['machines'], 1):
    path = DIAGRAM_DIR / f"state-{index:02d}-{machine['field'].lower().replace('.', '-')}.png"
    create_state_diagram(path, machine)
    STATE_DIAGRAM_PATHS.append((machine, path))

TRANSACTION_DIAGRAMS = [
    ('User Save', DIAGRAM_DIR / 'tx-user-save.png', ['User','NestJS','Worker','Git + SQLite','PostgreSQL','Preview'], [
        (0,1,'Save command'),(1,2,'idempotency + lease'),(2,3,'commit + online backup'),(3,2,'hashes + manifests'),(2,4,'AVAILABLE + next version CAS'),(4,5,'resume current generation'),(5,0,'saved version visible')],
     'Any capture/hash/manifest/health failure creates no AVAILABLE snapshot and does not advance current_snapshot_id or version number.'),
    ('Whole-app Restore', DIAGRAM_DIR / 'tx-restore.png', ['User','NestJS','Worker','Staging generation','PostgreSQL','Preview'], [
        (0,1,'Restore target'),(1,2,'owner + idempotency + lease'),(2,5,'withdraw route / quiesce'),(2,3,'materialize code + SQLite'),(3,2,'verify hashes/locks/health'),(2,4,'journaled pointer CAS'),(4,5,'activate exact generation'),(5,0,'verified result')],
     'Before-images and one provisional RESTORE_RESULT exist before mutation. Failure restores verified prior pointers or enters MANUAL_INTERVENTION with route withdrawn.'),
    ('Owner-only Preview Authentication', DIAGRAM_DIR / 'tx-preview-auth.png', ['Browser','NestJS','Redis','PostgreSQL','Preview Gateway','Preview Pod'], [
        (0,1,'request with opaque cookie + CSRF'),(1,2,'validate session + version + expiry'),(1,3,'persist jti/kid/aud/generation grant'),(1,0,'return one-time JWT <=60s'),(0,4,'exchange at exact preview host'),(4,3,'CAS ISSUED jti to ACTIVE cookie hash'),(4,0,'set host-only preview cookie'),(4,5,'strip platform credentials; proxy')],
     'Replay, wrong kid/audience/host/owner/session/project/generation, expiry or revocation fails closed. Redis session loss denies JWT issue; PostgreSQL remains replay authority.'),
    ('Durable Capacity Admission', DIAGRAM_DIR / 'tx-capacity-admission.png', ['User','NestJS','Capacity Worker','CAPACITY_POOL','Runtime Controller','K3s'], [
        (0,1,'start preview / run'),(1,2,'durable enqueue'),(2,3,'row lock + owner round-robin/FIFO'),(3,2,'reserve explicit slots'),(2,4,'admitted command'),(4,5,'create component resources'),(5,4,'UID/readiness observation'),(4,1,'READY or bounded failure')],
     'Queue depth, timeout, reserve and slot-sum limits are deterministic. Admission never oversubscribes silently or turns capacity exhaustion into login/generic failure.'),
]
for title, path, lanes, steps, guard in TRANSACTION_DIAGRAMS:
    create_sequence_diagram(path, title, lanes, steps, guard)

# ---------------- build PRD ----------------
prd = DocBuilder(
    'Vibe-Code Platform\nNo-Code POC PRD & Delivery Plan',
    'VIBE-CODE  /  POC PRD & DELIVERY PLAN',
    'Product scope, feature list, epics, stories, use cases, security baseline, acceptance and delivery plan'
)
prd.cover(
    'PRODUCT REQUIREMENTS DOCUMENT',
    'A business-readable definition of the Stage 1 conversational application-builder POC. Technical implementation is specified separately.',
    [('Status','DRAFT'),('Boundary','Stage 1 POC'),('Version','1.1'),('POC deadline','30 September 2026'),('Tentative contingency','9 October 2026'),('Related technical document','vibe-code-poc-technical-design-and-runtime-architecture.docx')],
    [('1. Executive summary','3'),('2. Scope and stage boundary','4'),('3. Actors','5'),('4. Feature list','6'),('5. Epic table','7'),('6. User Story table','8'),('7. Use Case summary and details','10'),('8. Security baseline','16'),('9. Acceptance criteria','17'),('10. Delivery Plan','18'),('11. Risks and decisions','19'),('12. Traceability and related artifacts','20')]
)

prd.para('1. Executive summary','Heading 1')
prd.para('Vibe Code lets a non-technical user create and improve a small application by describing business outcomes in a chatbot. The user reviews proposed changes, sees an authenticated live preview, accepts or rejects the result, saves numbered whole-app versions and can restore an earlier healthy version. Source code, runtime operations and infrastructure details remain behind the product boundary.')
prd.callout('Stage 1 proves the complete create → chat → change → review → live preview → accept/reject → recover journey. It does not include a Deploy button or production release lifecycle.', PALE_TEAL, TEAL)
prd.para('Product outcomes','Heading 2')
prd.table(['Outcome','Evidence'],[
    ('Clear no-code journey','A user completes the primary journey without seeing an IDE, terminal or infrastructure configuration.'),
    ('Controlled change','Every proposed change has visible progress, review, explicit decision and recoverable history.'),
    ('Real live preview','Eligible frontend and backend source edits update automatically; unsupported changes are classified honestly.'),
    ('Safe multi-user boundary','Project, chat, run, checkpoint and preview ownership tests fail closed.'),
    ('Operational honesty','Single-node K3s is described as process-level recovery, not hardware high availability.'),
], [2.0,5.0], 8)
prd.image(PRODUCT_FLOW, 7.05, 'Figure 1 — Stage 1 product journey and the boundary between mutable preview and Stage 2 deployment.')
prd.page()

prd.para('2. Scope and stage boundary','Heading 1')
prd.table(['Stage','Included','Excluded / gate'], stage_rows, [1.2,3.35,2.45], 7.5)
prd.para('Stage 1 product principles','Heading 2')
for item in [
    'Conversation is the primary interface; implementation detail is hidden from normal users.',
    'The fixed template reduces generation, preview, security and support variability.',
    'A mutable preview is not a deployed or published application.',
    'A saved version binds code and approved SQLite data as one recoverable whole; a running preview is disposable.',
    'Unknown, unsupported or unsafe changes fail closed with an actionable status.',
    'Platform restrictions are visible and explainable rather than silently enforced.',
]: prd.bullet(item)
prd.para('Explicit Stage 1 non-goals','Heading 2')
for item in [
    'No Deploy button, production release publication, stable application subdomain, promotion or rollback.',
    'No user-supplied infrastructure definition or privileged host access.',
    'No custom admin portal; restricted AdminJS is a developer/break-glass tool only.',
    'No multi-node high availability, service mesh, GitOps controller, custom operator or CRD.',
    'No simultaneous multi-user editing of one project and no multi-replica mutable preview.',
]: prd.bullet(item)
prd.page()

prd.para('3. Actors','Heading 1')
prd.table(['ID','Actor','Goal / responsibility','Boundary'], actors, [.65,1.45,2.9,2.0], 7.6)
prd.para('Normal product boundary','Heading 2')
prd.table(['User sees','User does not see'],[
    ('Projects, chatbot, progress, diff, preview, checkpoints, statuses and recovery actions','Source editor, terminal, infrastructure objects, raw ports, hostnames, commands, stack traces, provider credentials'),
    ('Plain-language restart/rebuild status','Runtime commands or infrastructure definitions'),
    ('Usage/cost estimate and approval prompts','Provider master key or unrestricted model configuration'),
], [3.5,3.5], 8)
prd.para('Vibe page layout','Heading 2')
prd.para('The Vibe page is a deliberate two-pane product experience: durable conversation and change control on the left, authenticated generated-application preview on the right, with project/version status above. The picture is an interaction contract, not a promise of final visual styling.')
prd.image(VIBE_LAYOUT, 7.05, 'Figure 2 — Vibe page layout: metadata, conversation, progress and decisions beside the owner-only live preview.')
prd.table(['Region','Required content','Forbidden content'],[
    ('Project header','Project identity, owner-safe status, current numbered version, last Save, History and full-preview action.','Raw namespace, Pod, host, port, command or infrastructure state.'),
    ('Conversation pane','Durable messages/events, prompt guidance, progress, Stop, safe diff decision, Save and Restore.','Source editor, terminal, unrestricted tool output or provider credential.'),
    ('Preview pane','Owner-only generated UI, readiness/live-update state and fixed application/API routes.','Arbitrary upstream, platform session cookie or Deploy/publish control.'),
    ('Status and recovery','Waiting, Recovering, Action required, autosave-versus-version distinction and correlation reference.','Generic success when a required verification is incomplete.'),
], [1.35,3.5,2.15], 7.2)

prd.para('4. Feature list','Heading 1')
prd.para('This is the Stage 1 product feature catalogue. Technical implementation references are intentionally kept in the separate Technical Design.')
prd.table(['ID','Feature','User outcome'], features, [.62,2.15,4.15], 7.4)
prd.page()

prd.para('5. Epic table','Heading 1')
prd.table(['ID','Epic','Outcome','Features'], epics, [.55,1.8,3.65,1.0], 7.4)
prd.callout('Dependency direction: owned project → durable conversation → generated proposal → review/live preview → checkpoint/recovery → operational evidence.', PALE_BLUE, BLUE)
prd.page()

prd.para('6. User Story table','Heading 1')
prd.para('Stories are short, user-visible outcomes. Runtime resources, live-update configuration and command details are tracked in the Technical Design.')
prd.table(['ID','Epic','Priority','User Story','Acceptance'], stories, [.52,.55,.55,4.65,.72], 6.9)
prd.page()

prd.para('7. Use Case summary','Heading 1')
prd.table(['ID','Use Case','Actors','Stories'], [(x[0],x[1],x[2],x[11]) for x in use_cases], [.55,3.1,1.4,1.6], 7.4)
prd.callout('Each detailed Use Case is a ten-field executable product contract. Business Rules are canonical BR references, not hidden prose in the flow.', PALE_TEAL, TEAL)
prd.page()
prd.para('Detailed Use Cases','Heading 2')
for index, use_case in enumerate(use_cases):
    uid,title,actor,module,layer,pre,flow,rules,security_guard,failure,evidence,story_ref = use_case
    if index:
        prd.page()
    prd.para(f'{uid} — {title}','Heading 2')
    prd.table(['Field','Definition'],[
        ('Actors',actor),
        ('Module',module),
        ('Layer',layer),
        ('Preconditions',pre),
        ('Main flow',flow),
        ('Business Rules',rules),
        ('Security Guard',security_guard),
        ('Failure / Recovery',failure),
        ('Acceptance evidence',evidence),
        ('Story trace',story_ref),
    ], [1.55,5.45], 7.5)
    prd.image(UC_FLOW_PATHS[uid], 7.05, f'Figure {uid} — Main, alternative/failure and evidence flow for {title}.')

prd.page()
prd.para('Business Rule Register','Heading 1')
prd.para('Business Rules constrain every Use Case independently of screen wording. A Use Case may reference only IDs in this register; Technical realization must trace the same IDs rather than redefine them.')
prd.table(['ID','Rule','Applies to','Enforcement / evidence'], business_rules, [.65,3.75,1.05,1.55], 6.8)

prd.landscape()
prd.para('Authoritative state diagrams','Heading 1')
prd.callout('These diagrams are rendered from the repository state contract that must exactly match Data Specification sheets “Enum Catalogue” and “State Transitions”. Filled dots mark initial states; ringed dots mark terminal states. Every unlisted edge fails closed.', PALE_AMBER, '92400E')
for index, (machine, path) in enumerate(STATE_DIAGRAM_PATHS, 1):
    if index > 1:
        prd.page()
    prd.para(machine['field'],'Heading 2')
    image = Image.open(path)
    state_height_cap = 5.0 if index == 1 else 5.85
    display_width = min(9.7, state_height_cap * image.width / image.height)
    image.close()
    prd.image(path, display_width, f"Figure S-{index:02d} — {machine['field']}: {len(machine['states'])} states and {len(machine['transitions'])} legal directed transitions.")
prd.portrait()

prd.para('8. Security baseline','Heading 1')
prd.para('The baseline is written as observable controls. Detailed runtime, container and network implementation belongs in the Technical Design.')
prd.table(['ID','Control area','Baseline requirement','Acceptance evidence'], security_baseline, [.58,1.35,3.1,2.0], 6.9)
prd.callout('Any unknown route, ownership, quota, secret, startup or lifecycle condition fails closed. The platform must not guess a successful state.', PALE_RED, '991B1B')
prd.para('Product architecture boundary','Heading 2')
prd.table(['Concern','Approved Stage 1 boundary'],[
    ('Platform control plane','NestJS + Prisma + PostgreSQL. This is the product platform, not generated-app code.'),
    ('Restricted administration','AdminJS exposes only allowlisted properties and named NestJS commands behind the same RBAC/CSRF/owner checks; no unrestricted direct CRUD.'),
    ('Browser identity','An opaque HttpOnly Secure session cookie references Redis server-side session state. Long-lived browser JWT/localStorage authentication is prohibited; Redis session loss requires fresh sign-in.'),
    ('Preview/service exchange','Only short-lived, audience-bound asymmetric JWTs with kid/jti/replay/revocation controls may cross a service boundary; generated preview code never receives platform session/JWT credentials.'),
    ('Generated application','React/Vite + FastAPI + SQLite in the project workspace. It cannot access platform PostgreSQL or Redis.'),
], [2.0,5.0], 7.6)
prd.page()

prd.para('9. Acceptance criteria','Heading 1')
prd.table(['ID','Acceptance evidence'], acceptance, [.65,6.35], 6.9)
prd.para('Capacity acceptance assumptions','Heading 2')
prd.table(['Measure','Planning target','Launch posture'],[
    ('Registered users','100','Named accounts; not 100 simultaneous previews'),
    ('Concurrent active users','50','Validate representative control-plane activity'),
    ('Active previews','25 planning target','Begin with 10 admitted previews'),
    ('Concurrent agent runs','8 planning target','Begin with 3 admitted runs'),
    ('Peak platform API','75 RPS','Load-test hypothesis, not production SLA'),
    ('Idle stop','60 minutes candidate','Confirm from usability and capacity evidence'),
], [2.0,2.35,2.65], 7.3)
prd.para('The planning targets are not procurement proof. Measured results must record host profile, test script, workload mix, duration, latency, error rate and resource use.', after=5)
prd.page()

prd.para('10. Delivery Plan','Heading 1')
prd.table(['Week','Dates','Business-readable outcome'], delivery, [.55,1.05,5.4], 7.5)
prd.para('Planning envelope','Heading 2')
prd.table(['Item','Value','Meaning'],[
    ('Team','2 developers','Capacity assumption'),
    ('Weekly capacity','64 team SP','32 SP per developer per week'),
    ('30 Sep capacity','358.4 SP','Integrated beta target; not the full seven-week allocation'),
    ('Full planning envelope','448 SP','Seven team-weeks'),
    ('Contingency date','9 Oct 2026','Used only if final planned week is required'),
], [1.5,1.35,4.15], 7.8)
prd.para('Technical delivery enablers','Heading 2')
prd.table(['ID','Enabler','Business exit evidence'],[
    ('EN-01','Cline SDK compatibility and recovery gate','Pinned integration supports progress, stop, page reconnect, heartbeat, stuck detection, bounded termination, isolation and file mutation.'),
    ('EN-02','Preview runtime foundation','Project preview and agent work can start, stop, recover and preserve durable work.'),
    ('EN-03','Authenticated preview gateway','Owner route supports application, progress and live-update traffic.'),
    ('EN-04','Live-update classification','Frontend refresh, backend refresh and restart/rebuild paths are demonstrably distinct.'),
    ('EN-05','Security baseline','Ownership, isolation, least privilege, network, secret and audit tests pass.'),
    ('EN-06','Load and recovery evidence','Admission, idle stop, Pod failure, restart and restore evidence is reproducible.'),
], [.65,2.2,4.15], 7.5)
prd.callout('Detailed runtime resources, commands, library versions and spike procedures are defined in the Technical Design, not duplicated here.', PALE_TEAL, TEAL)
prd.page()

prd.para('11. Risks and decisions','Heading 1')
prd.para('Top risks','Heading 2')
prd.table(['Risk','Likelihood / impact','Response'], risks, [2.0,1.3,3.7], 6.8)
prd.para('Decision register','Heading 2')
prd.table(['ID','Decision area','Current decision','Status'], decisions, [.62,1.6,3.8,1.0], 6.4)
prd.para('Open product decisions','Heading 2')
for item in [
    'Approved model gateway/provider, residency and hard cost cap.',
    'Project, chat, audit, checkpoint, workspace and backup retention durations.',
    'Exact launch admission limits after measured runtime load evidence.',
    'Any future Stage 2 sharing or delegated authorization model requires a separate product and security decision; Stage 1 remains owner-only.',
    'Future team workspace and project membership roles; OU/Department placement must never become an implicit sharing rule.',
]: prd.bullet(item)

prd.para('12. Traceability and related artifacts','Heading 1')
prd.table(['Artifact','Authority'],[
    ('This PRD & Delivery Plan','Product scope, actors, features, Epics, Stories, Use Cases, security baseline and acceptance.'),
    ('Technical Design & Runtime Architecture DOCX','K3s/runtime implementation, live-update mechanics, versions, networking, security implementation and spike evidence.'),
    ('Data Specification XLSX','Entities, columns, relationships, enums, transitions, error registry and governance contract.'),
], [2.75,4.25], 7.8)
prd.para('Traceability rule','Heading 2')
prd.para('Technical sections and tests reference Feature, Story, Use Case, Security and Acceptance IDs from this document. They do not redefine the product requirement independently.')
prd.save(PRD_OUT, 'Vibe-Code, PRD, POC, chatbot, live preview, K3s, delivery plan')

# ---------------- Technical Design data ----------------
tech_decisions = [
    ('TD-DEC-01','Orchestration','Single-node K3s v1.36.3+k3s1, stable-channel release resolved and frozen 26 Aug 2026','CONFIRMED','Exact install artifact/checksum is pinned; no floating channel at deployment. Replacement-node upgrade/rollback only.'),
    ('TD-DEC-02','Container runtime','Bundled containerd v2.3.2-k3s2 + standard runc v1.4.2','CONFIRMED','Versions inherit the pinned K3s release; shared-kernel limitation remains explicit.'),
    ('TD-DEC-03','Ingress','K3s-bundled Traefik v3.7.8','CONFIRMED','Exact bundled release; dashboard private/disabled; HTTP/SSE/WSS regression gate before switch.'),
    ('TD-DEC-04','Preview replica model','Zero or one replica per mutable project preview','CONFIRMED','Vite state, mutable workspace and SQLite make multi-replica preview unsafe.'),
    ('TD-DEC-05','Generated stack','React/Vite + FastAPI + SQLAlchemy/Alembic + SQLite','CONFIRMED','One fixed template for the POC.'),
    ('TD-DEC-06','Project runtime boundary','One K3s namespace per project','CONFIRMED','Namespace, quota, service account and NetworkPolicy form the Stage 1 isolation/cleanup unit.'),
    ('TD-DEC-07','Workspace persistence','Project PVC: local-path-retain, RWO, 2 GiB baseline, UID/GID/fsGroup 10001','CONFIRMED','Custom local-path StorageClass uses WaitForFirstConsumer + Retain; snapshot restore provisions a new PVC; no live SQLite file copy.'),
    ('TD-DEC-08','Deployment lifecycle','No Stage 1 Deploy button','CONFIRMED','Stage 2 owns immutable build, registry, stable route and rollback.'),
    ('TD-DEC-09','Cline session recovery','Persist/reconnect first; prove old Job/Pod write revocation before an in-place linked retry','CONFIRMED DIRECTION','Lease fencing metadata cannot revoke PVC writes; an unproven old writer keeps mutation blocked.'),
    ('TD-DEC-10','Organization hierarchy','Normalized Organization/OU/Department tree plus effective-dated user membership','CONFIRMED DIRECTION','Stage 1 authorization still uses PROJECT.owner_id only; future project membership is a separate contract.'),
    ('TD-DEC-11','Whole-app versioning','main mirrors latest saved version; workspace/<project-id> holds active work; immutable snapshot binds Git + SQLite','CONFIRMED DIRECTION','One branch per Save is rejected; autosaves rotate and Restore creates new history.'),
    ('TD-DEC-12','Platform application stack','NestJS + Prisma + PostgreSQL with restricted AdminJS','CONFIRMED','AdminJS is an allowlisted UI over NestJS services, never unrestricted direct CRUD or a second authorization path.'),
    ('TD-DEC-13','Session and token model','Opaque browser session in Redis; short-lived audience-bound JWT only for preview/service exchange','CONFIRMED DIRECTION','No long-lived browser JWT/localStorage. Session/cache failure, CSRF, rotation, revocation and replay fail closed.'),
]

components = [
    ('Product UI','React/Vite/TypeScript','Projects, chatbot, progress, diff, preview and checkpoints.','Stateless product frontend; no token in localStorage.'),
    ('Control plane','NestJS + Prisma','Identity, ownership, desired state, commands, API and admitted background workers.','PostgreSQL is authoritative; Prisma transactions wrap CAS/dedup operations.'),
    ('Administration/session','AdminJS + Redis','Restricted administration, opaque server-side sessions, CSRF state and bounded cache.','AdminJS actions call NestJS services; Redis session loss fails authentication closed.'),
    ('Platform database','PostgreSQL','Users, RBAC, projects, chat, runs, routes, state, token replay grants and audit.','Prisma-managed schema; not exposed to generated apps.'),
    ('Ingress','Traefik','TLS, app route, wildcard preview route, HTTP/SSE/WSS.','Dashboard private or disabled.'),
    ('Preview Gateway','Authenticated reverse proxy','Owner authorization, project route lookup and same-origin path routing.','Does not trust user-supplied upstreams or platform credentials.'),
    ('Agent adapter','Pinned Cline SDK service','Runs/cancels agent, persists events, emits heartbeat and maps progress/recovery.','Page reconnect never starts a new run; no provider master key in workspace.'),
    ('Agent Job','Kubernetes Job','Mounts project workspace and writes validated files.','One mutation/project; bounded termination grace.'),
    ('Preview Pod','Vite + Uvicorn containers','Mutable frontend/backend preview, zero or one replica.','Disposable; not a release artifact.'),
    ('Workspace PVC','local-path-retain RWO PVC','Generated source, Git working tree and SQLite app data.','2 GiB baseline; UID/GID/fsGroup 10001; Restore uses a verified new generation.'),
]

resource_model = [
    ('Project boundary','One namespace per project','Required','Namespace + quota + service account + NetworkPolicy is the isolation and cleanup unit.'),
    ('Workspace storage','local-path-retain PVC, RWO, 2 GiB','Required','Preview and one Agent Job mount with runAsUser/runAsGroup/fsGroup 10001; WaitForFirstConsumer; reclaim Retain.'),
    ('Preview workload','Deployment, replicas 0 or 1','Required','Scale to zero on stop/idle; one mutable active instance.'),
    ('Frontend runtime','Vite container, port 5173','Required','HMR development server; never published as production runtime.'),
    ('Backend runtime','Uvicorn/FastAPI container, port 8000','Required','Source reload enabled only in mutable preview.'),
    ('Preview route','Service + Preview Gateway','Required','Named frontend/backend targets; no host port per project.'),
    ('Agent mutation','Job mounting workspace PVC','Required','Bounded duration; heartbeat; one mutation lease/project; fenced recovery.'),
    ('Resource governance','ResourceQuota + LimitRange','Required','CPU, RAM, storage and object-count limits.'),
    ('Network governance','Default-deny capable NetworkPolicy','Required','Exact K3s policy provider confirmed by spike.'),
    ('Runtime hardening','Non-root, no privileges/socket/hostPath','Required','Standard runc shares the host kernel; not microVM isolation.'),
]

routing = [
    ('app.<domain>','/','Product UI service','Authenticated product experience.'),
    ('app.<domain>','/api/*','NestJS service','Platform API; session/CSRF/auth enforced.'),
    ('*.preview.<domain>','/','Preview Gateway → project frontend target','Vite HTML/assets and HMR endpoint.'),
    ('*.preview.<domain>','/api/*','Preview Gateway → project backend target','Same-project FastAPI only; path contract is immutable.'),
    ('*.preview.<domain>','WebSocket upgrade','Preview Gateway → Vite HMR','WSS, forwarded host/protocol and compatible idle timeout.'),
]

change_matrix = [
    ('React/TypeScript/CSS source','Vite HMR','No manual action','Preview updated','Automatic when module accepts HMR.'),
    ('Frontend source not accepted by HMR','Browser/module reload','No manual action','Preview refreshed','Still not a deployment.'),
    ('Python application source','Uvicorn process reload','No manual action','Backend refreshed','Frontend reconnect/retry is bounded.'),
    ('Frontend dependency/lockfile','Controlled install + preview restart/rebuild','Status shown','Updating dependencies','Allowlisted package policy and timeout.'),
    ('Python dependency/lockfile','Controlled install + backend restart/rebuild','Status shown','Updating dependencies','Allowlisted package policy and timeout.'),
    ('Environment value','Preview restart','Status shown','Restarting preview','Secret values never enter source or user-visible event.'),
    ('Additive Alembic migration','Controlled migration + backend restart','Status shown','Updating project data','Snapshot/checkpoint before apply.'),
    ('Destructive migration','Explicit approval','Approve/reject','Approval required','Fail closed on unknown impact.'),
    ('Vite/Uvicorn port contract','Reject or recreate from platform template','Status shown','Rebuild required','User cannot choose arbitrary upstream.'),
    ('Dockerfile/base runtime contract','Platform-controlled preview rebuild','Status shown','Rebuilding preview','Not an OCI release/publish action.'),
    ('User Kubernetes/Compose definition','Reject in Stage 1','None','Unsupported change','Platform owns allowlisted runtime resources.'),
]

lifecycle = [
    ('STOPPED','No active Preview Pod; durable workspace retained.','Explicit Open/Resume.'),
    ('QUEUED','Capacity/admission waiting.','Start when slot available or Cancel.'),
    ('STARTING','Preview resources created; readiness pending.','Become READY or recoverable ERROR.'),
    ('READY','Preview and route healthy.','Activity, mutation, idle warning or stop.'),
    ('BUSY','Agent/restart/rebuild active.','Return READY or terminal error.'),
    ('IDLE_WARNING','No recent activity; stop countdown visible.','Activity returns READY; timeout starts STOPPING.'),
    ('STOPPING','Route withdrawn and Preview Pod scaling down.','Become STOPPED or recoverable ERROR.'),
    ('ERROR / DEGRADED','Expected state does not match healthy observed state.','Explicit Retry, Reconcile, Revert or Stop.'),
]

security_detail = [
    ('Identity','NestJS guards use an opaque HttpOnly Secure browser cookie backed by Redis server-side session state, per-session CSRF protection and Prisma/PostgreSQL owner checks on every object. Long-lived browser JWTs are prohibited.','SEC-01, SEC-02'),
    ('AdminJS','Mounted behind the same NestJS authentication/CSRF/RBAC guards; resources, properties and actions are explicit allowlists. Secret fields, raw state edits and unrestricted CRUD are disabled.','SEC-01, SEC-07, SEC-09'),
    ('Organization hierarchy','Default Organization → OU → Department plus one effective-dated primary user membership.','SEC-01; PROJECT.owner_id remains the only normal Stage 1 project grant.'),
    ('Route authorization','Preview Gateway validates active session, owner, project state and route generation before proxying.','SEC-01, SEC-02'),
    ('Pod security','Non-root; no privileged mode; no hostNetwork, hostPID, hostIPC, hostPath or containerd socket.','SEC-04'),
    ('Filesystem','Project-scoped PVC; reject absolute, drive-qualified and parent-traversal paths; reject symlink ancestors/leaves (including broken links) and non-regular special nodes; resolve beneath the project root before Git, hash, backup or restore reads; use descriptor-relative/no-follow operations and fail closed on concurrent path mutation.','SEC-03, SEC-04'),
    ('Network','Default deny where supported; allow DNS, approved model/package endpoints and platform callbacks only.','SEC-03, SEC-06'),
    ('Resources','Requests/limits, PID control where available, quota, global admission and bounded Job timeout.','SEC-05'),
    ('Secrets','Secret references injected by platform; never persist provider key in PVC, Git, logs, event payload or SYS_CONFIG.','SEC-07'),
    ('Change safety','Mutation lease, pre-change checkpoint, validation, explicit destructive approval and fail-closed classifier.','SEC-08'),
    ('Audit','Correlation ID across request, run, Job, route and lifecycle commands; redacted evidence only.','SEC-09'),
    ('Recovery','Explicit Save has zero committed-data RPO; restore drill target is ≤5 minutes for the acceptance fixture. PVC reclaim is Retain and restore provisions a new PVC from verified Git+SQLite snapshot; autosave interval/retention remain separate decisions.','SEC-10'),
    ('Version integrity','Server-owned Git refs; allowlisted source paths; SQLite online backup; hashes and manifest; no secret/runtime data committed.','SEC-08, SEC-10; AC-VER-01–04'),
]

capacity = [
    ('Registered users','100','Identity and metadata only; not simultaneous runtime demand.'),
    ('Concurrent active users','50','Control-plane planning workload.'),
    ('Active previews','6 launch maximum','Each mutable preview is one replica and consumes 1 capacity slot.'),
    ('Concurrent Agent Jobs','2 launch maximum','Each Agent Job consumes 2 slots; one mutation per project.'),
    ('Admission pool','10 slots / 2 reserved / queue 100 / timeout 900 s','Serialized pool admits at most six previews plus two Agent Jobs; reserve protects control/recovery capacity.'),
    ('Peak platform API','75 RPS','Hypothesis to validate with representative workload.'),
    ('Preview idle stop','60-minute candidate','Measure UX and resource effect before approval.'),
    ('Host candidate','64 vCPU / 128 GiB / 2–4 TB NVMe','Planning candidate only; procurement requires benchmark evidence.'),
]

resource_profiles = [
    ('Preview frontend','100m / 256 MiB','500m / 512 MiB','128 / 1 GiB','1 Preview component slot'),
    ('Preview backend','100m / 256 MiB','500m / 512 MiB','128 / 1 GiB','included in the same Preview slot'),
    ('Agent Job','250m / 512 MiB','1000m / 1 GiB','256 / 2 GiB','2 Agent component slots'),
]

failure_modes = [
    ('NestJS Pod crash','Kubernetes restarts Pod; other stateless replica may serve new request.','Persist session/event truth externally; reconnect test.'),
    ('Preview Pod crash','Recreated against same PVC; route remains unready until health passes.','No loss of Git/checkpoint/SQLite; bounded readiness.'),
    ('Agent Job failure','Run fails safely; no accepted checkpoint; retry creates a linked Job/run.','Pre-change checkpoint and immutable terminal evidence.'),
    ('Cline heartbeat stalls','UI shows Recover agent session; watchdog/owner requests STOPPING and proves old Job/Pod write revocation before an in-place retry.','Lease fencing alone is insufficient; prove no overlapping writer/Job/run and preserve the last healthy checkpoint.'),
    ('HMR WebSocket disconnect','Browser reconnects; page reload available if module state cannot recover.','No auth bypass or silent Ready.'),
    ('Uvicorn reload failure','Backend target remains unready/degraded; UI shows recovery.','Last healthy checkpoint remains available.'),
    ('PVC mount/ownership failure','Preview does not become Ready.','Record safe error, Pod event evidence and remediation.'),
    ('Capacity exhausted','Command becomes queued or rejected with bounded retry.','Do not present login failure or generic server error.'),
    ('Single node fails','All K3s/platform/preview services stop.','Restore/rebuild from durable backup; no HA claim.'),
    ('PostgreSQL fails','Control plane cannot safely authorize or mutate.','Fail closed; restore procedure required.'),
    ('Redis fails','Current browser authentication and non-idempotent replay-protected calls fail closed; cache-only reads bypass to PostgreSQL.','Restore Redis, deny missing sessions, prove no unbounded-login or browser-JWT fallback, and require fresh sign-in after session loss.'),
    ('Snapshot part fails or hashes disagree','Snapshot remains non-AVAILABLE and PROJECT.current_snapshot_id is unchanged.','Delete/quarantine partial object; reconcile Git main to last AVAILABLE snapshot; retry with same idempotency key.'),
]

version_matrix = [
    ('K3s','v1.36.3+k3s1','Exact release artifact + checksum; stable channel resolved once on 26 Aug 2026','Fresh install, snapshot restore, replacement-node upgrade and rollback','CONFIRMED'),
    ('Traefik','v3.7.8 bundled','Pinned K3s manifest/chart','HTTP/SSE/WSS route regression before node switch','CONFIRMED'),
    ('containerd/runc','v2.3.2-k3s2 / v1.4.2','Pinned K3s distribution','Pod security, PVC and compatibility regression','CONFIRMED'),
    ('Python','TBD','Container image digest','FastAPI/Cline/tool compatibility','DECISION REQUIRED'),
    ('Node.js','22 target','Container image digest','Vite and @cline/sdk compatibility','PROVISIONAL'),
    ('@cline/sdk','0.0.79','Exact lockfile','Events, cancel, page reconnect, heartbeat/stuck recovery and isolation','CONFIRMED TARGET'),
    ('Vite','TBD','Frontend lockfile','HMR over wildcard WSS','DECISION REQUIRED'),
    ('FastAPI/Uvicorn','TBD','Generated-app Python lockfile','Reload, readiness and reconnect','DECISION REQUIRED'),
    ('NestJS','TBD pinned major/minor','Platform Node lockfile + image digest','API, guards, SSE/WSS, worker and shutdown compatibility','DECISION REQUIRED'),
    ('Prisma','TBD pinned version','Platform Node lockfile + migration history','PostgreSQL schema, interactive transactions, CAS row-count and rollback tests','DECISION REQUIRED'),
    ('AdminJS','TBD pinned version','Platform Node lockfile','NestJS/Prisma adapter, action/property allowlist and auth/CSRF regression','DECISION REQUIRED'),
    ('PostgreSQL','TBD pinned image digest','Platform deployment manifest','Migration, concurrency, backup/restore and failover procedure','DECISION REQUIRED'),
    ('Redis','TBD pinned image digest','Platform deployment manifest','Session TTL/revoke, cache fallback, replay and outage fail-closed tests','DECISION REQUIRED'),
    ('Git','TBD pinned distribution','Workspace image package lock','SHA format, ref CAS, path allowlist, crash recovery and repository ownership','DECISION REQUIRED'),
]

spikes = [
    ('SPK-01','K3s foundation','Install pinned K3s; prove Pod, Job, Service, PVC, quota and cleanup.','All resources reconcile; no host-level HA claim.'),
    ('SPK-02','PVC file events and ownership','Agent Job and Preview Pod mount one project PVC; file watches propagate.','Vite/Python change is detected; UID/reclaim rules documented.'),
    ('SPK-03','Authenticated wildcard routing','Traefik to Preview Gateway to project targets for HTTP, SSE and WSS.','Wrong owner/host/path denied; HMR reconnect passes.'),
    ('SPK-04','Reload classification','Exercise each row in the change matrix.','No dependency/migration/runtime change is falsely reported as HMR.'),
    ('SPK-05','Isolation and egress','Apply Pod security, quota and default-deny-capable policy.','Socket/hostPath/private network/cross-project negative tests pass.'),
    ('SPK-06','Admission and failure recovery','Launch-limit load, Pod kill, Job fail, PVC error and node-stop exercises.','Queue remains bounded; durable evidence/recovery paths verified.'),
    ('SPK-07','Cline session recovery','Refresh browser during run; suppress heartbeat; issue automatic/manual recovery; race duplicate clicks.','Reconnect creates no duplicate; old Pod UID/mount can no longer write before an in-place linked retry; lease fencing alone fails the gate.'),
    ('SPK-08','Whole-app versioning','Initialize Version 1; dirty/no-change autosave; concurrent Save; crash each capture stage; mutate SQLite in WAL mode; restore old code+DB.','Numbering/idempotency is exact; no live-file copy; partial snapshot never becomes current; restore creates new history and passes health.'),
]

# ---------------- build Technical Design ----------------
tech = DocBuilder(
    'Vibe-Code POC\nTechnical Design & Runtime Architecture',
    'VIBE-CODE  /  TECHNICAL DESIGN & RUNTIME ARCHITECTURE',
    'Stage 1 K3s runtime, live preview, HMR/reload, routing, security, capacity and verification design'
)
tech.cover(
    'TECHNICAL DESIGN DOCUMENT',
    'Implementation design for the Stage 1 mutable preview runtime. Product requirements remain authoritative in the PRD & Delivery Plan.',
    [('Status','DRAFT — spike validation required'),('Boundary','Stage 1 POC technical baseline'),('Version','0.1'),('Product source','vibe-code-no-code-poc-prd-and-delivery-plan.docx v1.1'),('Orchestrator','Single-node K3s'),('Runtime','containerd + standard runc'),('Ingress','Traefik')],
    [('1. Purpose and authority','3'),('2. Approved decisions and open choices','4'),('3. Architecture context','5'),('4. K3s project runtime model','6'),('5. End-to-end live update design','7'),('6. Preview routing and connectivity','8'),('7. Change, reload, restart and rebuild classification','9'),('8. Lifecycle and state ownership','10'),('9. Workspace, Git, SQLite and checkpoint consistency','11'),('10. Security implementation','12'),('11. Capacity, admission and resource profiles','13'),('12. Observability and operational evidence','14'),('13. Failure and recovery design','15'),('14. Technology and version matrix','16'),('15. Required spikes and technical acceptance','17'),('16. Stage 2 extension boundary','18'),('17. Open technical decisions','19'),('18. Traceability matrix','20')]
)

tech.para('1. Purpose and authority','Heading 1')
tech.para('This document specifies how the approved Stage 1 product requirements are intended to work. It may evolve through technical spikes without silently changing product scope. If this document conflicts with the PRD on user behavior, scope or acceptance, the PRD wins until the product decision is changed explicitly.')
tech.table(['Artifact','Owns'],[
    ('PRD & Delivery Plan','Why/what/who, product scope, Features, Epics, Stories, Use Cases, security baseline and acceptance.'),
    ('This Technical Design','How, component boundaries, K3s resources, routing, live update, versions, failure modes and spike evidence.'),
    ('Data Specification XLSX','Logical/physical data contracts, states, relationships, error map and governance.'),
], [2.2,4.8], 7.8)
tech.callout('Stage 1 is a mutable development preview. Stage 2 introduces Deploy, immutable build/publish, registry, stable application route and promotion/rollback.', PALE_AMBER, '92400E')
tech.page()

tech.para('2. Approved decisions and open implementation choices','Heading 1')
tech.table(['ID','Area','Decision','Status','Rationale / gate'], tech_decisions, [.75,1.25,2.25,1.05,2.1], 6.8)
tech.para('Decision discipline','Heading 2')
for item in [
    'Do not freeze a version, storage class, namespace model or NetworkPolicy provider without executable evidence.',
    'Prefer standard Kubernetes resources and a small platform-owned runtime contract.',
    'Generated projects cannot supply arbitrary Kubernetes or Compose definitions.',
    'A decision marked Provisional is implemented behind a small adapter and must be revisited after its named spike.',
]: tech.bullet(item)
tech.page()

tech.para('3. Architecture context','Heading 1')
tech.image(ARCHITECTURE, 7.1, 'Figure 1 — Stage 1 technical context. Service-level distribution exists on one host; node-level HA does not.')
tech.para('Component catalogue','Heading 2')
tech.table(['Component','Technology','Responsibility','Boundary'], components, [1.45,1.45,2.75,1.35], 6.8)
tech.para('Vibe page component and layout contract','Heading 2')
tech.para('The PRD owns the user-visible two-pane layout. This view defines which technical component owns each region and where credentials, commands and generated application traffic must not cross.')
tech.image(UI_COMPONENT_CONTRACT, 7.1, 'Figure 2 — Vibe page component ownership and trust boundaries. The generated preview never receives platform credentials or control-plane authority.')
tech.table(['Product region','Technical owner','Data / protocol contract','Failure posture'],[
    ('Project header and status','Product UI + NestJS','Owner-scoped project/version/status API; safe correlation reference.','Unknown or stale status shows Action required, never fabricated Ready.'),
    ('Conversation and controls','Product UI + NestJS + Agent adapter','Durable messages/events; idempotent Stop/Save/Restore/decision commands.','Reconnect existing run; uncertain mutation remains blocked.'),
    ('Preview pane','Preview Gateway + Preview Pod','One-time exact-audience JWT, host-only cookie, fixed frontend/API/WSS targets.','Replay, stale generation, wrong owner/session/host or unready route is denied.'),
    ('Full preview','Same Preview Gateway contract','New viewport, not a new authorization or runtime path.','No domain-wide cookie, arbitrary upstream or stable published route.'),
], [1.45,1.55,2.8,1.2], 6.9)
tech.para('Authentication, sessions and token boundaries','Heading 2')
tech.para('The first-party browser uses an opaque HttpOnly Secure session cookie with a Redis server-side session. A long-lived JWT is never stored in browser JavaScript or localStorage. PostgreSQL remains authoritative for users, RBAC, audit and one-time preview replay grants; Redis is disposable session/cache infrastructure, not the business system of record.')
tech.table(['Flow','Credential','Store / authority','Lifetime','Revocation / failure'],[
    ('Browser login','Opaque random session ID in `__Host-vibe_session`; HttpOnly, Secure, SameSite=Lax; regenerated at login/privilege change.','Redis server-side session; AUTH_USER/RBAC in PostgreSQL through Prisma.','Bounded idle + absolute TTL; exact values DECISION REQUIRED.','Logout, password/role/session_version change or admin revoke deletes all matching Redis sessions; Redis unavailable fails authentication closed.'),
    ('State-changing browser request','Per-session CSRF secret rendered as token and sent in a custom header; Origin/Referer allowlist.','CSRF hash and session metadata in Redis.','Same bound as session; rotate with session.','Missing/mismatch/wrong origin returns 403; no cookie-only mutation endpoint.'),
    ('AdminJS','Same opaque server-side session plus PLATFORM_ADMIN role and action permission; step-up hook for destructive/break-glass actions.','NestJS guard + Prisma/PostgreSQL RBAC; AdminJS is UI only.','Same session; shorter configurable idle policy may apply.','No direct model-wide CRUD bypass, secret field rendering, arbitrary action or raw state edit; accepted/denied actions audited.'),
    ('Preview JWT exchange','One-time short-lived JWT with `iss`, exact-host `aud`, `sub`, session hash, project, workspace generation, `jti`, `iat`, `exp`; signed by active asymmetric key.','AUTH_KEYSET public metadata + PostgreSQL PREVIEW_AUTH_GRANT one-time CAS; private key is a secret reference.','At most 60 seconds and one successful exchange.','Replay, wrong audience/session/project/generation, expired token or revoked grant fails closed.'),
    ('Preview browsing','Opaque host-only preview cookie; no platform session/JWT forwarded to generated code.','PostgreSQL PREVIEW_AUTH_GRANT authority; bounded Redis read-through cache may accelerate only positive state with short TTL.','Short configurable TTL no longer than platform session/grant.','Every HTTP/WSS handshake revalidates binding; cache failure falls back to PostgreSQL, never grants from stale cache.'),
    ('Internal service call','Short-lived JWT with service `sub`, exact `aud`, scoped permissions and unique `jti`; asymmetric key rotation through `kid`.','AUTH_KEYSET public metadata; issuer policy in NestJS; replay cache in Redis only where endpoint is non-idempotent.','Maximum five minutes; no refresh token.','Unknown `kid`, wrong audience/scope, expiry or replay fails closed; rotate keys with overlap and audited retirement.'),
], [1.1,2.4,1.8,1.2,1.5], 6.4)
tech.callout('Redis failure policy: authenticated browser requests fail closed because current sessions are unavailable. Cache-only reads bypass Redis and load authoritative PostgreSQL through Prisma. The platform never falls back to a self-contained browser JWT that cannot be revoked.', PALE_RED, '991B1B')
tech.page()

tech.para('4. K3s project runtime model','Heading 1')
tech.para('The preferred minimal runtime is shown below. Exact resource names and namespace/storage implementation remain provisional until the spikes pass.')
tech.table(['Concern','Preferred resource/model','Status','Contract'], resource_model, [1.35,1.9,1.0,2.75], 7.0)
tech.para('Preferred object relationship','Heading 2')
tech.callout('Project boundary → Workspace PVC → Preview Deployment (0/1) → one multi-port Preview Service. A bounded Agent Job mounts the same PVC and writes source under one mutation lease.', PALE_BLUE, BLUE)
tech.para('Why one Preview Pod and one multi-port Service','Heading 2')
for item in [
    'Vite and FastAPI share one project lifecycle and one workspace generation.',
    'A single Service can expose named frontend and backend ports to the Preview Gateway.',
    'It avoids unnecessary frontend/backend Deployments until independent scaling is required.',
    'The Agent Job remains separate because it is finite, admitted, auditable and independently terminated.',
]: tech.bullet(item)
tech.para('Load-balancing boundary','Heading 2')
tech.table(['Workload','Replica posture','Meaning'],[
    ('NestJS and Preview Gateway','Two same-node replicas where useful','Kubernetes Service can distribute new requests; shared state must be external.'),
    ('Each project Preview Pod','Zero or one replica','No horizontal balancing of one mutable preview.'),
    ('Agent Jobs','Several projects concurrently within global admission','One active mutating Job per project.'),
    ('PostgreSQL / K3s node','Single Stage 1 failure domain','No hardware or database HA claim.'),
], [2.1,2.3,2.6], 7.5)
tech.page()

tech.para('5. End-to-end live update design','Heading 1')
tech.image(RELOAD_FLOW, 7.1, 'Figure 2 — Explicit frontend HMR, backend reload and persistence flow.')
tech.para('Authoritative sequence','Heading 2')
for item in [
    'NestJS persists the user message and creates an admitted AGENT_RUN.',
    'The runtime controller creates one bounded Agent Job for the project generation.',
    'The Job mounts the project Workspace PVC and writes only permitted paths.',
    'Vite observes eligible frontend changes and sends an HMR update over the existing WSS route.',
    'Uvicorn observes supported Python source changes and reloads the backend process.',
    'The Preview Gateway continues same-project owner authorization and path routing during reconnect.',
    'Validation determines Ready, Restart required, Rebuild required, Approval required or Action required.',
    'Only an accepted result creates the accepted checkpoint/current-version transition.',
]: tech.number(item)
tech.page()

tech.para('6. Preview routing and connectivity','Heading 1')
tech.table(['Origin','Path / protocol','Target','Rule'], routing, [1.55,1.25,2.3,2.0], 7.2)
tech.para('Frontend HMR contract','Heading 2')
tech.para('The generated Vite configuration must listen inside the Pod and advertise the external secure WebSocket endpoint:')
tech.callout('server.host = "0.0.0.0"; server.port = 5173; HMR protocol = "wss"; client port = 443. Exact host derivation must use the authorized preview origin, not a user-supplied upstream.', LIGHT, NAVY)
tech.para('Backend reload contract','Heading 2')
tech.callout('uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload', LIGHT, NAVY)
tech.para('Uvicorn reload is allowed only in the mutable preview. A Stage 2 published application must use an immutable production command without --reload.', after=5)
tech.para('Preview Gateway authorization','Heading 2')
for item in [
    'The platform UI requests a one-time preview JWT through its opaque HttpOnly Secure session cookie backed by Redis. NestJS signs a short-lived JWT containing exact-host audience, user/session/project/workspace-generation binding and unique jti; the exact preview host exchanges it once for a separate short-lived HttpOnly, Secure, SameSite=Strict host-only preview cookie. Raw platform session cookies or JWTs are never domain-wide or forwarded to generated code.',
    'Every Preview Gateway replica uses PostgreSQL PREVIEW_AUTH_GRANT as the shared replay/revocation authority. Atomic row-version CAS consumes the hashed JWT jti once, replaces its credential hash with the preview-cookie hash, and binds exact host, owner, session and workspace generation; a replay or second replica loses the CAS and is denied.',
    'The Preview Gateway validates expiry, audience, owner, project, workspace generation and grant status on every HTTP and WebSocket handshake, rotates/revokes all grants for the platform session, and strips platform Cookie/Authorization headers before proxying. Raw JWTs/cookies are never persisted or logged.',
    'Resolve route key to the current workspace generation and owner in the control plane.',
    'Reject stopped, deleted, expired, wrong-owner, unknown-host and stale-generation requests.',
    'Route only fixed frontend and `/api/*` targets; never accept arbitrary upstream URL or forwarded authorization header from project code.',
    'Support WebSocket upgrade, SSE streaming, bounded request body and timeouts compatible with HMR/reload.',
]: tech.bullet(item)
tech.page()

tech.para('7. Change, reload, restart and rebuild classification','Heading 1')
tech.table(['Change type','Platform action','User action','User-visible state','Guard'], change_matrix, [1.75,1.75,1.0,1.5,1.5], 6.5)
tech.callout('A preview rebuild recreates a mutable development runtime. It is not the Stage 2 immutable Deploy/build/publish lifecycle.', PALE_AMBER, '92400E')
tech.page()

tech.para('8. Lifecycle and state ownership','Heading 1')
tech.table(['State','Meaning','Allowed next behavior'], lifecycle, [1.2,3.6,2.2], 7.3)
tech.para('State ownership rules','Heading 2')
for item in [
    'NestJS/PostgreSQL stores desired product state; K3s resources are observed runtime state.',
    'The runtime controller alone reconciles approved desired state to allowlisted K3s resources.',
    'The UI, scheduler and AdminJS issue named commands; they never forge runtime success.',
    'READY requires both workload health and authenticated route verification.',
    'Terminal attempts are immutable; retry creates a linked new attempt.',
    'Idle stop first records desired STOPPED, withdraws route and then scales Preview to zero.',
]: tech.bullet(item)
tech.para('Cline session continuity and stuck recovery','Heading 2')
for item in [
    'Refreshing the product page reloads persisted chat messages and AGENT_EVENT progress, then reconnects to the existing run; it never creates another Cline process, Agent Job or mutation lease.',
    'The Cline adapter emits a bounded heartbeat by appending safe run evidence and renewing only the matching mutation lease. Exact warning and stuck thresholds remain configuration decisions.',
    'When heartbeat is stale, the UI changes from Running to Recovering/Action required and exposes the owner-only Recover agent session command. Duplicate clicks share one idempotency key.',
    'Automatic or manual recovery first requests cancellation, moves the run through STOPPING and terminates the old Agent Job/Pod after a bounded grace period. Changing only a PostgreSQL lease or fencing token never removes filesystem write authority.',
    'The Runtime Controller persists exactly one authoritative Agent JOB RUNTIME_RESOURCE_REF per attempt: agent_run_id is a unique FK to that AGENT_RUN, observed UID is immutable, and write_revocation_json holds Job/Pod UID, deletion-observed time, failed-closed post-delete write probe, verifier digest and verification time.',
    'STOPPING→STOPPED and a new attempt with prior_attempt_id are rejected unless the prior attempt is terminal and the proof row has agent_run_id equal to that exact stopped/prior attempt. Missing, duplicate or another attempt’s valid proof fails closed. A staged retry may use only a separate immutable root; uncertain tool actions are never replayed.',
    'If write revocation cannot be proven, the platform fails closed, keeps the project mutation-blocked and requires bounded administrator reconciliation; lease expiry alone is insufficient evidence.',
]: tech.number(item)

tech.landscape()
tech.para('Authoritative state diagrams','Heading 2')
tech.callout('The repository state contract is shared with the PRD render and must equal the Data Specification Enum Catalogue and State Transitions sheets. Each diagram shows initial/terminal states, every legal directed edge, owner and the machine-level guard.', PALE_AMBER, '92400E')
for index, (machine, path) in enumerate(STATE_DIAGRAM_PATHS, 1):
    if index > 1:
        tech.page()
    tech.para(machine['field'],'Heading 2')
    image = Image.open(path)
    state_height_cap = 5.0 if index == 1 else 5.85
    display_width = min(9.7, state_height_cap * image.width / image.height)
    image.close()
    tech.image(path, display_width, f"Figure TS-{index:02d} — {machine['field']}: {len(machine['states'])} states and {len(machine['transitions'])} legal directed transitions.")
tech.portrait()

tech.para('9. Workspace, Git, SQLite and checkpoint consistency','Heading 1')
tech.table(['Asset','Authority','Stage 1 rule'],[
    ('Project metadata / lifecycle','PostgreSQL','PROJECT.current_snapshot_id is authoritative; only an AVAILABLE same-project snapshot may become current.'),
    ('Generated source / Git refs','Protected project workspace + backup','main mirrors the latest saved version; workspace/<project-id> is the active work branch; ref names are server-generated.'),
    ('Mutable working tree','Workspace PVC','Mounted by Preview and one Agent Job; Git excludes SQLite, secrets, caches, build output and runtime files.'),
    ('Generated application data','SQLite on project PVC','Single mutable preview; capture uses SQLite online backup/checkpoint under the project mutation lease, never a raw live-file copy.'),
    ('Snapshot manifest','Private snapshot store + PostgreSQL','Binds kind/version, Git commit/ref, SQLite object/hash, migration head, template/locks, actor and health evidence.'),
], [2.0,2.25,2.75], 7.4)
tech.para('Git reference contract','Heading 2')
tech.callout('Repository boundary: the platform source/delivery Git repository is separate from every generated-project repository. Project initialization runs only inside the generated project root: git init --initial-branch=main → add allowlisted template source → commit “1st version” → create workspace/<project-id>. Updating generated-project main never updates, checks out or pushes the platform source repository; normal users never receive a terminal.', LIGHT, NAVY)
for item in [
    'The interval worker commits changed allowlisted source on workspace/<project-id>; a clean tree creates no commit. It also captures a rotating AUTOSAVE whole-app recovery snapshot, but does not increment the user-visible version number.',
    'Save flushes any pending workspace commit, captures and verifies SQLite plus the manifest, then advances main to the exact commit through an ancestry check and compare-and-swap. It creates immutable checkpoint/<snapshot-id> evidence; it does not create one branch per Save.',
    'Git commit IDs are opaque 40–64 hexadecimal identifiers. The data contract must not assume SHA-1 forever.',
]: tech.bullet(item)
tech.para('Whole-app Save and Restore transaction','Heading 2')
for item in [
    'Acquire the project mutation lease; withdraw the preview route and quiesce/terminate the active generated-app Pod so no process retains the old SQLite inode or writable tree. Lease metadata alone is not proof of quiescence.',
    'Capture code commit, SQLite online backup, migration head and version locks; hash every external artifact and validate health before SNAPSHOT becomes AVAILABLE.',
    'Version numbering and idempotency are allocated by the control plane. Initial snapshot is Version 1; each successful user Save or Restore result creates exactly one next version.',
    'Restore creates a durable idempotent RESTORE_JOB, captures an AVAILABLE PRE_RESTORE snapshot, and materializes code plus SQLite into a new generation-specific staging root. The active generation remains immutable and authoritative while staged hashes, migration head, locks and health are verified.',
    'The ACQUIRING_LOCK→RESTORING_CODE transaction runs before code/DB mutation: capture AVAILABLE PRE_RESTORE; prior_active_generation, prior_snapshot_id, optional prior_route_generation and exact prior_route_status; allocate target_generation and exactly one provisional RESTORE_RESULT. These links/before-images become immutable, so early RESTORING_CODE/DB rollback and later rollback use the same result row.',
    'At cutover, one PostgreSQL transaction locks PROJECT, WORKSPACE and PREVIEW_ROUTE, re-checks generation/snapshot/route-generation/route-status before-images and row versions, then CAS-updates all pointers, marks the existing result AVAILABLE and writes journal_decision=COMMIT, journal_phase=COMMITTED and cutover_committed_at. Any mismatch rolls back the entire database transaction.',
    'Product Restore never leaves detached HEAD, rewrites published history or relies on git revert/checkout alone. Git checkout/switch may be used only inside a fenced temporary verification worktree.',
    'Any early or late failure marks the always-present provisional result INVALID/quarantined. ROLLED_BACK requires one verified transaction restoring generation, snapshot, route generation and exact prior route status; READY is restored only if it was the prior status and health passes. Rollback failure enters MANUAL_INTERVENTION, keeps route withdrawn/mutation blocked, and never claims success.',
    'A restore_journal_recovery worker runs on controller startup and every 30 seconds. It SELECTs non-terminal RESTORE_JOB rows stale after 60 seconds FOR UPDATE SKIP LOCKED and CAS-claims (restore_job_id, journal_phase, row_version). PENDING rolls back staging/result, committed COMMIT reconciles target, and ROLLBACK restores verified prior pointers; the third failed attempt enters MANUAL_INTERVENTION. The acceptance fixture must reach verified READY or MANUAL_INTERVENTION within 5 minutes.',
    'For commands that use AUDIT_LOG.command_id dedup, the unique command claim, aggregate row_version compare-and-swap, accepted/rejected result and immutable audit append commit in the same PostgreSQL transaction; either all commit or all roll back. command_id never substitutes for aggregate CAS, and retries return the stored outcome without replay.',
]: tech.number(item)

tech.landscape()
tech.para('Transactional sequence diagrams','Heading 2')
tech.callout('These four diagrams are normative realizations of Save, Restore, preview authentication and capacity admission. Their guards, authority stores and failure outcomes bind the prose above.')
for tx_no, (title, path, lanes, steps, guard) in enumerate(TRANSACTION_DIAGRAMS, 1):
    if tx_no > 1:
        tech.page()
    tech.para(title, 'Heading 3')
    tech.image(path, 9.5, f'Figure TX-{tx_no:02d} — {title}. {guard}')
tech.portrait()
tech.para('10. Security implementation','Heading 1')
tech.table(['Area','Implementation contract','PRD trace'], security_detail, [1.3,4.7,1.0], 7.1)
tech.para('Shared-kernel limitation','Heading 2')
tech.callout('Standard containerd/runc containers share the K3s host kernel. Namespace, NetworkPolicy and Pod security improve control and governance but do not provide microVM-grade isolation.', PALE_RED, '991B1B')
tech.para('Future isolation option','Heading 2')
tech.para('A gVisor RuntimeClass may be evaluated after Stage 1 compatibility evidence. It is not a current requirement. Kata/microVM approaches require a separate risk, cost and lifecycle decision.')
tech.page()

tech.para('11. Capacity, admission and resource profiles','Heading 1')
tech.table(['Measure','Assumption','Operational interpretation'], capacity, [1.85,2.25,2.9], 7.2)
tech.para('Stage 1 container profiles','Heading 2')
tech.table(['Component','CPU / memory request','CPU / memory limit','PID / ephemeral limit','Admission charge'], resource_profiles, [1.25,1.55,1.55,1.3,1.35], 6.7)
tech.para('Admission principles','Heading 2')
for item in [
    'CAPACITY_POOL stores required CPU/RAM/PID/storage profile, capacity/reserve slots, max queue depth, timeout and OWNER_ROUND_ROBIN/FIFO. Every PREVIEW and AGENT_JOB has a separate durable RESOURCE_LEASE component row; AGENT_JOB includes exact agent_run_id, while owner is derived through workspace/project.',
    'Under SELECT CAPACITY_POOL FOR UPDATE, enqueue assigns monotonic queue_sequence and rejects max depth. Dequeue derives owner via workspace/project, scans distinct queued owner IDs cyclically ascending after durable last_admitted_owner_id (NULL or absent owner starts at the smallest queued owner), then takes that owner’s minimum queue_sequence. QUEUED→RESERVED and cursor update commit together; after crash, committed cursor plus queued rows reconstruct the unique next owner. Reserve only when SUM(RESERVED slot_count)+requested fits. Agent completion releases only its component.',
    'Preserve login, project access and running previews before admitting new starts.',
    'Queue or reject new starts with a visible state; queue overflow and timeout are deterministic, and silent global oversubscription is prohibited.',
    'Use requests and limits per container plus namespace/global quotas.',
    'Raise launch limits only after representative load and failure tests pass.',
]: tech.bullet(item)
tech.page()

tech.para('12. Observability and operational evidence','Heading 1')
tech.table(['Signal','Required evidence','User exposure'],[
    ('Request/run correlation','Correlation ID across API request, AGENT_RUN, Agent Job, route and checkpoint.','Reference code and plain-language status.'),
    ('Lifecycle','Desired/observed state, Pod condition, readiness and bounded event reason.','Ready, Waiting, Recovering, Stopped or Action required.'),
    ('Live update','HMR connect/reconnect, Uvicorn reload, classifier decision and elapsed time.','Preview updated/refreshed/restarting/rebuilding.'),
    ('Agent continuity','Cline adapter heartbeat, last event time, Job UID, mutation lease holder, cancel/termination result and prior attempt.','Running, Recovering or Recover agent session; no raw process output.'),
    ('Capacity','Admitted previews/jobs, queue depth, CPU, RAM, disk and startup duration.','Waiting state; administrators see safe aggregate evidence.'),
    ('Security','Denied owner/route/network/secret/policy decisions.','Generic denial; technical evidence restricted and redacted.'),
    ('Cost/usage','Model alias, token estimate, duration and approved cap.','Safe estimate with provenance/freshness label.'),
], [1.45,3.75,1.8], 7.1)
tech.para('No high-cardinality project IDs or raw prompt content should be placed in unbounded metrics labels. Logs/events must exclude secrets and unrestricted command output.', after=5)
tech.page()

tech.para('13. Failure and recovery design','Heading 1')
tech.table(['Failure','Expected behavior','Verification / recovery'], failure_modes, [1.55,3.0,2.45], 7.0)
tech.callout('Single-node K3s handles process/Pod failure. A host, disk, network-interface or single PostgreSQL failure can stop the complete platform. This is an accepted Stage 1 limitation.', PALE_AMBER, '92400E')
tech.para('Stage 1 recovery baseline','Heading 2')
for item in [
    'restore_journal_recovery runs at controller startup and every 30 seconds. A row is stale after 60 seconds without owner heartbeat; one worker claims it with FOR UPDATE SKIP LOCKED + row_version CAS. Three failed recovery attempts enter MANUAL_INTERVENTION with route withdrawn and mutation blocked.',
    'An explicit User Save has zero committed-data RPO. For the acceptance fixture, Restore and journal crash recovery must reach verified READY or MANUAL_INTERVENTION within 5 minutes. Crash-recovery RPO between autosaves remains tied to the separately approved autosave interval.',
    'Before K3s change, capture K3s datastore, PostgreSQL and whole-app snapshots off-node and record hashes. Upgrade uses a replacement node on pinned artifacts; run Pod/PVC/route/restore regression, then switch traffic. Rollback discards the candidate and restores the prior pinned node plus verified snapshots; no in-place floating-channel upgrade.',
    'PVC recovery provisions a new local-path-retain RWO PVC, restores allowlisted Git files and SQLite through verified snapshot/online-backup data, applies UID/GID/fsGroup 10001, then health-checks before route activation. Retention duration remains DECISION REQUIRED.',
]: tech.bullet(item)
tech.page()

tech.para('14. Technology and version matrix','Heading 1')
tech.table(['Component','Selected/target version','Pin method','Required evidence','Status'], version_matrix, [1.35,1.45,1.35,2.2,1.1], 6.8)
tech.para('Version rule','Heading 2')
tech.para('TBD values remain decision gates. A version becomes Confirmed only when installation, compatibility, security and rollback evidence is recorded. Floating latest tags or unbounded package ranges are prohibited.')
tech.page()

tech.para('15. Required spikes and technical acceptance','Heading 1')
tech.table(['ID','Spike','Execution','Exit evidence'], spikes, [.75,1.55,3.15,1.55], 6.9)
tech.para('Minimum live-preview test set','Heading 2')
for item in [
    'React component and CSS edit updates through Vite HMR over authenticated WSS.',
    'Python endpoint edit triggers Uvicorn reload and browser/API reconnect without cross-owner leakage.',
    'Frontend and Python dependency changes select restart/rebuild rather than HMR.',
    'Environment and port/runtime-contract changes select the correct bounded action.',
    'Additive migration checkpoints and applies; destructive migration pauses for explicit approval.',
    'Agent Job failure, Preview Pod kill and HMR disconnect preserve durable state and produce safe evidence.',
    'Refreshing the web page during a running Cline session reconnects to the same AGENT_RUN and creates no duplicate Job or lease.',
    'A suppressed heartbeat exposes Recover agent session; duplicate recovery commands remain idempotent and one linked retry starts only after old-Job/Pod write revocation is proven; metadata lease fencing alone does not qualify.',
    'Provisioning creates the default Organization/OU/Department membership; two users in the same Department still fail all cross-owner project/preview/API tests.',
    'Project initialization creates main, Version 1 and workspace/<project-id>; dirty/no-change interval autosave and duplicate Save requests are deterministic.',
    'SQLite WAL writes during capture produce a valid online backup; injected failures never expose a partial snapshot or advance current_snapshot_id.',
    'Restoring Version N stages a generation, creates exactly one linked RESTORE_RESULT, then uses a journaled generation-pointer compare-and-swap; crash injection at every step either preserves the old generation or exposes the complete new one.',
]: tech.bullet(item)
tech.page()

tech.para('16. Stage 2 extension boundary','Heading 1')
tech.para('Stage 2 may introduce the following separate immutable release path:')
tech.callout('Accepted Git SHA → controlled BuildKit build → OCI image → private registry → published K3s Deployment → stable application route → health gate → promotion / rollback', PALE_BLUE, BLUE)
tech.table(['Stage 1 mutable preview','Stage 2 published release'],[
    ('Vite development server and HMR','Built static assets served by production web server'),
    ('Uvicorn --reload','Production process without reload watcher'),
    ('Workspace PVC is mutable','Image/artifact is immutable and content-addressed'),
    ('Zero or one preview replica','Replica/HA policy defined by production workload'),
    ('Owner-only temporary route','Stable approved application route'),
    ('Checkpoint/revert within project','Promotion and rollback of published releases'),
], [3.5,3.5], 7.7)
tech.para('Stage 1 must not implement hidden partial release behavior under the label “preview rebuild”.', after=5)
tech.page()

tech.para('17. Open technical decisions','Heading 1')
tech.table(['ID','Decision required','Evidence / owner'],[
    ('TD-OPEN-04','NetworkPolicy provider and default-deny enforcement in selected K3s profile.','Negative network tests; Security + Infrastructure.'),
    ('TD-OPEN-05','Implement and prove the fixed PostgreSQL PREVIEW_AUTH_GRANT one-time short-lived JWT → host-only preview-cookie contract, including asymmetric kid rotation, cross-replica jti CAS, expiry, exact audience, session revoke, WebSocket handshake and upstream stripping.','Route/auth/WebSocket/replay/rotation negative tests; Platform + Security. Data mechanism is fixed; this is implementation proof.'),
    ('TD-OPEN-06','Select queue worker library without changing the fixed CAPACITY_POOL row-lock, slot-sum, bounded-depth/timeout and OWNER_ROUND_ROBIN/FIFO contract.','Crash/retry/fairness/idempotency spike; Platform.'),
    ('TD-OPEN-09','Cline heartbeat warning/stuck threshold, cancellation grace and hard Job-termination timeout.','SPK-07 measurements; Platform + Product + Operations.'),
    ('TD-OPEN-10','Future team workspace, PROJECT_MEMBER roles, delegated administration and organization inheritance policy.','Separate Stage 2 product/security design; organization placement alone must never grant project access.'),
    ('TD-OPEN-11','Autosave interval/retention, Git version/SHA format, snapshot object-store implementation and long-term backup retention.','SPK-08 crash matrix plus Product/Data/Operations approval. Stage 1 explicit-Save RPO and 5-minute restore drill target are already fixed.'),
], [.9,3.45,2.65], 7.1)
tech.landscape()
tech.para('18. Traceability matrix','Heading 1')
tech.table(['Use Case','Stories / Acceptance','Business Rules','Technical realization','Verification'],[
    ('UC-01','US-01, US-02, US-27 / AC-ID-01–03',', '.join(UC_BUSINESS_RULES['UC-01']),'NestJS identity guards, Prisma AUTH_USER/RBAC, Redis opaque session and current primary organization validation.','Sign-in, CSRF/session revoke, organization placement and cross-owner project-list tests.'),
    ('UC-02','US-03–04 / AC-PRJ-01–02',', '.join(UC_BUSINESS_RULES['UC-02']),'Prisma command transaction initializes Project, Version 1 and workspace branch with command dedup + row CAS.','Idempotent project creation, initialization and owner-scoped open tests.'),
    ('UC-03','US-05–08, US-24–26 / AC-CHAT-01–06, AC-UX-02',', '.join(UC_BUSINESS_RULES['UC-03']),'NestJS persists chat/run; adapter launches UID-bound Agent Job and resumes the same durable run after refresh.','Prompt guidance, streaming/reconnect and SPK-07 UID-bound recovery tests.'),
    ('UC-04','US-09–11 / AC-CHG-01–03',', '.join(UC_BUSINESS_RULES['UC-04']),'Proposal/outcome rows and mutation lease enforce explicit decision and destructive-change approval.','Diff/preview, accept/reject/iterate and destructive-approval tests.'),
    ('UC-05','US-12–14, US-20 / AC-PRV-01–03, AC-SEC-01',', '.join(UC_BUSINESS_RULES['UC-05']),'Short-lived audience-bound JWT exchange, PostgreSQL jti CAS, host-only preview cookie, gateway stripping and generation binding.','Owner-only HTTP/WSS, cross-replica replay, revoke, HMR/Uvicorn and classifier tests.'),
    ('UC-06','US-15–16, US-28 / AC-VER-01–04',', '.join(UC_BUSINESS_RULES['UC-06']),'Snapshot service and restore journal stage, verify and atomically CAS whole-app generation pointers.','Version 1, autosave, idempotent Save and journaled history-preserving Restore tests.'),
    ('UC-07','US-17–19 / AC-LFC-01–02, AC-CAP-01',', '.join(UC_BUSINESS_RULES['UC-07']),'State owners, controller reconciliation and durable CAPACITY_POOL/RESOURCE_LEASE row locks.','Stop/scale-to-zero/resume plus global admission and bounded queue tests.'),
    ('UC-08','US-21, US-23, US-26 / AC-OPS-01, AC-UX-01, AC-CHAT-06',', '.join(UC_BUSINESS_RULES['UC-08']),'Evidence-backed observed state, typed error/audit rows and allowlisted recovery commands.','Failure recovery, exact prior-attempt proof, audit and actionable status tests.'),
    ('UC-09','US-21–22 / AC-OPS-01–02',', '.join(UC_BUSINESS_RULES['UC-09']),'Restricted AdminJS calls NestJS command services with RBAC, CSRF, property/action allowlists, CAS and audit.','Denied direct mutation, typed config bounds, last-admin protection and attributable allowlisted command tests.'),
], [1.0,2.2,1.2,2.9,2.4], 5.8)
tech.para('Every UC-01–UC-09 and its Business Rules are explicitly traceable above. Any product behavior not traceable to a PRD Feature/Story/Use Case/Business Rule/Acceptance ID is a proposal, not an approved requirement. Capability installation remains Stage 2/outside this Stage 1 contract.', after=5)
tech.save(TECH_OUT, 'Vibe-Code, technical design, NestJS, Prisma, AdminJS, Redis session, JWT, K3s, Traefik, Vite HMR, Uvicorn reload')

print({
    'prd': str(PRD_OUT),
    'technical_design': str(TECH_OUT),
    'prd_epics': len(epics),
    'prd_stories': len(stories),
    'prd_use_cases': len(use_cases),
    'prd_features': len(features),
    'security_controls': len(security_baseline),
    'acceptance_criteria': len(acceptance),
    'technical_components': len(components),
    'change_classifications': len(change_matrix),
    'spikes': len(spikes),
})
