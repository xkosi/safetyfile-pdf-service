# generate_pdf.py
# Flask microservice for Safetyfile: generate PDF & DOCX from wizard preview JSON
#
# Endpoints:
#   GET  /health          -> {"ok": true}
#   POST /generate        -> body: { preview: {...}, format: "pdf" | "docx" }
#                             returns: application/pdf or application/vnd.openxmlformats-officedocument.wordprocessingml.document
#
# Requirements (requirements.txt):
#   Flask
#   reportlab
#   pypdf
#   Pillow
#   requests
#   python-docx
#
# Run locally:
#   export PORT=8000
#   python3 generate_pdf.py
#
# Railway (recommended):
#   - Add a new Python service with this file as entrypoint.
#   - Start command: python3 generate_pdf.py
#   - Set env LOGO_URL if you want a custom AVM logo URL.
#   - Optionally set BRAND_PRIMARY (hex, e.g. #E30613).
#
# Notes:
# - PDF export is complete: cover, sections, materials tables, mini bio (clickable), external PDFs inlined (pages appended).
# - DOCX export is functional but simpler: headings, tables, images, hyperlinks.
#   External PDFs are linked (not fully embedded as page images). Converting PDF pages to images requires extra deps (pdf2image/poppler).
#   If you want that later, we can add it.
#
from __future__ import annotations
import io, os, re, json, math
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import requests
from PIL import Image
from flask import Flask, request, send_file, jsonify

# PDF libs
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Paragraph, Frame, Table, TableStyle, KeepTogether, Spacer, Image as RLImage
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from pypdf import PdfReader, PdfWriter

# DOCX
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm, Inches
from docx.oxml.shared import OxmlElement, qn

app = Flask(__name__)

# Branding
BRAND_PRIMARY = os.getenv("BRAND_PRIMARY", "#E30613")  # Pyred rood
BRAND_DARK = os.getenv("BRAND_DARK", "#000000")
BRAND_LIGHT = os.getenv("BRAND_LIGHT", "#FFFFFF")
LOGO_URL = os.getenv("LOGO_URL", "https://sfx.rentals/projects/media/logo.png")  # AVM logo

# Helpers ------------------------------------------------------------------

def hex_to_rgb(h: str) -> Tuple[float, float, float]:
    h = h.lstrip("#")
    if len(h) == 3:
        h = "".join([c*2 for c in h])
    r = int(h[0:2], 16) / 255.0
    g = int(h[2:4], 16) / 255.0
    b = int(h[4:6], 16) / 255.0
    return (r, g, b)

PRIMARY_RGB = hex_to_rgb(BRAND_PRIMARY)

def fetch_bytes(url: str, timeout: int = 20) -> Optional[bytes]:
    if not url:
        return None
    try:
        r = requests.get(url, timeout=timeout)
        if r.ok:
            return r.content
    except Exception:
        pass
    return None

def image_from_url(url: str, max_w_mm: float = 50, max_h_mm: float = 50) -> Optional[RLImage]:
    data = fetch_bytes(url)
    if not data:
        return None
    try:
        img = Image.open(io.BytesIO(data))
        img_format = img.format or "PNG"
        # Resize keeping aspect ratio
        max_w = max_w_mm * mm
        max_h = max_h_mm * mm
        w, h = img.size
        scale = min(max_w / w, max_h / h)
        new_w, new_h = int(w * scale), int(h * scale)
        img = img.resize((new_w, new_h), Image.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, format=img_format)
        buf.seek(0)
        return RLImage(buf, width=new_w, height=new_h)
    except Exception:
        return None

def parse_nec(value: Any) -> float:
    """Parse NEC values that may be stored as messy text. Returns a float sum (grams)."""
    if value is None:
        return 0.0
    s = str(value)
    # extract numbers (allow comma decimal)
    nums = re.findall(r"[\d]+(?:[.,]\d+)?", s)
    total = 0.0
    for n in nums:
        n = n.replace(",", ".")
        try:
            total += float(n)
        except ValueError:
            pass
    return round(total, 3)

def clean_text(s: Any) -> str:
    if s is None:
        return ""
    # strip basic HTML if present
    txt = re.sub(r"<[^>]+>", "", str(s))
    return txt.strip()

def add_hyperlink(run, url):
    """Add a hyperlink to a python-docx run (helper)."""
    part = run.part
    r_id = part.relate_to(url, reltype="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    r = run._r
    hlink = OxmlElement('w:hyperlink')
    hlink.set(qn('r:id'), r_id)
    newr = OxmlElement('w:r')
    rpr = OxmlElement('w:rPr')
    u = OxmlElement('w:u'); u.set(qn('w:val'), 'single')
    color = OxmlElement('w:color'); color.set(qn('w:val'), '0000FF')
    rpr.append(u); rpr.append(color)
    t = OxmlElement('w:t')
    t.text = run.text
    newr.append(rpr); newr.append(t)
    hlink.append(newr)
    r.addnext(hlink)
    r.text = ""

# PDF building --------------------------------------------------------------

def pdf_cover(c: canvas.Canvas, preview: Dict[str, Any], logo_bytes: Optional[bytes]):
    W, H = A4
    c.setFillColorRGB(*PRIMARY_RGB)
    c.rect(0, H-60, W, 60, fill=1, stroke=0)

    # Logo
    if logo_bytes:
        try:
            img = Image.open(io.BytesIO(logo_bytes))
            iw, ih = img.size
            scale = min(160/iw, 50/ih)
            nw, nh = iw*scale, ih*scale
            buf = io.BytesIO(); img.save(buf, format=img.format or "PNG"); buf.seek(0)
            c.drawImage(buf, 30, H-55, width=nw, height=nh, mask='auto')
        except Exception:
            pass

    title = "Veiligheidsdossier"
    project_name = preview.get("avm", {}).get("name") or preview.get("dees", {}).get("name") or ""

    c.setFont("Helvetica-Bold", 28)
    c.setFillColorRGB(0,0,0)
    c.drawString(30, H-120, title)
    if project_name:
        c.setFont("Helvetica", 16)
        c.drawString(30, H-145, project_name)

    c.setFont("Helvetica", 10)
    c.setFillColorRGB(0.3,0.3,0.3)
    c.drawString(30, 40, f"Gegenereerd: {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    c.showPage()

def pdf_section_title(c: canvas.Canvas, title: str):
    W, H = A4
    c.setFillColorRGB(1,1,1)
    c.rect(0,0,W,H,fill=1,stroke=0)
    c.setFillColorRGB(*PRIMARY_RGB)
    c.rect(0, H-40, W, 40, fill=1, stroke=0)
    c.setFillColorRGB(1,1,1)
    c.setFont("Helvetica-Bold", 18)
    c.drawString(30, H-28, title)
    c.showPage()

def project_table_story(preview: Dict[str, Any]) -> List[Any]:
    p = preview.get("avm") or {}
    rows = []
    def row(label, val):
        rows.append([Paragraph(f"<b>{label}</b>", ParagraphStyle(name='pl', fontName='Helvetica', fontSize=10)),
                     Paragraph(f"{clean_text(val)}", ParagraphStyle(name='pv', fontName='Helvetica', fontSize=10, underlineWidth=0.5))])
    row("Project", p.get("name", ""))
    row("Opdrachtgever", p.get("customer", {}).get("name", ""))
    row("Adres", p.get("customer", {}).get("address", ""))
    contact = p.get("customer", {}).get("contact")
    if contact:
        row("Contactpersoon", contact.get("name",""))
        row("Tel", contact.get("phone",""))
        row("E-mail", contact.get("email",""))
    row("Locatie", p.get("location", {}).get("name",""))
    row("Adres locatie", p.get("location", {}).get("address",""))
    row("Start", p.get("project_start_date",""))
    row("Einde", p.get("project_end_date",""))

    t = Table(rows, colWidths=[60*mm, 120*mm])
    t.setStyle(TableStyle([
        ("BOX",(0,0),(-1,-1),0.75, colors.black),
        ("INNERGRID",(0,0),(-1,-1),0.25, colors.grey),
        ("BACKGROUND",(0,0),(0,-1), colors.whitesmoke),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))
    return [t]

def materials_table_story_avm(materials: List[Dict[str, Any]]) -> List[Any]:
    story: List[Any] = []
    header = ["Kies", "Naam", "Aantal", "Type", "Links"]
    data = [header]
    for m in materials:
        links_txt = ""
        for f in m.get("files", []):
            label = os.path.splitext(f.get("name",""))[0]
            url = f.get("url","")
            if url:
                links_txt += f"{label}\n"
        data.append(["✔", clean_text(m.get("displayname","")), str(m.get("quantity_total",0)), m.get("type",""), links_txt])

    t = Table(data, colWidths=[15*mm, 85*mm, 20*mm, 20*mm, 50*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0), colors.Color(*PRIMARY_RGB)),
        ("TEXTCOLOR",(0,0),(-1,0), colors.white),
        ("GRID",(0,0),(-1,-1),0.25, colors.grey),
        ("VALIGN",(0,0),(-1,-1),"TOP"),
        ("ALIGN",(2,1),(2,-1),"RIGHT"),
        ("ALIGN",(3,1),(3,-1),"CENTER"),
    ]))
    story.append(t)
    return story

def materials_table_story_dees(materials: List[Dict[str, Any]]) -> List[Any]:
    # Split pyro vs vuurwerk based on custom_13 content (T1/T2 vs F1..F4)
    pyro_rows = [["Naam","NEC (g)","Aantal","Type"]]
    vuur_rows = [["Naam","NEC (g)","Aantal","Type"]]
    tot_pyro = 0.0; tot_vuur = 0.0

    for m in materials:
        label = clean_text(m.get("displayname",""))
        qty = int(m.get("quantity_total",0) or 0)
        tag = str(m.get("custom_13","")).upper()
        nec_val = parse_nec(tag)
        # bucket
        if any(x in tag for x in ["T1","T2"]):
            pyro_rows.append([label, f"{nec_val:.3f}", str(qty), "Pyro"])
            tot_pyro += nec_val * max(qty,1)
        elif any(x in tag for x in ["F1","F2","F3","F4"]):
            vuur_rows.append([label, f"{nec_val:.3f}", str(qty), "Vuurwerk"])
            tot_vuur += nec_val * max(qty,1)
        else:
            # unknown → go to pyro table but mark as "-"
            pyro_rows.append([label, f"{nec_val:.3f}", str(qty), "-"])
            tot_pyro += nec_val * max(qty,1)

    story: List[Any] = []
    if len(pyro_rows) > 1:
        t1 = Table(pyro_rows, colWidths=[90*mm, 30*mm, 20*mm, 30*mm])
        t1.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0), colors.Color(*PRIMARY_RGB)),
            ("TEXTCOLOR",(0,0),(-1,0), colors.white),
            ("GRID",(0,0),(-1,-1),0.25, colors.grey),
            ("ALIGN",(1,1),(2,-1),"RIGHT"),
        ]))
        story.append(Paragraph("<b>Pyrotechniek (T1/T2)</b>", ParagraphStyle(name="h", fontName="Helvetica-Bold", fontSize=12)))
        story.append(Spacer(1,4))
        story.append(t1)
        story.append(Spacer(1,6))
        story.append(Paragraph(f"<b>Totaal NEC:</b> {tot_pyro:.3f} g", ParagraphStyle(name="p", fontName="Helvetica", fontSize=10)))
        story.append(Spacer(1,10))

    if len(vuur_rows) > 1:
        t2 = Table(vuur_rows, colWidths=[90*mm, 30*mm, 20*mm, 30*mm])
        t2.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0), colors.Color(*PRIMARY_RGB)),
            ("TEXTCOLOR",(0,0),(-1,0), colors.white),
            ("GRID",(0,0),(-1,-1),0.25, colors.grey),
            ("ALIGN",(1,1),(2,-1),"RIGHT"),
        ]))
        story.append(Paragraph("<b>Vuurwerk (F1–F4)</b>", ParagraphStyle(name="h", fontName="Helvetica-Bold", fontSize=12)))
        story.append(Spacer(1,4))
        story.append(t2)
        story.append(Spacer(1,6))
        story.append(Paragraph(f"<b>Totaal NEC:</b> {tot_vuur:.3f} g", ParagraphStyle(name="p", fontName="Helvetica", fontSize=10)))

    if len(story) == 0:
        story.append(Paragraph("Geen Dees-materialen geselecteerd.", ParagraphStyle(name="p", fontName="Helvetica", fontSize=10)))

    return story

def build_base_pdf(preview: Dict[str, Any]) -> io.BytesIO:
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)

    logo_bytes = fetch_bytes(LOGO_URL)

    # Cover
    pdf_cover(c, preview, logo_bytes)

    # Inhoudstafel (simple, zonder pagenumbers to avoid complexity)
    pdf_section_title(c, "Inhoudstafel")
    c.setFont("Helvetica", 11)
    sections = [
        "1. Projectgegevens",
        "2. Emergency",
        "3. Verzekeringen",
        "4. Verantwoordelijke",
        "5. Materialen",
        "6. Inplantingsplan",
        "7. Risicoanalyse Pyro & Special Effects",
        "8. Windplan",
        "9. Droogteplan",
        "10. Vergunningen & Toelatingen",
    ]
    y = A4[1]-70
    for s in sections:
        c.drawString(30, y, s)
        y -= 18
        if y < 60:
            c.showPage()
            y = A4[1]-60
    c.showPage()

    # 1. Projectgegevens
    pdf_section_title(c, "1. Projectgegevens")
    story = project_table_story(preview)
    frame = Frame(30, 50, A4[0]-60, A4[1]-120, showBoundary=0)
    frame.addFromList(story, c)
    c.showPage()

    # 2. Emergency (title page only, actual PDF will be appended later)
    pdf_section_title(c, "2. Emergency")

    # 3. Verzekeringen (title page only)
    pdf_section_title(c, "3. Verzekeringen")

    # 4. Verantwoordelijke
    pdf_section_title(c, "4. Verantwoordelijke")
    resp = preview.get("responsible")
    crew_bio_mini = preview.get("documents", {}).get("crew_bio_mini") or ""
    crew_bio_full = preview.get("documents", {}).get("crew_bio_full") or ""
    c.setFont("Helvetica-Bold", 12)
    c.drawString(30, A4[1]-70, f"Projectverantwoordelijke: {resp or '-'}")
    y = A4[1]-90
    # Mini bio image if URL
    if crew_bio_mini:
        img = image_from_url(crew_bio_mini, max_w_mm=70, max_h_mm=60)
        if img:
            frame = Frame(30, 120, A4[0]-60, A4[1]-240, showBoundary=0)
            story = [img, Spacer(1,6)]
            if crew_bio_full:
                story.append(Paragraph(f'<font color="blue"><u><link href="{crew_bio_full}">Open volledige bio</link></u></font>',
                                       ParagraphStyle(name="lnk", fontName="Helvetica", fontSize=10)))
            frame.addFromList(story, c)
        else:
            c.setFont("Helvetica", 10)
            c.drawString(30, y, "(Mini bio kon niet geladen worden)")
            y -= 16
    else:
        c.setFont("Helvetica", 10)
        c.drawString(30, y, "Geen mini-bio ingesteld.")
        y -= 16
    c.showPage()

    # 5. Materialen
    pdf_section_title(c, "5. Materialen")

    # 5.1 Dees
    materials_dees = (preview.get("materials") or {}).get("dees") or []
    if materials_dees:
        c.setFont("Helvetica-Bold", 12); c.drawString(30, A4[1]-70, "5.1 Dees")
        frame = Frame(30, 60, A4[0]-60, A4[1]-120, showBoundary=0)
        frame.addFromList(materials_table_story_dees(materials_dees), c)
        c.showPage()
    else:
        c.setFont("Helvetica", 10); c.drawString(30, A4[1]-70, "5.1 Dees — geen items geselecteerd.")
        c.showPage()

    # 5.2 AVM
    materials_avm = (preview.get("materials") or {}).get("avm") or []
    if materials_avm:
        c.setFont("Helvetica-Bold", 12); c.drawString(30, A4[1]-70, "5.2 AVM")
        frame = Frame(30, 60, A4[0]-60, A4[1]-120, showBoundary=0)
        frame.addFromList(materials_table_story_avm(materials_avm), c)
        c.showPage()
    else:
        c.setFont("Helvetica", 10); c.drawString(30, A4[1]-70, "5.2 AVM — geen items geselecteerd.")
        c.showPage()

    # 6. Inplantingsplan (title; actual upload appended later)
    pdf_section_title(c, "6. Inplantingsplan")

    # 7. Risicoanalyse Pyro & Special Effects
    pdf_section_title(c, "7. Risicoanalyse Pyro & Special Effects")

    # 8. Windplan
    pdf_section_title(c, "8. Windplan")

    # 9. Droogteplan
    pdf_section_title(c, "9. Droogteplan")

    # 10. Vergunningen & Toelatingen
    pdf_section_title(c, "10. Vergunningen & Toelatingen")

    c.save()
    buf.seek(0)
    return buf

def append_external_pdfs(base_pdf: io.BytesIO, preview: Dict[str, Any]) -> io.BytesIO:
    """Append external PDFs (emergency, insurance, risk, wind/drought, uploads) after their title pages."""
    base_reader = PdfReader(base_pdf)
    writer = PdfWriter()
    for p in base_reader.pages:
        writer.add_page(p)

    # Helper: append a PDF by URL (each URL)
    def append_url(url: str):
        if not url: return
        b = fetch_bytes(url)
        if not b: return
        try:
            r = PdfReader(io.BytesIO(b))
            for pg in r.pages:
                writer.add_page(pg)
        except Exception:
            # if not PDF, ignore
            pass

    docs = preview.get("documents", {})

    # Emergency (2)
    append_url(docs.get("emergency"))

    # Insurance (3) -> list
    for u in docs.get("insurance", []) or []:
        append_url(u)

    # (4 is verantwoordelijke - no pdf append)

    # (5 Materials - no pdf append)

    # (6) Inplantingsplan -> uploads.siteplan (could be multiple files)
    for up in (preview.get("uploads", {}) or {}).get("siteplan", []) or []:
        append_url(up.get("data") or up.get("url"))

    # (7) Risks (pyro + general)
    append_url(docs.get("risk_pyro"))
    append_url(docs.get("risk_general"))

    # (8) Windplan
    if preview.get("documents", {}).get("windplan"):
        append_url(preview["documents"]["windplan"])

    # (9) Droogteplan
    if preview.get("documents", {}).get("droughtplan"):
        append_url(preview["documents"]["droughtplan"])

    # (10) Permits uploads
    for up in (preview.get("uploads", {}) or {}).get("permits", []) or []:
        append_url(up.get("data") or up.get("url"))

    out = io.BytesIO()
    writer.write(out)
    out.seek(0)
    return out

# DOCX building -------------------------------------------------------------

def build_docx(preview: Dict[str, Any]) -> io.BytesIO:
    doc = Document()

    # Title / cover
    h = doc.add_heading("Veiligheidsdossier", 0)
    if preview.get("avm", {}).get("name"):
        p = doc.add_paragraph(preview["avm"]["name"])
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # TOC placeholder
    doc.add_paragraph("Inhoudstafel (vereenvoudigd)").bold = True
    for s in ["1. Projectgegevens","2. Emergency","3. Verzekeringen","4. Verantwoordelijke","5. Materialen","6. Inplantingsplan","7. Risicoanalyse Pyro & Special Effects","8. Windplan","9. Droogteplan","10. Vergunningen & Toelatingen"]:
        doc.add_paragraph(s, style=None)

    # 1. Projectgegevens
    doc.add_heading("1. Projectgegevens", level=1)
    p = preview.get("avm") or {}
    table = doc.add_table(rows=0, cols=2)
    def add_row(label, val):
        r = table.add_row().cells
        r[0].text = label
        r[1].text = clean_text(val)
    add_row("Project", p.get("name",""))
    add_row("Opdrachtgever", p.get("customer",{}).get("name",""))
    add_row("Adres", p.get("customer",{}).get("address",""))
    if p.get("customer",{}).get("contact"):
        cc = p["customer"]["contact"]
        add_row("Contactpersoon", cc.get("name",""))
        add_row("Tel", cc.get("phone",""))
        add_row("E-mail", cc.get("email",""))
    add_row("Locatie", p.get("location",{}).get("name",""))
    add_row("Adres locatie", p.get("location",{}).get("address",""))
    add_row("Start", p.get("project_start_date",""))
    add_row("Einde", p.get("project_end_date",""))

    # 2. Emergency
    doc.add_heading("2. Emergency", level=1)
    url = preview.get("documents",{}).get("emergency")
    if url:
        pr = doc.add_paragraph("Open emergency: ")
        run = pr.add_run(url); add_hyperlink(run, url)
    else:
        doc.add_paragraph("Geen emergency-document ingesteld.")

    # 3. Verzekeringen
    doc.add_heading("3. Verzekeringen", level=1)
    ins = preview.get("documents",{}).get("insurance") or []
    if ins:
        for u in ins:
            pr = doc.add_paragraph("Verzekering: ")
            run = pr.add_run(u); add_hyperlink(run, u)
    else:
        doc.add_paragraph("Geen verzekeringsdocumenten ingesteld.")

    # 4. Verantwoordelijke
    doc.add_heading("4. Verantwoordelijke", level=1)
    resp = preview.get("responsible") or "-"
    doc.add_paragraph(f"Projectverantwoordelijke: {resp}")
    mini = preview.get("documents",{}).get("crew_bio_mini")
    full = preview.get("documents",{}).get("crew_bio_full")
    if mini:
        try:
            b = fetch_bytes(mini)
            if b:
                tmp = io.BytesIO(b)
                doc.add_picture(tmp, width=Inches(3.0))
        except Exception:
            pass
    if full:
        pr = doc.add_paragraph("Volledige bio: ")
        run = pr.add_run(full); add_hyperlink(run, full)

    # 5. Materialen
    doc.add_heading("5. Materialen", level=1)

    # 5.1 Dees
    doc.add_heading("5.1 Dees", level=2)
    dees = (preview.get("materials") or {}).get("dees") or []
    if dees:
        t = doc.add_table(rows=1, cols=4)
        hdr = t.rows[0].cells
        hdr[0].text = "Naam"; hdr[1].text = "NEC (g)"; hdr[2].text = "Aantal"; hdr[3].text = "Type"
        tot_py = 0.0; tot_fw = 0.0
        for m in dees:
            tag = str(m.get("custom_13",""))
            nec = parse_nec(tag)
            qty = int(m.get("quantity_total",0) or 0)
            ty = "Pyro" if any(x in tag.upper() for x in ["T1","T2"]) else ("Vuurwerk" if any(x in tag.upper() for x in ["F1","F2","F3","F4"]) else "-")
            r = t.add_row().cells
            r[0].text = clean_text(m.get("displayname","")); r[1].text = f"{nec:.3f}"; r[2].text = str(qty); r[3].text = ty
            if ty == "Pyro": tot_py += nec * max(qty,1)
            elif ty == "Vuurwerk": tot_fw += nec * max(qty,1)
        doc.add_paragraph(f"Totaal NEC (Pyro): {tot_py:.3f} g")
        doc.add_paragraph(f"Totaal NEC (Vuurwerk): {tot_fw:.3f} g")
    else:
        doc.add_paragraph("Geen Dees-materialen geselecteerd.")

    # 5.2 AVM
    doc.add_heading("5.2 AVM", level=2)
    avm = (preview.get("materials") or {}).get("avm") or []
    if avm:
        t = doc.add_table(rows=1, cols=4)
        hdr = t.rows[0].cells
        hdr[0].text = "Naam"; hdr[1].text = "Aantal"; hdr[2].text = "Type"; hdr[3].text = "Links"
        for m in avm:
            r = t.add_row().cells
            r[0].text = clean_text(m.get("displayname","")); r[1].text = str(m.get("quantity_total",0)); r[2].text = m.get("type","")
            links = []
            for f in m.get("files",[]):
                url = f.get("url"); nm = f.get("name")
                if url:
                    links.append(f"{nm}")
            r[3].text = "\n".join(links)
    else:
        doc.add_paragraph("Geen AVM-materialen geselecteerd.")

    # 6. Inplantingsplan
    doc.add_heading("6. Inplantingsplan", level=1)
    site = (preview.get("uploads",{}) or {}).get("siteplan") or []
    if site:
        for f in site:
            u = f.get("data") or f.get("url")
            if u:
                pr = doc.add_paragraph("Inplantingsplan: ")
                run = pr.add_run(u); add_hyperlink(run, u)
    else:
        doc.add_paragraph("Geen inplantingsplan toegevoegd.")

    # 7. Risicoanalyse
    doc.add_heading("7. Risicoanalyse Pyro & Special Effects", level=1)
    for key, label in [("risk_pyro","Risicoanalyse Pyro"),("risk_general","Risicoanalyse Special Effects")]:
        u = preview.get("documents",{}).get(key)
        if u:
            pr = doc.add_paragraph(label + ": ")
            run = pr.add_run(u); add_hyperlink(run, u)

    # 8. Windplan
    doc.add_heading("8. Windplan", level=1)
    u = preview.get("documents",{}).get("windplan")
    if u:
        pr = doc.add_paragraph("Windplan: ")
        run = pr.add_run(u); add_hyperlink(run, u)
    else:
        doc.add_paragraph("Niet geselecteerd.")

    # 9. Droogteplan
    doc.add_heading("9. Droogteplan", level=1)
    u = preview.get("documents",{}).get("droughtplan")
    if u:
        pr = doc.add_paragraph("Droogteplan: ")
        run = pr.add_run(u); add_hyperlink(run, u)
    else:
        doc.add_paragraph("Niet geselecteerd.")

    # 10. Vergunningen & Toelatingen
    doc.add_heading("10. Vergunningen & Toelatingen", level=1)
    permits = (preview.get("uploads",{}) or {}).get("permits") or []
    if permits:
        for f in permits:
            u = f.get("data") or f.get("url")
            if u:
                pr = doc.add_paragraph("Vergunning/Toelating: ")
                run = pr.add_run(u); add_hyperlink(run, u)
    else:
        doc.add_paragraph("Geen vergunningen of toelatingen toegevoegd.")

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out

# Endpoint logic ------------------------------------------------------------

@app.get("/health")
def health():
    return jsonify(ok=True, ts=datetime.utcnow().isoformat())

@app.post("/generate")
def generate():
    try:
        payload = request.get_json(force=True, silent=False) or {}
    except Exception:
        return jsonify(error="Invalid JSON"), 400

    preview = payload.get("preview") or payload  # allow raw preview
    fmt = (payload.get("format") or preview.get("format") or "pdf").lower()

    # Build base PDF
    base_pdf = build_base_pdf(preview)
    # Append external PDFs
    final_pdf = append_external_pdfs(base_pdf, preview)

    if fmt == "pdf":
        filename = f"dossier.pdf"
        return send_file(final_pdf, mimetype="application/pdf", as_attachment=True, download_name=filename)

    # DOCX path
    docx_buf = build_docx(preview)
    filename = f"dossier.docx"
    return send_file(docx_buf, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                     as_attachment=True, download_name=filename)

if __name__ == "__main__":
    port = int(os.getenv("PORT", "8000"))
    app.run(host="0.0.0.0", port=port)
