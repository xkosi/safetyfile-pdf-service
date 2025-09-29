
import io
import os
import json
import math
import datetime
import requests

from flask import Flask, request, send_file, jsonify

# PDF building
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph, Table, TableStyle, Spacer, SimpleDocTemplate, PageBreak
from reportlab.lib.enums import TA_LEFT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from PyPDF2 import PdfReader, PdfWriter, Transformation

# DOCX building
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT, WD_SECTION_START
from docx.oxml.ns import qn

# Render PDF -> image for DOCX
import fitz  # PyMuPDF

app = Flask(__name__)

# --------------------- helpers ---------------------

RED = colors.Color(0.80, 0.00, 0.00)
DARK = colors.black
LIGHTGRAY = colors.HexColor("#f2f2f2")

MARGIN_LEFT = 20*mm
MARGIN_RIGHT = 20*mm
MARGIN_TOP = 20*mm
MARGIN_BOTTOM = 20*mm
CONTENT_WIDTH = A4[0] - MARGIN_LEFT - MARGIN_RIGHT

def fmt_date(s):
    if not s:
        return ""
    try:
        # try ISO
        dt = datetime.datetime.fromisoformat(s.replace("Z","+00:00"))
    except Exception:
        try:
            dt = datetime.datetime.strptime(s[:10], "%Y-%m-%d")
        except Exception:
            return str(s)
    return dt.strftime("%d/%m/%Y")

def safe_get(d, path, default=""):
    cur = d
    for p in path.split("."):
        if cur is None:
            return default
        if isinstance(cur, dict):
            cur = cur.get(p)
        else:
            return default
    return cur if cur is not None else default

def fetch_binary(url_or_datauri):
    """Return bytes of a URL (Dropbox dl=1 supported) or data: URI."""
    if not url_or_datauri:
        return None
    if isinstance(url_or_datauri, bytes):
        return url_or_datauri
    if str(url_or_datauri).startswith("data:"):
        try:
            head, b64 = url_or_datauri.split(",", 1)
            import base64
            return base64.b64decode(b64)
        except Exception:
            return None
    try:
        r = requests.get(url_or_datauri, timeout=30)
        if r.ok:
            return r.content
    except Exception:
        return None
    return None

def add_header_bar(c, title, y_offset=0):
    """Draw red header bar with white title at current page top."""
    bar_h = 12*mm
    x = MARGIN_LEFT
    y = A4[1] - MARGIN_TOP - bar_h - y_offset
    c.setFillColor(RED)
    c.rect(x, y, CONTENT_WIDTH, bar_h, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x+4*mm, y + 4*mm, title)

def draw_kv_table(c, rows, start_y):
    """Draw a simple 2-col table (label,value) starting at start_y; returns last y used."""
    col1 = 40*mm
    col2 = CONTENT_WIDTH - col1
    row_h = 8*mm
    x = MARGIN_LEFT
    y = start_y
    c.setStrokeColor(colors.grey)
    c.setLineWidth(0.4)
    for label, value in rows:
        c.setFillColor(LIGHTGRAY)
        c.rect(x, y-row_h, col1, row_h, fill=1, stroke=1)
        c.setFillColor(colors.white)
        # top border already drawn as stroke
        c.setFillColor(colors.black)
        c.setFont("Helvetica", 9)
        c.drawString(x+2*mm, y-row_h+2.5*mm, str(label or ""))
        c.setFillColor(colors.black)
        c.rect(x+col1, y-row_h, col2, row_h, fill=0, stroke=1)
        # wrap not handled; simple drawString
        txt = str(value or "")
        c.drawString(x+col1+2*mm, y-row_h+2.5*mm, txt)
        y -= row_h
    return y

def add_materials_table(c, items, start_y):
    col_defs = [
        ("Naam", 85*mm),
        ("Aantal", 20*mm),
        ("Type", 25*mm),
        ("Links", CONTENT_WIDTH - (85*mm + 20*mm + 25*mm)),
    ]
    head_h = 8*mm
    row_h = 7*mm
    x = MARGIN_LEFT
    y = start_y

    # header
    c.setFillColor(RED)
    c.rect(x, y-head_h, CONTENT_WIDTH, head_h, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 9)
    cx = x
    for title, w in col_defs:
        c.drawString(cx+2*mm, y-head_h+2.5*mm, title)
        cx += w
    y -= head_h

    c.setFont("Helvetica", 8)
    c.setFillColor(colors.black)
    for it in items:
        cx = x
        vals = [
            it.get("displayname",""),
            str(it.get("quantity_total","") or ""),
            str(it.get("type","") or ""),
            build_links_text(it.get("links") or {}),
        ]
        c.setFillColor(colors.white)
        c.setStrokeColor(colors.grey)
        c.rect(x, y-row_h, CONTENT_WIDTH, row_h, fill=0, stroke=1)
        for val, (_, w) in zip(vals, col_defs):
            c.drawString(cx+2*mm, y-row_h+2.2*mm, val[:120])
            cx += w
        y -= row_h
        if y < 25*mm:
            c.showPage()
            add_header_bar(c, "5. Materialen (vervolg)")
            y = A4[1] - MARGIN_TOP - 20*mm
    return y

def build_links_text(links):
    out = []
    for key in ("ce","manual","msds"):
        v = links.get(key)
        if v:
            out.append(f"{key.upper()}: ja")
    return ", ".join(out)

def merge_pdf_firstpage_under_header(base_pdf_bytes, external_pdf_bytes, header_height_mm=20):
    """
    Take base PDF first page (with header drawn), and place the first page of external PDF scaled
    to fit below header inside margins. Append remaining pages afterwards.
    """
    base_reader = PdfReader(io.BytesIO(base_pdf_bytes))
    ext_reader = PdfReader(io.BytesIO(external_pdf_bytes))

    writer = PdfWriter()
    base_page = base_reader.pages[0]
    writer.add_page(base_page)

    # Place ext first page onto base first page
    page = writer.pages[0]
    # Compute target rect
    page_width = float(page.mediabox.width)
    page_height = float(page.mediabox.height)
    left = MARGIN_LEFT
    right = page_width - MARGIN_RIGHT
    top = page_height - MARGIN_TOP - (header_height_mm*mm) - 5  # a bit of spacing
    bottom = MARGIN_BOTTOM
    target_w = right - left
    target_h = top - bottom
    src_page = ext_reader.pages[0]
    src_w = float(src_page.mediabox.width)
    src_h = float(src_page.mediabox.height)
    scale = min(target_w/src_w, target_h/src_h)
    tx = left
    ty = bottom
    t = Transformation().scale(scale).translate(tx/scale, ty/scale)
# Compatibiliteit PyPDF2 2.x / 3.x
if hasattr(page, "merge_transformed_page"):
    page.merge_transformed_page(src_page, t)  # nieuwe naam (sinds PyPDF2 >= 3.0)
elif hasattr(page, "mergeTransformedPage"):
    page.mergeTransformedPage(src_page, t)    # oude naam
else:
    raise AttributeError("Geen geldige merge-functie in PyPDF2 gevonden")


    # Append the rest pages of base (if any beyond page 1)
    for i in range(1, len(base_reader.pages)):
        writer.add_page(base_reader.pages[i])

    # Append remaining pages of the external PDF
    for i in range(1, len(ext_reader.pages)):
        writer.add_page(ext_reader.pages[i])

    buf = io.BytesIO()
    writer.write(buf)
    buf.seek(0)
    return buf.getvalue()

# ---------------- PDF generator ----------------

def build_pdf(preview: dict) -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)

    # Title page
    add_header_bar(c, "Veiligheidsdossier")
    c.setFont("Helvetica-Bold", 18)
    c.setFillColor(colors.black)
    title = safe_get(preview, "avm.name") or "Project"
    c.drawString(MARGIN_LEFT, A4[1]-MARGIN_TOP-20*mm, title)

    # generated timestamp
    c.setFont("Helvetica", 8)
    c.setFillColor(colors.black)
    c.drawString(MARGIN_LEFT, MARGIN_BOTTOM, "Gegenereerd: " + datetime.datetime.now().strftime("%d/%m/%Y %H:%M"))

    # Contents header for next page
    c.setFont("Helvetica-Bold", 10)
    c.showPage()

    # 1. Projectgegevens
    add_header_bar(c, "1. Projectgegevens")
    y = A4[1] - MARGIN_TOP - 20*mm

    avm = preview.get("avm") or {}
    cust = avm.get("customer") or {}
    loc = avm.get("location") or {}
    contact = cust.get("contact") or {}

    rows = [
        ("Project", avm.get("name","")),
        ("Opdrachtgever", cust.get("name","")),
        ("Adres", cust.get("address","")),
        ("Contactpersoon", contact.get("name","")),
        ("Tel", contact.get("phone","")),
        ("E-mail", contact.get("email","")),
        ("Locatie", loc.get("name","")),
        ("Adres locatie", loc.get("address","")),
        ("Start", fmt_date(avm.get("start_date") or avm.get("project_start_date"))),
        ("Einde", fmt_date(avm.get("end_date") or avm.get("project_end_date"))),
    ]
    y = draw_kv_table(c, rows, y)

    # Verantwoordelijke (mini bio img if available)
    c.setFont("Helvetica-Bold", 11)
    c.setFillColor(colors.black)
    y -= 6*mm
    c.drawString(MARGIN_LEFT, y, "Projectverantwoordelijke")
    y -= 6*mm
    c.setFont("Helvetica", 10)
    resp_name = preview.get("responsible") or ""
    c.drawString(MARGIN_LEFT, y, f"Naam: {resp_name}")
    y -= 40
    mini_url = safe_get(preview, "documents.crew_bio_mini")
    if mini_url:
        img_bytes = fetch_binary(mini_url)
        if img_bytes:
            try:
                # draw image approx 35mm height
                img_h = 35*mm
                c.drawImage(io.BytesIO(img_bytes), MARGIN_LEFT, y, height=img_h, preserveAspectRatio=True, mask='auto')
                y -= img_h + 6
            except Exception:
                pass

    c.showPage()

    # 2. Emergency
    add_header_bar(c, "2. Emergency")
    emergency_url = safe_get(preview, "documents.emergency")
    c.showPage()  # placeholder page to merge onto
    emergency_cover = c.getpdfdata()  # not correct; getpdfdata flushes whole doc. We'll construct covers separately.

    # To avoid complexity with partial getpdfdata, we will build sections as separate PDFs and merge at the end.
    c.save()
    base_pdf = buf.getvalue()

    # We'll build the rest using a PDF writer approach:
    writer = PdfWriter()
    base_reader = PdfReader(io.BytesIO(base_pdf))
    for p in base_reader.pages:
        writer.add_page(p)

    def add_pdf_section_with_cover(title, url):
        if not url:
            return
        # 1) build a single-page PDF cover with the header bar 'title'
        cover_buf = io.BytesIO()
        cc = canvas.Canvas(cover_buf, pagesize=A4)
        add_header_bar(cc, title)
        cc.save()
        cover_pdf = cover_buf.getvalue()
        ext_bytes = fetch_binary(url)
        if not ext_bytes:
            # add only the cover
            cover_reader = PdfReader(io.BytesIO(cover_pdf))
            writer.add_page(cover_reader.pages[0])
            return

        # merge external first page under header and append rest
        merged = merge_pdf_firstpage_under_header(cover_pdf, ext_bytes, header_height_mm=12)
        merged_reader = PdfReader(io.BytesIO(merged))
        for p in merged_reader.pages:
            writer.add_page(p)

    # Append sections with external PDFs
    add_pdf_section_with_cover("2. Emergency", safe_get(preview, "documents.emergency"))

    # 3. Verzekeringen (can be list)
    ins_list = preview.get("documents", {}).get("insurance", []) or []
    if ins_list:
        for i, url in enumerate(ins_list, start=1):
            title = "3. Verzekeringen" if i == 1 else "3. Verzekeringen (vervolg)"
            add_pdf_section_with_cover(title, url)

    # 4. Verantwoordelijke (full CV NOT embedded – intentionally skipped)

    # 5. Materialen
    # Build a small PDF with the tables and append
    def materials_pdf(preview):
        mbuf = io.BytesIO()
        c2 = canvas.Canvas(mbuf, pagesize=A4)
        add_header_bar(c2, "5. Materialen")
        y = A4[1] - MARGIN_TOP - 20*mm

        # 5.1 Pyrotechnische materialen (Dees)
        c2.setFont("Helvetica-Bold", 11)
        c2.setFillColor(colors.black)
        c2.drawString(MARGIN_LEFT, y-6*mm, "5.1 Pyrotechnische materialen (Dees)")
        y -= 12*mm
        dees = (preview.get("materials") or {}).get("dees") or []
        if dees:
            y = add_materials_table(c2, dees, y)
        else:
            c2.setFont("Helvetica", 9)
            c2.drawString(MARGIN_LEFT, y, "Geen items geselecteerd.")
            y -= 10*mm

        # 5.2 Speciale effecten (AVM)
        c2.setFont("Helvetica-Bold", 11)
        c2.drawString(MARGIN_LEFT, y-6*mm, "5.2 Speciale effecten (AVM)")
        y -= 12*mm
        avm_items = (preview.get("materials") or {}).get("avm") or []
        if avm_items:
            y = add_materials_table(c2, avm_items, y)
        else:
            c2.setFont("Helvetica", 9)
            c2.drawString(MARGIN_LEFT, y, "Geen items geselecteerd.")
            y -= 10*mm

        c2.save()
        return mbuf.getvalue()

    mats_pdf = materials_pdf(preview)
    mats_reader = PdfReader(io.BytesIO(mats_pdf))
    for p in mats_reader.pages:
        writer.add_page(p)

    # 6. Inplantingsplan (uploads.siteplan) – put if provided
    siteplans = preview.get("uploads", {}).get("siteplan", []) or []
    if siteplans:
        for i, f in enumerate(siteplans, start=1):
            title = "6. Inplantingsplan" if i == 1 else "6. Inplantingsplan (vervolg)"
            data = fetch_binary(f.get("data") or f.get("url") or "")
            if data:
                merged = merge_pdf_firstpage_under_header(_cover_pdf(title), data, header_height_mm=12)
                mr = PdfReader(io.BytesIO(merged))
                for p in mr.pages:
                    writer.add_page(p)

    # 7. Risicoanalyse
    avm_has = bool((preview.get("materials") or {}).get("avm"))
    dees_has = bool((preview.get("materials") or {}).get("dees"))
    if avm_has or dees_has:
        # Add a simple text page
        rbuf = io.BytesIO()
        c3 = canvas.Canvas(rbuf, pagesize=A4)
        add_header_bar(c3, "7. Risicoanalyse")
        c3.setFont("Helvetica", 10)
        c3.setFillColor(colors.black)
        c3.drawString(MARGIN_LEFT, A4[1]-MARGIN_TOP-30*mm, "Deze risicoanalyse is opgesteld op basis van de geselecteerde materialen.")
        c3.save()
        rr = PdfReader(io.BytesIO(rbuf.getvalue()))
        for p in rr.pages:
            writer.add_page(p)

    # 8. Windplan
    add_pdf_section_with_cover("8. Windplan", safe_get(preview, "documents.windplan"))

    # 9. Droogteplan
    add_pdf_section_with_cover("9. Droogteplan", safe_get(preview, "documents.droughtplan"))

    # 10. Vergunningen
    permits = preview.get("uploads", {}).get("permits", []) or []
    if permits:
        for i, f in enumerate(permits, start=1):
            title = "10. Vergunningen" if i == 1 else "10. Vergunningen (vervolg)"
            data = fetch_binary(f.get("data") or f.get("url") or "")
            if data:
                merged = merge_pdf_firstpage_under_header(_cover_pdf(title), data, header_height_mm=12)
                mr = PdfReader(io.BytesIO(merged))
                for p in mr.pages:
                    writer.add_page(p)

    # Finalize
    out = io.BytesIO()
    writer.write(out)
    out.seek(0)
    return out.read()

def _cover_pdf(title: str) -> bytes:
    b = io.BytesIO()
    cc = canvas.Canvas(b, pagesize=A4)
    add_header_bar(cc, title)
    cc.save()
    return b.getvalue()

# ---------------- DOCX generator ----------------

def pdf_first_page_to_png_bytes(pdf_bytes: bytes, zoom=2.0) -> bytes:
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        if doc.page_count == 0:
            return None
        page = doc.load_page(0)
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        return pix.tobytes("png")
    except Exception:
        return None

def build_docx(preview: dict) -> bytes:
    doc = Document()
    # narrow margins
    for s in doc.sections:
        s.top_margin = Inches(0.6)
        s.bottom_margin = Inches(0.6)
        s.left_margin = Inches(0.7)
        s.right_margin = Inches(0.7)

    def h(title):
        p = doc.add_paragraph()
        run = p.add_run(title)
        run.bold = True
        p.style = doc.styles['Heading 1']

    # Title
    h("Veiligheidsdossier")
    doc.add_paragraph(preview.get("avm", {}).get("name",""))

    # 1 Projectgegevens
    h("1. Projectgegevens")
    avm = preview.get("avm") or {}
    cust = avm.get("customer") or {}
    loc = avm.get("location") or {}
    contact = cust.get("contact") or {}
    rows = [
        ("Project", avm.get("name","")),
        ("Opdrachtgever", cust.get("name","")),
        ("Adres", cust.get("address","")),
        ("Contactpersoon", contact.get("name","")),
        ("Tel", contact.get("phone","")),
        ("E-mail", contact.get("email","")),
        ("Locatie", loc.get("name","")),
        ("Adres locatie", loc.get("address","")),
        ("Start", fmt_date(avm.get("start_date") or avm.get("project_start_date"))),
        ("Einde", fmt_date(avm.get("end_date") or avm.get("project_end_date"))),
    ]
    t = doc.add_table(rows=1, cols=2)
    hdr_cells = t.rows[0].cells
    hdr_cells[0].text = "Veld"
    hdr_cells[1].text = "Waarde"
    for a,b in rows:
        row_cells = t.add_row().cells
        row_cells[0].text = a or ""
        row_cells[1].text = b or ""

    # Verantwoordelijke (mini)
    p = doc.add_paragraph()
    p.add_run("\nProjectverantwoordelijke: ").bold = True
    p.add_run(preview.get("responsible") or "")
    mini = safe_get(preview, "documents.crew_bio_mini")
    if mini:
        img = fetch_binary(mini)
        if img:
            doc.add_picture(io.BytesIO(img), width=Inches(1.4))

    # 2 Emergency
    h("2. Emergency")
    emer = safe_get(preview, "documents.emergency")
    if emer:
        data = fetch_binary(emer)
        if data:
            png = pdf_first_page_to_png_bytes(data)
            if png:
                doc.add_picture(io.BytesIO(png), width=Inches(6.0))

    # 3 Verzekeringen
    ins_list = preview.get("documents", {}).get("insurance", []) or []
    if ins_list:
        h("3. Verzekeringen")
        for url in ins_list:
            data = fetch_binary(url)
            if data:
                png = pdf_first_page_to_png_bytes(data)
                if png:
                    doc.add_picture(io.BytesIO(png), width=Inches(6.0))

    # 5 Materialen
    h("5. Materialen")
    doc.add_paragraph("5.1 Pyrotechnische materialen (Dees)")
    dees = (preview.get("materials") or {}).get("dees") or []
    if dees:
        tbl = doc.add_table(rows=1, cols=4)
        tbl.rows[0].cells[0].text = "Naam"
        tbl.rows[0].cells[1].text = "Aantal"
        tbl.rows[0].cells[2].text = "Type"
        tbl.rows[0].cells[3].text = "Links"
        for it in dees:
            row = tbl.add_row().cells
            row[0].text = str(it.get("displayname",""))
            row[1].text = str(it.get("quantity_total",""))
            row[2].text = str(it.get("type",""))
            row[3].text = build_links_text(it.get("links") or {})
    else:
        doc.add_paragraph("Geen items geselecteerd.")

    doc.add_paragraph("5.2 Speciale effecten (AVM)")
    avm_items = (preview.get("materials") or {}).get("avm") or []
    if avm_items:
        tbl = doc.add_table(rows=1, cols=4)
        tbl.rows[0].cells[0].text = "Naam"
        tbl.rows[0].cells[1].text = "Aantal"
        tbl.rows[0].cells[2].text = "Type"
        tbl.rows[0].cells[3].text = "Links"
        for it in avm_items:
            row = tbl.add_row().cells
            row[0].text = str(it.get("displayname",""))
            row[1].text = str(it.get("quantity_total",""))
            row[2].text = str(it.get("type",""))
            row[3].text = build_links_text(it.get("links") or {})

    # 8 Windplan
    h("8. Windplan")
    wp = safe_get(preview, "documents.windplan")
    if wp:
        data = fetch_binary(wp)
        if data:
            png = pdf_first_page_to_png_bytes(data)
            if png:
                doc.add_picture(io.BytesIO(png), width=Inches(6.0))

    # 9 Droogteplan
    h("9. Droogteplan")
    dp = safe_get(preview, "documents.droughtplan")
    if dp:
        data = fetch_binary(dp)
        if data:
            png = pdf_first_page_to_png_bytes(data)
            if png:
                doc.add_picture(io.BytesIO(png), width=Inches(6.0))

    # 10 Vergunningen
    h("10. Vergunningen")
    for f in (preview.get("uploads", {}).get("permits") or []):
        data = fetch_binary(f.get("data") or f.get("url"))
        if data:
            png = pdf_first_page_to_png_bytes(data)
            if png:
                doc.add_picture(io.BytesIO(png), width=Inches(6.0))

    # 11 Inplantingsplan
    h("11. Inplantingsplan")
    for f in (preview.get("uploads", {}).get("siteplan") or []):
        data = fetch_binary(f.get("data") or f.get("url"))
        if data:
            png = pdf_first_page_to_png_bytes(data)
            if png:
                doc.add_picture(io.BytesIO(png), width=Inches(6.0))

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.read()

# ---------------- Flask route ----------------

@app.route("/generate", methods=["POST"])
def generate_route():
    try:
        body = request.get_json(force=True)
    except Exception:
        return jsonify(error="Invalid JSON"), 400

    preview = body.get("preview") if isinstance(body, dict) else None
    if not preview:
        # backward compatibility: allow body to be preview directly
        preview = body

    fmt = (body.get("format") if isinstance(body, dict) else None) or "pdf"
    fmt = fmt.lower()

    if fmt == "pdf":
        data = build_pdf(preview)
        return send_file(io.BytesIO(data), mimetype="application/pdf", as_attachment=True, download_name="generate_pdf.pdf")
    elif fmt in ("docx","word","doc"):
        data = build_docx(preview)
        return send_file(io.BytesIO(data), mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document", as_attachment=True, download_name="generate_pdf.docx")
    else:
        return jsonify(error="Unsupported format"), 400

def create_app():
    return app

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000)
