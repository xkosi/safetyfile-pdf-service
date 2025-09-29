
# -*- coding: utf-8 -*-
"""
Veiligheidsdossier generator (PDF & DOCX)

POST /generate
{
  "preview": { ...wizard json... },
  "format": "pdf" | "docx"
}
"""

import io, base64, datetime, re
import requests

from flask import Flask, request, send_file, jsonify

# -------- PDF libs --------
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph, Table, TableStyle, Frame, Spacer

from PyPDF2 import PdfReader, PdfWriter, Transformation

# -------- DOCX libs --------
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# -------- PDF->Image for DOCX --------
import fitz  # PyMuPDF

app = Flask(__name__)

# =========================
# Layout constants
# =========================
W, H = A4
MARGIN_L = 22
MARGIN_R = 22
MARGIN_T = 22
MARGIN_B = 34

BANNER_H = 48
CONTENT_TOP_Y = H - (BANNER_H + 26)

RED = colors.HexColor("#B00000")
BLACK = colors.black
GRAY = colors.HexColor("#BBBBBB")

styles = getSampleStyleSheet()
P = ParagraphStyle("Body", parent=styles["Normal"], fontName="Helvetica", fontSize=10, leading=13, textColor=BLACK)
P_H = ParagraphStyle("H", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=12, leading=14, textColor=BLACK, spaceAfter=6)
P_TOC = ParagraphStyle("TOC", parent=styles["Normal"], fontName="Helvetica", fontSize=11, leading=15, textColor=BLACK)

# =========================
# Helpers
# =========================

def _safe(x, default=""):
    return default if x is None else str(x)

def _fmt_date(s):
    if not s: return ""
    try:
        iso = str(s).replace("Z","+00:00").replace("T"," ")
        return datetime.datetime.fromisoformat(iso).strftime("%d/%m/%Y")
    except Exception:
        try:
            return datetime.datetime.strptime(str(s)[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
        except Exception:
            return str(s)

def _dataurl_to_bytes(u):
    if not u or not isinstance(u,str): return None
    if u.startswith("data:"):
        try:
            return base64.b64decode(u.split(",",1)[1])
        except Exception:
            return None
    return None

def fetch_binary(item):
    """Return bytes from data URL or http(s) URL; None on failure."""
    if not item: return None
    if isinstance(item, bytes): return item
    if isinstance(item, str):
        b = _dataurl_to_bytes(item)
        if b: return b
        try:
            r = requests.get(item, timeout=25)
            if r.ok: return r.content
        except Exception:
            return None
    return None

def draw_banner(c, title):
    c.setFillColor(RED)
    c.rect(0, H-BANNER_H, W, BANNER_H, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(MARGIN_L, H-BANNER_H+16, _safe(title))

def draw_marker(c, key):
    c.setFillColor(colors.white)
    c.setFont("Helvetica", 1)
    c.drawString(2,2,f"[SEC::{key}]")

def start_section(c, key, title):
    c.showPage()
    draw_banner(c, title)
    draw_marker(c, key)

def content_frame():
    return Frame(MARGIN_L, MARGIN_B, W-(MARGIN_L+MARGIN_R), CONTENT_TOP_Y-MARGIN_B, showBoundary=0)

def project_table(preview):
    avm = (preview or {}).get("avm") or {}
    cust = (avm.get("customer") or {})
    contact = (cust.get("contact") or {})
    loc = (avm.get("location") or {})

    rows = [
        ["Project", _safe(avm.get("name"))],
        ["Opdrachtgever", _safe(cust.get("name"))],
        ["Adres", _safe(cust.get("address"))],
    ]
    if any(contact.values()):
        rows += [
            ["Contactpersoon", _safe(contact.get("name"))],
            ["Tel", _safe(contact.get("phone"))],
            ["E-mail", _safe(contact.get("email"))],
        ]
    rows += [
        ["Locatie", _safe(loc.get("name"))],
        ["Adres locatie", _safe(loc.get("address"))],
        ["Start", _fmt_date(avm.get("start_date") or avm.get("project_start_date"))],
        ["Einde", _fmt_date(avm.get("end_date") or avm.get("project_end_date"))],
    ]

    table = Table(rows, colWidths=[120, W-(MARGIN_L+MARGIN_R)-120])
    table.setStyle(TableStyle([
        ("FONT",(0,0),(-1,-1),"Helvetica",10),
        ("FONTNAME",(0,0),(0,-1),"Helvetica-Bold"),
        ("TEXTCOLOR",(0,0),(-1,-1),BLACK),
        ("GRID",(0,0),(-1,-1),0.25,GRAY),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("BACKGROUND",(0,0),(0,-1),colors.HexColor("#F5F5F5")),
    ]))
    return table

def responsible_story(preview):
    name = _safe((preview or {}).get("responsible") or "")
    elems = []
    if not name:
        elems.append(Paragraph("Geen verantwoordelijke geselecteerd.", P))
        return elems
    elems.append(Paragraph(f"Projectverantwoordelijke: <b>{name}</b>", P))

    mini = ((preview or {}).get("documents") or {}).get("crew_bio_mini")
    if mini:
        img = fetch_binary(mini)
        if img:
            # Render as image under the text (approx 35mm height)
            from reportlab.platypus import Image
            try:
                elems.append(Spacer(1,6))
                elems.append(Image(io.BytesIO(img), height=100, width=None, hAlign="LEFT"))
            except Exception:
                pass
    return elems

def materials_story(preview):
    mats = (preview or {}).get("materials") or {}
    avm_items = mats.get("avm") or []
    dees_items = mats.get("dees") or []

    def rows(items):
        r = [["Naam","Aantal","Type","CE","Manual","MSDS"]]
        for m in (items or []):
            ln = (m.get("links") or {})
            r.append([
                _safe(m.get("displayname")),
                _safe(m.get("quantity_total")),
                _safe(m.get("type")),
                "ja" if ln.get("ce") else "",
                "ja" if ln.get("manual") else "",
                "ja" if ln.get("msds") else "",
            ])
        return r

    story = []
    story.append(Paragraph("5.1 Pyrotechnische materialen", P_H))
    if not dees_items:
        story.append(Paragraph("Geen items geselecteerd.", P))
    else:
        t = Table(rows(dees_items), colWidths=[220,45,80,90,90,90])
        t.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,GRAY),("FONT",(0,0),(-1,-1),"Helvetica",9),
                               ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F1E3E3"))]))
        story.append(t)

    story.append(Spacer(1,10))
    story.append(Paragraph("5.2 Speciale effecten", P_H))
    if not avm_items:
        story.append(Paragraph("Geen items geselecteerd.", P))
    else:
        t = Table(rows(avm_items), colWidths=[220,45,80,90,90,90])
        t.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,GRAY),("FONT",(0,0),(-1,-1),"Helvetica",9),
                               ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F1E3E3"))]))
        story.append(t)

    return story

def build_sections(preview):
    mats = (preview or {}).get("materials") or {}
    has_pyro = bool(mats.get("dees"))
    has_sfx = bool(mats.get("avm"))

    keys = ["project","emergency","insurance","responsible","materials","siteplan"]
    if has_pyro: keys.append("risk_pyro")
    if has_sfx: keys.append("risk_sfx")
    keys += ["wind","drought","permits"]

    title_map = {
        "project":"Projectgegevens",
        "emergency":"Emergency",
        "insurance":"Verzekeringen",
        "responsible":"Verantwoordelijke",
        "materials":"Materialen",
        "siteplan":"Inplantingsplan",
        "risk_pyro":"Risicoanalyse Pyro",
        "risk_sfx":"Risicoanalyse Speciale effecten",
        "wind":"Windplan",
        "drought":"Droogteplan",
        "permits":"Vergunningen & Toelatingen",
    }
    sections = []
    for i, k in enumerate(keys, 1):
        sections.append({"key":k, "title": f"{i}. {title_map[k]}"})
    return sections

def draw_cover(c, preview):
    draw_banner(c, "Veiligheidsdossier")
    avm = (preview or {}).get("avm") or {}
    pname = avm.get("name") or (preview or {}).get("project",{}).get("name") or ""
    c.setFillColor(BLACK)
    c.setFont("Helvetica-Bold", 22)
    if pname:
        c.drawString(MARGIN_L, H-BANNER_H-36, pname)
    c.setFont("Helvetica", 9)
    c.drawString(MARGIN_L, MARGIN_B-12, "Gegenereerd: %s" % datetime.datetime.now().strftime("%d/%m/%Y %H:%M"))

def build_base_pdf(preview):
    sections = build_sections(preview)
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)

    # cover
    draw_cover(c, preview)

    # TOC
    c.showPage()
    draw_banner(c, "Inhoudstafel")
    y = CONTENT_TOP_Y
    fr = content_frame()
    story = []
    for s in sections:
        story.append(Paragraph(_safe(s["title"]), P_TOC))
    fr.addFromList(story, c)

    # Base sections with internal content
    for s in sections:
        start_section(c, s["key"], s["title"])
        fr = content_frame()
        if s["key"] == "project":
            fr.addFromList([project_table(preview)], c)
        elif s["key"] == "responsible":
            fr.addFromList(responsible_story(preview), c)
        elif s["key"] == "materials":
            fr.addFromList(materials_story(preview), c)
        else:
            # External-contents sections get a marker page but are filled during merge
            pass

    c.save()
    return buf.getvalue(), sections

# ---------- External PDF merging -----------
def find_section_pages(reader):
    mapping = {}
    for idx, pg in enumerate(reader.pages):
        try:
            text = pg.extract_text() or ""
        except Exception:
            text = ""
        for key in re.findall(r"\[SEC::([^\]]+)\]", text):
            mapping[key] = idx
    return mapping

def scale_merge_first_page_under_banner(writer, page_index, ext_reader):
    """Place ext_reader page 1 under the red header on the targeted page; append the rest pages afterwards."""
    if not ext_reader or len(ext_reader.pages) == 0: return
    dst = writer.pages[page_index]
    src = ext_reader.pages[0]

    # compute available rect under banner
    avail_w = W - (MARGIN_L + MARGIN_R)
    avail_h = CONTENT_TOP_Y - MARGIN_B

    sw = float(src.mediabox.width)
    sh = float(src.mediabox.height)
    scale = min(avail_w/sw, avail_h/sh, 1.0)

    tx = MARGIN_L
    ty = MARGIN_B

    t = Transformation().scale(scale).translate(tx/scale, ty/scale)

    # PyPDF2 compatibility
    if hasattr(dst, "merge_transformed_page"):
        dst.merge_transformed_page(src, t)
    elif hasattr(dst, "mergeTransformedPage"):
        dst.mergeTransformedPage(src, t)
    else:
        raise AttributeError("Geen geldige merge-methode gevonden in PyPDF2")

    # append remaining pages of ext
    for i in range(1, len(ext_reader.pages)):
        writer.add_page(ext_reader.pages[i])

def collect_externals(preview):
    docs = (preview or {}).get("documents") or {}
    uploads = (preview or {}).get("uploads") or {}
    def listify(v):
        return v if isinstance(v, list) else ([v] if v else [])
    out = {
        "emergency": [fetch_binary(x) for x in listify(docs.get("emergency")) if x],
        "insurance": [fetch_binary(x) for x in listify(docs.get("insurance")) if x],
        "wind": [fetch_binary(docs.get("windplan"))] if docs.get("windplan") else [],
        "drought": [fetch_binary(docs.get("droughtplan"))] if docs.get("droughtplan") else [],
        "permits": [fetch_binary(f.get("data")) for f in uploads.get("permits",[]) if f.get("data")],
        "siteplan": [fetch_binary(f.get("data")) for f in uploads.get("siteplan",[]) if f.get("data")],
        # risk templates (if provided)
        "risk_pyro": [fetch_binary(docs.get("risk_pyro"))] if docs.get("risk_pyro") else [],
        "risk_sfx": [fetch_binary(docs.get("risk_general"))] if docs.get("risk_general") else [],
        # crew bios for responsible section (PDF mini-bio ignored on purpose; we only embed image in story)
        "responsible_bio": []  # not used (spec: only mini image)
    }
    return out

def make_cover_pdf(title):
    b = io.BytesIO()
    c = canvas.Canvas(b, pagesize=A4)
    draw_banner(c, title)
    c.save()
    return b.getvalue()

def merge_externals(base_bytes, sections, preview):
    reader = PdfReader(io.BytesIO(base_bytes))
    writer = PdfWriter()
    for p in reader.pages:
        writer.add_page(p)

    page_map = find_section_pages(reader)
    blobs = collect_externals(preview)

    # mapping section key -> list of bytes arrays (pdfs)
    plan = {
        "emergency": blobs.get("emergency", []),
        "insurance": blobs.get("insurance", []),
        "siteplan": blobs.get("siteplan", []),
        "risk_pyro": blobs.get("risk_pyro", []),
        "risk_sfx": blobs.get("risk_sfx", []),
        "wind": blobs.get("wind", []),
        "drought": blobs.get("drought", []),
        "permits": blobs.get("permits", []),
    }

    # Merge from back to front (page indices stable)
    items = [(k, page_map[k]) for k in plan.keys() if k in page_map and plan[k]]
    items.sort(key=lambda x: x[1], reverse=True)

    # Helper to build cover + merge each external
    def add_with_cover(key, idx, title):
        pdf_list = plan.get(key) or []
        if not pdf_list: return
        cover = make_cover_pdf(title)
        for j, blob in enumerate(pdf_list):
            try:
                ext = PdfReader(io.BytesIO(blob))
            except Exception:
                continue
            if j == 0:
                # merge first page under banner of existing target page
                cover_reader = PdfReader(io.BytesIO(cover))
                # replace target page with cover (so we have the banner) then place external page into it
                # actually we already have a banner page in base; so we just overlay external page onto that page
                scale_merge_first_page_under_banner(writer, idx, ext)
            else:
                # subsequent items: create a new cover then merge
                merged = PdfWriter()
                cov_reader = PdfReader(io.BytesIO(cover))
                # start with a fresh cover page
                merged.add_page(cov_reader.pages[0])
                # now overlay first page of the external onto this fresh page
                # write merged interim
                tmp_writer = PdfWriter()
                tmp_writer.add_page(cov_reader.pages[0])
                # we need to scale/overlay too:
                # easiest: reuse same routine on a small writer
                tmp = PdfWriter()
                tmp.add_page(cov_reader.pages[0])
                scale_merge_first_page_under_banner(tmp, 0, ext)
                # append resulting pages to main writer
                for p in tmp.pages:
                    writer.add_page(p)

    # Titles from sections
    title_map = {s["key"]: s["title"] for s in sections}

    for key, idx in items:
        add_with_cover(key, idx, title_map.get(key, key.title()))

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()

# =========================
# DOCX helpers
# =========================

def pdf_first_page_to_png(pdf_bytes, zoom=2.0):
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        if doc.page_count == 0: return None
        page = doc.load_page(0)
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        return pix.tobytes("png")
    except Exception:
        return None

def build_docx(preview):
    doc = Document()

    # Title
    doc.add_heading("Veiligheidsdossier", 0)
    avm = (preview or {}).get("avm") or {}
    doc.add_paragraph(_safe(avm.get("name")))

    # 1. Projectgegevens
    doc.add_heading("1. Projectgegevens", level=1)
    cust = (avm.get("customer") or {})
    contact = (cust.get("contact") or {})
    loc = (avm.get("location") or {})
    rows = [
        ("Project", _safe(avm.get("name"))),
        ("Opdrachtgever", _safe(cust.get("name"))),
        ("Adres", _safe(cust.get("address"))),
        ("Contactpersoon", _safe(contact.get("name"))),
        ("Tel", _safe(contact.get("phone"))),
        ("E-mail", _safe(contact.get("email"))),
        ("Locatie", _safe(loc.get("name"))),
        ("Adres locatie", _safe(loc.get("address"))),
        ("Start", _fmt_date(avm.get("start_date") or avm.get("project_start_date"))),
        ("Einde", _fmt_date(avm.get("end_date") or avm.get("project_end_date"))),
    ]
    tbl = doc.add_table(rows=1, cols=2)
    hdr = tbl.rows[0].cells
    hdr[0].text = "Veld"; hdr[1].text = "Waarde"
    for a,b in rows:
        r = tbl.add_row().cells
        r[0].text = a; r[1].text = b

    # 4. Verantwoordelijke (met mini-bio image)
    doc.add_heading("4. Verantwoordelijke", level=1)
    resp = _safe((preview or {}).get("responsible") or "")
    doc.add_paragraph(f"Projectverantwoordelijke: {resp}")
    mini = ((preview or {}).get("documents") or {}).get("crew_bio_mini")
    if mini:
        img = fetch_binary(mini)
        if img:
            try:
                doc.add_picture(io.BytesIO(img), width=Inches(1.4))
            except Exception:
                pass

    # 5. Materialen
    doc.add_heading("5. Materialen", level=1)

    doc.add_paragraph("5.1 Pyrotechnische materialen (Dees)")
    dees = ((preview or {}).get("materials") or {}).get("dees") or []
    if dees:
        t = doc.add_table(rows=1, cols=4)
        t.rows[0].cells[0].text = "Naam"; t.rows[0].cells[1].text = "Aantal"
        t.rows[0].cells[2].text = "Type"; t.rows[0].cells[3].text = "Links"
        for it in dees:
            row = t.add_row().cells
            row[0].text = _safe(it.get("displayname"))
            row[1].text = _safe(it.get("quantity_total"))
            row[2].text = _safe(it.get("type"))
            ln = it.get("links") or {}
            row[3].text = ", ".join([k.upper() for k in ["ce","manual","msds"] if ln.get(k)])
    else:
        doc.add_paragraph("Geen items geselecteerd.")

    doc.add_paragraph("5.2 Speciale effecten (AVM)")
    avm_items = ((preview or {}).get("materials") or {}).get("avm") or []
    if avm_items:
        t = doc.add_table(rows=1, cols=4)
        t.rows[0].cells[0].text = "Naam"; t.rows[0].cells[1].text = "Aantal"
        t.rows[0].cells[2].text = "Type"; t.rows[0].cells[3].text = "Links"
        for it in avm_items:
            row = t.add_row().cells
            row[0].text = _safe(it.get("displayname"))
            row[1].text = _safe(it.get("quantity_total"))
            row[2].text = _safe(it.get("type"))
            ln = it.get("links") or {}
            row[3].text = ", ".join([k.upper() for k in ["ce","manual","msds"] if ln.get(k)])
    else:
        doc.add_paragraph("Geen items geselecteerd.")

    # External PDFs as images
    def add_pdf_img_section(title, blobs):
        doc.add_heading(title, level=1)
        for b in blobs:
            png = pdf_first_page_to_png(b)
            if png:
                try:
                    doc.add_picture(io.BytesIO(png), width=Inches(6.0))
                except Exception:
                    pass

    blobs = collect_externals(preview)
    if blobs.get("emergency"):
        add_pdf_img_section("2. Emergency", blobs["emergency"])
    if blobs.get("insurance"):
        add_pdf_img_section("3. Verzekeringen", blobs["insurance"])
    if blobs.get("wind"):
        add_pdf_img_section("8. Windplan", blobs["wind"])
    if blobs.get("drought"):
        add_pdf_img_section("9. Droogteplan", blobs["drought"])
    if blobs.get("permits"):
        add_pdf_img_section("10. Vergunningen", blobs["permits"])
    if blobs.get("siteplan"):
        add_pdf_img_section("6. Inplantingsplan", blobs["siteplan"])

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

# =========================
# Flask route
# =========================

@app.route("/generate", methods=["POST"])
def generate():
    try:
        payload = request.get_json(force=True, silent=False) or {}
    except Exception:
        return jsonify({"error":"Invalid JSON"}), 400

    preview = payload.get("preview") or payload
    fmt = (payload.get("format") or "pdf").lower()

    if fmt == "docx":
        try:
            data = build_docx(preview)
            return send_file(io.BytesIO(data),
                mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                as_attachment=True, download_name="dossier.docx")
        except Exception as e:
            return jsonify({"error":"DOCX generation failed","detail":str(e)}), 500

    try:
        base_bytes, sections = build_base_pdf(preview)
        final_bytes = merge_externals(base_bytes, sections, preview)
        return send_file(io.BytesIO(final_bytes), mimetype="application/pdf",
            as_attachment=True, download_name="dossier.pdf")
    except Exception as e:
        return jsonify({"error":"PDF generation failed","detail":str(e)}), 500

def create_app():
    return app

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000)
