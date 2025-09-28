
import io
import base64
import datetime
import requests

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, Table, TableStyle, Frame, Spacer

from pypdf import PdfReader, PdfWriter, Transformation

# ------------------------------
# Helpers
# ------------------------------
W, H = A4
MARGIN_LR = 30
MARGIN_BOTTOM = 40
BANNER_H = 60
CONTENT_TOP_Y = H - (BANNER_H + 30)

styles = getSampleStyleSheet()
P = ParagraphStyle(
    "Body", parent=styles["Normal"], fontName="Helvetica", fontSize=10, leading=13, spaceAfter=4
)
P_SMALL = ParagraphStyle(
    "Small", parent=styles["Normal"], fontName="Helvetica", fontSize=9, leading=11, spaceAfter=2
)
P_H = ParagraphStyle(
    "H", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=12, leading=14, spaceAfter=6
)

def _safe(val, default=""):
    return "" if val is None else str(val)

def _fmt_date(val):
    if not val: return ""
    try:
        # accept 'YYYY-MM-DD' or ISO
        d = datetime.datetime.fromisoformat(str(val).replace("Z","").replace("T"," ")).date()
        return d.strftime("%d/%m/%Y")
    except Exception:
        return str(val)

def _dataurl_to_bytes(u):
    # support 'data:application/pdf;base64,...'
    if not u or not isinstance(u, str):
        return None
    if u.startswith("data:"):
        try:
            b64 = u.split(",", 1)[1]
            return base64.b64decode(b64)
        except Exception:
            return None
    return None

def _fetch_pdf_bytes(url_or_dataurl, timeout=20):
    if not url_or_dataurl:
        return None
    b = _dataurl_to_bytes(url_or_dataurl)
    if b is not None:
        return b
    # http(s)
    try:
        r = requests.get(url_or_dataurl, timeout=timeout)
        if r.ok and r.content:
            return r.content
    except Exception:
        return None
    return None

# ------------------------------
# Drawing primitives
# ------------------------------
def draw_banner(c, title):
    c.setFillColor(colors.HexColor("#222"))
    c.rect(0, H - BANNER_H, W, BANNER_H, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(MARGIN_LR, H - BANNER_H + 20, _safe(title))

def start_section(c, title):
    # Always begin a new page for a section, draw the banner,
    # and keep the cursor on the *same* page for content.
    c.showPage()
    draw_banner(c, title)

def draw_cover(c, preview):
    c.setFillColor(colors.white)
    c.rect(0, 0, W, H, stroke=0, fill=1)

    # top bar
    draw_banner(c, "Veiligheidsdossier")

    # project name (big)
    avm = (preview or {}).get("avm") or {}
    project_name = avm.get("name") or (preview or {}).get("project", {}).get("name") or ""
    c.setFillColor(colors.HexColor("#111"))
    c.setFont("Helvetica-Bold", 24)
    c.drawString(MARGIN_LR, H - BANNER_H - 40, _safe(project_name))

    # meta
    c.setFont("Helvetica", 10)
    c.setFillColor(colors.HexColor("#333"))
    c.drawString(MARGIN_LR, H - BANNER_H - 60, f"Gegenereerd: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}")
    # NOTE: Add logo if you pass preview['logo'] as dataurl/png; omitted if not provided.

def _frame_for_content():
    return Frame(MARGIN_LR, MARGIN_BOTTOM, W - 2*MARGIN_LR, CONTENT_TOP_Y - MARGIN_BOTTOM, showBoundary=0)

# ------------------------------
# Story builders
# ------------------------------
def story_project(preview):
    avm = (preview or {}).get("avm") or {}
    customer = (avm.get("customer") or {})
    contact  = (customer.get("contact") or {})
    location = (avm.get("location") or {})

    rows = [
        ["Project",       _safe(avm.get("name"))],
        ["Opdrachtgever", _safe(customer.get("name"))],
        ["Adres",         _safe(customer.get("address"))],
    ]
    if contact:
        rows += [
            ["Contactpersoon", _safe(contact.get("name"))],
            ["Tel",            _safe(contact.get("phone"))],
            ["E-mail",         _safe(contact.get("email"))],
        ]
    rows += [
        ["Locatie",       _safe(location.get("name"))],
        ["Adres locatie", _safe(location.get("address"))],
        ["Start",         _fmt_date(avm.get("project_start_date"))],
        ["Einde",         _fmt_date(avm.get("project_end_date"))],
    ]

    t = Table(rows, colWidths=[120, W - 2*MARGIN_LR - 120])
    t.setStyle(TableStyle([
        ("FONT", (0,0), (-1,-1), "Helvetica", 10),
        ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"),
        ("ALIGN", (0,0), (-1,-1), "LEFT"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("GRID", (0,0), (-1,-1), 0.3, colors.HexColor("#999")),
        ("ROWBACKGROUNDS", (0,0), (-1,-1), [colors.white, colors.HexColor("#f7f7f7")]),
    ]))
    return [t]

def _materials_rows(items):
    rows = [["Naam", "Aantal", "Type", "CE", "Manual", "MSDS"]]
    for m in (items or []):
        lnks = (m.get("links") or {})
        rows.append([
            _safe(m.get("displayname")),
            _safe(m.get("quantity_total")),
            _safe(m.get("type")),
            _safe(lnks.get("ce") or ""),
            _safe(lnks.get("manual") or ""),
            _safe(lnks.get("msds") or ""),
        ])
    return rows

def story_materials(preview):
    mats = (preview or {}).get("materials") or {}
    avm_items  = (mats.get("avm")  or [])
    dees_items = (mats.get("dees") or [])

    story = []
    story.append(Paragraph("5.1 Pyrotechnische materialen", P_H))
    if not dees_items:
        story.append(Paragraph("Geen items geselecteerd.", P))
    else:
        t1 = Table(_materials_rows(dees_items), colWidths=[230, 60, 80, 100, 100, 100])
        t1.setStyle(TableStyle([
            ("FONT", (0,0), (-1,-1), "Helvetica", 9),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#bbb")),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f9f9f9")]),
        ]))
        story.append(t1)

    story.append(Spacer(1, 12))
    story.append(Paragraph("5.2 Speciale effecten", P_H))
    if not avm_items:
        story.append(Paragraph("Geen items geselecteerd.", P))
    else:
        t2 = Table(_materials_rows(avm_items), colWidths=[230, 60, 80, 100, 100, 100])
        t2.setStyle(TableStyle([
            ("FONT", (0,0), (-1,-1), "Helvetica", 9),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#bbb")),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f9f9f9")]),
        ]))
        story.append(t2)

    return story

def story_responsible(preview):
    resp_key = (preview or {}).get("responsible") or ""
    if not resp_key:
        return [Paragraph("Geen verantwoordelijke geselecteerd.", P)]
    # get mini-bio if available
    bio = ((preview or {}).get("documents") or {}).get("crew_bio_mini")
    parts = [Paragraph(f"Projectverantwoordelijke: <b>{_safe(resp_key)}</b>", P)]
    if bio and isinstance(bio, str) and bio.startswith("http"):
        parts.append(Paragraph("(Volledige bio wordt ingevoegd als PDF in de volgende pagina's.)", P_SMALL))
    return parts

# ------------------------------
# TOC
# ------------------------------
def build_toc(preview):
    mats = (preview or {}).get("materials") or {}
    has_pyro = bool(mats.get("dees"))
    has_sfx  = bool(mats.get("avm"))

    sections = []
    def add(label): sections.append(label)

    n = 1
    add(f"{n}. Projectgegevens"); n += 1
    add(f"{n}. Emergency"); n += 1
    add(f"{n}. Verzekeringen"); n += 1
    add(f"{n}. Verantwoordelijke"); n += 1
    add(f"{n}. Materialen"); materials_n = n; n += 1

    add(f"{n}. Inplantingsplan"); inplan_n = n; n += 1

    risk_pyro_label = risk_sfx_label = None
    if has_pyro:
        risk_pyro_label = f"{n}. Risicoanalyse Pyrotechniek"; n += 1
        add(risk_pyro_label)
    if has_sfx:
        risk_sfx_label = f"{n}. Risicoanalyse Speciale effecten"; n += 1
        add(risk_sfx_label)

    wind_label = f"{n}. Windplan"; n += 1; add(wind_label)
    drought_label = f"{n}. Droogteplan"; n += 1; add(drought_label)
    permits_label = f"{n}. Vergunningen & Toelatingen"; n += 1; add(permits_label)

    return {
        "list": sections,
        "labels": {
            "materials": materials_n,
            "inplan": inplan_n,
            "risk_pyro": risk_pyro_label,
            "risk_sfx": risk_sfx_label,
            "wind": wind_label,
            "drought": drought_label,
            "permits": permits_label
        }
    }

# ------------------------------
# Build base PDF (cover + all section headers + inline content pages)
# ------------------------------
def build_base_pdf(preview):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)

    # COVER
    draw_cover(c, preview)

    # TOC
    toc = build_toc(preview)
    start_section(c, "Inhoudstafel")
    c.setFont("Helvetica", 11)
    y = CONTENT_TOP_Y
    for item in toc["list"]:
        c.drawString(MARGIN_LR, y, item)
        y -= 16
        if y < 80:
            c.showPage()
            draw_banner(c, "Inhoudstafel")
            y = CONTENT_TOP_Y

    # 1. Projectgegevens
    start_section(c, toc["list"][0])
    story = story_project(preview)
    _frame_for_content().addFromList(story, c)

    # 2. Emergency (banner only, PDF will be merged after this page if provided)
    start_section(c, toc["list"][1])

    # 3. Verzekeringen (banner only, PDFs will be merged/added)
    start_section(c, toc["list"][2])

    # 4. Verantwoordelijke (with short text; full bio PDF will be inserted next)
    start_section(c, toc["list"][3])
    _frame_for_content().addFromList(story_responsible(preview), c)

    # 5. Materialen
    start_section(c, toc["list"][4])
    _frame_for_content().addFromList(story_materials(preview), c)

    # 6. Inplantingsplan
    start_section(c, f"{build_toc(preview)['labels']['inplan']}")

    # Risk (conditional)
    labels = build_toc(preview)["labels"]
    if labels["risk_pyro"]:
        start_section(c, labels["risk_pyro"])
    if labels["risk_sfx"]:
        start_section(c, labels["risk_sfx"])

    # Wind + Drought + Permits
    start_section(c, labels["wind"])
    start_section(c, labels["drought"])
    start_section(c, labels["permits"])

    c.save()
    return buf.getvalue(), toc

# ------------------------------
# Insert/merge external PDFs at their sections
# ------------------------------
def _scale_to_fit(epage, max_w, max_h):
    w = float(epage.mediabox.width)
    h = float(epage.mediabox.height)
    sx = max_w / w
    sy = max_h / h
    s = min(sx, sy, 1.0)  # never upscale
    new_w, new_h = w * s, h * s
    return s, new_w, new_h

def _merge_first_page_under_banner(writer, base_page_index, ext_reader):
    # Merge first page of ext_reader under the banner on writer.pages[base_page_index]
    if not ext_reader or len(ext_reader.pages) == 0:
        return
    dst = writer.pages[base_page_index]
    src = ext_reader.pages[0]

    # available rect under banner
    avail_w = W - 2*MARGIN_LR
    avail_h = CONTENT_TOP_Y - MARGIN_BOTTOM
    s, new_w, new_h = _scale_to_fit(src, avail_w, avail_h)

    # position: left margin, bottom margin
    tx = MARGIN_LR
    ty = MARGIN_BOTTOM

    tr = Transformation().scale(s).translate(tx/s, ty/s)
    dst.merge_transformed_page(src, tr)

    # remaining pages -> append as-is
    for i in range(1, len(ext_reader.pages)):
        writer.add_page(ext_reader.pages[i])

def merge_and_append_at_markers(base_pdf_bytes, preview, toc):
    reader = PdfReader(io.BytesIO(base_pdf_bytes))
    writer = PdfWriter()

    # Prepare all byte blobs
    docs = (preview or {}).get("documents") or {}
    uploads = (preview or {}).get("uploads") or {}

    emergency_urls = []
    if isinstance(docs.get("emergency"), str) and docs.get("emergency"):
        emergency_urls = [docs.get("emergency")]
    elif isinstance(docs.get("emergency"), list):
        emergency_urls = [u for u in docs.get("emergency") if u]

    insurance_urls = [u for u in (docs.get("insurance") or []) if u]
    wind_url = docs.get("windplan")
    drought_url = docs.get("droughtplan")

    # uploads
    permit_blobs = []
    for f in (uploads.get("permits") or []):
        data = _dataurl_to_bytes(f.get("data"))
        if data: permit_blobs.append(data)
    siteplan_blobs = []
    for f in (uploads.get("siteplan") or []):
        data = _dataurl_to_bytes(f.get("data"))
        if data: siteplan_blobs.append(data)

    # track labels so we know where to merge/append
    labels = toc["labels"]
    # For detection we read text of each base page and match the label substring
    def page_has_label(pg, label_str):
        try:
            txt = (pg.extract_text() or "").replace("\\n", " ")
            return label_str.split(".",1)[1].strip() in txt  # match by section name only
        except Exception:
            return False

    for idx, pg in enumerate(reader.pages):
        writer.add_page(pg)  # add the base page
        # After adding, maybe merge/append
        # Emergency
        if labels and page_has_label(pg, f"{labels['materials']-3}. Emergency"):
            # NOTE: labels['materials'] is section number of "Materialen"; Emergency is always 2
            pass  # handled below by direct string check

        if page_has_label(pg, "Emergency"):
            # merge first emergency pdf if any
            for j, u in enumerate(emergency_urls):
                b = _fetch_pdf_bytes(u)
                if not b: continue
                ext = PdfReader(io.BytesIO(b))
                if j == 0:
                    _merge_first_page_under_banner(writer, len(writer.pages)-1, ext)
                else:
                    # all pages appended
                    for p in ext.pages:
                        writer.add_page(p)

        if page_has_label(pg, "Verzekeringen"):
            first = True
            for u in insurance_urls:
                b = _fetch_pdf_bytes(u)
                if not b: continue
                ext = PdfReader(io.BytesIO(b))
                if first:
                    _merge_first_page_under_banner(writer, len(writer.pages)-1, ext)
                    first = False
                else:
                    for p in ext.pages:
                        writer.add_page(p)

        if page_has_label(pg, "Inplantingsplan"):
            # siteplan uploads first (if any)
            first = True
            for b in siteplan_blobs:
                ext = PdfReader(io.BytesIO(b))
                if first:
                    _merge_first_page_under_banner(writer, len(writer.pages)-1, ext)
                    first = False
                else:
                    for p in ext.pages:
                        writer.add_page(p)

        if labels.get("risk_pyro") and page_has_label(pg, "Risicoanalyse Pyrotechniek"):
            # insert pyro risk PDFs if you ever provide them by link in preview['documents']['risk_pyro_*']
            # Placeholder: nothing to insert now.
            pass

        if labels.get("risk_sfx") and page_has_label(pg, "Risicoanalyse Speciale effecten"):
            # insert sfx risk PDFs – placeholder
            pass

        if page_has_label(pg, "Windplan") and wind_url:
            b = _fetch_pdf_bytes(wind_url)
            if b:
                ext = PdfReader(io.BytesIO(b))
                _merge_first_page_under_banner(writer, len(writer.pages)-1, ext)
                for p in ext.pages[1:]:
                    writer.add_page(p)

        if page_has_label(pg, "Droogteplan") and drought_url:
            b = _fetch_pdf_bytes(drought_url)
            if b:
                ext = PdfReader(io.BytesIO(b))
                _merge_first_page_under_banner(writer, len(writer.pages)-1, ext)
                for p in ext.pages[1:]:
                    writer.add_page(p)

        if page_has_label(pg, "Vergunningen & Toelatingen"):
            # append permits (uploads) here
            first = True
            for b in permit_blobs:
                ext = PdfReader(io.BytesIO(b))
                if first:
                    _merge_first_page_under_banner(writer, len(writer.pages)-1, ext)
                    first = False
                else:
                    for p in ext.pages:
                        writer.add_page(p)

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()

# ------------------------------
# DOCX (kept simple; same route can return DOCX)
# ------------------------------
def build_docx(preview):
    try:
        from docx import Document
        from docx.shared import Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH
    except Exception as e:
        raise RuntimeError("python-docx not installed on the server") from e

    doc = Document()
    # Title
    avm = (preview or {}).get("avm") or {}
    project_name = avm.get("name") or (preview or {}).get("project", {}).get("name") or "Veiligheidsdossier"
    h = doc.add_heading("Veiligheidsdossier", level=0)
    h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    doc.add_paragraph(project_name)

    # Project table
    tbl = doc.add_table(rows=0, cols=2)
    for k, v in [
        ("Project", avm.get("name")),
        ("Opdrachtgever", ((avm.get("customer") or {}).get("name"))),
        ("Adres", ((avm.get("customer") or {}).get("address"))),
        ("Locatie", ((avm.get("location") or {}).get("name"))),
        ("Adres locatie", ((avm.get("location") or {}).get("address"))),
        ("Start", _fmt_date(avm.get("project_start_date"))),
        ("Einde", _fmt_date(avm.get("project_end_date"))),
    ]:
        row = tbl.add_row().cells
        row[0].text = _safe(k); row[1].text = _safe(v)

    # Materials
    doc.add_heading("Materialen", level=1)
    mats = (preview or {}).get("materials") or {}
    for label, items in [("Pyrotechnische materialen", mats.get("dees")), ("Speciale effecten", mats.get("avm"))]:
        doc.add_heading(label, level=2)
        if not items:
            doc.add_paragraph("Geen items geselecteerd.")
            continue
        t = doc.add_table(rows=1, cols=6)
        hdr = t.rows[0].cells
        for i, h in enumerate(["Naam","Aantal","Type","CE","Manual","MSDS"]):
            hdr[i].text = h
        for m in items:
            r = t.add_row().cells
            lnks = (m.get("links") or {})
            r[0].text = _safe(m.get("displayname"))
            r[1].text = _safe(m.get("quantity_total"))
            r[2].text = _safe(m.get("type"))
            r[3].text = _safe(lnks.get("ce") or "")
            r[4].text = _safe(lnks.get("manual") or "")
            r[5].text = _safe(lnks.get("msds") or "")

    bio_key = _safe((preview or {}).get("responsible") or "")
    if bio_key:
        doc.add_heading("Verantwoordelijke", level=1)
        doc.add_paragraph(f"Projectverantwoordelijke: {bio_key}")

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

# ------------------------------
# Flask app
# ------------------------------
from flask import Flask, request, send_file, jsonify
app = Flask(__name__)

@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json(silent=True) or {}
    preview = data.get("preview") or data  # support both shapes
    fmt = (data.get("format") or "pdf").lower()

    if fmt == "docx":
        try:
            docx_bytes = build_docx(preview)
            return send_file(
                io.BytesIO(docx_bytes),
                mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                as_attachment=True,
                download_name="dossier.docx"
            )
        except Exception as e:
            return jsonify({"error":"docx failed", "detail": str(e)}), 500

    # PDF
    try:
        base_bytes, toc = build_base_pdf(preview)
        final_bytes = merge_and_append_at_markers(base_bytes, preview, toc)
        return send_file(
            io.BytesIO(final_bytes),
            mimetype="application/pdf",
            as_attachment=True,
            download_name="dossier.pdf"
        )
    except Exception as e:
        return jsonify({"error":"pdf failed", "detail": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000)
