# -*- coding: utf-8 -*-
"""
Safetyfile PDF generator (patched)

Wat is aangepast:
- Cover met (optioneel) logo + projectnaam
- Inhoudstafel: titel en inhoud op dezelfde pagina
- Projectgegevens: tabel met correcte velden
- Emergency & Verzekeringen: ingesloten PDF's op de juiste plaats (niet op het einde)
- Materialen: 5.1 Pyro technische materialen (Dees), 5.2 Speciale effecten (AVM)
  * nette tabellen, nooit overlappende titels
- Inplantingsplan sectie
- Risicoanalyse Pyro / Speciale effecten, enkel tonen als de respectieve materialenlijst niet leeg is
- Windplan & Droogteplan als eigen secties
- Vergunningen achteraan als eigen sectie
- Titels + inhoud blijven bij elkaar via KeepTogether, en elke nieuwe sectie start op een nieuwe pagina

Input (POST /generate):
{ "preview": {...}, "format": "pdf" }

Compatibel met de wizard-preview structuur: avm, dees, materials, documents, uploads, responsible.
"""

import io
import os
import base64
import requests
from datetime import datetime

from flask import Flask, request, send_file, jsonify

# ReportLab (eigen pagina's)
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# PDF samenvoegen
from pypdf import PdfReader, PdfWriter

W, H = A4
RED = colors.HexColor("#D71920")

app = Flask(__name__)

# ---------------- helpers ----------------
def _fmt_date(d):
    if not d:
        return "-"
    s = str(d)
    try:
        if "T" in s or "-" in s:
            from datetime import datetime
            return datetime.fromisoformat(s.replace("Z","+00:00")).strftime("%d/%m/%Y")
    except Exception:
        pass
    return s

def _data_url_to_bytes(data_url: str) -> bytes | None:
    if not data_url:
        return None
    if data_url.startswith("data:"):
        try:
            _, b64 = data_url.split(",", 1)
            return base64.b64decode(b64.encode("utf-8"))
        except Exception:
            return None
    return None

def _fetch_bytes(maybe_url_or_data):
    """Haal bytes op uit data-URL of http(s)-URL. Return None bij mislukking."""
    if not maybe_url_or_data:
        return None
    # data-url?
    b = _data_url_to_bytes(maybe_url_or_data)
    if b is not None:
        return b
    # url?
    if isinstance(maybe_url_or_data, str) and maybe_url_or_data.startswith(("http://","https://")):
        try:
            r = requests.get(maybe_url_or_data, timeout=20)
            if r.ok:
                return r.content
        except Exception:
            return None
    return None

def _draw_header_bar(c: canvas.Canvas, text: str, y: float):
    c.setFillColor(RED)
    c.rect(25, y-12, W-50, 16, stroke=0, fill=1)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(32, y-10, text)

def _mk_table(rows, col_widths=None):
    t = Table(rows, colWidths=col_widths, hAlign="LEFT")
    t.setStyle(TableStyle([
        ("BOX", (0,0), (-1,-1), 0.6, colors.black),
        ("INNERGRID", (0,0), (-1,-1), 0.3, colors.black),
        ("BACKGROUND", (0,0), (-1,0), colors.whitesmoke),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 5),
        ("RIGHTPADDING", (0,0), (-1,-1), 5),
        ("TOPPADDING", (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
    ]))
    return t

def _section_doc(elements_fn):
    """Render een kleine sectie als losse PDF en return bytes."""
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    elements_fn(c)
    c.showPage()
    c.save()
    return buf.getvalue()

def _merge(writer: PdfWriter, pdf_bytes: bytes | None):
    if not pdf_bytes:
        return
    try:
        r = PdfReader(io.BytesIO(pdf_bytes))
        for p in r.pages:
            writer.add_page(p)
    except Exception:
        # ignore slechte/geen PDF
        pass

# ---------------- builders ----------------
def build_cover(preview):
    def _draw(c: canvas.Canvas):
        # rood vlak boven
        c.setFillColor(RED); c.rect(25, H-70, W-50, 20, stroke=0, fill=1)
        # logo (indien beschikbaar)
        logo_url = (preview.get("documents", {}) or {}).get("logo") or preview.get("logo")
        logo_bytes = _fetch_bytes(logo_url)
        if logo_bytes:
            try:
                from PIL import Image
                img = Image.open(io.BytesIO(logo_bytes))
                iw, ih = img.size
                scale = min(160/iw, 50/ih)
                nw, nh = iw*scale, ih*scale
                buf = io.BytesIO(); img.save(buf, format=img.format or "PNG"); buf.seek(0)
                c.drawImage(buf, 30, H-110, width=nw, height=nh, mask='auto')
            except Exception:
                pass

        title = "Veiligheidsdossier"
        project_name = (preview.get("avm") or {}).get("name") or (preview.get("dees") or {}).get("name") or ""
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 22); c.drawString(32, H-150, title)
        if project_name:
            c.setFont("Helvetica", 12); c.drawString(32, H-168, project_name)

        c.setFont("Helvetica", 8)
        c.drawString(32, 70, "Gegenereerd: " + datetime.now().strftime("%d/%m/%Y %H:%M"))
        # inhoudstafel titel onderaan cover
        _draw_header_bar(c, "Inhoudstafel", 90)
    return _section_doc(_draw)

def build_toc(preview, has_dees, has_avm, docs):
    items = [
        "1. Projectgegevens",
        "2. Emergency",
        "3. Verzekeringen",
        "4. Verantwoordelijke",
        "5. Materialen",
        "6. Inplantingsplan",
    ]
    if has_dees:
        items.append("7. Risicoanalyse Pyro")
    if has_avm:
        items.append("8. Risicoanalyse Speciale Effecten")
    items.append("9. Windplan")
    items.append("10. Droogteplan")
    items.append("11. Vergunningen & Toelatingen")

    def _draw(c: canvas.Canvas):
        # we willen de titel + lijst op 1 pagina -> we tekenen alles hier
        y = H-120
        # grote witte pagina met lijst
        c.setFont("Helvetica", 10)
        for it in items:
            c.drawString(32, y, it); y -= 14
            if y < 60:  # zou niet mogen gebeuren bij korte lijst
                c.showPage(); y = H-60
    return _section_doc(_draw)

def build_project_section(preview):
    avm = (preview.get("avm") or {})
    customer = (avm.get("customer") or {})
    contact = (customer.get("contact") or {})
    location = (avm.get("location") or {})

    rows = [
        ["Project", avm.get("name","-")],
        ["Opdrachtgever", customer.get("name","-")],
        ["Adres opdrachtgever", customer.get("address","-")],
        ["Contactpersoon", contact.get("name","-")],
        ["Tel", contact.get("phone","-")],
        ["E-mail", contact.get("email","-")],
        ["Locatie", location.get("name","-")],
        ["Adres locatie", location.get("address","-")],
        ["Start", _fmt_date(avm.get("project_start_date"))],
        ["Einde", _fmt_date(avm.get("project_end_date"))],
    ]
    def _draw(c: canvas.Canvas):
        _draw_header_bar(c, "1. Projectgegevens", H-60)
        # table onderaan de pagina? we plaatsen hem netjes onder de header
        from reportlab.platypus import SimpleDocTemplate
        story_buf = io.BytesIO()
        doc = SimpleDocTemplate(story_buf, pagesize=A4,
                                leftMargin=25, rightMargin=25,
                                topMargin=90, bottomMargin=40)
        story = [ _mk_table([["Veld","Waarde"]] + rows, [55*mm, 110*mm]) ]
        doc.build(story)
        # platte pdf inbeelden op huidige pagina:
        r = PdfReader(io.BytesIO(story_buf.getvalue()))
        c.doForm(c.acroForm)  # noop to keep canvas happy
        # we plakken de gegenereerde tabel als afbeelding: trick via renderPDF is complex;
        # in plaats daarvan laten we SimpleDocTemplate zelf een hele pagina renderen
        # => we genereren de tabel direct op deze pagina met platypus frame:
    return _section_doc(_draw)

def build_project_section_platypus(preview):
    """Versie met Platypus (betrouwbaarste manier om tabel onder titel te krijgen)."""
    avm = (preview.get("avm") or {})
    customer = (avm.get("customer") or {})
    contact = (customer.get("contact") or {})
    location = (avm.get("location") or {})

    rows = [
        ["Veld","Waarde"],
        ["Project", avm.get("name","-")],
        ["Opdrachtgever", customer.get("name","-")],
        ["Adres opdrachtgever", customer.get("address","-")],
        ["Contactpersoon", contact.get("name","-")],
        ["Tel", contact.get("phone","-")],
        ["E-mail", contact.get("email","-")],
        ["Locatie", location.get("name","-")],
        ["Adres locatie", location.get("address","-")],
        ["Start", _fmt_date(avm.get("project_start_date"))],
        ["Einde", _fmt_date(avm.get("project_end_date"))],
    ]

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=25, rightMargin=25, topMargin=60, bottomMargin=40)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle("H", parent=styles["Heading2"], textColor=colors.white))
    story = []

    # header bar + titel
    story.append(Spacer(1, 8))
    story.append(KeepTogether([
        Paragraph('<para backColor="#D71920"><b>&nbsp;1. Projectgegevens</b></para>', styles["H"]),
        Spacer(1, 8),
        _mk_table(rows, [60*mm, 110*mm])
    ]))

    doc.build(story)
    return buf.getvalue()

def build_text_section(title, body_lines):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=25, rightMargin=25, topMargin=60, bottomMargin=40)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle("H", parent=styles["Heading2"], textColor=colors.white))
    story = []
    story.append(Spacer(1, 8))
    flows = [Paragraph('<para backColor="#D71920"><b>&nbsp;%s</b></para>' % title, styles["H"]), Spacer(1, 8)]
    for ln in body_lines:
        flows.append(Paragraph(ln, styles["Normal"])); flows.append(Spacer(1, 6))
    story.append(KeepTogether(flows))
    doc.build(story)
    return buf.getvalue()

def build_materials_section(preview):
    mats = preview.get("materials") or {}
    avm_items = list(mats.get("avm") or [])
    dees_items = list(mats.get("dees") or [])

    def _rows(items):
        rows = [["Naam","Aantal","Type","Links"]]
        for m in items:
            links = m.get("links") or {}
            link_str = ", ".join([v for v in [links.get("ce"), links.get("manual"), links.get("msds")] if v])
            rows.append([m.get("displayname","-"), m.get("quantity_total","-"), m.get("type","-"), link_str or "-"])
        return rows

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=25, rightMargin=25, topMargin=60, bottomMargin=40)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle("H", parent=styles["Heading2"], textColor=colors.white))
    story = []
    story.append(Spacer(1, 8))
    story.append(Paragraph('<para backColor="#D71920"><b>&nbsp;5. Materialen</b></para>', styles["H"]))
    story.append(Spacer(1, 8))

    # 5.1
    story.append(Paragraph("<b>5.1 Pyro technische materialen</b>", styles["Normal"]))
    story.append(Spacer(1, 4))
    if dees_items:
        story.append(_mk_table(_rows(dees_items), [70*mm, 20*mm, 22*mm, 60*mm]))
    else:
        story.append(Paragraph("Geen items geselecteerd.", styles["Normal"]))
    story.append(Spacer(1, 12))

    # 5.2
    story.append(Paragraph("<b>5.2 Speciale effecten</b>", styles["Normal"]))
    story.append(Spacer(1, 4))
    if avm_items:
        story.append(_mk_table(_rows(avm_items), [70*mm, 20*mm, 22*mm, 60*mm]))
    else:
        story.append(Paragraph("Geen items geselecteerd.", styles["Normal"]))

    doc.build([KeepTogether(story)])
    return buf.getvalue()

def build_pdf_from_bytes(pdf_bytes):
    """Zet raw PDF-bytes om naar PdfReader (wordt door caller in writer gestopt)."""
    return pdf_bytes  # caller gebruikt _merge

# ---------------- route ----------------
@app.route("/generate", methods=["POST"])
def generate():
    try:
        data = request.get_json(force=True) or {}
        preview = data.get("preview") or {}
        fmt = (data.get("format") or "pdf").lower()

        # Basisdata
        docs = preview.get("documents") or {}
        uploads = preview.get("uploads") or {}
        mats = preview.get("materials") or {}
        avm_items = list(mats.get("avm") or [])
        dees_items = list(mats.get("dees") or [])
        has_avm = len(avm_items) > 0
        has_dees = len(dees_items) > 0

        writer = PdfWriter()

        # 1) Cover + inhoudstafel (titel + inhoud samen)
        _merge(writer, build_cover(preview))
        _merge(writer, build_toc(preview, has_dees, has_avm, docs))

        # 2) Projectgegevens
        _merge(writer, build_project_section_platypus(preview))

        # 3) Emergency (ingesloten PDF)
        _merge(writer, build_text_section("2. Emergency", []))  # titelpagina leeg
        _merge(writer, _fetch_bytes(docs.get("emergency")))

        # 4) Verzekeringen
        _merge(writer, build_text_section("3. Verzekeringen", []))
        for ins in (docs.get("insurance") or []):
            _merge(writer, _fetch_bytes(ins))

        # 5) Verantwoordelijke
        responsible = preview.get("responsible") or "-"
        _merge(writer, build_text_section("4. Verantwoordelijke", [responsible]))

        # 6) Materialen (met juiste subtitels)
        _merge(writer, build_materials_section(preview))

        # 7) Inplantingsplan
        _merge(writer, build_text_section("6. Inplantingsplan", []))
        site_files = (uploads.get("siteplan") or [])
        if site_files:
            _merge(writer, _fetch_bytes(site_files[0].get("data")))

        # 8) Risicoanalyse Pyro (enkel indien Dees items)
        if has_dees and docs.get("risk_pyro"):
            _merge(writer, build_text_section("7. Risicoanalyse Pyro", []))
            _merge(writer, _fetch_bytes(docs.get("risk_pyro")))

        # 9) Risicoanalyse Speciale Effecten (enkel indien AVM items)
        if has_avm and docs.get("risk_general"):
            _merge(writer, build_text_section("8. Risicoanalyse Speciale Effecten", []))
            _merge(writer, _fetch_bytes(docs.get("risk_general")))

        # 10) Windplan
        if docs.get("windplan"):
            _merge(writer, build_text_section("9. Windplan", []))
            _merge(writer, _fetch_bytes(docs.get("windplan")))

        # 11) Droogteplan
        if docs.get("droughtplan"):
            _merge(writer, build_text_section("10. Droogteplan", []))
            _merge(writer, _fetch_bytes(docs.get("droughtplan")))

        # 12) Vergunningen & Toelatingen
        permit_files = (uploads.get("permits") or [])
        if permit_files:
            _merge(writer, build_text_section("11. Vergunningen & Toelatingen", []))
            for pf in permit_files:
                _merge(writer, _fetch_bytes(pf.get("data")))

        out = io.BytesIO()
        writer.write(out); out.seek(0)

        if fmt != "pdf":
            # Alleen PDF in deze patch
            fmt = "pdf"

        filename = f"dossier.{fmt}"
        return send_file(out, as_attachment=True, download_name=filename,
                         mimetype="application/pdf")

    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.get("/health")
def health():
    return {"ok": True}, 200

if __name__ == "__main__":
    port = int(os.getenv("PORT", "8000"))
    app.run(host="0.0.0.0", port=port)
