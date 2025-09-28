# -*- coding: utf-8 -*-
"""
Patched generate_pdf.py

- Vult projectvelden, verantwoordelijke, materialen en documenten
- Titels met bijhorende inhoud blijven bij elkaar (geen titel op pagina X en inhoud op pagina Y)
- Elk volgend hoofdstuk begint op een nieuwe pagina
"""

from flask import Flask, request, send_file, jsonify
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, PageBreak, KeepTogether)
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm
import io
from datetime import datetime

app = Flask(__name__)

# ---------- helpers ----------
def fmt(d):
    """Formatteer ISO-datum naar dd/mm/yyyy; anders geef het origineel of '-'."""
    if not d:
        return "-"
    try:
        # laat datetime zelf ISO/parsen en formatteer BE-stijl
        return datetime.fromisoformat(str(d).replace('Z','+00:00')).strftime("%d/%m/%Y")
    except Exception:
        return str(d)

def build_two_col_table(rows, col_widths=(60*mm, 100*mm)):
    t = Table(rows, colWidths=col_widths, hAlign="LEFT")
    t.setStyle(TableStyle([
        ("BOX", (0,0), (-1,-1), 0.6, colors.black),
        ("INNERGRID", (0,0), (-1,-1), 0.3, colors.black),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("BACKGROUND", (0,0), (-1,0), colors.whitesmoke),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
        ("RIGHTPADDING", (0,0), (-1,-1), 6),
        ("TOPPADDING", (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
    ]))
    return t

def build_materials_table(items):
    header = ["Naam", "Aantal", "Type", "CE", "Manual", "MSDS"]
    data = [header]
    for m in items:
        links = m.get("links", {}) or {}
        data.append([
            m.get("displayname", "-"),
            m.get("quantity_total", "-"),
            m.get("type", "-"),
            links.get("ce") or "-",
            links.get("manual") or "-",
            links.get("msds") or "-",
        ])
    t = Table(data, colWidths=[70*mm, 18*mm, 18*mm, 24*mm, 24*mm, 24*mm], hAlign="LEFT")
    t.setStyle(TableStyle([
        ("BOX", (0,0), (-1,-1), 0.6, colors.black),
        ("INNERGRID", (0,0), (-1,-1), 0.3, colors.black),
        ("BACKGROUND", (0,0), (-1,0), colors.whitesmoke),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 4),
        ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("TOPPADDING", (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
    ]))
    return t

def add_section(elements, title, inner_flows, styles, start_on_new_page=False):
    """
    Voegt een hoofdstuk toe.
    - Als start_on_new_page=True: PageBreak VOOR de titel (niet erna).
    - Titel + inhoud worden met KeepTogether gegroepeerd zodat ze niet gesplitst worden.
    """
    if start_on_new_page:
        elements.append(PageBreak())
    section = []
    section.append(Paragraph(title, styles["H2"]))
    section.append(Spacer(1, 4))
    section.extend(inner_flows if isinstance(inner_flows, list) else [inner_flows])
    elements.append(KeepTogether(section))
    elements.append(Spacer(1, 8))


@app.route("/generate", methods=["POST"])
def generate():
    try:
        data = request.get_json(force=True) or {}
        preview = data.get("preview", {}) or {}

        project = preview.get("project", {}) or {}
        materials = preview.get("materials", {}) or {}
        # responsible en documents kunnen ook in preview zitten of op top-level; neem beide veilig mee
        responsible = preview.get("responsible") or data.get("responsible")
        documents = preview.get("documents", {}) or {}

        # ---------- layout / styles ----------
        buf = io.BytesIO()
        doc = SimpleDocTemplate(
            buf,
            pagesize=A4,
            leftMargin=18*mm,
            rightMargin=18*mm,
            topMargin=18*mm,
            bottomMargin=18*mm,
            title="Veiligheidsdossier",
        )

        styles = getSampleStyleSheet()
        styles.add(ParagraphStyle("H1", parent=styles["Heading1"], spaceAfter=10, keepWithNext=True))
        styles.add(ParagraphStyle("H2", parent=styles["Heading2"], spaceBefore=0, spaceAfter=6, keepWithNext=True))
        styles.add(ParagraphStyle("Small", parent=styles["Normal"], fontSize=9))

        story = []

        # Cover
        story.append(Paragraph("Veiligheidsdossier", styles["H1"]))
        story.append(Paragraph("Gegenereerd: {}".format(datetime.now().strftime("%d/%m/%Y %H:%M")), styles["Small"]))
        story.append(Spacer(1, 10))

        # ---------- Projectgegevens (zelfde pagina) ----------
        proj_rows = [
            ["Project", project.get("name", "-")],
            ["Opdrachtgever", (project.get("customer") or {}).get("name", "-")],
            ["Adres opdrachtgever", (project.get("customer") or {}).get("address", "-")],
            ["Contactpersoon", ((project.get("customer") or {}).get("contact") or {}).get("name", "-")],
            ["Tel", ((project.get("customer") or {}).get("contact") or {}).get("phone", "-")],
            ["E-mail", ((project.get("customer") or {}).get("contact") or {}).get("email", "-")],
            ["Locatie", (project.get("location") or {}).get("name", "-")],
            ["Adres locatie", (project.get("location") or {}).get("address", "-")],
            ["Startdatum", fmt(project.get("project_start_date"))],
            ["Einddatum", fmt(project.get("project_end_date"))],
        ]
        add_section(story, "Projectgegevens", build_two_col_table(proj_rows), styles, start_on_new_page=False)

        # ---------- Verantwoordelijke (nieuwe pagina) ----------
        add_section(story, "Projectverantwoordelijke",
                    Paragraph(responsible or "-", styles["Normal"]),
                    styles, start_on_new_page=True)

        # ---------- Materialen (nieuwe pagina) ----------
        avm_items = list(materials.get("avm", []) or [])
        dees_items = list(materials.get("dees", []) or [])

        mats_flow = []
        if avm_items:
            mats_flow.append(Paragraph("AVM", styles["Heading3"]))
            mats_flow.append(Spacer(1, 4))
            mats_flow.append(build_materials_table(avm_items))
            mats_flow.append(Spacer(1, 8))

        if dees_items:
            mats_flow.append(Paragraph("Dees", styles["Heading3"]))
            mats_flow.append(Spacer(1, 4))
            mats_flow.append(build_materials_table(dees_items))

        if not mats_flow:
            mats_flow.append(Paragraph("Geen materialen geselecteerd.", styles["Normal"]))

        add_section(story, "Materialen", mats_flow, styles, start_on_new_page=True)

        # ---------- Documenten (nieuwe pagina) ----------
        docs_rows = [["Type", "Link"]]
        for key, val in (documents or {}).items():
            if isinstance(val, list):
                for v in val:
                    docs_rows.append([key, v or "-"])
            else:
                docs_rows.append([key, val or "-"])

        add_section(story, "Documenten", build_two_col_table(docs_rows, col_widths=(40*mm, 120*mm)),
                    styles, start_on_new_page=True)

        # Build
        doc.build(story)
        buf.seek(0)
        return send_file(buf, as_attachment=True, download_name="dossier.pdf", mimetype="application/pdf")

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/")
def root():
    return "Service is running", 200
