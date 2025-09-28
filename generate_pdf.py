# generate_pdf.py
# Unified generator voor PDF en Word dossiers met correcte layout & projectvelden.

from flask import Flask, request, send_file, jsonify
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from docx import Document
import io, base64, datetime

app = Flask(__name__)

styles = getSampleStyleSheet()
style_h1 = ParagraphStyle('Heading1', parent=styles['Heading1'], fontSize=16, textColor=colors.white, backColor=colors.red, spaceAfter=12)
style_normal = styles['Normal']

def add_project_info_pdf(story, project):
    data = [
        ["Project", project.get("name", "")],
        ["Opdrachtgever", project.get("customer", {}).get("name", "")],
        ["Adres", project.get("customer", {}).get("address", "")],
        ["Locatie", project.get("location", {}).get("name", "")],
        ["Adres locatie", project.get("location", {}).get("address", "")],
        ["Start", project.get("project_start_date", "")],
        ["Einde", project.get("project_end_date", "")],
        ["Verantwoordelijke", project.get("responsible", "")],
    ]
    t = Table(data, colWidths=[150, 300])
    t.setStyle(TableStyle([('GRID',(0,0),(-1,-1),0.5,colors.black)]))
    story.append(t)

@app.route("/generate", methods=["POST"])
def generate():
    body = request.json
    fmt = body.get("format", "pdf")
    preview = body.get("preview", {})

    if fmt == "pdf":
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        story = []

        # Titelpagina
        story.append(Paragraph("Veiligheidsdossier", style_h1))
        story.append(Spacer(1, 20))
        story.append(Paragraph(f"Gegenereerd: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}", style_normal))
        story.append(PageBreak())

        # Projectgegevens
        story.append(Paragraph("1. Projectgegevens", style_h1))
        add_project_info_pdf(story, preview.get("avm", {}))
        story.append(PageBreak())

        # Emergency
        story.append(Paragraph("2. Emergency", style_h1))
        emergency = preview.get("documents", {}).get("emergency")
        if emergency:
            story.append(Paragraph(f"Nooddocument: {emergency}", style_normal))
        story.append(PageBreak())

        # Verzekeringen
        story.append(Paragraph("3. Verzekeringen", style_h1))
        ins = preview.get("documents", {}).get("insurance", [])
        for link in ins:
            story.append(Paragraph(f"Document: {link}", style_normal))
        story.append(PageBreak())

        # Materialen
        story.append(Paragraph("4. Materialen", style_h1))
        for mat in preview.get("materials", {}).get("avm", []):
            story.append(Paragraph(f"- {mat.get('displayname','')} ({mat.get('quantity_total','')})", style_normal))

        doc.build(story)
        buffer.seek(0)
        return send_file(buffer, mimetype="application/pdf", as_attachment=True, download_name="dossier.pdf")

    elif fmt == "docx":
        document = Document()
        document.add_heading("Veiligheidsdossier", 0)
        document.add_paragraph(f"Gegenereerd: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}")

        document.add_heading("1. Projectgegevens", level=1)
        p = preview.get("avm", {})
        document.add_paragraph(f"Project: {p.get('name','')}")
        document.add_paragraph(f"Opdrachtgever: {p.get('customer',{}).get('name','')}")
        document.add_paragraph(f"Adres: {p.get('customer',{}).get('address','')}")
        document.add_paragraph(f"Locatie: {p.get('location',{}).get('name','')}")
        document.add_paragraph(f"Adres locatie: {p.get('location',{}).get('address','')}")
        document.add_paragraph(f"Start: {p.get('project_start_date','')}")
        document.add_paragraph(f"Einde: {p.get('project_end_date','')}")
        document.add_paragraph(f"Verantwoordelijke: {p.get('responsible','')}")

        document.add_heading("2. Emergency", level=1)
        emergency = preview.get("documents", {}).get("emergency")
        if emergency:
            document.add_paragraph(f"Nooddocument: {emergency}")

        document.add_heading("3. Verzekeringen", level=1)
        ins = preview.get("documents", {}).get("insurance", [])
        for link in ins:
            document.add_paragraph(f"Document: {link}")

        document.add_heading("4. Materialen", level=1)
        for mat in preview.get("materials", {}).get("avm", []):
            document.add_paragraph(f"- {mat.get('displayname','')} ({mat.get('quantity_total','')})")

        buffer = io.BytesIO()
        document.save(buffer)
        buffer.seek(0)
        return send_file(buffer, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document", as_attachment=True, download_name="dossier.docx")

    return jsonify({"error": "Unsupported format"}), 400

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
