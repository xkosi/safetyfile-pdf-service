from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from flask import Flask, request, send_file, jsonify
import io

app = Flask(__name__)

@app.route('/generate', methods=['POST'])
def generate():
    try:
        data = request.get_json(force=True)
        preview = data.get("preview", {})
        project = preview.get("project", {})
        materials = preview.get("materials", {})
        documents = preview.get("documents", {})
        responsible = preview.get("responsible", None)

        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        elements = []
        styles = getSampleStyleSheet()
        h1 = styles['Heading1']
        h2 = styles['Heading2']
        normal = styles['Normal']

        # Titel
        elements.append(Paragraph("Veiligheidsdossier", h1))
        elements.append(Spacer(1, 12))

        # Projectgegevens
        elements.append(Paragraph("Projectgegevens", h2))
        proj_table = [
            ["Project", project.get("name", "-")],
            ["Opdrachtgever", project.get("customer", {}).get("name", "-")],
            ["Adres opdrachtgever", project.get("customer", {}).get("address", "-")],
            ["Contactpersoon", project.get("customer", {}).get("contact", {}).get("name", "-")],
            ["Tel", project.get("customer", {}).get("contact", {}).get("phone", "-")],
            ["E-mail", project.get("customer", {}).get("contact", {}).get("email", "-")],
            ["Locatie", project.get("location", {}).get("name", "-")],
            ["Adres locatie", project.get("location", {}).get("address", "-")],
            ["Startdatum", project.get("project_start_date", "-")],
            ["Einddatum", project.get("project_end_date", "-")]
        ]
        table = Table(proj_table, colWidths=[150, 300])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('BOX', (0,0), (-1,-1), 0.5, colors.black),
            ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black)
        ]))
        elements.append(table)
        elements.append(Spacer(1, 12))

        # Verantwoordelijke
        elements.append(Paragraph("Projectverantwoordelijke", h2))
        elements.append(Paragraph(responsible or "-", normal))
        elements.append(Spacer(1, 12))

        # Materialen
        elements.append(Paragraph("Materialen", h2))
        mat_table_data = [["Naam", "Aantal", "Type", "CE", "Manual", "MSDS"]]
        for m in materials.get("avm", []) + materials.get("dees", []):
            links = m.get("links", {})
            mat_table_data.append([
                m.get("displayname", "-"),
                m.get("quantity_total", "-"),
                m.get("type", "-"),
                links.get("ce", "-"),
                links.get("manual", "-"),
                links.get("msds", "-")
            ])
        mat_table = Table(mat_table_data, colWidths=[150,50,50,100,100,100])
        mat_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('BOX', (0,0), (-1,-1), 0.5, colors.black),
            ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black)
        ]))
        elements.append(mat_table)
        elements.append(Spacer(1, 12))

        # Documenten
        elements.append(Paragraph("Documenten", h2))
        doc_table_data = [["Type", "Link"]]
        for key, val in documents.items():
            if isinstance(val, list):
                for v in val:
                    doc_table_data.append([key, v or "-"])
            else:
                doc_table_data.append([key, val or "-"])
        doc_table = Table(doc_table_data, colWidths=[150, 300])
        doc_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('BOX', (0,0), (-1,-1), 0.5, colors.black),
            ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black)
        ]))
        elements.append(doc_table)

        doc.build(elements)
        buffer.seek(0)
        return send_file(buffer, as_attachment=True, download_name="dossier.pdf", mimetype="application/pdf")

    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/')
def root():
    return "Service is running", 200
