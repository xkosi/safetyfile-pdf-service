# -*- coding: utf-8 -*-
# Simplified generate_pdf service (PDF & DOCX)
# See previous version for full details

import io, datetime
from flask import Flask, request, send_file, jsonify
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors

try:
    from docx import Document
    DOCX_AVAILABLE = True
except Exception:
    DOCX_AVAILABLE = False

app = Flask(__name__)
W, H = A4
RED = colors.HexColor("#B00000")
BLACK = colors.black

def draw_banner(c, title):
    c.setFillColor(RED)
    c.rect(0, H-50, W, 50, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(30, H-35, title)

def build_pdf(preview):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    draw_banner(c, "Veiligheidsdossier")
    pname = (preview.get("avm") or {}).get("name") or preview.get("project",{}).get("name","")
    c.setFillColor(BLACK)
    c.setFont("Helvetica-Bold", 20)
    c.drawString(30, H-100, pname)
    c.setFont("Helvetica", 9)
    c.drawString(30, 40, "Gegenereerd: %s" % datetime.datetime.now().strftime("%d/%m/%Y %H:%M"))
    c.showPage()
    draw_banner(c, "Inhoud")
    c.setFillColor(BLACK)
    c.setFont("Helvetica", 12)
    c.drawString(30, H-100, "Voorbeeld inhoudspagina")
    c.save()
    return buf.getvalue()

def build_docx(preview):
    if not DOCX_AVAILABLE:
        raise RuntimeError("python-docx not available")
    doc = Document()
    doc.add_heading("Veiligheidsdossier", 0)
    pname = (preview.get("avm") or {}).get("name") or preview.get("project",{}).get("name","")
    doc.add_paragraph(pname)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

@app.route("/generate", methods=["POST"])
def generate():
    payload = request.get_json(force=True, silent=False) or {}
    preview = payload.get("preview") or payload
    fmt = (payload.get("format") or "pdf").lower()
    if fmt=="docx":
        try:
            return send_file(io.BytesIO(build_docx(preview)),
                mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                as_attachment=True, download_name="dossier.docx")
        except Exception as e:
            return jsonify({"error":"DOCX generation failed","detail":str(e)}), 500
    try:
        pdf_bytes = build_pdf(preview)
        return send_file(io.BytesIO(pdf_bytes), mimetype="application/pdf",
                         as_attachment=True, download_name="dossier.pdf")
    except Exception as e:
        return jsonify({"error":"PDF generation failed","detail":str(e)}), 500

if __name__=="__main__":
    app.run(host="0.0.0.0", port=8000)
