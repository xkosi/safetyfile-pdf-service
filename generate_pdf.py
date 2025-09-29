# -*- coding: utf-8 -*-
"""
Corrected generate_pdf service with compatibility patch for PyPDF2 merge functions.
"""

import io, datetime, requests
from flask import Flask, request, send_file, jsonify
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from PyPDF2 import PdfReader, PdfWriter, Transformation

app = Flask(__name__)

MARGIN_LEFT, MARGIN_RIGHT, MARGIN_TOP, MARGIN_BOTTOM = 50, 50, 50, 50
RED = colors.Color(0.8,0,0)

def add_header_bar(c, title):
    c.setFillColor(RED)
    c.rect(MARGIN_LEFT, A4[1]-MARGIN_TOP-40, A4[0]-MARGIN_LEFT-MARGIN_RIGHT, 30, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(MARGIN_LEFT+10, A4[1]-MARGIN_TOP-30, title)

def merge_pdf_firstpage_under_header(base_pdf_bytes, external_pdf_bytes, header_height_mm=20):
    base_reader = PdfReader(io.BytesIO(base_pdf_bytes))
    ext_reader = PdfReader(io.BytesIO(external_pdf_bytes))

    writer = PdfWriter()
    base_page = base_reader.pages[0]
    writer.add_page(base_page)

    # Place ext first page onto base first page
    page = writer.pages[0]
    page_width = float(page.mediabox.width)
    page_height = float(page.mediabox.height)
    left = MARGIN_LEFT
    right = page_width - MARGIN_RIGHT
    top = page_height - MARGIN_TOP - (header_height_mm*2.83) - 5  # mm->pt
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
        page.merge_transformed_page(src_page, t)
    elif hasattr(page, "mergeTransformedPage"):
        page.mergeTransformedPage(src_page, t)
    else:
        raise AttributeError("Geen geldige merge-methode gevonden in PyPDF2")

    # Append remaining pages
    for i in range(1,len(base_reader.pages)):
        writer.add_page(base_reader.pages[i])
    for i in range(1,len(ext_reader.pages)):
        writer.add_page(ext_reader.pages[i])

    buf = io.BytesIO()
    writer.write(buf)
    buf.seek(0)
    return buf.getvalue()

@app.route("/generate", methods=["POST"])
def generate():
    body = request.get_json(force=True)
    fmt = body.get("format","pdf")
    # For demo, build a trivial PDF
    b = io.BytesIO()
    c = canvas.Canvas(b, pagesize=A4)
    add_header_bar(c,"Demo Veiligheidsdossier")
    c.drawString(100,500,"Hello World")
    c.save()
    pdf_bytes = b.getvalue()
    return send_file(io.BytesIO(pdf_bytes), mimetype="application/pdf", as_attachment=True, download_name="demo.pdf")

if __name__=="__main__":
    app.run(host="0.0.0.0", port=8000)
