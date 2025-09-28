# generate_pdf.py
# ------------------------------------------------------------
# Bouwt een compleet Veiligheidsdossier als PDF volgens
# jouw voorbeelddossier-structuur, met Pyred-branding en AVM-logo.
#
# Vereisten:
#   pip install reportlab pypdf requests pillow
#
# Gebruik:
#   python generate_pdf.py --preview preview.json --out dossier.pdf
# ------------------------------------------------------------

import io
import os
import sys
import json
import math
import argparse
from typing import List, Dict, Any, Optional, Tuple

# Externe libs
import requests
from PIL import Image

# PDF libs
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak, Flowable
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# pypdf (voorkeur) of PyPDF2 fallback
try:
    from pypdf import PdfReader, PdfWriter, PageObject
    HAS_PYPDF = True
except Exception:
    from PyPDF2 import PdfReader, PdfWriter
    HAS_PYPDF = False

# ------------------------------------------------------------
# Branding / layout
# ------------------------------------------------------------
PAGE_SIZE = A4
PAGE_W, PAGE_H = PAGE_SIZE

# Pyred kleuren (default; kan je finetunen vanuit style.css later)
PYRED_PRIMARY = colors.HexColor("#E30613")
PYRED_TEXT = colors.HexColor("#111111")
PYRED_LIGHT = colors.HexColor("#F5F5F7")
LINE_GRAY = colors.HexColor("#DDDDDD")

MARGIN_L = 20*mm
MARGIN_R = 20*mm
MARGIN_T = 20*mm
MARGIN_B = 20*mm

STYLES = getSampleStyleSheet()
STYLES.add(ParagraphStyle(name="TitlePyred", fontName="Helvetica-Bold", fontSize=22, leading=26, textColor=PYRED_PRIMARY, spaceAfter=8))
STYLES.add(ParagraphStyle(name="H1", fontName="Helvetica-Bold", fontSize=16, leading=20, textColor=PYRED_PRIMARY, spaceAfter=8))
STYLES.add(ParagraphStyle(name="H2", fontName="Helvetica-Bold", fontSize=13, leading=16, textColor=PYRED_PRIMARY, spaceAfter=6))
STYLES.add(ParagraphStyle(name="Body", fontName="Helvetica", fontSize=10.5, leading=14, textColor=PYRED_TEXT))
STYLES.add(ParagraphStyle(name="SmallGray", fontName="Helvetica", fontSize=9, leading=12, textColor=colors.grey))

# ------------------------------------------------------------
# Utilities
# ------------------------------------------------------------

def http_get_bytes(url: str, timeout: int = 25) -> Optional[bytes]:
    if not url:
        return None
    try:
        r = requests.get(url, timeout=timeout)
        if r.ok:
            return r.content
        return None
    except Exception:
        return None

def datauri_to_bytes(data_uri: str) -> Optional[bytes]:
    # verwacht 'data:application/pdf;base64,....'
    try:
        import base64
        if not data_uri or not data_uri.startswith("data:"):
            return None
        header, b64 = data_uri.split(",", 1)
        return base64.b64decode(b64)
    except Exception:
        return None

def get_image_for_rl(url_or_data: str, max_w: float) -> Optional[RLImage]:
    # haalt beeld op (http of data:) en maakt RLImage met max breedte
    raw = None
    if not url_or_data:
        return None
    if url_or_data.startswith("data:"):
        raw = datauri_to_bytes(url_or_data)
    else:
        raw = http_get_bytes(url_or_data)
    if not raw:
        return None
    try:
        im = Image.open(io.BytesIO(raw)).convert("RGB")
        w, h = im.size
        scale = min(max_w / w, 1.0)
        new_w = int(w * scale)
        new_h = int(h * scale)
        im = im.resize((new_w, new_h), Image.LANCZOS)
        buf = io.BytesIO()
        im.save(buf, format="PNG")
        buf.seek(0)
        return RLImage(buf, width=new_w, height=new_h)
    except Exception:
        return None

def write_rl_pdf(flowables: List[Flowable]) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=PAGE_SIZE,
        leftMargin=MARGIN_L, rightMargin=MARGIN_R,
        topMargin=MARGIN_T, bottomMargin=MARGIN_B,
        title="Veiligheidsdossier"
    )
    doc.build(flowables)
    return buf.getvalue()

def append_pdf(writer: "PdfWriter", pdf_bytes: bytes, scale_to_a4: bool = True, start_on_new_page: bool = True):
    """Voegt alle pagina's van pdf_bytes toe aan writer.
       Als pypdf beschikbaar is, schaalt naar A4 (portret)."""
    reader = PdfReader(io.BytesIO(pdf_bytes))
    for i, page in enumerate(reader.pages):
        p = page
        if HAS_PYPDF and scale_to_a4:
            # Probeer de pagina te schalen naar A4 (595x842 pt)
            try:
                # pypdf: PageObject heeft methodes scale_to
                # Fallback: calc factor en transformeer mediabox
                pw = float(p.mediabox.width)
                ph = float(p.mediabox.height)
                sx = PAGE_W / pw
                sy = PAGE_H / ph
                s = min(sx, sy)
                # transformeer pagina
                p.scale_by(s)
                # Centreer op A4
                # p.trimbox en p.mediabox worden niet automatisch herzet; we 'embedden' hem in blanco A4
                new_page = writer.add_blank_page(PAGE_W, PAGE_H)
                # merge scaled page at centered position
                tx = (PAGE_W - float(p.mediabox.width)) / 2.0
                ty = (PAGE_H - float(p.mediabox.height)) / 2.0
                new_page.merge_translated_page(p, tx, ty)
                continue  # volgende pagina
            except Exception:
                pass
        # Zonder schalen: gewoon toevoegen
        writer.add_page(p)

def add_titled_section_page(title: str) -> bytes:
    # simpele titelpagina voor begin van een sectie
    flow = []
    flow.append(Spacer(1, 10*mm))
    flow.append(Paragraph(title, STYLES["H1"]))
    # subtiele rode lijn
    tbl = Table([[""]], colWidths=[PAGE_W - MARGIN_L - MARGIN_R], rowHeights=[2])
    tbl.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,-1), PYRED_PRIMARY)]))
    flow.append(tbl)
    flow.append(Spacer(1, 8*mm))
    return write_rl_pdf(flow)

def build_toc_page(toc_items: List[Tuple[str,int]]) -> bytes:
    # 1 pagina (of meerdere) met inhoudstafel. Simpel layout.
    flow = []
    flow.append(Spacer(1, 6*mm))
    flow.append(Paragraph("Inhoudstafel", STYLES["H1"]))
    data = []
    for title, pageno in toc_items:
        # dotted leader effect
        dots = "." * max(2, 80 - len(title))
        data.append([Paragraph(title, STYLES["Body"]), Paragraph(f"{pageno}", STYLES["Body"])])
    tbl = Table(data, colWidths=[(PAGE_W - MARGIN_L - MARGIN_R)-30*mm, 30*mm])
    tbl.setStyle(TableStyle([
        ("LINEBELOW", (0,0), (-1,-1), 0.25, LINE_GRAY),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("RIGHTPADDING", (-1,0), (-1,-1), 0),
        ("LEFTPADDING", (0,0), (0,-1), 0),
    ]))
    flow.append(tbl)
    return write_rl_pdf(flow)

# ------------------------------------------------------------
# Secties (ReportLab)
# ------------------------------------------------------------

def section_cover(avm_logo_url: Optional[str], project_name: str) -> bytes:
    flow = []
    # Boven balk in Pyred
    flow.append(Spacer(1, 4*mm))
    flow.append(Paragraph("Veiligheidsdossier", STYLES["TitlePyred"]))
    flow.append(Spacer(1, 2*mm))
    flow.append(Paragraph(project_name or "", STYLES["H2"]))
    flow.append(Spacer(1, 20*mm))

    # Logo
    if avm_logo_url:
        img = get_image_for_rl(avm_logo_url, max_w=(PAGE_W - MARGIN_L - MARGIN_R) * 0.35)
        if img:
            flow.append(img)
            flow.append(Spacer(1, 10*mm))

    # subtiele blok
    flow.append(Paragraph("Pyred • AVM Special Effects", STYLES["SmallGray"]))
    return write_rl_pdf(flow)

def section_project_table(project: Dict[str, Any]) -> bytes:
    # Maak tabel zoals in het voorbeelddossier: brede lijnen / form-look
    # We gebruiken simpele key/value regels.
    rows = []
    def row(label, value):
        rows.append([Paragraph(f"<b>{label}</b>", STYLES["Body"]), Paragraph(value or "-", STYLES["Body"])])

    row("Project", project.get("name") or project.get("displayname") or "-")
    cust = project.get("customer") or {}
    row("Opdrachtgever", cust.get("name") or "-")
    row("Adres opdrachtgever", cust.get("address") or "-")
    loc = project.get("location") or {}
    row("Locatie", loc.get("name") or "-")
    row("Adres locatie", loc.get("address") or "-")
    row("Periode", f'{project.get("project_start_date") or "-"}  →  {project.get("project_end_date") or "-"}')

    col_w = [(PAGE_W - MARGIN_L - MARGIN_R) * 0.28, (PAGE_W - MARGIN_L - MARGIN_R) * 0.72]
    tbl = Table(rows, colWidths=col_w, hAlign="LEFT")
    tbl.setStyle(TableStyle([
        ("LINEABOVE", (0,0), (-1,0), 1.2, LINE_GRAY),
        ("LINEBELOW", (0,0), (-1,-1), 0.8, LINE_GRAY),
        ("LINEBELOW", (0,0), (-1,0), 1.2, LINE_GRAY),
        ("LINEBELOW", (0,-1), (-1,-1), 1.2, LINE_GRAY),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING", (0,0), (-1,-1), 6),
        ("BOTTOMPADDING", (0,0), (-1,-1), 8),
    ]))

    flow = []
    flow.append(Paragraph("Projectgegevens", STYLES["H1"]))
    flow.append(tbl)
    return write_rl_pdf(flow)

def section_responsible(mini_url: Optional[str], name: str, full_bio_url: Optional[str]) -> bytes:
    flow = []
    flow.append(Paragraph("Verantwoordelijke", STYLES["H1"]))
    # mini foto + naam + link
    row = []
    img = get_image_for_rl(mini_url, max_w=60*mm) if mini_url else None
    if img:
        row.append(img)
    else:
        # lege placeholder
        placeholder = Table([[" "]], colWidths=[60*mm], rowHeights=[40*mm])
        placeholder.setStyle(TableStyle([("BOX",(0,0),(-1,-1),0.25,LINE_GRAY)]))
        row.append(placeholder)

    # Tekstkolom
    bio_parts = [Paragraph(f"<b>{name or '-'}</b>", STYLES["Body"])]
    if full_bio_url:
        bio_parts.append(Spacer(1, 2*mm))
        bio_parts.append(Paragraph(f'<font color="#1a0dab">Open volledige bio</font><br/>{full_bio_url}', STYLES["SmallGray"]))
    right = []
    right.extend(bio_parts)

    t = Table([[row[0], right]], colWidths=[60*mm, (PAGE_W - MARGIN_L - MARGIN_R) - 60*mm])
    t.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP")]))
    flow.append(t)
    return write_rl_pdf(flow)

def _material_links_cell(files: List[Dict[str,Any]]) -> Paragraph:
    if not files:
        return Paragraph("-", STYLES["Body"])
    # toon labels; we gebruiken displayname als label, of heuristisch CE/Manual/MSDS in naam
    labels = []
    for f in files:
        name = (f.get("name") or f.get("displayname") or "").strip()
        url = f.get("url") or ""
        if not url:
            continue
        if "manual" in name.lower():
            lbl = "Manual"
        elif "conform" in name.lower() or "declaration" in name.lower() or "CE" in name:
            lbl = "CE"
        elif "msds" in name.lower() or "sds" in name.lower():
            lbl = "MSDS"
        else:
            lbl = name[:24] + ("…" if len(name)>24 else "")
        labels.append(f'<a href="{url}" color="blue">{lbl}</a>')
    if not labels:
        return Paragraph("-", STYLES["Body"])
    return Paragraph(" | ".join(labels), STYLES["Body"])

def section_materials_avm(avm_items: List[Dict[str,Any]]) -> bytes:
    flow = []
    flow.append(Paragraph("Materialen — AVM", STYLES["H1"]))
    if not avm_items:
        flow.append(Paragraph("Geen AVM-materialen geselecteerd.", STYLES["Body"]))
        return write_rl_pdf(flow)

    data = [[Paragraph("<b>Naam</b>", STYLES["Body"]),
             Paragraph("<b>Aantal</b>", STYLES["Body"]),
             Paragraph("<b>Type</b>", STYLES["Body"]),
             Paragraph("<b>Links</b>", STYLES["Body"])]]

    for it in avm_items:
        # kleine foto links van naam (indien image)
        name_flow = []
        if it.get("image"):
            img = get_image_for_rl(it["image"], max_w=26*mm)
            if img:
                name_flow.append(img)
                name_flow.append(Spacer(1,2*mm))
        name_flow.append(Paragraph(it.get("displayname","-"), STYLES["Body"]))

        links_par = _material_links_cell(it.get("files") or [])

        row = [name_flow, str(it.get("quantity_total",0)), it.get("type","-"), links_par]
        data.append(row)

    col_w = [ (PAGE_W - MARGIN_L - MARGIN_R)*0.50,
              (PAGE_W - MARGIN_L - MARGIN_R)*0.12,
              (PAGE_W - MARGIN_L - MARGIN_R)*0.14,
              (PAGE_W - MARGIN_L - MARGIN_R)*0.24 ]

    tbl = Table(data, colWidths=col_w, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.25, LINE_GRAY),
        ("BACKGROUND", (0,0), (-1,0), PYRED_LIGHT),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
        ("RIGHTPADDING", (0,0), (-1,-1), 6),
    ]))

    flow.append(tbl)
    return write_rl_pdf(flow)

def section_materials_dees(dees_items: List[Dict[str,Any]]) -> bytes:
    # Verwachting: elk item heeft custom_13 met T1/T2/F1..F4 etc, en optioneel NEC/CE info in files/naam
    flow = []
    flow.append(Paragraph("Materialen — Dees (Pyro & Vuurwerk)", STYLES["H1"]))

    if not dees_items:
        flow.append(Paragraph("Geen Dees-materialen geselecteerd.", STYLES["Body"]))
        return write_rl_pdf(flow)

    # Split op pyro (T1/T2) en vuurwerk (F1..F4)
    pyro = [x for x in dees_items if str(x.get("custom_13","")).upper().startswith("T")]
    vuurwerk = [x for x in dees_items if str(x.get("custom_13","")).upper().startswith("F")]

    def build_table(items: List[Dict[str,Any]], title: str) -> Flowable:
        data = [[Paragraph("<b>Class</b>", STYLES["Body"]),
                 Paragraph("<b>Product</b>", STYLES["Body"]),
                 Paragraph("<b>CE</b>", STYLES["Body"]),
                 Paragraph("<b>NEC</b>", STYLES["Body"]),
                 Paragraph("<b>Aantal</b>", STYLES["Body"]),
                 Paragraph("<b>Totaal</b>", STYLES["Body"])]]

        total_all = 0.0
        for it in items:
            cls = it.get("custom_13","-")
            name = it.get("displayname","-")
            qty = float(it.get("quantity_total",0) or 0)
            # heuristiek: CE / NEC uit bestandsnamen indien beschikbaar
            ce = "-"
            nec = 0.0
            for f in (it.get("files") or []):
                fname = (f.get("name") or "").lower()
                if "nec" in fname:
                    # simpel: haal getal vóór 'g' uit naam
                    import re
                    m = re.search(r'(\d+(?:[\.,]\d+)?)\s*g', fname)
                    if m:
                        nec = float(m.group(1).replace(",", "."))
                if "ce" in fname or "conform" in fname or "declaration" in fname:
                    ce = "CE"
            total = nec * qty
            total_all += total
            data.append([cls, name, ce, f"{nec:g} g", f"{int(qty)}", f"{total:g} g"])

        # voeg eindtotaal toe
        data.append(["", "", "", "", Paragraph("<b>Totaal</b>", STYLES["Body"]), Paragraph(f"<b>{total_all:g} g</b>", STYLES["Body"])])

        col_w = [20*mm,
                 (PAGE_W - MARGIN_L - MARGIN_R) - (20+22+22+20+26)*mm,  # rest voor naam
                 22*mm, 22*mm, 20*mm, 26*mm]

        tbl = Table(data, colWidths=col_w, repeatRows=1)
        tbl.setStyle(TableStyle([
            ("GRID", (0,0), (-1,-1), 0.25, LINE_GRAY),
            ("BACKGROUND", (0,0), (-1,0), PYRED_LIGHT),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("LEFTPADDING", (0,0), (-1,-1), 6),
            ("RIGHTPADDING", (0,0), (-1,-1), 6),
        ]))
        wrapper = []
        wrapper.append(Paragraph(title, STYLES["H2"]))
        wrapper.append(tbl)
        return wrapper

    flow.extend(build_table(pyro, "Pyro (T1/T2)"))
    flow.append(Spacer(1, 6*mm))
    flow.extend(build_table(vuurwerk, "Vuurwerk (F1–F4)"))
    return write_rl_pdf(flow)

# ------------------------------------------------------------
# Hoofdcompositie met TOC en externe PDF-embeds
# ------------------------------------------------------------

def bytes_for_uploaded_list(file_list: List[Dict[str,Any]]) -> List[Tuple[str, bytes]]:
    out = []
    for f in file_list or []:
        name = f.get("name") or "upload.pdf"
        data = f.get("data") or ""
        raw = None
        if data.startswith("data:"):
            raw = datauri_to_bytes(data)
        else:
            raw = http_get_bytes(data)
        if raw:
            out.append((name, raw))
    return out

def fetch_pdf(url_or_datauri: str) -> Optional[bytes]:
    if not url_or_datauri:
        return None
    if url_or_datauri.startswith("data:"):
        return datauri_to_bytes(url_or_datauri)
    return http_get_bytes(url_or_datauri)

def generate_pdf(preview: Dict[str,Any], out_path: str):
    # 1) Bouw alle sectie-PDFs en houd bij hoeveel pagina’s elke sectie heeft
    sections: List[Tuple[str, bytes]] = []  # (titel, pdf_bytes)
    toc: List[Tuple[str,int]] = []          # (titel, absolute_start_page)

    # Cover
    cover = section_cover(
        avm_logo_url="avm-logo.png",  # pas aan als je een absolute URL gebruikt
        project_name=(preview.get("avm") or {}).get("name") or (preview.get("dees") or {}).get("name") or ""
    )
    sections.append(("Cover", cover))

    # Projectgegevens
    project_obj = preview.get("avm") or {}
    sections.append(("1. Projectgegevens", section_project_table(project_obj)))

    # 2. Emergency (embed PDF)
    emergency_url = ((preview.get("documents") or {}).get("emergency") or "")
    if emergency_url:
        sections.append(("2. Emergency", add_titled_section_page("2. Emergency")))
        emb = fetch_pdf(emergency_url)
        if emb:
            sections.append(("2. Emergency (bijlage)", emb))

    # 3. Verzekeringen (1 of 2 PDF’s)
    insurance = (preview.get("documents") or {}).get("insurance") or []
    if insurance:
        sections.append(("3. Verzekeringen", add_titled_section_page("3. Verzekeringen")))
        for i, url in enumerate(insurance, start=1):
            emb = fetch_pdf(url)
            if emb:
                sections.append((f"3.{i} Verzekering", emb))

    # 4. Verantwoordelijke (mini + link naar full)
    crew_key = preview.get("responsible")
    crew_bios = (preview.get("documents") or {}).get("crew_bio_mini"), (preview.get("documents") or {}).get("crew_bio_full")
    mini = (preview.get("documents") or {}).get("crew_bio_mini")
    fullb = (preview.get("documents") or {}).get("crew_bio_full")
    sections.append(("4. Verantwoordelijke", section_responsible(mini, crew_key or "-", fullb)))

    # 5. Materialen (sub: Dees, AVM) — alleen als via wizard gekozen
    mats = preview.get("materials") or {}
    if (mats.get("dees") or []) or (mats.get("avm") or []):
        sections.append(("5. Materialen", add_titled_section_page("5. Materialen")))
        if mats.get("dees"):
            sections.append(("5.1 Dees", section_materials_dees(mats.get("dees"))))
        if mats.get("avm"):
            sections.append(("5.2 AVM", section_materials_avm(mats.get("avm"))))

    # 6. Inplantingsplan (uploads.siteplan — embed)
    uploads = (preview.get("uploads") or {})
    siteplan_list = bytes_for_uploaded_list(uploads.get("siteplan") or [])
    if siteplan_list:
        sections.append(("6. Inplantingsplan", add_titled_section_page("6. Inplantingsplan")))
        for name, raw in siteplan_list:
            sections.append((f"6.* {name}", raw))

    # 7. Risicoanalyse Pyro & Special Effects (documentlinks in KV)
    risk_pyro = ((preview.get("documents") or {}).get("risk_pyro") or "")
    risk_general = ((preview.get("documents") or {}).get("risk_general") or "")
    if risk_pyro or risk_general:
        sections.append(("7. Risicoanalyses", add_titled_section_page("7. Risicoanalyses")))
        if risk_pyro:
            raw = fetch_pdf(risk_pyro)
            if raw:
                sections.append(("7.1 Risicoanalyse Pyro", raw))
        if risk_general:
            raw = fetch_pdf(risk_general)
            if raw:
                sections.append(("7.2 Risicoanalyse Special Effects", raw))

    # 8. Windplan (optioneel)
    if ((preview.get("documents") or {}).get("windplan")):
        raw = fetch_pdf(preview["documents"]["windplan"])
        if raw:
            sections.append(("8. Windplan", add_titled_section_page("8. Windplan")))
            sections.append(("8.* Windplan (bijlage)", raw))

    # 9. Droogteplan (optioneel)
    if ((preview.get("documents") or {}).get("droughtplan")):
        raw = fetch_pdf(preview["documents"]["droughtplan"])
        if raw:
            sections.append(("9. Droogteplan", add_titled_section_page("9. Droogteplan")))
            sections.append(("9.* Droogteplan (bijlage)", raw))

    # 10. Vergunningen & Toelatingen (uploads.permits — embed)
    permit_list = bytes_for_uploaded_list(uploads.get("permits") or [])
    if permit_list:
        sections.append(("10. Vergunningen & Toelatingen", add_titled_section_page("10. Vergunningen & Toelatingen")))
        for name, raw in permit_list:
            sections.append((f"10.* {name}", raw))

    # -------------------------------
    # Twee-pass: eerst secties samenvoegen & paginanummers verzamelen,
    # dan TOC-pagina bouwen, en alles opnieuw samenstellen.
    # -------------------------------
    # Pass 1: concat & tel pagina’s per sectie
    tmp_writer = PdfWriter()
    page_cursor = 1
    section_starts: List[Tuple[str,int]] = []  # titel -> startpagina

    for title, pdfb in sections:
        section_starts.append((title, page_cursor))
        append_pdf(tmp_writer, pdfb, scale_to_a4=True, start_on_new_page=True)
        # update cursor
        added = len(PdfReader(io.BytesIO(pdfb)).pages)
        page_cursor += added

    # Bouw inhoudstafel-items op basis van alleen hoofdsecties (met nummering aan begin)
    def is_chapter(t: str) -> bool:
        # Hoofdstukregels: beginnen met "N. " (bv "5. Materialen")
        return any(t.startswith(f"{i}. ") for i in range(1, 11)) or t == "Cover"

    toc_items: List[Tuple[str,int]] = [(t, p) for (t, p) in section_starts if is_chapter(t)]

    toc_pdf = build_toc_page(toc_items)

    # Pass 2: definitieve writer — Cover, TOC, rest
    final_writer = PdfWriter()
    # Cover (eerste sectie is cover)
    cover_title, cover_bytes = sections[0]
    append_pdf(final_writer, cover_bytes, scale_to_a4=True)

    # TOC
    append_pdf(final_writer, toc_pdf, scale_to_a4=True)

    # Rest
    for idx, (title, pdfb) in enumerate(sections[1:], start=1):
        append_pdf(final_writer, pdfb, scale_to_a4=True)

    # Wegschrijven
    with open(out_path, "wb") as f:
        final_writer.write(f)

# ------------------------------------------------------------
# CLI
# ------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description="Genereer Veiligheidsdossier PDF uit preview.json")
    ap.add_argument("--preview", required=True, help="Pad naar preview.json")
    ap.add_argument("--out", default="dossier.pdf", help="Uitvoerbestand (PDF)")
    args = ap.parse_args()

    with open(args.preview, "r", encoding="utf-8") as fh:
        preview = json.load(fh)

    generate_pdf(preview, args.out)
    print(f"✅ Klaar: {args.out}")

if __name__ == "__main__":
    main()
