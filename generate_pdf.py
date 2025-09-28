import io
import base64
import math
import re
from typing import List, Dict, Any, Optional, Tuple

import requests
from fastapi import FastAPI, Body
from fastapi.responses import Response, PlainTextResponse
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Paragraph
from pypdf import PdfReader, PdfWriter

# ---------------------------
# Pyred / AVM huisstijl
# ---------------------------
PYRED_RED = colors.HexColor("#E30613")
TEXT = colors.HexColor("#111111")
MUTED = colors.HexColor("#666666")
TABLE_HEADER_BG = colors.HexColor("#F3F3F4")

PAGE_W, PAGE_H = A4
MARGIN = 18 * mm
LINE_H = 7.5 * mm
FONT = "Helvetica"
FONT_B = "Helvetica-Bold"

# ---------------------------
# Webservice
# ---------------------------
app = FastAPI(title="Safetyfile PDF Service", version="1.0.0")


# =========================================================
# Helpers
# =========================================================

def _new_canvas(buf: io.BytesIO) -> canvas.Canvas:
    c = canvas.Canvas(buf, pagesize=A4)
    c.setTitle("Veiligheidsdossier")
    return c


def _draw_header_bar(c: canvas.Canvas, title: str):
    c.setFillColor(PYRED_RED)
    c.rect(0, PAGE_H - 20 * mm, PAGE_W, 20 * mm, stroke=0, fill=1)
    c.setFillColor(colors.white)
    c.setFont(FONT_B, 18)
    c.drawString(MARGIN, PAGE_H - 13 * mm, title)


def _draw_logo(c: canvas.Canvas, logo_url: Optional[str], x: float, y: float, w: float, h: float):
    if not logo_url:
        return
    try:
        if logo_url.startswith("data:"):
            head, data = logo_url.split(",", 1)
            img = ImageReader(io.BytesIO(base64.b64decode(data)))
        else:
            resp = requests.get(logo_url, timeout=15)
            resp.raise_for_status()
            img = ImageReader(io.BytesIO(resp.content))
        c.drawImage(img, x, y, w, h, preserveAspectRatio=True, mask="auto")
    except Exception:
        pass


def _text(c: canvas.Canvas, x: float, y: float, txt: str, size=11, bold=False, color=TEXT):
    c.setFillColor(color)
    c.setFont(FONT_B if bold else FONT, size)
    c.drawString(x, y, txt)


def _para(c: canvas.Canvas, x: float, y: float, w: float, txt: str, size=11):
    style = ParagraphStyle(
        "body", fontName=FONT, fontSize=size, leading=size * 1.2, textColor=TEXT
    )
    p = Paragraph(txt, style)
    a_w = w
    a_h = 5000
    a = p.wrap(a_w, a_h)
    p.drawOn(c, x, y - a[1])
    return a[1]


def _underline_field(c: canvas.Canvas, label: str, value: str, x: float, y: float, w: float):
    _text(c, x, y + 3, label, size=10, bold=True, color=PYRED_RED)
    c.setStrokeColor(MUTED)
    c.line(x, y, x + w, y)
    _text(c, x + 2, y - 12, value or "—", size=12, bold=False, color=TEXT)


def _draw_table(
    c: canvas.Canvas,
    x: float,
    y: float,
    colspec: List[Tuple[str, float]],
    rows: List[List[str]],
    row_h: float = 8 * mm,
    header_bg=TABLE_HEADER_BG,
):
    # header
    c.setFillColor(header_bg)
    c.rect(x, y - row_h, sum(w for _, w in colspec), row_h, stroke=0, fill=1)
    c.setStrokeColor(colors.black)
    c.setLineWidth(0.5)
    c.rect(x, y - row_h, sum(w for _, w in colspec), row_h, stroke=1, fill=0)
    c.setFillColor(TEXT)
    c.setFont(FONT_B, 10)
    cur_x = x + 2
    for name, width in colspec:
        c.drawString(cur_x, y - row_h + 2.5, name)
        cur_x += width
    # rows
    c.setFont(FONT, 10)
    y_row = y - row_h
    for r in rows:
        y_row -= row_h
        c.setStrokeColor(colors.black)
        c.rect(x, y_row, sum(w for _, w in colspec), row_h, stroke=1, fill=0)
        cur_x = x + 2
        for i, (_, w) in enumerate(colspec):
            val = r[i] if i < len(r) else ""
            c.drawString(cur_x, y_row + 2.5, str(val))
            cur_x += w
    return y_row


def _fetch_image_reader(url_or_data: Optional[str]) -> Optional[ImageReader]:
    if not url_or_data:
        return None
    try:
        if url_or_data.startswith("data:"):
            b64 = url_or_data.split(",", 1)[1]
            return ImageReader(io.BytesIO(base64.b64decode(b64)))
        r = requests.get(url_or_data, timeout=20)
        r.raise_for_status()
        return ImageReader(io.BytesIO(r.content))
    except Exception:
        return None


def _pdf_bytes_from_url_or_data(ref: str) -> Optional[bytes]:
    try:
        if ref.startswith("data:"):
            return base64.b64decode(ref.split(",", 1)[1])
        r = requests.get(ref, timeout=25)
        r.raise_for_status()
        return r.content
    except Exception:
        return None


def _chapter_page(title: str, subtitle: Optional[str] = None, logo_url: Optional[str] = None) -> bytes:
    buf = io.BytesIO()
    c = _new_canvas(buf)
    _draw_header_bar(c, title)
    if subtitle:
        _text(c, MARGIN, PAGE_H - 35 * mm, subtitle, size=14, bold=False)
    _draw_logo(c, logo_url, PAGE_W - 50 * mm, PAGE_H - 18 * mm, 40 * mm, 12 * mm)
    c.showPage()
    c.save()
    return buf.getvalue()


def _append_pdf(writer: PdfWriter, pdf_bytes: bytes):
    reader = PdfReader(io.BytesIO(pdf_bytes))
    for page in reader.pages:
        writer.add_page(page)


# =========================================================
# Secties – genereren als losse PDF-bytes
# =========================================================

def section_cover(preview: Dict[str, Any]) -> bytes:
    title = "VEILIGHEIDSDOSSIER"
    project_name = (preview.get("avm") or {}).get("name") or (preview.get("project") or {}).get("name") or ""
    logo = preview.get("branding", {}).get("logo") or preview.get("logo") or "https://sfx.rentals/projects/media/logo.png"

    buf = io.BytesIO()
    c = _new_canvas(buf)

    # Pyred accent
    c.setFillColor(PYRED_RED)
    c.rect(0, PAGE_H * 0.72, PAGE_W, PAGE_H * 0.28, stroke=0, fill=1)

    # AVM logo
    _draw_logo(c, logo, MARGIN, PAGE_H - 40 * mm, 50 * mm, 16 * mm)

    # Titel
    c.setFillColor(colors.white)
    c.setFont(FONT_B, 36)
    c.drawString(MARGIN, PAGE_H * 0.78, title)

    # Projectnaam
    c.setFillColor(colors.white)
    c.setFont(FONT_B, 18)
    c.drawString(MARGIN, PAGE_H * 0.78 - 20 * mm, project_name or "—")

    c.showPage()
    c.save()
    return buf.getvalue()


def section_toc(simple_items: List[str]) -> bytes:
    buf = io.BytesIO()
    c = _new_canvas(buf)
    _draw_header_bar(c, "INHOUDSTAFEL")
    y = PAGE_H - 30 * mm
    c.setFont(FONT, 12)
    c.setFillColor(TEXT)
    for i, item in enumerate(simple_items, 1):
        c.drawString(MARGIN, y, f"{i}. {item}")
        y -= 8 * mm
        if y < 40 * mm:
            c.showPage()
            _draw_header_bar(c, "INHOUDSTAFEL (vervolg)")
            y = PAGE_H - 30 * mm
    c.showPage()
    c.save()
    return buf.getvalue()


def section_project_info(preview: Dict[str, Any]) -> bytes:
    p = preview.get("avm") or preview.get("project") or {}
    customer = p.get("customer") or {}
    contact = customer.get("contact") or {}
    location = p.get("location") or {}

    buf = io.BytesIO()
    c = _new_canvas(buf)
    _draw_header_bar(c, "PROJECTGEGEVENS (AVM)")

    x = MARGIN
    y = PAGE_H - 35 * mm

    # Tabel-stijl (brede onderstreepte regels)
    col_w = (PAGE_W - 2 * MARGIN)
    _underline_field(c, "Project", p.get("name", ""), x, y, col_w); y -= 18 * mm
    _underline_field(c, "Opdrachtgever", customer.get("name", ""), x, y, col_w); y -= 18 * mm
    _underline_field(c, "Adres", customer.get("address", ""), x, y, col_w); y -= 18 * mm
    _underline_field(c, "Contactpersoon", contact.get("name", ""), x, y, col_w); y -= 18 * mm
    _underline_field(c, "Tel", contact.get("phone", ""), x, y, col_w); y -= 18 * mm
    _underline_field(c, "E-mail", contact.get("email", ""), x, y, col_w); y -= 18 * mm
    _underline_field(c, "Locatie", location.get("name", ""), x, y, col_w); y -= 18 * mm
    _underline_field(c, "Adres locatie", location.get("address", ""), x, y, col_w); y -= 18 * mm

    # Datums
    start = p.get("project_start_date") or p.get("start_date") or ""
    end = p.get("project_end_date") or p.get("end_date") or ""
    _underline_field(c, "Start", start, x, y, col_w * 0.48)
    _underline_field(c, "Einde", end, x + col_w * 0.52, y, col_w * 0.48); y -= 18 * mm

    c.showPage()
    c.save()
    return buf.getvalue()


def section_responsible(preview: Dict[str, Any]) -> bytes:
    docs = preview.get("documents", {}) or {}
    mini_url = docs.get("crew_bio_mini")
    full_url = docs.get("crew_bio_full")
    responsible = preview.get("responsible") or "Verantwoordelijke"

    buf = io.BytesIO()
    c = _new_canvas(buf)
    _draw_header_bar(c, "VERANTWOORDELIJKE")

    y = PAGE_H - 35 * mm
    _text(c, MARGIN, y, responsible, size=16, bold=True); y -= 10 * mm

    # Mini bio afbeelding
    img = _fetch_image_reader(mini_url)
    if img:
        c.drawImage(img, MARGIN, y - 60 * mm, width=60 * mm, height=60 * mm, preserveAspectRatio=True, mask="auto")
    _para(
        c,
        MARGIN + 65 * mm,
        y,
        PAGE_W - MARGIN * 2 - 65 * mm,
        "De mini-bio hierboven is aanklikbaar in het digitale dossier. De volledige bio is beschikbaar via onderstaande link."
    )
    y -= 70 * mm

    if full_url:
        _text(c, MARGIN, y, "Volledige bio:", size=12, bold=True); y -= 6 * mm
        c.setFillColor(PYRED_RED)
        c.setFont(FONT, 11)
        c.linkURL(full_url, (MARGIN, y - 2, MARGIN + 400, y + 12))
        c.drawString(MARGIN, y, full_url)
        y -= 10 * mm

    c.showPage()
    c.save()
    return buf.getvalue()


def _classify_dees(item: Dict[str, Any]) -> str:
    # verwacht custom_13 met T1/T2 of F1..F4
    c = (item.get("custom_13") or item.get("custom", {}).get("custom_13") or "").strip().upper()
    if re.fullmatch(r"T[12]", c):
        return c
    if re.fullmatch(r"F[1-4]", c):
        return c
    return ""


def section_materials_dees(preview: Dict[str, Any]) -> bytes:
    items: List[Dict[str, Any]] = (preview.get("materials") or {}).get("dees") or []
    if not items:
        # lege sectie
        return _chapter_page("MATERIALEN – DEES", "Geen materialen geselecteerd.")

    # opsplitsen
    pyro = [i for i in items if _classify_dees(i).startswith("T")]
    fireworks = [i for i in items if _classify_dees(i).startswith("F")]

    def build_table_for(label: str, rows_src: List[Dict[str, Any]]) -> bytes:
        buf = io.BytesIO()
        c = _new_canvas(buf)
        _draw_header_bar(c, f"MATERIALEN – DEES ({label})")
        y = PAGE_H - 30 * mm

        colspec = [
            ("Code", 30 * mm),
            ("Naam", 90 * mm),
            ("Class", 20 * mm),
            ("NEC", 20 * mm),
            ("Aantal", 20 * mm),
            ("Totaal NEC", 25 * mm),
        ]

        rows = []
        total_nec = 0.0

        for it in rows_src:
            nec = it.get("nec") or it.get("custom", {}).get("nec") or ""
            try:
                nec_val = float(str(nec).replace(",", "."))
            except Exception:
                nec_val = 0.0
            qty = int(it.get("quantity_total") or it.get("qty") or 0)
            rows.append([
                it.get("code") or "—",
                it.get("displayname") or "—",
                _classify_dees(it) or "—",
                str(nec or "—"),
                str(qty),
                f"{nec_val * max(qty,1):.2f}" if nec_val else "—"
            ])
            total_nec += nec_val * max(qty, 1)

        _draw_table(c, MARGIN, y, colspec, rows)
        y = 40 * mm
        _text(c, MARGIN, y, f"Totaal NEC voor {label}: {total_nec:.2f}" if rows else "—", size=12, bold=True)
        c.showPage()
        c.save()
        return buf.getvalue()

    parts: List[bytes] = []
    if pyro:
        parts.append(build_table_for("Pyro (T1/T2)", pyro))
    if fireworks:
        parts.append(build_table_for("Vuurwerk (F1–F4)", fireworks))
    if not parts:
        parts.append(_chapter_page("MATERIALEN – DEES", "Geen herkenbare T/F-classificatie."))

    # merge terug
    writer = PdfWriter()
    for p in parts:
        _append_pdf(writer, p)
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


def section_materials_avm(preview: Dict[str, Any]) -> bytes:
    items: List[Dict[str, Any]] = (preview.get("materials") or {}).get("avm") or []
    if not items:
        return _chapter_page("MATERIALEN – AVM", "Geen materialen geselecteerd.")

    buf = io.BytesIO()
    c = _new_canvas(buf)
    _draw_header_bar(c, "MATERIALEN – AVM")
    y = PAGE_H - 30 * mm

    colspec = [
        ("Naam", 90 * mm),
        ("Aantal", 20 * mm),
        ("Type", 20 * mm),
        ("Links (CE/Manual/MSDS)", 60 * mm),
    ]

    rows: List[List[str]] = []
    for it in items:
        files = it.get("files") or []
        def pick(label: str):
            # heuristisch op bestandsnaam
            for f in files:
                name = (f.get("name") or f.get("displayname") or "").lower()
                if label in name:
                    return f.get("url")
            return None

        ce = pick("conform") or pick("ce")
        man = pick("manual") or pick("handleid")
        msds = pick("msds") or pick("sds")
        linktxt = " ".join([
            f"[CE]" if ce else "",
            f"[Manual]" if man else "",
            f"[MSDS]" if msds else "",
        ]).strip() or "—"

        rows.append([
            it.get("displayname") or "—",
            str(it.get("quantity_total") or it.get("qty") or 0),
            it.get("type") or "—",
            linktxt
        ])

    y_end = _draw_table(c, MARGIN, y, colspec, rows)

    # visuele link-annotaties onder de tabel
    y_links = y_end - 10
    c.setFont(FONT, 10)
    for it in items:
        files = it.get("files") or []
        links_line = []
        for label in ["CE", "MANUAL", "MSDS"]:
            url = None
            for f in files:
                name = (f.get("name") or f.get("displayname") or "").lower()
                if (label == "CE" and ("conform" in name or "ce" in name)) or \
                   (label == "MANUAL" and ("manual" in name or "handleid" in name)) or \
                   (label == "MSDS" and ("msds" in name or "sds" in name)):
                    url = f.get("url"); break
            if url:
                links_line.append((label, url))
        if links_line:
            x = MARGIN
            c.setFillColor(PYRED_RED)
            for lab, url in links_line:
                txt = f"{lab}: {url}"
                c.drawString(x, y_links, txt)
                c.linkURL(url, (x, y_links - 2, x + c.stringWidth(txt, FONT, 10), y_links + 10))
                y_links -= 6 * mm

    c.showPage()
    c.save()
    return buf.getvalue()


def section_embed_pdf(title: str, url_or_data: str, logo: Optional[str] = None) -> bytes:
    """Maakt eerst een titelpagina, daarna plakt het de externe PDF erachter."""
    ch = _chapter_page(title, None, logo)
    ext = _pdf_bytes_from_url_or_data(url_or_data)
    writer = PdfWriter()
    _append_pdf(writer, ch)
    if ext:
        _append_pdf(writer, ext)
    else:
        # foutpagina
        _append_pdf(writer, _chapter_page(title, "⚠️ Kon het document niet laden.", logo))
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


def section_embed_uploads(title: str, uploads: List[Dict[str, Any]], logo: Optional[str] = None) -> bytes:
    """Embedt meerdere geüploade PDF's als hoofdstuk."""
    if not uploads:
        return _chapter_page(title, "Geen uploads.", logo)
    writer = PdfWriter()
    _append_pdf(writer, _chapter_page(title, None, logo))
    for up in uploads:
        data_url = up.get("data")
        if not data_url:
            continue
        try:
            b = base64.b64decode(data_url.split(",", 1)[1])
        except Exception:
            b = None
        if not b:
            continue
        # als PDF: direct append; als niet-PDF (bv. image), maak 1 pagina en plak afbeelding
        if (up.get("type") or "").lower().endswith("pdf"):
            try:
                _append_pdf(writer, b)
            except Exception:
                pass
        else:
            # afbeelding -> eigen pagina
            buf = io.BytesIO()
            c = _new_canvas(buf)
            _draw_header_bar(c, up.get("name") or "Bijlage")
            img = ImageReader(io.BytesIO(b))
            iw, ih = img.getSize()
            max_w = PAGE_W - 2 * MARGIN
            max_h = PAGE_H - 50 * mm
            scale = min(max_w / iw, max_h / ih)
            w = iw * scale
            h = ih * scale
            c.drawImage(img, (PAGE_W - w) / 2, (PAGE_H - h) / 2, w, h, preserveAspectRatio=True, mask="auto")
            c.showPage(); c.save()
            _append_pdf(writer, buf.getvalue())

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


# =========================================================
# Orchestratie – volledige PDF
# =========================================================

def build_full_pdf(preview: Dict[str, Any]) -> bytes:
    branding = preview.get("branding", {}) or {}
    logo = branding.get("logo") or preview.get("logo") or "https://sfx.rentals/projects/media/logo.png"

    # Inhoudstafel (tekstueel, zonder paginanummers)
    toc_items = [
        "Projectgegevens",
        "Emergency",
        "Verzekeringen",
        "Verantwoordelijke",
        "Materialen – Dees",
        "Materialen – AVM",
        "Inplantingsplan",
        "Risicoanalyse Pyro & Special Effects",
        "Windplan",
        "Droogteplan",
        "Vergunningen & Toelatingen",
    ]

    parts: List[bytes] = []
    parts.append(section_cover(preview))
    parts.append(section_toc(toc_items))
    parts.append(section_project_info(preview))

    docs = preview.get("documents", {}) or {}
    # 2 Emergency
    if docs.get("emergency"):
        parts.append(section_embed_pdf("EMERGENCY", docs["emergency"], logo))
    else:
        parts.append(_chapter_page("EMERGENCY", "Geen document opgegeven.", logo))

    # 3 Verzekeringen (één hoofdstuk met beide PDFs achter elkaar)
    ins = docs.get("insurance") or []
    if ins:
        writer = PdfWriter()
        _append_pdf(writer, _chapter_page("VERZEKERINGEN", None, logo))
        for u in ins:
            b = _pdf_bytes_from_url_or_data(u)
            if b:
                _append_pdf(writer, b)
        out = io.BytesIO(); writer.write(out)
        parts.append(out.getvalue())
    else:
        parts.append(_chapter_page("VERZEKERINGEN", "Geen documenten.", logo))

    # 4 Verantwoordelijke
    parts.append(section_responsible(preview))

    # 5.1 Dees
    parts.append(section_materials_dees(preview))

    # 5.2 AVM
    parts.append(section_materials_avm(preview))

    # 6 Inplantingsplan (uploads.siteplan -> één bestand)
    upl = (preview.get("uploads") or {})
    siteplan_list = upl.get("siteplan") or []
    parts.append(section_embed_uploads("INPLANTINGSPLAN", siteplan_list, logo))

    # 7 Risicoanalyses
    ra_writer = PdfWriter()
    _append_pdf(ra_writer, _chapter_page("RISICOANALYSE PYRO & SPECIAL EFFECTS", None, logo))
    any_ra = False
    for key in ["risk_pyro", "risk_general"]:
        if docs.get(key):
            b = _pdf_bytes_from_url_or_data(docs[key])
            if b:
                _append_pdf(ra_writer, b); any_ra = True
    if any_ra:
        tmp = io.BytesIO(); ra_writer.write(tmp); parts.append(tmp.getvalue())
    else:
        parts.append(_chapter_page("RISICOANALYSE PYRO & SPECIAL EFFECTS", "Geen documenten.", logo))

    # 8 Windplan
    if docs.get("windplan"):
        parts.append(section_embed_pdf("WINDPLAN", docs["windplan"], logo))
    else:
        parts.append(_chapter_page("WINDPLAN", "Niet van toepassing of niet toegevoegd.", logo))

    # 9 Droogteplan
    if docs.get("droughtplan"):
        parts.append(section_embed_pdf("DROOGTEPLAN", docs["droughtplan"], logo))
    else:
        parts.append(_chapter_page("DROOGTEPLAN", "Niet van toepassing of niet toegevoegd.", logo))

    # 10 Vergunningen & Toelatingen (uploads.permits[])
    permit_list = upl.get("permits") or []
    parts.append(section_embed_uploads("VERGUNNINGEN & TOELATINGEN", permit_list, logo))

    # -------- Merge alle delen in de juiste volgorde --------
    writer = PdfWriter()
    for p in parts:
        _append_pdf(writer, p)
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


# =========================================================
# API Endpoints
# =========================================================

@app.get("/", response_class=PlainTextResponse)
def root():
    return "Safetyfile PDF Service – OK"


@app.post("/generate")
def generate(preview: Dict[str, Any] = Body(...)):
    """
    Verwacht de 'preview' JSON van de wizard.
    Returned: application/pdf (binary).
    """
    try:
        pdf_bytes = build_full_pdf(preview)
        headers = {
            "Content-Disposition": 'attachment; filename="veiligheidsdossier.pdf"'
        }
        return Response(content=pdf_bytes, media_type="application/pdf", headers=headers)
    except Exception as e:
        return PlainTextResponse(str(e), status_code=500)


# Railway/Render start: python generate_pdf.py
if __name__ == "__main__":
    import uvicorn
    uvicorn.run("generate_pdf:app", host="0.0.0.0", port=8000, reload=False)
