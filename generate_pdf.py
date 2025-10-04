# -*- coding: utf-8 -*-
"""
Finalized generator for Veiligheidsdossier (PDF & DOCX).
Adds:
 - Projectgegevens populated from preview.avm
 - Responsible populated from preview.responsible
 - Crew bio PDFs inserted under responsible
 - External PDFs (emergency, insurance, wind, drought, permits, siteplan, risks) merged under sections
"""

import io, base64, datetime, re, requests
from flask import Flask, request, send_file, jsonify
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph, Table, TableStyle, Frame, Spacer
from pypdf import PdfReader, PdfWriter, Transformation

try:
    from docx import Document
    DOCX_AVAILABLE = True
except:
    DOCX_AVAILABLE = False

app = Flask(__name__)

# Layout
W,H=A4
MARGIN_L=22; MARGIN_R=22; MARGIN_B=34; BANNER_H=48; CONTENT_TOP_Y=H-(BANNER_H+26)
RED=colors.HexColor("#B00000"); BLACK=colors.black

styles=getSampleStyleSheet()
P=ParagraphStyle("Body",parent=styles["Normal"],fontName="Helvetica",fontSize=10,leading=13,textColor=BLACK)
P_H=ParagraphStyle("H",parent=styles["Heading2"],fontName="Helvetica-Bold",fontSize=12,leading=14,textColor=BLACK,spaceAfter=6)

def _safe(val,default=""): return default if val is None else str(val)
def _fmt_date(val):
    if not val: return ""
    try:
        iso=str(val).replace("Z","").replace("T"," ")
        return datetime.datetime.fromisoformat(iso).strftime("%d/%m/%Y")
    except: return str(val)
def _dataurl_to_bytes(u):
    if not u or not isinstance(u,str): return None
    if u.startswith("data:"):
        try: return base64.b64decode(u.split(",",1)[1])
        except: return None
    return None
def _fetch_pdf_bytes(item):
    if not item: return None
    if isinstance(item,bytes): return item
    if isinstance(item,str):
        b=_dataurl_to_bytes(item)
        if b: return b
        try:
            r=requests.get(item,timeout=15)
            if r.ok: return r.content
        except: return None
    return None

# PDF drawing
def draw_banner(c,title):
    c.setFillColor(RED); c.rect(0,H-BANNER_H,W,BANNER_H,fill=1,stroke=0)
    c.setFillColor(colors.white); c.setFont("Helvetica-Bold",14)
    c.drawString(MARGIN_L,H-BANNER_H+16,_safe(title))
def draw_marker(c,key):
    c.setFillColor(colors.white); c.setFont("Helvetica",1); c.drawString(2,2,f"[SEC::{key}]")
def start_section(c,key,title):
    c.showPage(); draw_banner(c,title); draw_marker(c,key)
def content_frame():
    return Frame(MARGIN_L,MARGIN_B,W-(MARGIN_L+MARGIN_R),CONTENT_TOP_Y-MARGIN_B,showBoundary=0)

# Content
def story_project(preview):
    avm=(preview or {}).get("avm") or {}
    cust=(avm.get("customer") or {}); contact=(cust.get("contact") or {}); loc=(avm.get("location") or {})
    rows=[["Project",_safe(avm.get("name"))],
          ["Opdrachtgever",_safe(cust.get("name"))],
          ["Adres",_safe(cust.get("address"))]]
    if any(contact.values()):
        rows+=[["Contactpersoon",_safe(contact.get("name"))],
               ["Tel",_safe(contact.get("phone"))],
               ["E-mail",_safe(contact.get("email"))]]
    rows+=[["Locatie",_safe(loc.get("name"))],
           ["Adres locatie",_safe(loc.get("address"))],
           ["Start",_fmt_date(avm.get("project_start_date"))],
           ["Einde",_fmt_date(avm.get("project_end_date"))]]
    table=Table(rows,colWidths=[120,W-(MARGIN_L+MARGIN_R)-120])
    table.setStyle(TableStyle([("FONT",(0,0),(-1,-1),"Helvetica",10),
        ("FONTNAME",(0,0),(0,-1),"Helvetica-Bold"),("TEXTCOLOR",(0,0),(-1,-1),BLACK),
        ("GRID",(0,0),(-1,-1),0.25,colors.HexColor("#BBBBBB"))]))
    return [table]

def story_responsible(preview):
    name=_safe((preview or {}).get("responsible") or "","")
    if not name: return [Paragraph("Geen verantwoordelijke geselecteerd.",P)]
    return [Paragraph(f"Projectverantwoordelijke: <b>{name}</b>",P)]

def _materials_rows(items):
    rows=[["Naam","Aantal","Type","CE","Manual","MSDS"]]
    for m in (items or []):
        lnks=(m.get("links") or {})
        rows.append([_safe(m.get("displayname")),
            _safe(m.get("quantity_total")),
            _safe(m.get("type")),
            _safe(lnks.get("ce") or ""),
            _safe(lnks.get("manual") or ""),
            _safe(lnks.get("msds") or "")])
    return rows

def story_materials(preview):
    mats=(preview or {}).get("materials") or {}
    avm_items=mats.get("avm") or []; dees_items=mats.get("dees") or []
    story=[Paragraph("5.1 Pyrotechnische materialen",P_H)]
    if not dees_items: story.append(Paragraph("Geen items geselecteerd.",P))
    else: story.append(Table(_materials_rows(dees_items),colWidths=[220,45,80,90,90,90]))
    story.append(Spacer(1,10))
    story.append(Paragraph("5.2 Speciale effecten",P_H))
    if not avm_items: story.append(Paragraph("Geen items geselecteerd.",P))
    else: story.append(Table(_materials_rows(avm_items),colWidths=[220,45,80,90,90,90]))
    return story

# Sections & TOC
def build_sections(preview):
    mats=(preview or {}).get("materials") or {}
    has_pyro=bool(mats.get("dees")); has_sfx=bool(mats.get("avm"))
    base=["project","emergency","insurance","responsible","materials","siteplan"]
    if has_pyro: base.append("risk_pyro")
    if has_sfx: base.append("risk_sfx")
    base+=["wind","drought","permits"]
    sections=[]
    for i,key in enumerate(base,1):
        title_map={"project":"Projectgegevens","emergency":"Emergency","insurance":"Verzekeringen",
                   "responsible":"Verantwoordelijke","materials":"Materialen","siteplan":"Inplantingsplan",
                   "risk_pyro":"Risicoanalyse Pyro","risk_sfx":"Risicoanalyse Speciale effecten",
                   "wind":"Windplan","drought":"Droogteplan","permits":"Vergunningen & Toelatingen"}
        sections.append({"key":key,"title":f"{i}. {title_map[key]}"})
    return sections

def draw_cover(c,preview):
    c.setFillColor(colors.white); c.rect(0,0,W,H,fill=1,stroke=0)
    draw_banner(c,"Veiligheidsdossier")
    avm=(preview or {}).get("avm") or {}
    pname=avm.get("name") or (preview or {}).get("project",{}).get("name") or ""
    c.setFillColor(BLACK); c.setFont("Helvetica-Bold",22)
    if pname: c.drawString(MARGIN_L,H-BANNER_H-36,pname)
    c.setFont("Helvetica",9); c.setFillColor(BLACK)
    c.drawString(MARGIN_L,MARGIN_B-12,"Gegenereerd: %s"%datetime.datetime.now().strftime("%d/%m/%Y %H:%M"))

def build_base_pdf(preview):
    sections=build_sections(preview)
    buf=io.BytesIO(); c=canvas.Canvas(buf,pagesize=A4)
    draw_cover(c,preview)
    c.showPage(); draw_banner(c,"Inhoudstafel")
    y=CONTENT_TOP_Y; c.setFont("Helvetica",11); c.setFillColor(BLACK)
    for s in sections:
        c.drawString(MARGIN_L,y,s["title"]); y-=16
        if y<MARGIN_B+30: c.showPage(); draw_banner(c,"Inhoudstafel"); y=CONTENT_TOP_Y
    for s in sections:
        start_section(c,s["key"],s["title"]); fr=content_frame()
        if s["key"]=="project": fr.addFromList(story_project(preview),c)
        elif s["key"]=="responsible": fr.addFromList(story_responsible(preview),c)
        elif s["key"]=="materials": fr.addFromList(story_materials(preview),c)
    c.save(); return buf.getvalue(),sections

# Merge externals
def find_section_pages(reader):
    mapping={}
    for idx,pg in enumerate(reader.pages):
        try: text=pg.extract_text() or ""
        except: text=""
        for key in re.findall(r"\[SEC::([^\]]+)\]",text):
            mapping[key]=idx
    return mapping

def scale_merge_first_page_under_banner(writer,page_index,ext_reader):
    if not ext_reader or len(ext_reader.pages)==0: return
    dst=writer.pages[page_index]; src=ext_reader.pages[0]
    avail_w=W-(MARGIN_L+MARGIN_R); avail_h=CONTENT_TOP_Y-MARGIN_B
    fw=float(src.mediabox.width); fh=float(src.mediabox.height)
    s=min(avail_w/fw,avail_h/fh,1.0)
    tx=MARGIN_L; ty=MARGIN_B
    op=Transformation().scale(s).translate(tx/s,ty/s)
    dst.merge_transformed_page(src,op)
    for i in range(1,len(ext_reader.pages)): writer.add_page(ext_reader.pages[i])

def collect_pdf_lists(preview):
    docs=(preview or {}).get("documents") or {}; uploads=(preview or {}).get("uploads") or {}
    def list_from(v): return v if isinstance(v,list) else ([v] if v else [])
    data={}
    data["emergency"]=[_fetch_pdf_bytes(u) for u in list_from(docs.get("emergency")) if u]
    data["insurance"]=[_fetch_pdf_bytes(u) for u in list_from(docs.get("insurance")) if u]
    data["wind"]=[_fetch_pdf_bytes(docs.get("windplan"))] if docs.get("windplan") else []
    data["drought"]=[_fetch_pdf_bytes(docs.get("droughtplan"))] if docs.get("droughtplan") else []
    bios=[]
    for k in ("crew_bio_full","crew_bio_mini"):
        u=docs.get(k); b=_fetch_pdf_bytes(u)
        if b: bios.append(b)
    data["responsible_bio"]=bios
    for rk,dk in (("risk_pyro","risk_pyro"),("risk_sfx","risk_general")):
        u=docs.get(dk); b=_fetch_pdf_bytes(u)
        data[rk]=[b] if b else []
    data["siteplan"]=[_dataurl_to_bytes(f.get("data")) for f in uploads.get("siteplan",[]) if f.get("data")]
    data["permits"]=[_dataurl_to_bytes(f.get("data")) for f in uploads.get("permits",[]) if f.get("data")]
    return data

def merge_externals(base_bytes,sections,preview):
    reader=PdfReader(io.BytesIO(base_bytes)); writer=PdfWriter()
    for p in reader.pages: writer.add_page(p)
    page_map=find_section_pages(reader); blobs=collect_pdf_lists(preview)
    plan={"emergency":blobs.get("emergency",[]),"insurance":blobs.get("insurance",[]),
          "responsible":blobs.get("responsible_bio",[]),"siteplan":blobs.get("siteplan",[]),
          "risk_pyro":blobs.get("risk_pyro",[]),"risk_sfx":blobs.get("risk_sfx",[]),
          "wind":blobs.get("wind",[]),"drought":blobs.get("drought",[]),
          "permits":blobs.get("permits",[])}
    sortable=[(k,page_map[k]) for k in plan.keys() if k in page_map and plan[k]]
    sortable.sort(key=lambda x:x[1],reverse=True)
    for key,page_index in sortable:
        items=plan[key]; first=True
        for b in items:
            try: ext=PdfReader(io.BytesIO(b))
            except: continue
            if first: scale_merge_first_page_under_banner(writer,page_index,ext); first=False
            else:
                for p in ext.pages: writer.add_page(p)
    out=io.BytesIO(); writer.write(out); return out.getvalue()

# DOCX simplified

from docxtpl import DocxTemplate

def build_docx(preview):
    import requests, io
    lang = (preview.get("language") or "nl").lower()
    url = f"https://sfx.rentals/safetyfile/templates/dossier_{lang}.docx"
    try:
        r = requests.get(url, timeout=20)
        r.raise_for_status()
    except Exception as e:
        raise Exception(f"Kon template niet ophalen: {e}")
    tpl = DocxTemplate(io.BytesIO(r.content))
    ctx = preview or {}
    try:
        tpl.render(ctx)
    except Exception as e:
        raise Exception(f"Fout bij vullen template: {e}")
    out = io.BytesIO()
    tpl.save(out)
    return out.getvalue()


@app.route("/generate",methods=["POST"])
def generate():
    payload=request.get_json(force=True,silent=False) or {}
    preview=payload.get("preview") or payload
    fmt=(payload.get("format") or "pdf").lower()
    if fmt=="docx":
        try: return send_file(io.BytesIO(build_docx(preview)),
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,download_name="dossier.docx")
        except Exception as e: return jsonify({"error":"DOCX generation failed","detail":str(e)}),500
    try:
        base_bytes,sections=build_base_pdf(preview)
        final_bytes=merge_externals(base_bytes,sections,preview)
        return send_file(io.BytesIO(final_bytes),mimetype="application/pdf",
            as_attachment=True,download_name="dossier.pdf")
    except Exception as e:
        return jsonify({"error":"PDF generation failed","detail":str(e)}),500

if __name__=="__main__": app.run(host="0.0.0.0",port=8000)
