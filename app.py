import os, re, json, uuid, tempfile
from urllib.parse import urlparse
import requests
from bs4 import BeautifulSoup
from flask import Flask, request, jsonify, send_from_directory
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

app = Flask(__name__)
OUTPUT_DIR = os.path.join(os.getcwd(), "generated")
os.makedirs(OUTPUT_DIR, exist_ok=True)

@app.route("/")
def home():
    return "Moajam HTML DOCX Render Service Running"

@app.route("/generate-docx", methods=["POST"])
def generate_docx():
    data = request.get_json(force=True, silent=True) or {}
    job_number = clean_filename(data.get("job_number") or "MOAJAM-JOB")
    final_html = data.get("final_html") or data.get("translated_html") or ""
    translated_text = data.get("translated_text") or ""
    letterhead_url = data.get("letterhead_image_link") or ""

    if not letterhead_url:
        return jsonify({"status":"error","message":"Missing letterhead_image_link"}), 400
    if not final_html and not translated_text:
        return jsonify({"status":"error","message":"Missing final_html/translated_text"}), 400

    try:
        letterhead_path = download_file(letterhead_url)
        doc = Document()
        setup_page(doc, letterhead_path)
        set_document_defaults(doc)
        if final_html:
            add_html_to_docx(doc, final_html)
        else:
            add_paragraph(doc, translated_text)
        filename = f"{job_number}-Final-Translation-{uuid.uuid4().hex[:8]}.docx"
        output_path = os.path.join(OUTPUT_DIR, filename)
        doc.save(output_path)
        base_url = request.host_url.rstrip("/")
        return jsonify({"status":"success","download_url":f"{base_url}/download/{filename}","filename":filename})
    except Exception as exc:
        return jsonify({"status":"error","message":str(exc)}), 500

@app.route("/download/<path:filename>")
def download(filename):
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)

def setup_page(doc, letterhead_path):
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(4.4)
    section.bottom_margin = Cm(2.6)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.header_distance = Cm(0)
    section.footer_distance = Cm(0)
    add_letterhead_background(section, letterhead_path)

def add_letterhead_background(section, image_path):
    header = section.header
    paragraph = header.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.space_before = Pt(0)
    run = paragraph.add_run()
    inline_shape = run.add_picture(image_path, width=Cm(21), height=Cm(29.7))
    inline = inline_shape._inline
    anchor = OxmlElement("wp:anchor")
    for k, v in {"distT":"0","distB":"0","distL":"0","distR":"0","simplePos":"0","relativeHeight":"0","behindDoc":"1","locked":"0","layoutInCell":"1","allowOverlap":"1"}.items():
        anchor.set(k, v)
    simple_pos = OxmlElement("wp:simplePos"); simple_pos.set("x","0"); simple_pos.set("y","0"); anchor.append(simple_pos)
    position_h = OxmlElement("wp:positionH"); position_h.set("relativeFrom","page"); pos_h = OxmlElement("wp:posOffset"); pos_h.text = "0"; position_h.append(pos_h); anchor.append(position_h)
    position_v = OxmlElement("wp:positionV"); position_v.set("relativeFrom","page"); pos_v = OxmlElement("wp:posOffset"); pos_v.text = "0"; position_v.append(pos_v); anchor.append(position_v)
    for tag in ["wp:extent","wp:effectExtent","wp:docPr","wp:cNvGraphicFramePr","a:graphic"]:
        el = inline.find(qn(tag))
        if el is not None:
            if tag == "wp:docPr": anchor.append(OxmlElement("wp:wrapNone"))
            anchor.append(el)
    if anchor.find(qn("wp:wrapNone")) is None:
        anchor.append(OxmlElement("wp:wrapNone"))
    inline.getparent().replace(inline, anchor)

def set_document_defaults(doc):
    normal = doc.styles["Normal"]
    normal.font.name = "Sakkal Majalla"
    normal.font.size = Pt(14)
    normal._element.rPr.rFonts.set(qn("w:cs"), "Sakkal Majalla")
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Sakkal Majalla")

def add_html_to_docx(doc, html):
    soup = BeautifulSoup(str(html), "html.parser")
    body = soup.body if soup.body else soup
    for element in body.children:
        add_html_element(doc, element)

def add_html_element(doc, element):
    if getattr(element, "name", None) is None:
        text = str(element).strip()
        if text: add_paragraph(doc, text)
        return
    name = element.name.lower()
    if name in ["style","script","meta","head"]: return
    if name in ["h1","h2","h3"]:
        size = 18 if name == "h1" else 16 if name == "h2" else 15
        add_paragraph(doc, element.get_text("\n", strip=True), bold=True, size=size, align=get_align(element, "center" if name == "h1" else "right"))
        return
    if name in ["p","div","section","article"]:
        if element.find(["table","ul","ol","h1","h2","h3"], recursive=False):
            for child in element.children: add_html_element(doc, child)
        else:
            text = element.get_text("\n", strip=True)
            if text: add_paragraph(doc, text, align=get_align(element,"right"), bold=has_bold(element), color=get_color(element))
        return
    if name == "br": doc.add_paragraph(""); return
    if name in ["ul","ol"]:
        for li in element.find_all("li", recursive=False): add_paragraph(doc, "• " + li.get_text(" ", strip=True))
        return
    if name == "table": add_html_table(doc, element); return
    text = element.get_text("\n", strip=True)
    if text: add_paragraph(doc, text)

def add_html_table(doc, table_el):
    rows = table_el.find_all("tr")
    parsed_rows = []
    max_cols = 0
    for tr in rows:
        cells = tr.find_all(["th","td"], recursive=False)
        row = []
        for cell in cells:
            row.append((cell.get_text("\n", strip=True), cell.name.lower()=="th", get_bg(cell), get_color(cell)))
        if row:
            parsed_rows.append(row); max_cols = max(max_cols, len(row))
    if not parsed_rows or max_cols == 0: return
    table = doc.add_table(rows=0, cols=max_cols)
    table.alignment = WD_TABLE_ALIGNMENT.RIGHT
    table.style = "Table Grid"
    set_table_rtl(table)
    for row in parsed_rows:
        cells = table.add_row().cells
        for i in range(max_cols):
            text, is_header, bg, color = row[i] if i < len(row) else ("", False, None, None)
            set_cell(cells[i], text, bold=is_header, bg=bg, color=color)
    doc.add_paragraph("")

def add_paragraph(doc, text, bold=False, size=14, align="right", color=None):
    text = str(text or "").strip()
    if not text: return
    for part in [p.strip() for p in text.split("\n") if p.strip()]:
        p = doc.add_paragraph(); set_paragraph_rtl(p)
        p.alignment = {"center":WD_ALIGN_PARAGRAPH.CENTER,"left":WD_ALIGN_PARAGRAPH.LEFT,"justify":WD_ALIGN_PARAGRAPH.JUSTIFY}.get(align, WD_ALIGN_PARAGRAPH.RIGHT)
        p.paragraph_format.space_after = Pt(4); p.paragraph_format.line_spacing = 1.1
        run = p.add_run(part); run.bold = bool(bold); run.font.name = "Sakkal Majalla"; run.font.size = Pt(size)
        rgb = parse_hex_color(color) if color else None
        if rgb: run.font.color.rgb = RGBColor(*rgb)
        run._element.rPr.rFonts.set(qn("w:cs"), "Sakkal Majalla"); run._element.rPr.rFonts.set(qn("w:eastAsia"), "Sakkal Majalla")

def set_cell(cell, text, bold=False, bg=None, color=None):
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER; cell.text = ""
    if bg:
        rgb_hex = normalize_hex(bg)
        if rgb_hex:
            tc_pr = cell._tc.get_or_add_tcPr(); shd = tc_pr.find(qn("w:shd"))
            if shd is None: shd = OxmlElement("w:shd"); tc_pr.append(shd)
            shd.set(qn("w:fill"), rgb_hex)
    p = cell.paragraphs[0]; set_paragraph_rtl(p); p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run(str(text or "")); run.bold = bool(bold); run.font.name = "Sakkal Majalla"; run.font.size = Pt(12)
    rgb = parse_hex_color(color) if color else None
    if rgb: run.font.color.rgb = RGBColor(*rgb)
    run._element.rPr.rFonts.set(qn("w:cs"), "Sakkal Majalla"); run._element.rPr.rFonts.set(qn("w:eastAsia"), "Sakkal Majalla")

def set_paragraph_rtl(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    bidi = pPr.find(qn("w:bidi"))
    if bidi is None: bidi = OxmlElement("w:bidi"); pPr.append(bidi)
    bidi.set(qn("w:val"), "1")
    jc = pPr.find(qn("w:jc"))
    if jc is None: jc = OxmlElement("w:jc"); pPr.append(jc)
    jc.set(qn("w:val"), "right")

def set_table_rtl(table):
    table.alignment = WD_TABLE_ALIGNMENT.RIGHT
    tblPr = table._tbl.tblPr
    bidi_visual = tblPr.find(qn("w:bidiVisual"))
    if bidi_visual is None: bidi_visual = OxmlElement("w:bidiVisual"); tblPr.append(bidi_visual)

def get_style(element): return element.get("style", "") or ""
def get_align(element, default="right"):
    style = get_style(element).lower()
    if "text-align:center" in style or "text-align: center" in style: return "center"
    if "text-align:left" in style or "text-align: left" in style: return "left"
    if "text-align:justify" in style or "text-align: justify" in style: return "justify"
    align = (element.get("align") or "").lower()
    return align if align in ["right","left","center","justify"] else default

def has_bold(element):
    style = get_style(element).lower()
    return bool(element.find(["b","strong"])) or "font-weight:bold" in style or "font-weight: bold" in style

def get_color(element):
    m = re.search(r"(?<!-)color\s*:\s*(#[0-9a-fA-F]{3,6})", get_style(element))
    return m.group(1) if m else None

def get_bg(element):
    m = re.search(r"background(?:-color)?\s*:\s*(#[0-9a-fA-F]{3,6})", get_style(element))
    return m.group(1) if m else None

def normalize_hex(value):
    if not value: return None
    value = value.strip()
    if not value.startswith("#"): return None
    value = value[1:]
    if len(value) == 3: value = "".join([c*2 for c in value])
    if len(value) != 6: return None
    return value.upper()

def parse_hex_color(value):
    hx = normalize_hex(value)
    if not hx: return None
    return (int(hx[0:2],16), int(hx[2:4],16), int(hx[4:6],16))

def download_file(url):
    parsed = urlparse(url); ext = os.path.splitext(parsed.path)[1].lower() or ".png"
    if ext not in [".png",".jpg",".jpeg"]: ext = ".png"
    response = requests.get(url, timeout=60); response.raise_for_status()
    fd, path = tempfile.mkstemp(suffix=ext)
    with os.fdopen(fd, "wb") as f: f.write(response.content)
    return path

def clean_filename(value):
    value = str(value or "MOAJAM").strip()
    safe = "".join(ch for ch in value if ch.isalnum() or ch in ("-","_"))
    return safe or "MOAJAM"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
