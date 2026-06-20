import os
import re
import uuid
import tempfile
from urllib.parse import urlparse

import requests
from bs4 import BeautifulSoup, NavigableString, Tag
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
    return "Moajam Pure Python HTML DOCX Service Running"


@app.route("/generate-docx", methods=["POST"])
def generate_docx():
    data = request.get_json(force=True, silent=True) or {}

    job_number = clean_filename(data.get("job_number") or "MOAJAM-JOB")
    final_html = data.get("final_html") or data.get("translated_html") or ""
    translated_text = data.get("translated_text") or ""
    letterhead_url = data.get("letterhead_image_link") or ""

    
    if not final_html and not translated_text:
        return jsonify({"status": "error", "message": "Missing final_html/translated_text"}), 400

    try:
        letterhead_path = download_file(letterhead_url)

        doc = Document(os.path.join(os.path.dirname(__file__), "letterhead_template.docx"))
        set_document_defaults(doc)

        html = final_html if final_html else text_to_html(translated_text)
        add_html_to_docx(doc, html)

        filename = f"{job_number}-Final-Translation-{uuid.uuid4().hex[:8]}.docx"
        output_path = os.path.join(OUTPUT_DIR, filename)
        doc.save(output_path)

        base_url = request.host_url.rstrip("/")
        return jsonify({
            "status": "success",
            "download_url": f"{base_url}/download/{filename}",
            "filename": filename
        })

    except Exception as exc:
        return jsonify({"status": "error", "message": str(exc)}), 500


@app.route("/download/<path:filename>")
def download(filename):
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)


def setup_doc(doc, letterhead_path):
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(4.5)
    section.bottom_margin = Cm(3.5)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)
    section.header_distance = Cm(0)
    section.footer_distance = Cm(0)


def add_letterhead_to_header(section, image_path):
    header = section.header
    paragraph = header.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.space_before = Pt(0)

    run = paragraph.add_run()
    inline_shape = run.add_picture(image_path, width=Cm(21), height=Cm(29.7))

    inline = inline_shape._inline
    inline.tag = qn("wp:anchor")

    inline.set("behindDoc", "1")
    inline.set("locked", "0")
    inline.set("layoutInCell", "1")
    inline.set("allowOverlap", "1")
    inline.set("simplePos", "0")
    inline.set("relativeHeight", "0")
    inline.set("distT", "0")
    inline.set("distB", "0")
    inline.set("distL", "0")
    inline.set("distR", "0")

    position_h = OxmlElement("wp:positionH")
    position_h.set("relativeFrom", "page")
    pos_h = OxmlElement("wp:posOffset")
    pos_h.text = "0"
    position_h.append(pos_h)

    position_v = OxmlElement("wp:positionV")
    position_v.set("relativeFrom", "page")
    pos_v = OxmlElement("wp:posOffset")
    pos_v.text = "0"
    position_v.append(pos_v)

    wrap_none = OxmlElement("wp:wrapNone")

    inline.insert(0, position_h)
    inline.insert(1, position_v)
    inline.insert(2, wrap_none)
    
def set_document_defaults(doc):
    normal = doc.styles["Normal"]
    normal.font.name = "Sakkal Majalla"
    normal.font.size = Pt(14)
    normal._element.rPr.rFonts.set(qn("w:cs"), "Sakkal Majalla")
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Sakkal Majalla")


def add_html_to_docx(doc, html):
    soup = BeautifulSoup(str(html), "html5lib")
    body = soup.body if soup.body else soup

    for child in body.children:
        add_node(doc, child)


def add_node(doc, node):
    if isinstance(node, NavigableString):
        text = str(node).strip()
        if text:
            add_paragraph(doc, text)
        return

    if not isinstance(node, Tag):
        return

    name = node.name.lower()

    if name in ["html", "body", "section", "article", "main", "div"]:
        # If div is a simple text container, keep it as one paragraph.
        direct_complex = node.find(["table", "h1", "h2", "h3", "ul", "ol"], recursive=False)
        if direct_complex:
            for child in node.children:
                add_node(doc, child)
        else:
            text = clean_text(node.get_text("\n", strip=True))
            if text:
                add_paragraph(
                    doc,
                    text,
                    bold=element_is_bold(node),
                    align=get_align(node),
                    color=get_color(node),
                    size=get_font_size(node, 14),
                )
        return

    if name in ["h1", "h2", "h3", "h4"]:
        size = {"h1": 18, "h2": 16, "h3": 15, "h4": 14}.get(name, 14)
        align = "center" if name == "h1" else "right"
        add_paragraph(doc, node.get_text(" ", strip=True), bold=True, size=size, align=get_align(node, align), color=get_color(node))
        return

    if name in ["p"]:
        text = clean_text(node.get_text("\n", strip=True))
        if text:
            add_paragraph(doc, text, bold=element_is_bold(node), align=get_align(node), color=get_color(node), size=get_font_size(node, 14))
        return

    if name == "br":
        doc.add_paragraph("")
        return

    if name in ["ul", "ol"]:
        for li in node.find_all("li", recursive=False):
            add_paragraph(doc, "• " + clean_text(li.get_text(" ", strip=True)), align="right")
        return

    if name == "table":
        add_table(doc, node)
        return

    if name in ["style", "script", "meta", "head", "title"]:
        return

    text = clean_text(node.get_text("\n", strip=True))
    if text:
        add_paragraph(doc, text)


def add_paragraph(doc, text, bold=False, size=14, align="right", color=None):
    text = clean_text(text)
    if not text:
        return

    lines = [line.strip() for line in str(text).split("\n") if line.strip()]
    for line in lines:
        p = doc.add_paragraph()
        set_paragraph_rtl(p)
        p.alignment = {
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }.get(align, WD_ALIGN_PARAGRAPH.RIGHT)

        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = 1.1

        run = p.add_run(line)
        run.bold = bool(bold)
        run.font.name = "Sakkal Majalla"
        run.font.size = Pt(size)

        rgb = parse_color(color)
        if rgb:
            run.font.color.rgb = RGBColor(*rgb)

        run._element.rPr.rFonts.set(qn("w:cs"), "Sakkal Majalla")
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Sakkal Majalla")
        


def add_table(doc, table_el):
    rows = []
    max_cols = 0

    for tr in table_el.find_all("tr"):
        cells = tr.find_all(["th", "td"], recursive=False)
        if not cells:
            continue

        row = []
        for cell in cells:
            text = clean_text(cell.get_text("\n", strip=True))
            row.append({
                "text": text,
                "header": cell.name.lower() == "th" or bool(tr.find("th")),
                "bg": get_background(cell) or get_background(tr),
                "color": get_color(cell),
                "bold": element_is_bold(cell) or cell.name.lower() == "th",
            })

        rows.append(row)
        max_cols = max(max_cols, len(row))

    if not rows or max_cols < 1:
        return

    table = doc.add_table(rows=0, cols=max_cols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.RIGHT
    set_table_rtl(table)

    for row in rows:
        cells = table.add_row().cells
        for i in range(max_cols):
            item = row[i] if i < len(row) else {"text": "", "header": False, "bg": None, "color": None, "bold": False}
            bg = item["bg"] or ("#d9eaf7" if item["header"] else None)
            set_cell(cells[i], item["text"], bold=item["bold"] or item["header"], bg=bg, color=item["color"])

    doc.add_paragraph("")


def set_cell(cell, text, bold=False, bg=None, color=None):
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    cell.text = ""

    if bg:
        fill = normalize_hex(bg)
        if fill:
            tc_pr = cell._tc.get_or_add_tcPr()
            shd = tc_pr.find(qn("w:shd"))
            if shd is None:
                shd = OxmlElement("w:shd")
                tc_pr.append(shd)
            shd.set(qn("w:fill"), fill)

    p = cell.paragraphs[0]
    set_paragraph_rtl(p)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)

    run = p.add_run(str(text or ""))
    run.bold = bool(bold)
    run.font.name = "Sakkal Majalla"
    run.font.size = Pt(11)

    rgb = parse_color(color)
    if rgb:
        run.font.color.rgb = RGBColor(*rgb)

    run._element.rPr.rFonts.set(qn("w:cs"), "Sakkal Majalla")
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Sakkal Majalla")


def set_paragraph_rtl(paragraph):
    pPr = paragraph._p.get_or_add_pPr()

    bidi = pPr.find(qn("w:bidi"))
    if bidi is None:
        bidi = OxmlElement("w:bidi")
        pPr.append(bidi)
    bidi.set(qn("w:val"), "1")

    jc = pPr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        pPr.append(jc)

    jc.set(qn("w:val"), "both")


def set_table_rtl(table):
    tblPr = table._tbl.tblPr
    bidi_visual = tblPr.find(qn("w:bidiVisual"))
    if bidi_visual is None:
        bidi_visual = OxmlElement("w:bidiVisual")
        tblPr.append(bidi_visual)


def get_style(el):
    return (el.get("style") or "") if isinstance(el, Tag) else ""


def get_align(el, default="right"):
    style = get_style(el).lower().replace(" ", "")
    if "text-align:center" in style:
        return "center"
    if "text-align:left" in style:
        return "left"
    if "text-align:justify" in style:
        return "justify"

    align = (el.get("align") or "").lower() if isinstance(el, Tag) else ""
    if align in ["right", "left", "center", "justify"]:
        return align

    return default


def element_is_bold(el):
    style = get_style(el).lower().replace(" ", "")
    if "font-weight:bold" in style or "font-weight:700" in style:
        return True
    if isinstance(el, Tag) and el.find(["b", "strong"]):
        return True
    return False


def get_color(el):
    style = get_style(el)
    m = re.search(r"(?<!-)color\s*:\s*(#[0-9a-fA-F]{3,6})", style)
    return m.group(1) if m else None


def get_background(el):
    style = get_style(el)
    m = re.search(r"background(?:-color)?\s*:\s*(#[0-9a-fA-F]{3,6})", style)
    return m.group(1) if m else None


def get_font_size(el, default):
    style = get_style(el)
    m = re.search(r"font-size\s*:\s*([0-9.]+)pt", style)
    if not m:
        return default
    try:
        value = float(m.group(1))
        return max(8, min(24, value))
    except Exception:
        return default


def normalize_hex(value):
    if not value:
        return None
    value = str(value).strip()
    if not value.startswith("#"):
        return None
    value = value[1:]
    if len(value) == 3:
        value = "".join(ch * 2 for ch in value)
    if len(value) != 6:
        return None
    return value.upper()


def parse_color(value):
    hx = normalize_hex(value)
    if not hx:
        return None
    try:
        return (int(hx[0:2], 16), int(hx[2:4], 16), int(hx[4:6], 16))
    except Exception:
        return None


def clean_text(text):
    text = str(text or "")
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def text_to_html(text):
    lines = []
    for line in str(text or "").splitlines():
        line = clean_text(line)
        if line:
            lines.append(f"<p>{escape_html(line)}</p>")
    return '<div dir="rtl" style="direction:rtl;text-align:right">' + "\n".join(lines) + "</div>"


def escape_html(value):
    return str(value).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def download_file(url):
    parsed = urlparse(url)
    ext = os.path.splitext(parsed.path)[1].lower() or ".png"

    if ext not in [".png", ".jpg", ".jpeg"]:
        ext = ".png"

    response = requests.get(url, timeout=60)
    response.raise_for_status()

    fd, path = tempfile.mkstemp(suffix=ext)
    with os.fdopen(fd, "wb") as f:
        f.write(response.content)

    return path


def clean_filename(value):
    value = str(value or "MOAJAM").strip()
    safe = "".join(ch for ch in value if ch.isalnum() or ch in ("-", "_"))
    return safe or "MOAJAM"


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
