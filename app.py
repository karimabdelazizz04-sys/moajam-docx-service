import os
import uuid
import tempfile
from urllib.parse import urlparse

import requests
from flask import Flask, request, jsonify, send_from_directory
from docx import Document
from docx.shared import Cm, Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT


app = Flask(__name__)

OUTPUT_DIR = os.path.join(os.getcwd(), "generated")
os.makedirs(OUTPUT_DIR, exist_ok=True)


@app.route("/")
def home():
    return "Moajam DOCX Service Running"


@app.route("/generate-docx", methods=["POST"])
def generate_docx():
    data = request.get_json(force=True, silent=True) or {}

    job_number = clean_filename(data.get("job_number") or "MOAJAM-JOB")
    translated_text = data.get("translated_text") or ""
    layout_plan = data.get("layout_plan_json") or {}
    letterhead_url = data.get("letterhead_image_link") or ""

    if not letterhead_url:
        return jsonify({"status": "error", "message": "Missing letterhead_image_link"}), 400

    try:
        letterhead_path = download_file(letterhead_url)
        doc = build_docx(job_number, translated_text, layout_plan, letterhead_path)

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


def build_docx(job_number, translated_text, layout_plan, letterhead_path):
    doc = Document()

    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)

    section.top_margin = Cm(5.2)
    section.bottom_margin = Cm(2.8)
    section.left_margin = Cm(1.4)
    section.right_margin = Cm(1.4)
    section.header_distance = Cm(0)
    section.footer_distance = Cm(0)

    add_letterhead_to_header(section, letterhead_path)

    set_document_defaults(doc)

    blocks = []
    if isinstance(layout_plan, dict):
        blocks = layout_plan.get("blocks") or []

    if blocks:
        for block in blocks:
            add_block(doc, block)
    else:
        add_paragraph(doc, translated_text)

    return doc


def add_letterhead_to_header(section, image_path):
    header = section.header
    paragraph = header.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    run.add_picture(image_path, width=Cm(21), height=Cm(29.7))

    # Put header image behind text as much as python-docx allows.
    # The safe margins above keep body content away from header/footer/frame.
    paragraph.paragraph_format.space_after = Pt(0)


def set_document_defaults(doc):
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Sakkal Majalla"
    normal.font.size = Pt(14)
    normal._element.rPr.rFonts.set(qn("w:cs"), "Sakkal Majalla")
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Sakkal Majalla")


def add_block(doc, block):
    if not isinstance(block, dict):
        add_paragraph(doc, str(block))
        return

    block_type = (block.get("type") or "paragraph").strip().lower()
    text = block.get("text") or block.get("content") or ""

    if block_type == "title":
        add_paragraph(doc, text, bold=True, size=18, align="center")
    elif block_type == "subtitle":
        add_paragraph(doc, text, bold=True, size=15, align="center")
    elif block_type == "section_heading":
        add_paragraph(doc, text, bold=True, size=15, align="right")
    elif block_type == "field_table":
        add_field_table(doc, block)
    elif block_type == "data_table":
        add_data_table(doc, block)
    elif block_type == "signature_block":
        add_paragraph(doc, text, size=14, align="right")
    elif block_type == "page_break":
        doc.add_page_break()
    elif block_type == "spacer":
        doc.add_paragraph("")
    else:
        add_paragraph(doc, text, size=14, align="right")


def add_paragraph(doc, text, bold=False, size=14, align="right"):
    text = str(text or "").strip()
    if not text:
        return

    for part in text.split("\n"):
        part = part.strip()
        if not part:
            continue

        p = doc.add_paragraph()
        set_paragraph_rtl(p)

        if align == "center":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == "left":
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.15

        run = p.add_run(part)
        run.bold = bold
        run.font.name = "Sakkal Majalla"
        run.font.size = Pt(size)
        run._element.rPr.rFonts.set(qn("w:cs"), "Sakkal Majalla")
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Sakkal Majalla")


def add_field_table(doc, block):
    rows = normalize_rows(block.get("rows") or block.get("fields") or [])
    if not rows:
        return

    table = doc.add_table(rows=0, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.RIGHT
    table.style = "Table Grid"
    set_table_rtl(table)

    for row in rows:
        label, value = row
        cells = table.add_row().cells
        cells[0].width = Cm(5)
        cells[1].width = Cm(12.5)

        set_cell(cells[0], label, bold=True)
        set_cell(cells[1], value)

    doc.add_paragraph("")


def add_data_table(doc, block):
    headers = block.get("headers") or []
    rows = block.get("rows") or []

    if not rows and not headers:
        return

    col_count = max(len(headers), max([len(r) if isinstance(r, list) else 1 for r in rows], default=1))
    table = doc.add_table(rows=0, cols=col_count)
    table.alignment = WD_TABLE_ALIGNMENT.RIGHT
    table.style = "Table Grid"
    set_table_rtl(table)

    if headers:
        cells = table.add_row().cells
        for i in range(col_count):
            set_cell(cells[i], headers[i] if i < len(headers) else "", bold=True)

    for r in rows:
        if not isinstance(r, list):
            r = [r]
        cells = table.add_row().cells
        for i in range(col_count):
            set_cell(cells[i], r[i] if i < len(r) else "")

    doc.add_paragraph("")


def set_cell(cell, text, bold=False):
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    cell.text = ""

    p = cell.paragraphs[0]
    set_paragraph_rtl(p)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    run = p.add_run(str(text or ""))
    run.bold = bold
    run.font.name = "Sakkal Majalla"
    run.font.size = Pt(12)
    run._element.rPr.rFonts.set(qn("w:cs"), "Sakkal Majalla")
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Sakkal Majalla")


def set_paragraph_rtl(paragraph):
    pPr = paragraph._p.get_or_add_pPr()

    bidi = pPr.find(qn("w:bidi"))
    if bidi is None:
        bidi = OxmlElement("w:bidi")
        pPr.append(bidi)

    jc = pPr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        pPr.append(jc)
    jc.set(qn("w:val"), "right")


def set_table_rtl(table):
    tblPr = table._tbl.tblPr
    bidi_visual = tblPr.find(qn("w:bidiVisual"))
    if bidi_visual is None:
        bidi_visual = OxmlElement("w:bidiVisual")
        tblPr.append(bidi_visual)


def normalize_rows(rows):
    output = []

    if isinstance(rows, dict):
        for k, v in rows.items():
            output.append([k, v])
        return output

    for row in rows:
        if isinstance(row, dict):
            label = row.get("label") or row.get("key") or row.get("0") or ""
            value = row.get("value") or row.get("text") or row.get("1") or ""
            output.append([label, value])
        elif isinstance(row, list):
            if len(row) >= 2:
                output.append([row[0], row[1]])
            elif len(row) == 1:
                output.append(["", row[0]])
        else:
            output.append(["", row])

    return output


def download_file(url):
    parsed = urlparse(url)
    ext = os.path.splitext(parsed.path)[1].lower() or ".png"

    if ext not in [".png", ".jpg", ".jpeg"]:
        ext = ".png"

    response = requests.get(url, timeout=30)
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
