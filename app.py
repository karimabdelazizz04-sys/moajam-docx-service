import os
import io
import re
import json
import uuid
import base64
import tempfile
from urllib.parse import urlparse

import fitz
import requests
from flask import Flask, request, jsonify, send_from_directory
from openai import OpenAI
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


app = Flask(__name__)
client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))

OUTPUT_DIR = os.path.join(os.getcwd(), "generated")
os.makedirs(OUTPUT_DIR, exist_ok=True)


@app.route("/")
def home():
    return "Moajam DOCX Vision Service Running"


@app.route("/generate-docx", methods=["POST"])
def generate_docx():
    data = request.get_json(force=True, silent=True) or {}

    job_number = clean_filename(data.get("job_number") or "MOAJAM-JOB")
    source_file_link = data.get("source_file_link") or ""
    source_text = data.get("source_text") or ""
    letterhead_url = data.get("letterhead_image_link") or ""

    if not source_file_link:
        return jsonify({"status": "error", "message": "Missing source_file_link"}), 400

    if not letterhead_url:
        return jsonify({"status": "error", "message": "Missing letterhead_image_link"}), 400

    try:
        source_path = download_file(source_file_link)
        letterhead_path = download_file(letterhead_url)

        page_images = source_to_images(source_path)
        if not page_images:
            return jsonify({"status": "error", "message": "Could not convert source file to images"}), 500

        ai_data = analyze_with_vision(data, page_images)

        translated_text = ai_data.get("translated_text", "")
        layout_plan = ai_data.get("layout_plan_json", {})

        doc = build_docx(job_number, translated_text, layout_plan, letterhead_path)

        filename = f"{job_number}-Final-Translation-{uuid.uuid4().hex[:8]}.docx"
        output_path = os.path.join(OUTPUT_DIR, filename)
        doc.save(output_path)

        base_url = request.host_url.rstrip("/")
        return jsonify({
            "status": "success",
            "download_url": f"{base_url}/download/{filename}",
            "filename": filename,
            "matched_collection": ai_data.get("matched_collection", ""),
            "matched_document_type": ai_data.get("matched_document_type", ""),
            "matched_sample": ai_data.get("matched_sample", ""),
            "layout_type": ai_data.get("layout_type", "structured_legal_translation"),
            "translated_text": translated_text,
            "layout_plan_json": layout_plan
        })

    except Exception as exc:
        return jsonify({"status": "error", "message": str(exc)}), 500


@app.route("/download/<path:filename>")
def download(filename):
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)


def analyze_with_vision(job_data, page_images):
    content = [{
        "type": "input_text",
        "text": build_vision_prompt(job_data)
    }]

    for img_path in page_images[:1]:
        content.append({
            "type": "input_image",
            "image_url": image_to_data_url(img_path),
            "detail": "low"
        })

    response = client.responses.create(
        model=os.environ.get("OPENAI_MODEL", "gpt-4.1-mini"),
        input=[{
            "role": "user",
            "content": content
        }]
    )

    text = getattr(response, "output_text", "") or ""

    if not text:
        text = str(response)

    return extract_json(text)


def build_vision_prompt(job):
    source_text = str(job.get("source_text") or "").strip()

    return f"""
You are UAE MOJ Legal Translation Assistant.

You can see the FIRST PAGE of the original document as an image.
Use the image ONLY to understand visual layout style, table style, headings, stamps/signatures if visible, and general document type.

Use SOURCE OCR TEXT below as the main authority for the full translation content.
Translate all SOURCE OCR TEXT into formal UAE legal Arabic.
Do not invent any names, dates, numbers, amounts, banks, parties, stamps, signatures or facts.
Preserve identifiers exactly.
If a value is unclear in the OCR, write [غير واضح].

SOURCE OCR TEXT:
{source_text}

Return ONLY valid JSON. No markdown.

Job data:
job_number: {job.get("job_number", "")}
selected_collection: {job.get("selected_collection", "Auto Detect")}
target_language: {job.get("target_language", "Arabic")}
font_family: Sakkal Majalla
font_size: 14pt
direction: rtl
alignment: right

Required JSON:
{{
  "translated_text": "",
  "translated_html": "",
  "matched_collection": "",
  "matched_document_type": "",
  "matched_sample": "",
  "layout_type": "structured_legal_translation",
  "layout_plan_json": {{
    "matched_collection": "",
    "matched_document_type": "",
    "matched_sample": "",
    "layout_type": "structured_legal_translation",
    "font_family": "Sakkal Majalla",
    "font_size": "14pt",
    "direction": "rtl",
    "alignment": "right",
    "line_height": "1.4",
    "use_frame": true,
    "content_width": 10000,
    "blocks": []
  }}
}}

Allowed blocks:
title, subtitle, section_heading, paragraph, field_table, data_table, signature_block, page_break, spacer.

Rules for layout_plan_json.blocks:
- Use the FIRST PAGE IMAGE as visual layout guidance.
- Use SOURCE OCR TEXT for complete content.
- Do not create one page per field.
- Do not add page_break after every block.
- Use field_table for key-value fields.
- Use data_table for multi-column tables.
- Keep related fields in the same table where possible.
- Use signature_block only for visible or OCR-mentioned stamps, seals, signatures.
- Output must fit inside a legal translation letterhead/frame.

def extract_json(text):
    text = text.strip()
    text = re.sub(r"^```json\s*", "", text, flags=re.I)
    text = re.sub(r"^```\s*", "", text, flags=re.I)
    text = re.sub(r"```$", "", text).strip()

    try:
        return json.loads(text)
    except Exception:
        pass

    start = text.find("{")
    end = text.rfind("}")
    if start >= 0 and end > start:
        return json.loads(text[start:end + 1])

    raise ValueError("OpenAI did not return valid JSON")


def source_to_images(path):
    ext = os.path.splitext(path)[1].lower()
    out = []

    if ext in [".png", ".jpg", ".jpeg"]:
        return [path]

    if ext == ".pdf":
        doc = fitz.open(path)
        max_pages = 1

        for i in range(max_pages):
            page = doc[i]

            pix = page.get_pixmap(
                matrix=fitz.Matrix(0.8, 0.8),
                alpha=False
            )

            img_path = os.path.join(
                tempfile.gettempdir(),
                f"moajam_page_{uuid.uuid4().hex}_{i}.jpg"
            )

            pix.save(img_path)
            out.append(img_path)

        doc.close()
        return out

    return []


def build_docx(job_number, translated_text, layout_plan, letterhead_path):
    doc = Document()

    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(4.2)
    section.bottom_margin = Cm(2.2)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
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
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.space_before = Pt(0)


def set_document_defaults(doc):
    normal = doc.styles["Normal"]
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

    parts = [p.strip() for p in text.split("\n") if p.strip()]
    for part in parts:
        p = doc.add_paragraph()
        set_paragraph_rtl(p)

        if align == "center":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == "left":
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.1

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

    for label, value in rows:
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
            label = row.get("label") or row.get("key") or ""
            value = row.get("value") or row.get("text") or ""
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
    ext = os.path.splitext(parsed.path)[1].lower() or ".bin"

    response = requests.get(url, timeout=60)
    response.raise_for_status()

    fd, path = tempfile.mkstemp(suffix=ext)
    with os.fdopen(fd, "wb") as f:
        f.write(response.content)

    return path


def image_to_data_url(path):
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("utf-8")
    return f"data:image/png;base64,{b64}"


def clean_filename(value):
    value = str(value or "MOAJAM").strip()
    safe = "".join(ch for ch in value if ch.isalnum() or ch in ("-", "_"))
    return safe or "MOAJAM"


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
