import os
import uuid
import tempfile
import subprocess
from urllib.parse import urlparse

import requests
from flask import Flask, request, jsonify, send_from_directory
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)
OUTPUT_DIR = os.path.join(os.getcwd(), "generated")
os.makedirs(OUTPUT_DIR, exist_ok=True)

@app.route("/")
def home():
    return "Moajam Pandoc HTML DOCX Service Running"

@app.route("/generate-docx", methods=["POST"])
def generate_docx():
    data = request.get_json(force=True, silent=True) or {}
    job_number = clean_filename(data.get("job_number") or "MOAJAM-JOB")
    final_html = data.get("final_html") or data.get("translated_html") or ""
    translated_text = data.get("translated_text") or ""
    letterhead_url = data.get("letterhead_image_link") or ""

    if not letterhead_url:
        return jsonify({"status": "error", "message": "Missing letterhead_image_link"}), 400
    if not final_html and not translated_text:
        return jsonify({"status": "error", "message": "Missing final_html/translated_text"}), 400

    try:
        letterhead_path = download_file(letterhead_url)
        if not final_html:
            final_html = text_to_html(translated_text)

        html_path = write_full_html(final_html)
        filename = f"{job_number}-Final-Translation-{uuid.uuid4().hex[:8]}.docx"
        output_path = os.path.join(OUTPUT_DIR, filename)

        convert_html_to_docx(html_path, output_path)
        apply_letterhead_and_margins(output_path, letterhead_path)

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

def write_full_html(inner_html):
    html = """<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<style>
@page { size: A4; margin: 4.4cm 1.5cm 2.6cm 1.5cm; }
body {
  direction: rtl;
  text-align: right;
  font-family: "Sakkal Majalla", "Arial", "Tahoma", sans-serif;
  font-size: 14pt;
  line-height: 1.35;
}
table {
  width: 100%;
  border-collapse: collapse;
  direction: rtl;
  margin: 10px 0 16px 0;
  page-break-inside: auto;
}
tr { page-break-inside: avoid; page-break-after: auto; }
th, td {
  border: 1px solid #555;
  padding: 5px 7px;
  vertical-align: top;
  text-align: right;
  font-size: 11.5pt;
}
th { background-color: #d9eaf7; font-weight: bold; }
h1, h2, h3 {
  direction: rtl;
  text-align: right;
  font-weight: bold;
  margin: 12px 0 8px 0;
}
h1 { font-size: 18pt; text-align: center; }
h2 { font-size: 16pt; color: #1f4e79; }
h3 { font-size: 14.5pt; color: #1f4e79; }
p { margin: 5px 0; }
ul, ol { direction: rtl; text-align: right; }
.page-break { page-break-before: always; }
</style>
</head>
<body>
""" + str(inner_html) + """
</body>
</html>"""
    fd, path = tempfile.mkstemp(suffix=".html")
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        f.write(html)
    return path

def convert_html_to_docx(html_path, output_path):
    try:
        result = subprocess.run(
            ["pandoc", html_path, "-o", output_path, "--from=html", "--to=docx"],
            capture_output=True,
            text=True,
            timeout=180
        )
    except FileNotFoundError:
        raise RuntimeError("pandoc command not found. Install pandoc in Render build command.")
    except Exception as exc:
        raise RuntimeError(str(exc))

    if result.returncode != 0 or not os.path.exists(output_path):
        msg = (result.stderr or result.stdout or "unknown pandoc error").strip()
        raise RuntimeError("HTML to DOCX conversion failed: " + msg)

def apply_letterhead_and_margins(docx_path, letterhead_path):
    doc = Document(docx_path)
    for section in doc.sections:
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(4.4)
        section.bottom_margin = Cm(2.6)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)
        section.header_distance = Cm(0)
        section.footer_distance = Cm(0)
        add_letterhead_to_header(section, letterhead_path)
    set_normal_style(doc)
    doc.save(docx_path)

def add_letterhead_to_header(section, image_path):
    header = section.header
    paragraph = header.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.space_before = Pt(0)
    run = paragraph.add_run()
    run.add_picture(image_path, width=Cm(21), height=Cm(29.7))

def set_normal_style(doc):
    normal = doc.styles["Normal"]
    normal.font.name = "Sakkal Majalla"
    normal.font.size = Pt(14)

def text_to_html(text):
    parts = []
    for line in str(text or "").splitlines():
        line = line.strip()
        if line:
            parts.append("<p>" + escape_html(line) + "</p>")
    return '<div dir="rtl" style="direction:rtl;text-align:right">' + "\n".join(parts) + "</div>"

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
