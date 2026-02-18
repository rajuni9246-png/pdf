import io
import os
import zipfile
from datetime import datetime

from flask import Flask, jsonify, make_response, request

try:
    import fitz  # PyMuPDF (optional on some hosts)
except Exception:
    fitz = None

import pdfplumber
from docx import Document
from openpyxl import Workbook
from PyPDF2 import PdfFileMerger

app = Flask(__name__)


def _cors(resp):
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Access-Control-Allow-Methods"] = "GET,POST,OPTIONS"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
    return resp


@app.after_request
def add_cors_headers(resp):
    return _cors(resp)


@app.route("/api/<path:_>", methods=["OPTIONS"])
def preflight(_):
    return ("", 204)


@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({
        "ok": True,
        "python": "3.6-compatible build",
        "pymupdf_loaded": bool(fitz),
        "time": datetime.utcnow().isoformat() + "Z"
    })


@app.route("/api/pdf-to-word", methods=["POST"])
def pdf_to_word():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file uploaded"}), 400

    try:
        pdf_bytes = f.read()
        if not pdf_bytes:
            return jsonify({"error": "Uploaded file is empty"}), 400

        text_parts = []
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for idx, page in enumerate(pdf.pages, start=1):
                txt = page.extract_text() or ""
                text_parts.append("[Page {}]".format(idx))
                text_parts.append(txt)
                text_parts.append("")

        doc = Document()
        doc.add_heading("Converted from PDF", level=1)
        for block in text_parts:
            doc.add_paragraph(block)

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)

        filename = os.path.splitext(f.filename or "converted")[0] + ".docx"
        resp = make_response(buf.getvalue())
        resp.headers["Content-Type"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        resp.headers["Content-Disposition"] = 'attachment; filename="{}"'.format(filename)
        return resp
    except Exception as ex:
        return jsonify({"error": str(ex)}), 500


@app.route("/api/pdf-to-excel", methods=["POST"])
def pdf_to_excel():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file uploaded"}), 400

    try:
        pdf_bytes = f.read()
        if not pdf_bytes:
            return jsonify({"error": "Uploaded file is empty"}), 400

        wb = Workbook()
        first_sheet = True

        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page_idx, page in enumerate(pdf.pages, start=1):
                tables = page.extract_tables() or []
                if not tables:
                    continue

                for t_idx, table in enumerate(tables, start=1):
                    title = "P{}_T{}".format(page_idx, t_idx)
                    ws = wb.active if first_sheet else wb.create_sheet()
                    first_sheet = False
                    ws.title = title[:31]

                    for r_idx, row in enumerate(table, start=1):
                        if row is None:
                            continue
                        for c_idx, cell in enumerate(row, start=1):
                            ws.cell(row=r_idx, column=c_idx, value=(cell or "").strip())

        if first_sheet:
            ws = wb.active
            ws.title = "NoTables"
            ws["A1"] = "No tables detected in PDF."

        out = io.BytesIO()
        wb.save(out)
        out.seek(0)

        filename = os.path.splitext(f.filename or "converted")[0] + ".xlsx"
        resp = make_response(out.getvalue())
        resp.headers["Content-Type"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        resp.headers["Content-Disposition"] = 'attachment; filename="{}"'.format(filename)
        return resp
    except Exception as ex:
        return jsonify({"error": str(ex)}), 500


@app.route("/api/pdf-to-image", methods=["POST"])
def pdf_to_image():
    if fitz is None:
        return jsonify({"error": "PDF to image unavailable on this host (PyMuPDF not installed for Python 3.6)."}), 501

    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file uploaded"}), 400

    try:
        pdf_bytes = f.read()
        if not pdf_bytes:
            return jsonify({"error": "Uploaded file is empty"}), 400

        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        out = io.BytesIO()
        with zipfile.ZipFile(out, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            for i in range(doc.page_count):
                page = doc.loadPage(i)
                pix = page.getPixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                zf.writestr("page-{}.png".format(i + 1), pix.getImageData("png"))

        doc.close()
        out.seek(0)

        filename = os.path.splitext(f.filename or "converted")[0] + "_images.zip"
        resp = make_response(out.getvalue())
        resp.headers["Content-Type"] = "application/zip"
        resp.headers["Content-Disposition"] = 'attachment; filename="{}"'.format(filename)
        return resp
    except Exception as ex:
        return jsonify({"error": str(ex)}), 500


@app.route("/api/pdf-shrink", methods=["POST"])
def pdf_shrink():
    if fitz is None:
        return jsonify({"error": "PDF shrink unavailable on this host (PyMuPDF not installed for Python 3.6)."}), 501

    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file uploaded"}), 400

    level = (request.form.get("level") or "medium").lower()
    scale = 0.9 if level == "medium" else 0.7

    try:
        original = f.read()
        if not original:
            return jsonify({"error": "Uploaded file is empty"}), 400

        src = fitz.open(stream=original, filetype="pdf")
        if src.page_count > 120:
            src.close()
            return jsonify({"error": "PDF has too many pages for online shrink (max 120 pages)."}), 400

        fast = io.BytesIO()
        src.save(fast, garbage=3, deflate=True)
        fast_bytes = fast.getvalue()
        best_bytes = fast_bytes if len(fast_bytes) < len(original) else original

        if level == "high":
            out_doc = fitz.open()
            for i in range(src.page_count):
                page = src.loadPage(i)
                pix = page.getPixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                jpg = pix.getImageData("jpg")
                new_page = out_doc.newPage(width=page.rect.width, height=page.rect.height)
                new_page.insertImage(new_page.rect, stream=jpg)

            rb = io.BytesIO()
            out_doc.save(rb, garbage=3, deflate=True)
            raster = rb.getvalue()
            out_doc.close()
            if len(raster) < len(best_bytes):
                best_bytes = raster

        src.close()

        filename = os.path.splitext(f.filename or "compressed")[0] + "_shrunk.pdf"
        resp = make_response(best_bytes)
        resp.headers["Content-Type"] = "application/pdf"
        resp.headers["Content-Disposition"] = 'attachment; filename="{}"'.format(filename)
        return resp
    except Exception as ex:
        return jsonify({"error": str(ex)}), 500


@app.route("/api/pdf-merge", methods=["POST"])
def pdf_merge():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "No files uploaded"}), 400

    pdf_files = [x for x in files if x and (x.filename or "").lower().endswith(".pdf")]
    if len(pdf_files) < 2:
        return jsonify({"error": "Upload at least 2 PDF files"}), 400

    merger = PdfFileMerger(strict=False)
    try:
        for f in pdf_files:
            merger.append(io.BytesIO(f.read()))

        out = io.BytesIO()
        merger.write(out)
        out.seek(0)

        resp = make_response(out.getvalue())
        resp.headers["Content-Type"] = "application/pdf"
        resp.headers["Content-Disposition"] = 'attachment; filename="merged.pdf"'
        return resp
    except Exception as ex:
        return jsonify({"error": str(ex)}), 500
    finally:
        try:
            merger.close()
        except Exception:
            pass


if __name__ == "__main__":
      app.run(host="0.0.0.0", port=5000, debug=False)
