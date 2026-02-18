import io
import os
import tempfile
import zipfile
from datetime import datetime

from flask import Flask, Response, jsonify, make_response, request
from openpyxl import Workbook
from pdf2docx import Converter
import pdfplumber
import fitz

app = Flask(__name__)


def _cors(resp):
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Access-Control-Allow-Methods"] = "GET,POST,OPTIONS"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
    return resp


@app.after_request
def add_cors_headers(resp):
    return _cors(resp)


@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"ok": True, "time": datetime.utcnow().isoformat() + "Z"})


@app.route("/api/<path:_>", methods=["OPTIONS"])
def preflight(_):
    return ("", 204)


@app.route("/api/pdf-to-word", methods=["POST"])
def pdf_to_word():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file uploaded"}), 400

    src_path = None
    dst_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as src:
            src.write(f.read())
            src_path = src.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as dst:
            dst_path = dst.name

        cv = Converter(src_path)
        cv.convert(dst_path)
        cv.close()

        with open(dst_path, "rb") as out:
            data = out.read()

        filename = os.path.splitext(f.filename or "converted")[0] + ".docx"
        resp = make_response(data)
        resp.headers["Content-Type"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        resp.headers["Content-Disposition"] = f'attachment; filename="{filename}"'
        return resp
    except Exception as ex:
        return jsonify({"error": str(ex)}), 500
    finally:
        for p in [src_path, dst_path]:
            if p and os.path.exists(p):
                try:
                    os.remove(p)
                except Exception:
                    pass


@app.route("/api/pdf-to-excel", methods=["POST"])
def pdf_to_excel():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file uploaded"}), 400

    src_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as src:
            src.write(f.read())
            src_path = src.name

        wb = Workbook()
        first_sheet = True

        with pdfplumber.open(src_path) as pdf:
            for page_idx, page in enumerate(pdf.pages, start=1):
                tables = page.extract_tables() or []
                if not tables:
                    continue

                for t_idx, table in enumerate(tables, start=1):
                    title = f"P{page_idx}_T{t_idx}"
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

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        filename = os.path.splitext(f.filename or "converted")[0] + ".xlsx"
        resp = make_response(buf.getvalue())
        resp.headers["Content-Type"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        resp.headers["Content-Disposition"] = f'attachment; filename="{filename}"'
        return resp
    except Exception as ex:
        return jsonify({"error": str(ex)}), 500
    finally:
        if src_path and os.path.exists(src_path):
            try:
                os.remove(src_path)
            except Exception:
                pass


@app.route("/api/pdf-to-image", methods=["POST"])
def pdf_to_image():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file uploaded"}), 400

    src_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as src:
            src.write(f.read())
            src_path = src.name

        doc = fitz.open(src_path)
        out = io.BytesIO()
        with zipfile.ZipFile(out, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            for i, page in enumerate(doc, start=1):
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                png = pix.tobytes("png")
                zf.writestr(f"page-{i}.png", png)
        doc.close()

        out.seek(0)
        filename = os.path.splitext(f.filename or "converted")[0] + "_images.zip"
        resp = make_response(out.getvalue())
        resp.headers["Content-Type"] = "application/zip"
        resp.headers["Content-Disposition"] = f'attachment; filename="{filename}"'
        return resp
    except Exception as ex:
        return jsonify({"error": str(ex)}), 500
    finally:
        if src_path and os.path.exists(src_path):
            try:
                os.remove(src_path)
            except Exception:
                pass



@app.route("/api/pdf-shrink", methods=["POST"])
def pdf_shrink():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file uploaded"}), 400

    level = (request.form.get("level") or "medium").lower()
    raster_scale = 0.9

    original_bytes = f.read()
    if not original_bytes:
        return jsonify({"error": "Uploaded file is empty"}), 400

    src = None
    out_doc = None
    try:
        # Fast pass: optimize PDF objects only.
        src = fitz.open(stream=original_bytes, filetype="pdf")
        if src.page_count > 120:
            return jsonify({"error": "PDF has too many pages for online shrink (max 120 pages)."}), 400

        optimized_buf = io.BytesIO()
        src.save(optimized_buf, garbage=3, deflate=True)
        optimized_bytes = optimized_buf.getvalue()
        src.close()
        src = None

        best_bytes = optimized_bytes if len(optimized_bytes) <= len(original_bytes) else original_bytes

        # Slow raster pass only for explicit high compression.
        if level == "high":
            src = fitz.open(stream=original_bytes, filetype="pdf")
            out_doc = fitz.open()
            for page in src:
                pix = page.get_pixmap(matrix=fitz.Matrix(raster_scale, raster_scale), alpha=False)
                jpg = pix.tobytes("jpg")
                new_page = out_doc.new_page(width=page.rect.width, height=page.rect.height)
                new_page.insert_image(new_page.rect, stream=jpg)

            raster_buf = io.BytesIO()
            out_doc.save(raster_buf, garbage=3, deflate=True)
            raster_bytes = raster_buf.getvalue()
            if len(raster_bytes) < len(best_bytes):
                best_bytes = raster_bytes

        filename = os.path.splitext(f.filename or "compressed")[0] + "_shrunk.pdf"
        resp = make_response(best_bytes)
        resp.headers["Content-Type"] = "application/pdf"
        resp.headers["Content-Disposition"] = f'attachment; filename="{filename}"'
        return resp
    except Exception as ex:
        return jsonify({"error": str(ex)}), 500
    finally:
        if src is not None:
            try:
                src.close()
            except Exception:
                pass
        if out_doc is not None:
            try:
                out_doc.close()
            except Exception:
                pass

@app.route("/api/pdf-merge", methods=["POST"])
def pdf_merge():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "No files uploaded"}), 400

    valid_files = [f for f in files if f and (f.filename or "").lower().endswith(".pdf")]
    if len(valid_files) < 2:
        return jsonify({"error": "Upload at least 2 PDF files"}), 400

    merged = fitz.open()
    opened_docs = []
    try:
        for f in valid_files:
            doc = fitz.open(stream=f.read(), filetype="pdf")
            opened_docs.append(doc)
            merged.insert_pdf(doc)

        buf = io.BytesIO()
        merged.save(buf, garbage=4, deflate=True, clean=True, linear=True)
        buf.seek(0)

        resp = make_response(buf.getvalue())
        resp.headers["Content-Type"] = "application/pdf"
        resp.headers["Content-Disposition"] = 'attachment; filename="merged.pdf"'
        return resp
    except Exception as ex:
        return jsonify({"error": str(ex)}), 500
    finally:
        for doc in opened_docs:
            try:
                doc.close()
            except Exception:
                pass
        try:
            merged.close()
        except Exception:
            pass
if __name__ == "__main__":
      port = int(os.environ.get("PORT", 5000))
      app.run(host="0.0.0.0", port=port, debug=False)



