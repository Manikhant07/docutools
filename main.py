from pathlib import Path
import shutil
import os
import uuid
from flask import Flask, render_template, request, send_file, jsonify

from tools import (
    pdf_to_image, image_to_pdf, merge_pdfs, split_pdf, compress_pdf,
    pdf_to_text, lock_pdf, word_to_pdf, ppt_to_pdf, excel_to_pdf,
    pdf_to_word, pdf_to_ppt, pdf_to_excel, crop_pdf, sign_pdf, add_page_numbers
)

# =========================================================
# CONFIGURATION
# =========================================================

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

# =========================================================
# ROUTES & HELPERS
# =========================================================

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/health")
def health():
    return jsonify({
        "status": "ok",
        "message": "DocuTools backend is running",
        "tools": 16
    })

def save_uploaded_file(file):
    if not file or not file.filename:
        raise ValueError("No file selected.")
    original_name = Path(file.filename).name
    file_path = UPLOAD_DIR / f"{uuid.uuid4().hex}_{original_name}"
    file.save(file_path)
    return file_path

def cleanup_file(file_path):
    try:
        if file_path and Path(file_path).exists():
            Path(file_path).unlink()
    except OSError:
        pass

def send_result(file_path, download_name=None):
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError("Output file was not created.")
    return send_file(
        file_path,
        as_attachment=True,
        download_name=download_name if download_name else file_path.name
    )

# =========================================================
# API TOOLS
# =========================================================

# 1. PDF -> IMAGE
@app.route("/api/pdf-to-image", methods=["POST"])
def api_pdf_to_image():
    input_file = None
    try:
        input_file = save_uploaded_file(request.files.get("file"))
        image_format = request.form.get("format", "png").lower()
        temp_output_dir = OUTPUT_DIR / f"pdf_images_{uuid.uuid4().hex}"
        generated_files = pdf_to_image(input_file, temp_output_dir, image_format)
        if not generated_files:
            raise RuntimeError("No images were generated.")
        return send_result(generated_files[0], generated_files[0].name)
    except Exception as error:
        return jsonify({"success": False, "error": str(error)}), 400
    finally:
        cleanup_file(input_file)

# 2. IMAGE -> PDF
@app.route("/api/image-to-pdf", methods=["POST"])
def api_image_to_pdf():
    uploaded_files = []
    try:
        files = request.files.getlist("files")
        if not files:
            raise ValueError("Please select at least one image.")
        for file in files:
            if file and file.filename:
                uploaded_files.append(save_uploaded_file(file))
        if not uploaded_files:
            raise ValueError("No valid image files selected.")
        output_file = OUTPUT_DIR / f"images_{uuid.uuid4().hex}.pdf"
        result = image_to_pdf(uploaded_files, output_file)
        return send_result(result, "DocuTools_Images.pdf")
    except Exception as error:
        return jsonify({"success": False, "error": str(error)}), 400
    finally:
        for file in uploaded_files:
            cleanup_file(file)

# 3. MERGE PDF
@app.route("/api/merge-pdf", methods=["POST"])
def api_merge_pdf():
    uploaded_files = []
    try:
        files = request.files.getlist("files")
        if len(files) < 2:
            raise ValueError("Please select at least two PDF files.")
        for file in files:
            if file and file.filename:
                uploaded_files.append(save_uploaded_file(file))
        output_file = OUTPUT_DIR / f"merged_{uuid.uuid4().hex}.pdf"
        result = merge_pdfs(uploaded_files, output_file)
        return send_result(result, "DocuTools_Merged.pdf")
    except Exception as error:
        return jsonify({"success": False, "error": str(error)}), 400
    finally:
        for file in uploaded_files:
            cleanup_file(file)

# 4. SPLIT PDF
@app.route("/api/split-pdf", methods=["POST"])
def api_split_pdf():
    input_file = None
    try:
        input_file = save_uploaded_file(request.files.get("file"))
        result_dir = OUTPUT_DIR / f"split_{uuid.uuid4().hex}"
        generated_files = split_pdf(input_file, result_dir)
        if not generated_files:
            raise RuntimeError("No pages were generated.")
        return send_result(generated_files[0], generated_files[0].name)
    except Exception as error:
        return jsonify({"success": False, "error": str(error)}), 400
    finally:
        cleanup_file(input_file)

# 5. COMPRESS PDF
@app.route("/api/compress-pdf", methods=["POST"])
def api_compress_pdf():
    input_file = None
    try:
        input_file = save_uploaded_file(request.files.get("file"))
        output_file = OUTPUT_DIR / f"compressed_{uuid.uuid4().hex}.pdf"
        result = compress_pdf(input_file, output_file)
        return send_result(result, "DocuTools_Compressed.pdf")
    except Exception as error:
        return jsonify({"success": False, "error": str(error)}), 400
    finally:
        cleanup_file(input_file)

# 6. PDF -> TEXT
@app.route("/api/pdf-to-text", methods=["POST"])
def api_pdf_to_text():
    input_file = None
    try:
        input_file = save_uploaded_file(request.files.get("file"))
        output_file = OUTPUT_DIR / f"text_{uuid.uuid4().hex}.txt"
        result = pdf_to_text(input_file, output_file)
        return send_result(result, "DocuTools_Text.txt")
    except Exception as error:
        return jsonify({"success": False, "error": str(error)}), 400
    finally:
        cleanup_file(input_file)

# 7. LOCK PDF
@app.route("/api/lock-pdf", methods=["POST"])
def api_lock_pdf():
    input_file = None
    try:
        password = request.form.get("password", "")
        if not password:
            raise ValueError("Please enter a password.")
        input_file = save_uploaded_file(request.files.get("file"))
        output_file = OUTPUT_DIR / f"protected_{uuid.uuid4().hex}.pdf"
        result = lock_pdf(input_file, password, output_file)
        return send_result(result, "DocuTools_Protected.pdf")
    except Exception as error:
        return jsonify({"success": False, "error": str(error)}), 400
    finally:
        cleanup_file(input_file)

# 8. WORD -> PDF
@app.route("/api/word-to-pdf", methods=["POST"])
def api_word_to_pdf():
    input_file = None
    try:
        input_file = save_uploaded_file(request.files.get("file"))
        result = word_to_pdf(input_file, OUTPUT_DIR)
        return send_result(result, "DocuTools_Word.pdf")
    except Exception as error:
        return jsonify({"success": False, "error": str(error)}), 400
    finally:
        cleanup_file(input_file)

# 9. POWERPOINT -> PDF
@app.route("/api/powerpoint-to-pdf", methods=["POST"])
def api_powerpoint_to_pdf():
    input_file = None
    try:
        input_file = save_uploaded_file(request.files.get("file"))
        result = ppt_to_pdf(input_file, OUTPUT_DIR)
        return send_result(result, "DocuTools_Presentation.pdf")
    except Exception as error:
        return jsonify({"success": False, "error": str(error)}), 400
    finally:
        cleanup_file(input_file)

# 10. EXCEL -> PDF
@app.route("/api/excel-to-pdf", methods=["POST"])
def api_excel_to_pdf():
    input_file = None
    try:
        input_file = save_uploaded_file(request.files.get("file"))
        result = excel_to_pdf(input_file, OUTPUT_DIR)
        return send_result(result, "DocuTools_Excel.pdf")
    except Exception as error:
        return jsonify({"success": False, "error": str(error)}), 400
    finally:
        cleanup_file(input_file)

# 11. PDF -> WORD
@app.route("/api/pdf-to-word", methods=["POST"])
def api_pdf_to_word():
    input_file = None
    try:
        input_file = save_uploaded_file(request.files.get("file"))
        output_file = OUTPUT_DIR / f"word_{uuid.uuid4().hex}.docx"
        result = pdf_to_word(input_file, output_file)
        return send_result(result, "DocuTools_Word.docx")
    except Exception as error:
        return jsonify({"success": False, "error": str(error)}), 400
    finally:
        cleanup_file(input_file)

# 12. PDF -> POWERPOINT
@app.route("/api/pdf-to-powerpoint", methods=["POST"])
def api_pdf_to_powerpoint():
    input_file = None
    try:
        input_file = save_uploaded_file(request.files.get("file"))
        output_file = OUTPUT_DIR / f"powerpoint_{uuid.uuid4().hex}.pptx"
        result = pdf_to_ppt(input_file, output_file)
        return send_result(result, "DocuTools_Presentation.pptx")
    except Exception as error:
        return jsonify({"success": False, "error": str(error)}), 400
    finally:
        cleanup_file(input_file)

# 13. PDF -> EXCEL
@app.route("/api/pdf-to-excel", methods=["POST"])
def api_pdf_to_excel():
    input_file = None
    try:
        input_file = save_uploaded_file(request.files.get("file"))
        output_file = OUTPUT_DIR / f"excel_{uuid.uuid4().hex}.xlsx"
        result = pdf_to_excel(input_file, output_file)
        return send_result(result, "DocuTools_Excel.xlsx")
    except Exception as error:
        return jsonify({"success": False, "error": str(error)}), 400
    finally:
        cleanup_file(input_file)

# 14. CROP PDF
@app.route("/api/crop-pdf", methods=["POST"])
def api_crop_pdf():
    input_file = None
    try:
        input_file = save_uploaded_file(request.files.get("file"))
        left = request.form.get("left", 0)
        top = request.form.get("top", 0)
        right = request.form.get("right", 0)
        bottom = request.form.get("bottom", 0)
        output_file = OUTPUT_DIR / f"cropped_{uuid.uuid4().hex}.pdf"
        result = crop_pdf(input_file, output_file, left, top, right, bottom)
        return send_result(result, "DocuTools_Cropped.pdf")
    except Exception as error:
        return jsonify({"success": False, "error": str(error)}), 400
    finally:
        cleanup_file(input_file)

# 15. SIGN PDF
@app.route("/api/sign-pdf", methods=["POST"])
def api_sign_pdf():
    input_file = None
    signature_file = None
    try:
        signature = request.files.get("signature")
        if not signature or not signature.filename:
            raise ValueError("Please upload a signature image.")
        input_file = save_uploaded_file(request.files.get("file"))
        signature_file = save_uploaded_file(signature)
        page_number = request.form.get("page", 1)
        x = request.form.get("x", 50)
        y = request.form.get("y", 50)
        width = request.form.get("width", 150)
        height = request.form.get("height", 60)
        output_file = OUTPUT_DIR / f"signed_{uuid.uuid4().hex}.pdf"
        result = sign_pdf(input_file, signature_file, output_file, page_number, x, y, width, height)
        return send_result(result, "DocuTools_Signed.pdf")
    except Exception as error:
        return jsonify({"success": False, "error": str(error)}), 400
    finally:
        cleanup_file(input_file)
        cleanup_file(signature_file)

# 16. ADD PAGE NUMBERS
@app.route("/api/add-page-numbers", methods=["POST"])
def api_add_page_numbers():
    input_file = None
    try:
        input_file = save_uploaded_file(request.files.get("file"))
        position = request.form.get("position", "bottom-center")
        start_number = request.form.get("start_number", 1)
        font_size = request.form.get("font_size", 10)
        output_file = OUTPUT_DIR / f"numbered_{uuid.uuid4().hex}.pdf"
        result = add_page_numbers(input_file, output_file, position, start_number, font_size)
        return send_result(result, "DocuTools_Page_Numbers.pdf")
    except Exception as error:
        return jsonify({"success": False, "error": str(error)}), 400
    finally:
        cleanup_file(input_file)

# =========================================================
# ERROR HANDLERS & SERVER
# =========================================================

@app.errorhandler(413)
def file_too_large(error):
    return jsonify({"success": False, "error": "File is too large. Maximum size is 50 MB."}), 413

@app.errorhandler(500)
def internal_error(error):
    return jsonify({"success": False, "error": "Something went wrong on the server."}), 500

if __name__ == "__main__":
    print("=" * 55)
    print("DocuTools")
    print("=" * 55)
    port = int(os.environ.get("PORT", "5000"))
    print(f"Server port: {port}")
    print("Health: /health")
    print("Tools: 16")
    print("=" * 55)
    app.run(host="0.0.0.0", port=port, debug=False)