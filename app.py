import os
import uuid
import tempfile
import logging
from flask import Flask, request, jsonify
from flask_cors import CORS
import requests
import fitz  # PyMuPDF
from PIL import Image, ImageWin
import win32print
import win32ui
from dotenv import load_dotenv

# -------------------- CONFIG --------------------
load_dotenv()

PORT = int(os.environ.get("PRINT_AGENT_PORT", 5000))
CORS_ORIGINS = os.environ.get("PRINT_AGENT_CORS_ORIGINS", "").split(",")

# -------------------- LOGGING --------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

# -------------------- FLASK --------------------
app = Flask(__name__)
if CORS_ORIGINS and CORS_ORIGINS != [""]:
    CORS(app, origins=[o.strip() for o in CORS_ORIGINS])
else:
    CORS(app)  # default: allow all for local testing

# -------------------- PRINT FUNCTION --------------------
def do_print(filepath, printer=None):
    """Print PDF silently to Windows printer using rendered bitmap pages."""
    if os.name != "nt":
        return False, "unsupported_os", "Only Windows printing is supported."

    logging.info(f"Starting print: {filepath}")
    try:
        doc = fitz.open(filepath)
        target_printer = printer or win32print.GetDefaultPrinter()

        hprinter = win32print.OpenPrinter(target_printer)
        hdc = win32ui.CreateDC()
        hdc.CreatePrinterDC(target_printer)
        hdc.StartDoc(filepath)

        for page_num in range(doc.page_count):
            page = doc.load_page(page_num)
            pix = page.get_pixmap(dpi=203)  # 203 DPI pro štítky
            mode = "RGB" if pix.n < 4 else "RGBA"
            img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)

            # Resize to printable area
            printable_area = hdc.GetDeviceCaps(8), hdc.GetDeviceCaps(10)  # HORZRES, VERTRES
            img = img.resize(printable_area)

            dib = ImageWin.Dib(img)
            hdc.StartPage()
            dib.draw(hdc.GetHandleOutput(), (0, 0, img.width, img.height))
            hdc.EndPage()

        hdc.EndDoc()
        hdc.DeleteDC()
        win32print.ClosePrinter(hprinter)

        logging.info(f"Printed {filepath} to printer {target_printer} successfully.")
        return True, "render_print", None

    except Exception as e:
        logging.error(f"Print failed: {e}", exc_info=True)
        return False, "render_error", str(e)

# -------------------- ROUTES --------------------
@app.route("/", methods=["GET"])
def index():
    return jsonify({"status": "ok", "message": "print-agent running"})

@app.route("/print", methods=["POST"])
def print_pdf():
    printer = request.values.get("printer")
    tmp_dir = tempfile.gettempdir()
    filename = f"print_agent_{uuid.uuid4()}.pdf"
    path = os.path.join(tmp_dir, filename)

    # Handle file upload
    if "file" in request.files:
        f = request.files["file"]
        f.save(path)
    else:
        url = request.values.get("url")
        if not url:
            return jsonify({"ok": False, "error": "No file uploaded and no 'url' provided"}), 400
        try:
            r = requests.get(url, stream=True, timeout=20)
            r.raise_for_status()
            with open(path, "wb") as fh:
                for chunk in r.iter_content(8192):
                    fh.write(chunk)
        except Exception as e:
            logging.error(f"Failed to download URL: {e}")
            return jsonify({"ok": False, "error": f"Failed to download URL: {e}"}), 400

    # Validate file
    if os.path.getsize(path) == 0:
        return jsonify({"ok": False, "error": "Uploaded file is empty"}), 400

    ok, method, extra = do_print(path, printer=printer)

    # Cleanup
    try:
        os.remove(path)
    except Exception:
        pass

    if ok:
        return jsonify({"ok": True, "method": method})
    else:
        return jsonify({"ok": False, "method": method, "detail": extra}), 500

# -------------------- MAIN --------------------
if __name__ == "__main__":
    logging.info(f"Starting print-agent on port {PORT}")
    app.run(host="0.0.0.0", port=PORT, debug=False)
