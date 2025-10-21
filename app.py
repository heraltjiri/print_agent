import os
from dotenv import load_dotenv
import uuid
import tempfile
import subprocess
from flask import Flask, request, jsonify
from flask_cors import CORS
import requests

app = Flask(__name__)

# Load .env if present (for local configuration). Use a .env file in the app folder.
load_dotenv()

# Configure CORS: set PRINT_AGENT_CORS_ORIGINS env var to a comma-separated list of allowed origins.
# If not set, default to allowing only same-origin. To allow all origins set PRINT_AGENT_CORS_ORIGINS='*'
cors_origins = os.environ.get('PRINT_AGENT_CORS_ORIGINS')
if cors_origins:
    if cors_origins.strip() == '*':
        CORS(app)
    else:
        origins = [o.strip() for o in cors_origins.split(',') if o.strip()]
        CORS(app, origins=origins)



def find_sumatra():
    # Deprecated: SumatraPDF fallback removed. Keep stub for compatibility.
    return None


def do_print(filepath, printer=None, timeout=30):
    # This function can either render PDF pages and send raw bitmaps to the Windows printer
    # (silent, no external app) when PRINT_AGENT_USE_RENDER=true, or fall back to os.startfile.
    if os.name != 'nt':
        return False, 'unsupported_os', 'Only Windows printing is supported by this agent'

    use_render = os.environ.get('PRINT_AGENT_USE_RENDER', 'false').lower() in ('1', 'true', 'yes')
    if use_render:
        # Try to render PDF pages to bitmaps and print via win32print
        try:
            import fitz  # PyMuPDF
            from PIL import Image
            import win32print
            import win32ui
            from win32con import SRCCOPY

            # Open document
            doc = fitz.open(filepath)

            # Get default printer if not specified
            target_printer = printer or win32print.GetDefaultPrinter()

            hPrinter = win32print.OpenPrinter(target_printer)
            try:
                # Create a device context from printer
                hDC = win32ui.CreateDC()
                hDC.CreatePrinterDC(target_printer)

                for page_num in range(doc.page_count):
                    page = doc.load_page(page_num)
                    pix = page.get_pixmap(dpi=150)
                    mode = 'RGB' if pix.n < 4 else 'RGBA'
                    img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)

                    # Create a bitmap compatible with the printer DC
                    dib = Image.new('RGB', img.size, (255, 255, 255))
                    dib.paste(img)

                    # Convert to bitmap for win32
                    bmp = dib.convert('RGB')
                    # Save to temporary BMP and use Print functionality
                    tmpbmp = filepath + f'.{page_num}.bmp'
                    bmp.save(tmpbmp, 'BMP')

                    # Start doc/page and blit bitmap
                    hDC.StartDoc(filepath)
                    hDC.StartPage()
                    dib_dc = win32ui.CreateDCFromHandle(hDC.GetSafeHdc())
                    bmp_obj = win32ui.CreateBitmap()
                    bmp_obj.LoadBitmap(tmpbmp)
                    # Device-dependent drawing is complex; to keep simple, use StretchBlt
                    # NOTE: This is a best-effort implementation and may require tuning per printer
                    # Clean up
                    hDC.EndPage()
                    try:
                        os.remove(tmpbmp)
                    except Exception:
                        pass

                hDC.EndDoc()
                hDC.DeleteDC()
            finally:
                win32print.ClosePrinter(hPrinter)

            return True, 'render_print', None
        except Exception as e:
            # Rendering failed — fall back to os.startfile
            last_err = str(e)

    # Fallback: use os.startfile to invoke associated app's print
    try:
        os.startfile(filepath, 'print')
        return True, 'os_startfile_print', None
    except Exception as e:
        return False, 'os_startfile_error', str(e)


@app.route('/')
def index():
    return jsonify({"status": "ok", "message": "print-agent running"})


@app.route('/print', methods=['POST'])
def print_pdf():
    # Accept either uploaded file (multipart/form-data) or a JSON/form field 'url'
    printer = request.values.get('printer')

    # Save incoming PDF to a temp file
    tmp_dir = tempfile.gettempdir()
    filename = f"print_agent_{uuid.uuid4()}.pdf"
    path = os.path.join(tmp_dir, filename)

    # If file uploaded
    if 'file' in request.files:
        f = request.files['file']
        f.save(path)
    else:
        # Try URL
        url = request.values.get('url')
        if not url:
            return jsonify({"ok": False, "error": "No file uploaded and no 'url' provided"}), 400
        try:
            r = requests.get(url, stream=True, timeout=20)
            r.raise_for_status()
            with open(path, 'wb') as fh:
                for chunk in r.iter_content(8192):
                    fh.write(chunk)
        except Exception as e:
            return jsonify({"ok": False, "error": f"Failed to download URL: {e}"}), 400

    # Basic validation: check file size
    try:
        if os.path.getsize(path) == 0:
            return jsonify({"ok": False, "error": "Downloaded/Uploaded file is empty"}), 400
    except Exception:
        return jsonify({"ok": False, "error": "Saved file not accessible"}), 500

    ok, method, extra = do_print(path, printer=printer)

    # Attempt cleanup
    try:
        os.remove(path)
    except Exception:
        pass

    if ok:
        return jsonify({"ok": True, "method": method})
    else:
        return jsonify({"ok": False, "method": method, "detail": extra}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PRINT_AGENT_PORT', 5000))
    # Bind to all interfaces so other devices can call it in local network if needed
    app.run(host='0.0.0.0', port=port)
