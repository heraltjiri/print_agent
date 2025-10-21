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
    # This function simply forwards the PDF to the OS printing subsystem on Windows.
    if os.name != 'nt':
        return False, 'unsupported_os', 'Only Windows printing is supported by this agent'

    # If pywin32 is available, prefer ShellExecute with 'printto' when a specific printer is requested.
    if printer:
        try:
            from win32com.shell import shell as win_shell
            from win32com.client import Dispatch
            # Use WScript.Shell to invoke printto
            Dispatch('WScript.Shell').ShellExecute(filepath, f'"{printer}"', None, 'printto', 0)
            return True, 'printto_shell', None
        except Exception:
            # Fall back to os.startfile (cannot specify printer)
            try:
                os.startfile(filepath, 'print')
                return True, 'os_startfile_print_no_printer_specified', None
            except Exception as e:
                return False, 'os_startfile_error', str(e)

    # No printer specified: just call default print
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
