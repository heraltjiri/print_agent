# Print Agent

Simple Python print agent that exposes an HTTP API to silently print PDF files on Windows.

Features
- POST /print to upload a PDF or provide a URL to a PDF
 - Forwards PDF to the Windows OS printing subsystem (uses os.startfile or Windows COM when available).
 - Helper script to add the agent to Windows startup.

Usage
1. Install dependencies:

   pip install -r requirements.txt

2. This agent forwards received PDFs to the Windows OS printing subsystem. No external PDF CLI is required by default.

Optional silent rendering printing
- If you want the agent to render PDF pages and send raw bitmaps to the printer (no external PDF viewer, more silent), enable `PRINT_AGENT_USE_RENDER=true` in your `.env` and install the optional dependencies:

   pip install PyMuPDF Pillow pywin32

   Note: Rendering-based printing uses Windows printing APIs and may require tuning per printer. It's experimental here and should be tested on the target machine.

3. Run the agent:

   python app.py

4. API examples
- Upload file (multipart/form-data): POST http://localhost:5000/print with form field `file`.
- Print from URL: POST http://localhost:5000/print with form value `url=https://example.com/file.pdf`.
- Optional: add `printer=Printer_Name` to target a specific printer.

Startup
Use `install_startup.py` on Windows to create a shortcut in the Startup folder so the agent runs on user login. You may need to run it as the target user.

Security
This is intentionally minimal. If exposing on a network, protect it with a firewall or add authentication.
