"""
app.py — Local server untuk Dashboard Nominasi Operator RTG TPS
=================================================================
Jalankan dengan:
    python app.py

Lalu buka browser ke:
    http://localhost:5000

Tidak butuh internet (kecuali untuk load SheetJS dari CDN di index.html).
Semua file statis (index.html, style.css, dashboard.js) di-serve dari
folder yang sama dengan app.py ini.

Requirements:
    pip install flask          (opsi 1 — direkomendasikan)
    -- atau --
    python -m http.server 5000 (opsi 2 — built-in, tanpa install)
"""

import os
import sys

# ── Coba pakai Flask dulu, fallback ke built-in HTTP server ──────────
try:
    from flask import Flask, send_from_directory

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    app = Flask(__name__, static_folder=BASE_DIR)

    @app.route('/')
    def index():
        return send_from_directory(BASE_DIR, 'index.html')

    @app.route('/<path:filename>')
    def static_files(filename):
        return send_from_directory(BASE_DIR, filename)

    if __name__ == '__main__':
        port = int(os.environ.get('PORT', 5000))
        print(f"\n✅  Dashboard RTG TPS berjalan di http://localhost:{port}")
        print("    Tekan Ctrl+C untuk berhenti.\n")
        app.run(host='0.0.0.0', port=port, debug=False)

except ImportError:
    # Flask tidak terinstall — pakai built-in HTTP server
    import http.server
    import socketserver
    import webbrowser
    import threading

    PORT = 5000
    DIRECTORY = os.path.dirname(os.path.abspath(__file__))

    class Handler(http.server.SimpleHTTPRequestHandler):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, directory=DIRECTORY, **kwargs)

        def log_message(self, format, *args):
            pass  # suppress request logs

    def open_browser():
        import time
        time.sleep(0.5)
        webbrowser.open(f'http://localhost:{PORT}')

    print(f"\n⚠️  Flask tidak ditemukan. Menggunakan built-in HTTP server.")
    print(f"    Install Flask untuk fitur lebih: pip install flask\n")
    print(f"✅  Dashboard RTG TPS berjalan di http://localhost:{PORT}")
    print("    Tekan Ctrl+C untuk berhenti.\n")

    threading.Thread(target=open_browser, daemon=True).start()

    with socketserver.TCPServer(("", PORT), Handler) as httpd:
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\n\nServer dihentikan.")
            sys.exit(0)
