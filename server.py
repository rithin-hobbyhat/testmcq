"""
Local Development & LAN Server for MCQ Practice Arena.
Automatically finds local IP, exports questions if needed, launches HTTP server,
and opens browser for instant local testing and smartphone access over WiFi.
"""
import os
import sys
import socket
import webbrowser
import http.server
import socketserver
import subprocess

PORT = 8000
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
WEB_DIR = os.path.join(BASE_DIR, "web")
QUESTIONS_JSON = os.path.join(WEB_DIR, "data", "questions.json")

def ensure_data_exported():
    """Ensure questions have been exported to JSON."""
    if not os.path.exists(QUESTIONS_JSON):
        print("Data files not found in web/data. Running export_data.py...")
        from export_data import export_all
        export_all()

def get_lan_ip():
    """Find the computer's local area network (LAN) IP for smartphone access."""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        # Connecting to a public IP doesn't actually send packets but resolves local interface IP
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
    except Exception:
        ip = '127.0.0.1'
    finally:
        s.close()
    return ip

class CustomHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=WEB_DIR, **kwargs)

    def end_headers(self):
        # Enable CORS and disable aggressive caching for development
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
        super().end_headers()

def run_server():
    ensure_data_exported()
    lan_ip = get_lan_ip()
    local_url = f"http://localhost:{PORT}"
    mobile_url = f"http://{lan_ip}:{PORT}"

    print("=" * 65)
    print("       🎯 MCQ PRACTICE ARENA - LOCAL & MOBILE SERVER")
    print("=" * 65)
    print(f"\n💻 On this PC:")
    print(f"   Open in Browser:  {local_url}\n")
    print(f"📱 On your Smartphone (Connected to same Wi-Fi / Hotspot):")
    print(f"   Open in Chrome/Safari:  {mobile_url}\n")
    print("=" * 65)
    print("Press Ctrl + C in this window to stop the server at any time.\n")

    # Open PC browser automatically
    webbrowser.open(local_url)

    # Allow port reuse
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("", PORT), CustomHandler) as httpd:
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServer stopped gracefully.")

if __name__ == "__main__":
    run_server()
