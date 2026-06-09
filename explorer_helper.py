"""
explorer_helper.py
Tiny helper server — called by the Outlook Items button in search.html.
Runs on port 9999. When it receives GET /open-emails, it opens the
Emails folder (next to this script) in Windows File Explorer.
"""
import os
import subprocess
from http.server import BaseHTTPRequestHandler, HTTPServer

# The Emails folder sits next to this script (same EmailBackup folder)
EMAILS_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Emails")

class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path == "/open-emails":
            subprocess.Popen(f'explorer "{EMAILS_FOLDER}"')
        # Always respond 200 so the browser fetch() doesn't error
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()

    def log_message(self, *args):
        pass  # suppress console noise

if __name__ == "__main__":
    server = HTTPServer(("localhost", 9999), Handler)
    print(f"[+] Explorer helper running on port 9999")
    print(f"[+] Emails folder: {EMAILS_FOLDER}")
    server.serve_forever()
