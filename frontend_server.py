import sys
import os
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent
OUT_DIR = ROOT_DIR / "out"

class CleanSPAHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(OUT_DIR), **kwargs)

    def do_GET(self):
        url_path = self.path.split("?")[0].strip("/")
        query = ("?" + self.path.split("?")[1]) if "?" in self.path else ""
        if url_path:
            target = OUT_DIR / f"{url_path}.html"
            if target.is_file():
                self.path = f"/{url_path}.html{query}"
            elif (OUT_DIR / url_path / "index.html").is_file():
                self.path = f"/{url_path}/index.html{query}"
        return super().do_GET()

    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "*")
        self.send_header("Access-Control-Allow-Headers", "*")
        self.send_header("Permissions-Policy", "camera=*, geolocation=*, microphone=*, clipboard-read=*, clipboard-write=*, display-capture=*")
        self.send_header("Feature-Policy", "camera *; geolocation *; microphone *; clipboard-read *; clipboard-write *")
        super().end_headers()

    def log_message(self, format, *args):
        # Suppress verbose log spam
        pass

if __name__ == "__main__":
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 3000
    server = ThreadingHTTPServer(("0.0.0.0", port), CleanSPAHandler)
    print(f"Frontend static server listening on http://0.0.0.0:{port}")
    server.serve_forever()
