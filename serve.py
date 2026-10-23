"""Run a local server for the static reconciliation dashboard."""

from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from pathlib import Path

ROOT = Path(__file__).resolve().parent

if __name__ == "__main__":
    import os
    os.chdir(ROOT)
    server = ThreadingHTTPServer(("127.0.0.1", 8000), SimpleHTTPRequestHandler)
    print("Dashboard running at http://127.0.0.1:8000")
    server.serve_forever()
