from __future__ import annotations

import cgi
import json
import os
import shutil
import tempfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

from processor import complete_stock


ROOT = Path(__file__).resolve().parent
INDEX = ROOT / "index.html"
HOST = os.environ.get("HOST", "0.0.0.0")
PORT = int(os.environ.get("PORT", "8765"))


class StockHandler(BaseHTTPRequestHandler):
    server_version = "StockUploader/1.0"

    def log_message(self, format, *args):
        print("%s - %s" % (self.address_string(), format % args))

    def send_bytes(self, content: bytes, content_type: str, status: int = 200, headers=None):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Content-Length", str(len(content)))
        if headers:
            for key, value in headers.items():
                self.send_header(key, value)
        self.end_headers()
        self.wfile.write(content)

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_GET(self):
        path = urlparse(self.path).path
        if path in {"/", "/index.html"}:
            self.send_bytes(INDEX.read_bytes(), "text/html; charset=utf-8")
            return
        self.send_error(404, "No encontrado")

    def do_POST(self):
        if urlparse(self.path).path != "/procesar":
            self.send_error(404, "No encontrado")
            return

        content_type = self.headers.get("Content-Type", "")
        if not content_type.startswith("multipart/form-data"):
            self.send_error(400, "El formulario debe ser multipart/form-data")
            return

        with tempfile.TemporaryDirectory(prefix="stock_upload_") as temp_name:
            temp = Path(temp_name)
            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={
                    "REQUEST_METHOD": "POST",
                    "CONTENT_TYPE": content_type,
                    "CONTENT_LENGTH": self.headers.get("Content-Length", "0"),
                },
            )

            stock_item = form["stock"] if "stock" in form else None
            if stock_item is None or not getattr(stock_item, "filename", ""):
                self.send_error(400, "Falta subir STOCK.xlsx")
                return

            stock_path = temp / "STOCK.xlsx"
            with stock_path.open("wb") as out:
                shutil.copyfileobj(stock_item.file, out)

            report_items = form["reports"] if "reports" in form else []
            if not isinstance(report_items, list):
                report_items = [report_items]

            reports_dir = temp / "reportes"
            reports_dir.mkdir()
            report_paths = []
            for item in report_items:
                if not getattr(item, "filename", ""):
                    continue
                name = Path(item.filename).name
                if Path(name).suffix.lower() not in {".xls", ".xlsx"}:
                    continue
                target = reports_dir / name
                with target.open("wb") as out:
                    shutil.copyfileobj(item.file, out)
                report_paths.append(target)

            if not report_paths:
                self.send_error(400, "Falta subir al menos un reporte .xls o .xlsx")
                return

            output_path = temp / "STOCK_COMPLETADO.xlsx"
            try:
                summary = complete_stock(stock_path, report_paths, output_path)
            except Exception as exc:
                message = json.dumps({"error": str(exc)}, ensure_ascii=False).encode("utf-8")
                self.send_bytes(message, "application/json; charset=utf-8", status=500)
                return

            headers = {
                "Content-Disposition": 'attachment; filename="STOCK_COMPLETADO.xlsx"',
                "X-Stock-Summary": json.dumps(summary, ensure_ascii=False),
            }
            self.send_bytes(
                output_path.read_bytes(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                headers=headers,
            )


def main():
    server = ThreadingHTTPServer((HOST, PORT), StockHandler)
    print(f"Servidor listo en http://{HOST}:{PORT}")
    print("Presioná Ctrl+C para detenerlo.")
    server.serve_forever()


if __name__ == "__main__":
    main()
