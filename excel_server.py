#!/usr/bin/env python3
import hmac
import json
import os
import tempfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from python_multipart.exceptions import MultipartParseError
from python_multipart.multipart import File as MultipartFile
from python_multipart.multipart import MultipartParser, parse_options_header

from excel_compiler import MAX_FILES, MAX_UPLOAD_BYTES as MAX_EXCEL_UPLOAD_BYTES, compile_workbooks

MAX_EXCEL_REQUEST_BYTES = MAX_EXCEL_UPLOAD_BYTES + 2 * 1024 * 1024


def make_handler(token):
    class ExcelAPIHandler(BaseHTTPRequestHandler):
        server_version = "POT-Excel-Compiler"

        def log_message(self, fmt, *args):
            super().log_message(fmt, *args)

        def send_json(self, status, body):
            payload = json.dumps(body).encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(payload)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(payload)

        def authorize(self):
            supplied = self.headers.get("Authorization", "")
            prefix = "Bearer "
            if not supplied.startswith(prefix) or not hmac.compare_digest(supplied[len(prefix):], token):
                self.send_json(401, {"error": "Token API tidak valid"})
                return False
            return True

        def do_POST(self):
            if self.path == "/api/excel/compile":
                self.compile_excel()
                return
            self.send_json(404, {"error": "Endpoint tidak ditemukan"})

        def do_GET(self):
            self.send_json(404, {"error": "Endpoint tidak ditemukan"})

        def do_PUT(self):
            self.send_json(404, {"error": "Endpoint tidak ditemukan"})

        def compile_excel(self):
            if not self.authorize():
                return
            try:
                content_length = int(self.headers.get("Content-Length", "0"))
            except (ValueError, TypeError):
                self.send_json(400, {"error": "Ukuran upload tidak valid"})
                return
            if content_length <= 0:
                self.send_json(400, {"error": "Pilih file Excel atau CSV untuk digabung"})
                return
            if content_length > MAX_EXCEL_REQUEST_BYTES:
                self.send_json(413, {"error": "Total upload melebihi batas 500 MB"})
                return

            uploads = []
            files = []
            output_path = None
            try:
                content_type, options = parse_options_header(self.headers.get("Content-Type"))
                boundary = options.get(b"boundary")
                if content_type != b"multipart/form-data" or not boundary:
                    raise ValueError("Format upload multipart tidak valid")

                state = {
                    "headers": {},
                    "header_name": bytearray(),
                    "header_value": bytearray(),
                    "upload": None,
                    "filename": None,
                }

                def on_part_begin():
                    state["headers"] = {}
                    state["header_name"].clear()
                    state["header_value"].clear()
                    state["upload"] = None
                    state["filename"] = None

                def on_header_field(data, start, end):
                    state["header_name"].extend(data[start:end])

                def on_header_value(data, start, end):
                    state["header_value"].extend(data[start:end])

                def on_header_end():
                    name = bytes(state["header_name"]).lower()
                    state["headers"][name] = bytes(state["header_value"])
                    state["header_name"].clear()
                    state["header_value"].clear()

                def on_headers_finished():
                    disposition, parameters = parse_options_header(
                        state["headers"].get(b"content-disposition")
                    )
                    if disposition != b"form-data" or parameters.get(b"name") != b"files":
                        return
                    filename = parameters.get(b"filename")
                    if not filename:
                        return
                    if len(uploads) >= MAX_FILES:
                        raise ValueError("Maksimal 20 file per proses")
                    decoded_name = filename.decode("utf-8", errors="replace")
                    upload = MultipartFile(
                        filename,
                        b"files",
                        config={"MAX_MEMORY_FILE_SIZE": 0, "UPLOAD_KEEP_EXTENSIONS": True},
                    )
                    uploads.append(upload)
                    state["upload"] = upload
                    state["filename"] = decoded_name

                def on_part_data(data, start, end):
                    upload = state["upload"]
                    if upload is not None:
                        upload.write(data[start:end])

                def on_part_end():
                    upload = state["upload"]
                    if upload is None:
                        return
                    upload.finalize()
                    source = upload.file_object.getvalue() if upload.in_memory else Path(upload.file_object.name)
                    files.append((state["filename"], source))
                    state["upload"] = None

                parser = MultipartParser(
                    boundary,
                    callbacks={
                        "on_part_begin": on_part_begin,
                        "on_header_field": on_header_field,
                        "on_header_value": on_header_value,
                        "on_header_end": on_header_end,
                        "on_headers_finished": on_headers_finished,
                        "on_part_data": on_part_data,
                        "on_part_end": on_part_end,
                    },
                    max_size=MAX_EXCEL_REQUEST_BYTES,
                )
                remaining = content_length
                while remaining:
                    chunk = self.rfile.read(min(64 * 1024, remaining))
                    if not chunk:
                        raise ValueError("Upload terputus sebelum selesai")
                    if parser.write(chunk) != len(chunk):
                        raise ValueError("Total upload melebihi batas 500 MB")
                    remaining -= len(chunk)
                parser.finalize()

                with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as output:
                    output_path = Path(output.name)
                compile_workbooks(files, output_path=output_path)
                output_size = output_path.stat().st_size
                self.send_response(200)
                self.send_header(
                    "Content-Type",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                self.send_header("Content-Length", str(output_size))
                self.send_header("Content-Disposition", 'attachment; filename="compiled.xlsx"')
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                with output_path.open("rb") as output:
                    while chunk := output.read(1024 * 1024):
                        self.wfile.write(chunk)
            except ValueError as error:
                self.send_json(400, {"error": str(error)})
            except MultipartParseError:
                self.send_json(400, {"error": "Format upload multipart tidak valid"})
            except (OSError, RuntimeError) as error:
                self.log_error("Excel compile failed: %s", error)
                self.send_json(500, {"error": "Gagal memproses file Excel"})
            finally:
                for upload in uploads:
                    upload.close()
                if output_path is not None:
                    output_path.unlink(missing_ok=True)

    return ExcelAPIHandler


def main():
    token = os.environ.get("EXCEL_API_TOKEN", "")
    if len(token) < 32:
        raise SystemExit("EXCEL_API_TOKEN harus diatur dan memiliki minimal 32 karakter")
    server = ThreadingHTTPServer(("0.0.0.0", 8000), make_handler(token))
    server.serve_forever()


if __name__ == "__main__":
    main()
