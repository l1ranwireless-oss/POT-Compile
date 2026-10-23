#!/usr/bin/env python3
import hmac
import json
import os
import re
import tempfile
import threading
import uuid
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote

MAX_UPLOAD_BYTES = 100 * 1024 * 1024
ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".csv"}
GENERATION_PATTERN = re.compile(r"^[a-f0-9]{32}\.bin$")


def safe_filename(value):
    name = Path(unquote(value).replace("\\", "/")).name
    name = re.sub(r"[^A-Za-z0-9._-]", "_", name)
    if not name or Path(name).suffix.lower() not in ALLOWED_EXTENSIONS:
        raise ValueError("Format CMDB harus .xlsx, .xls, atau .csv")
    return name


def make_handler(data_dir, token):
    root = Path(data_dir)
    root.mkdir(parents=True, exist_ok=True)
    lock = threading.Lock()

    class CMDBHandler(BaseHTTPRequestHandler):
        server_version = "POT-CMDB"

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
                self.send_json(401, {"error": "Token CMDB tidak valid"})
                return False
            return True

        def read_metadata(self):
            metadata_path = root / "metadata.json"
            if not metadata_path.is_file():
                return None
            try:
                metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                raise RuntimeError("Metadata CMDB server tidak dapat dibaca")
            if not GENERATION_PATTERN.fullmatch(metadata.get("file", "")):
                raise RuntimeError("Metadata CMDB server tidak valid")
            if not (root / metadata["file"]).is_file():
                raise RuntimeError("File CMDB server tidak ditemukan")
            return metadata

        def do_GET(self):
            if self.path not in {"/api/cmdb", "/api/cmdb/file"}:
                self.send_json(404, {"error": "Endpoint tidak ditemukan"})
                return
            if not self.authorize():
                return
            try:
                with lock:
                    metadata = self.read_metadata()
                if self.path == "/api/cmdb":
                    if metadata is None:
                        self.send_json(200, {"exists": False})
                    else:
                        self.send_json(200, {
                            "exists": True,
                            "name": metadata["name"],
                            "size": metadata["size"],
                            "updated_at": metadata["updated_at"],
                        })
                    return
                if metadata is None:
                    self.send_json(404, {"error": "CMDB belum tersimpan di server"})
                    return
                source = root / metadata["file"]
                if not source.is_file():
                    self.send_json(500, {"error": "File CMDB server tidak ditemukan"})
                    return
                size = source.stat().st_size
                self.send_response(200)
                self.send_header("Content-Type", "application/octet-stream")
                self.send_header("Content-Length", str(size))
                self.send_header("Content-Disposition", f'attachment; filename="{metadata["name"]}"')
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                with source.open("rb") as stream:
                    while chunk := stream.read(1024 * 1024):
                        self.wfile.write(chunk)
            except RuntimeError as error:
                self.send_json(500, {"error": str(error)})
            except OSError:
                self.send_json(500, {"error": "Gagal membaca file CMDB server"})

        def do_POST(self):
            self.store_source("save")

        def do_PUT(self):
            self.store_source("update")

        def store_source(self, action):
            expected_path = "/api/cmdb/save" if action == "save" else "/api/cmdb/update"
            if self.path != expected_path:
                self.send_json(404, {"error": "Endpoint tidak ditemukan"})
                return
            if not self.authorize():
                return
            try:
                filename = safe_filename(self.headers.get("X-CMDB-Filename", ""))
                content_length = int(self.headers.get("Content-Length", "0"))
            except (ValueError, TypeError):
                self.send_json(400, {"error": "Nama file atau ukuran upload tidak valid"})
                return
            if content_length <= 0 or content_length > MAX_UPLOAD_BYTES:
                self.send_json(413, {"error": "Ukuran CMDB harus antara 1 byte dan 100 MB"})
                return

            temporary_path = None
            metadata_temp = None
            stored_path = None
            committed = False
            try:
                with lock:
                    previous = self.read_metadata()
                    if action == "save" and previous is not None:
                        self.send_json(409, {"error": "CMDB sudah tersimpan; gunakan Update CMDB"})
                        return
                    if action == "update" and previous is None:
                        self.send_json(404, {"error": "CMDB belum tersimpan; gunakan Simpan CMDB"})
                        return

                    with tempfile.NamedTemporaryFile(dir=root, prefix=".upload-", delete=False) as temporary:
                        temporary_path = Path(temporary.name)
                        remaining = content_length
                        while remaining:
                            chunk = self.rfile.read(min(1024 * 1024, remaining))
                            if not chunk:
                                raise OSError("Upload CMDB terputus")
                            temporary.write(chunk)
                            remaining -= len(chunk)
                        temporary.flush()
                        os.fsync(temporary.fileno())

                    generation = f"{uuid.uuid4().hex}.bin"
                    stored_path = root / generation
                    os.replace(temporary_path, stored_path)
                    temporary_path = None
                    metadata = {
                        "file": generation,
                        "name": filename,
                        "size": content_length,
                        "updated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                    }
                    with tempfile.NamedTemporaryFile(
                        mode="w", encoding="utf-8", dir=root, prefix=".metadata-", delete=False
                    ) as temporary:
                        metadata_temp = Path(temporary.name)
                        json.dump(metadata, temporary)
                        temporary.flush()
                        os.fsync(temporary.fileno())
                    os.replace(metadata_temp, root / "metadata.json")
                    metadata_temp = None
                    committed = True
                    if previous is not None:
                        old_path = root / previous["file"]
                        if old_path != stored_path:
                            try:
                                old_path.unlink(missing_ok=True)
                            except OSError as error:
                                self.log_error("Could not remove replaced CMDB generation: %s", error)
                self.send_json(200, {
                    "exists": True,
                    "name": metadata["name"],
                    "size": metadata["size"],
                    "updated_at": metadata["updated_at"],
                })
            except (OSError, RuntimeError) as error:
                self.send_json(500, {"error": f"Gagal menyimpan CMDB: {error}"})
            finally:
                if temporary_path is not None:
                    temporary_path.unlink(missing_ok=True)
                if metadata_temp is not None:
                    metadata_temp.unlink(missing_ok=True)
                if stored_path is not None and not committed:
                    stored_path.unlink(missing_ok=True)

    return CMDBHandler


def main():
    token = os.environ.get("CMDB_API_TOKEN", "")
    if len(token) < 32:
        raise SystemExit("CMDB_API_TOKEN harus diatur dan memiliki minimal 32 karakter")
    data_dir = os.environ.get("CMDB_DATA_DIR", "/data/cmdb")
    server = ThreadingHTTPServer(("0.0.0.0", 8000), make_handler(data_dir, token))
    server.serve_forever()


if __name__ == "__main__":
    main()
