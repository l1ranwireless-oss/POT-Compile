#!/usr/bin/env python3
import json
import logging
import math
import os
import re
import sqlite3
import tempfile
import threading
import time
import uuid
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote

import psycopg
from psycopg.types.json import Jsonb
from python_multipart.exceptions import MultipartParseError
from python_multipart.multipart import File as MultipartFile
from python_multipart.multipart import MultipartParser, parse_options_header

from excel_compiler import MAX_FILES, MAX_UPLOAD_BYTES as MAX_EXCEL_UPLOAD_BYTES, compile_workbooks

MAX_EXCEL_REQUEST_BYTES = MAX_EXCEL_UPLOAD_BYTES + 2 * 1024 * 1024
MAX_CMDB_UPLOAD_BYTES = 100 * 1024 * 1024
MAX_STAGING_CHUNK_BYTES = 8 * 1024 * 1024
MAX_STAGING_CHUNK_ROWS = 5000
MAX_STAGING_SESSION_BYTES = 256 * 1024 * 1024
CMDB_ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".csv"}
LEGACY_GENERATION_PATTERN = re.compile(r"^[a-f0-9]{32}\.bin$")
STAGING_UUID_PATTERN = r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"


def safe_cmdb_filename(value):
    name = Path(unquote(value).replace("\\", "/")).name
    name = re.sub(r"[^A-Za-z0-9._-]", "_", name)
    suffix = Path(name).suffix.lower()
    if not name or suffix not in CMDB_ALLOWED_EXTENSIONS:
        raise ValueError("Format CMDB harus .xlsx, .xls, atau .csv")
    if len(name) > 255:
        name = name[:255 - len(suffix) - 1] + name[-len(suffix) - 1:]
    return name


def initialize_cmdb_database(database_path):
    path = Path(database_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(path)
    try:
        connection.execute("PRAGMA journal_mode=WAL")
        connection.execute("""
            CREATE TABLE IF NOT EXISTS cmdb_files (
                id INTEGER PRIMARY KEY CHECK (id = 1),
                name TEXT NOT NULL,
                content BLOB NOT NULL,
                size INTEGER NOT NULL CHECK (size > 0),
                updated_at TEXT NOT NULL,
                CHECK (length(content) = size)
            )
        """)
        connection.commit()
        migrate_legacy_cmdb(connection, path.parent)
    finally:
        connection.close()


def initialize_staging_database(database_url):
    with psycopg.connect(database_url, connect_timeout=10) as connection:
        connection.execute("""
            CREATE TABLE IF NOT EXISTS cmdb_staging_sessions (
                session_id UUID PRIMARY KEY,
                created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
                expires_at TIMESTAMPTZ NOT NULL,
                staging_bytes BIGINT NOT NULL DEFAULT 0 CHECK (staging_bytes >= 0)
            )
        """)
        connection.execute("""
            CREATE TABLE IF NOT EXISTS cmdb_staging_chunks (
                session_id UUID NOT NULL REFERENCES cmdb_staging_sessions(session_id) ON DELETE CASCADE,
                chunk_index INTEGER NOT NULL CHECK (chunk_index >= 0),
                rows JSONB NOT NULL,
                row_count INTEGER NOT NULL CHECK (row_count > 0),
                payload_bytes INTEGER NOT NULL CHECK (payload_bytes > 0),
                PRIMARY KEY (session_id, chunk_index)
            )
        """)
        connection.execute(
            "ALTER TABLE cmdb_staging_sessions ADD COLUMN IF NOT EXISTS "
            "staging_bytes BIGINT NOT NULL DEFAULT 0"
        )
        connection.execute(
            "ALTER TABLE cmdb_staging_chunks ADD COLUMN IF NOT EXISTS "
            "payload_bytes INTEGER NOT NULL DEFAULT 1"
        )
        connection.execute(
            "CREATE INDEX IF NOT EXISTS cmdb_staging_sessions_expires_at "
            "ON cmdb_staging_sessions (expires_at)"
        )
        connection.execute(
            "DELETE FROM cmdb_staging_sessions WHERE expires_at <= now()"
        )


def start_staging_cleanup_worker(database_url, interval_seconds=900):
    def cleanup_loop():
        while True:
            time.sleep(interval_seconds)
            try:
                with psycopg.connect(database_url, connect_timeout=10) as connection:
                    result = connection.execute(
                        "DELETE FROM cmdb_staging_sessions WHERE expires_at <= now()"
                    )
                if result.rowcount:
                    logging.info("Removed %s expired PostgreSQL staging sessions", result.rowcount)
            except psycopg.Error:
                logging.exception("Expired PostgreSQL staging cleanup failed")

    threading.Thread(target=cleanup_loop, name="cmdb-staging-cleanup", daemon=True).start()


def parse_staging_rows(value):
    if not isinstance(value, list) or not 1 <= len(value) <= MAX_STAGING_CHUNK_ROWS:
        raise ValueError(f"Chunk harus berisi 1 sampai {MAX_STAGING_CHUNK_ROWS} baris")
    for entry in value:
        if (
            not isinstance(entry, list)
            or len(entry) != 2
            or not isinstance(entry[0], str)
            or not entry[0]
            or len(entry[0]) > 256
            or not isinstance(entry[1], list)
        ):
            raise ValueError("Format baris CMDB staging tidak valid")
        for cell in entry[1]:
            if cell is not None and not isinstance(cell, (str, int, float, bool)):
                raise ValueError("Nilai kolom CMDB staging harus berupa nilai sederhana")
            if isinstance(cell, float) and not math.isfinite(cell):
                raise ValueError("Nilai numerik CMDB staging tidak valid")
    return value


def migrate_legacy_cmdb(connection, data_dir):
    if connection.execute("SELECT 1 FROM cmdb_files WHERE id = 1").fetchone():
        return
    metadata_path = data_dir / "metadata.json"
    if not metadata_path.is_file():
        return
    try:
        metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise RuntimeError("Metadata CMDB lama tidak dapat dibaca") from error
    if not isinstance(metadata, dict):
        raise RuntimeError("Metadata CMDB lama tidak valid")

    generation = metadata.get("file")
    size = metadata.get("size")
    updated_at = metadata.get("updated_at")
    try:
        filename = safe_cmdb_filename(metadata.get("name", ""))
    except (AttributeError, TypeError, ValueError) as error:
        raise RuntimeError("Metadata CMDB lama tidak valid") from error
    if (
        not isinstance(generation, str)
        or not LEGACY_GENERATION_PATTERN.fullmatch(generation)
        or not isinstance(size, int)
        or isinstance(size, bool)
        or not 0 < size <= MAX_CMDB_UPLOAD_BYTES
        or not isinstance(updated_at, str)
    ):
        raise RuntimeError("Metadata CMDB lama tidak valid")

    source_path = data_dir / generation
    if not source_path.is_file() or source_path.stat().st_size != size:
        raise RuntimeError("File CMDB lama tidak ditemukan atau ukurannya tidak cocok")
    connection.execute(
        "INSERT INTO cmdb_files (id, name, content, size, updated_at) "
        "VALUES (1, ?, zeroblob(?), ?, ?)",
        (filename, size, size, updated_at),
    )
    with connection.blobopen("cmdb_files", "content", 1, readonly=False) as blob:
        with source_path.open("rb") as source:
            while chunk := source.read(1024 * 1024):
                blob.write(chunk)
    connection.commit()
    logging.getLogger(__name__).info("Migrated saved CMDB file into SQLite")


def make_handler(database_path, staging_database_url=None, staging_ttl_seconds=86400):
    initialize_cmdb_database(database_path)
    if staging_database_url:
        initialize_staging_database(staging_database_url)
    database_path = Path(database_path)

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

        def do_POST(self):
            if self.path == "/api/excel/compile":
                self.compile_excel()
                return
            if self.path == "/api/cmdb/save":
                self.store_cmdb("save")
                return
            if self.path == "/api/cmdb/staging":
                self.create_staging_session()
                return
            staging_match = re.fullmatch(
                rf"/api/cmdb/staging/({STAGING_UUID_PATTERN})/chunks", self.path
            )
            if staging_match:
                self.store_staging_chunk(staging_match.group(1))
                return
            self.send_json(404, {"error": "Endpoint tidak ditemukan"})

        def do_GET(self):
            if self.path == "/api/health":
                self.get_health()
                return
            if self.path == "/api/cmdb":
                self.get_cmdb_status()
                return
            if self.path == "/api/cmdb/file":
                self.download_cmdb()
                return
            if self.path == "/api/cmdb/staging":
                self.get_staging_status()
                return
            self.send_json(404, {"error": "Endpoint tidak ditemukan"})

        def do_PUT(self):
            if self.path == "/api/cmdb/update":
                self.store_cmdb("update")
                return
            self.send_json(404, {"error": "Endpoint tidak ditemukan"})

        def do_DELETE(self):
            staging_match = re.fullmatch(
                rf"/api/cmdb/staging/({STAGING_UUID_PATTERN})", self.path
            )
            if staging_match:
                self.delete_staging_session(staging_match.group(1))
                return
            self.send_json(404, {"error": "Endpoint tidak ditemukan"})

        def get_staging_status(self):
            self.send_json(200, {
                "enabled": bool(staging_database_url),
                "retention_seconds": staging_ttl_seconds,
            })

        def create_staging_session(self):
            if not staging_database_url:
                self.send_json(503, {"error": "PostgreSQL staging belum dikonfigurasi"})
                return
            session_id = uuid.uuid4()
            try:
                with psycopg.connect(staging_database_url, connect_timeout=10) as connection:
                    connection.execute(
                        "DELETE FROM cmdb_staging_sessions WHERE expires_at <= now()"
                    )
                    connection.execute(
                        "INSERT INTO cmdb_staging_sessions (session_id, expires_at) "
                        "VALUES (%s, now() + (%s * interval '1 second'))",
                        (session_id, staging_ttl_seconds),
                    )
                self.send_json(201, {
                    "session_id": str(session_id),
                    "retention_seconds": staging_ttl_seconds,
                })
            except psycopg.Error as error:
                self.log_error("PostgreSQL staging session creation failed: %s", error)
                self.send_json(500, {"error": "Gagal membuat sesi staging PostgreSQL"})

        def read_staging_json(self):
            if self.headers.get_content_type() != "application/json":
                raise ValueError("Content-Type harus application/json")
            try:
                content_length = int(self.headers.get("Content-Length", "0"))
            except (TypeError, ValueError) as error:
                raise ValueError("Ukuran request staging tidak valid") from error
            if content_length <= 0 or content_length > MAX_STAGING_CHUNK_BYTES:
                raise OverflowError("Ukuran chunk PostgreSQL melebihi batas 8 MiB")
            body = self.rfile.read(content_length)
            if len(body) != content_length:
                raise ValueError("Upload chunk terputus sebelum selesai")
            try:
                return json.loads(body), content_length
            except (UnicodeDecodeError, ValueError) as error:
                raise ValueError("JSON chunk staging tidak valid") from error

        def store_staging_chunk(self, session_id):
            if not staging_database_url:
                self.send_json(503, {"error": "PostgreSQL staging belum dikonfigurasi"})
                return
            try:
                payload, payload_bytes = self.read_staging_json()
                if not isinstance(payload, dict):
                    raise ValueError("Payload chunk staging tidak valid")
                chunk_index = payload.get("chunk_index")
                if (
                    not isinstance(chunk_index, int)
                    or isinstance(chunk_index, bool)
                    or chunk_index < 0
                    or chunk_index > 2_147_483_647
                ):
                    raise ValueError("Nomor chunk staging tidak valid")
                rows = parse_staging_rows(payload.get("rows"))
            except OverflowError as error:
                self.send_json(413, {"error": str(error)})
                return
            except ValueError as error:
                self.send_json(400, {"error": str(error)})
                return

            try:
                session_uuid = uuid.UUID(session_id)
                with psycopg.connect(staging_database_url, connect_timeout=10) as connection:
                    active_session = connection.execute(
                        "SELECT staging_bytes FROM cmdb_staging_sessions "
                        "WHERE session_id = %s AND expires_at > now() FOR UPDATE",
                        (session_uuid,),
                    ).fetchone()
                    if active_session is None:
                        self.send_json(404, {"error": "Sesi staging tidak ditemukan atau kedaluwarsa"})
                        return
                    existing_chunk = connection.execute(
                        "SELECT payload_bytes FROM cmdb_staging_chunks "
                        "WHERE session_id = %s AND chunk_index = %s",
                        (session_uuid, chunk_index),
                    ).fetchone()
                    updated_size = active_session[0] - (existing_chunk[0] if existing_chunk else 0) + payload_bytes
                    if updated_size > MAX_STAGING_SESSION_BYTES:
                        self.send_json(413, {"error": "Sesi staging melebihi batas total 256 MiB"})
                        return
                    connection.execute(
                        "INSERT INTO cmdb_staging_chunks "
                        "(session_id, chunk_index, rows, row_count, payload_bytes) "
                        "VALUES (%s, %s, %s, %s, %s) "
                        "ON CONFLICT (session_id, chunk_index) DO UPDATE "
                        "SET rows = EXCLUDED.rows, row_count = EXCLUDED.row_count, "
                        "payload_bytes = EXCLUDED.payload_bytes",
                        (session_uuid, chunk_index, Jsonb(rows), len(rows), payload_bytes),
                    )
                    connection.execute(
                        "UPDATE cmdb_staging_sessions SET staging_bytes = %s WHERE session_id = %s",
                        (updated_size, session_uuid),
                    )
                self.send_json(201, {
                    "chunk_index": chunk_index,
                    "row_count": len(rows),
                })
            except (ValueError, psycopg.Error) as error:
                self.log_error("PostgreSQL staging chunk failed: %s", error)
                self.send_json(500, {"error": "Gagal menyimpan chunk ke PostgreSQL"})

        def delete_staging_session(self, session_id):
            if not staging_database_url:
                self.send_json(503, {"error": "PostgreSQL staging belum dikonfigurasi"})
                return
            try:
                session_uuid = uuid.UUID(session_id)
                with psycopg.connect(staging_database_url, connect_timeout=10) as connection:
                    connection.execute(
                        "DELETE FROM cmdb_staging_sessions WHERE session_id = %s",
                        (session_uuid,),
                    )
                self.send_json(200, {"deleted": True})
            except (ValueError, psycopg.Error) as error:
                self.log_error("PostgreSQL staging cleanup failed: %s", error)
                self.send_json(500, {"error": "Gagal menghapus sesi staging PostgreSQL"})

        def get_health(self):
            connection = None
            try:
                connection = sqlite3.connect(database_path)
                connection.execute("SELECT 1 FROM cmdb_files LIMIT 1").fetchone()
                self.send_json(200, {"status": "ok"})
            except sqlite3.Error as error:
                self.log_error("API health check failed: %s", error)
                self.send_json(500, {"status": "error", "error": "Database CMDB tidak tersedia"})
            finally:
                if connection is not None:
                    connection.close()

        def get_cmdb_status(self):
            connection = None
            try:
                connection = sqlite3.connect(database_path)
                row = connection.execute(
                    "SELECT name, size, updated_at FROM cmdb_files WHERE id = 1"
                ).fetchone()
                if row is None:
                    self.send_json(200, {"exists": False})
                    return
                self.send_json(200, {
                    "exists": True,
                    "name": row[0],
                    "size": row[1],
                    "updated_at": row[2],
                })
            except sqlite3.Error as error:
                self.log_error("CMDB status query failed: %s", error)
                self.send_json(500, {"error": "Gagal membaca status CMDB"})
            finally:
                if connection is not None:
                    connection.close()

        def download_cmdb(self):
            connection = None
            blob = None
            response_started = False
            try:
                connection = sqlite3.connect(database_path)
                row = connection.execute(
                    "SELECT id, name, size FROM cmdb_files WHERE id = 1"
                ).fetchone()
                if row is None:
                    self.send_json(404, {"error": "CMDB belum tersimpan di server"})
                    return
                _, name, size = row
                blob = connection.blobopen("cmdb_files", "content", row[0], readonly=True)
                self.send_response(200)
                self.send_header("Content-Type", "application/octet-stream")
                self.send_header("Content-Length", str(size))
                self.send_header("Content-Disposition", f'attachment; filename="{name}"')
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                response_started = True
                while chunk := blob.read(1024 * 1024):
                    self.wfile.write(chunk)
            except (OSError, sqlite3.Error) as error:
                self.log_error("CMDB download failed: %s", error)
                if not response_started:
                    self.send_json(500, {"error": "Gagal membaca file CMDB server"})
            finally:
                if blob is not None:
                    blob.close()
                if connection is not None:
                    connection.close()

        def store_cmdb(self, action):
            try:
                filename = safe_cmdb_filename(self.headers.get("X-CMDB-Filename", ""))
                content_length = int(self.headers.get("Content-Length", "0"))
            except (ValueError, TypeError):
                self.send_json(400, {"error": "Nama file atau ukuran upload tidak valid"})
                return
            if content_length <= 0 or content_length > MAX_CMDB_UPLOAD_BYTES:
                self.send_json(413, {"error": "Ukuran CMDB harus antara 1 byte dan 100 MB"})
                return

            temporary_path = None
            connection = None
            try:
                with tempfile.NamedTemporaryFile(delete=False) as temporary:
                    temporary_path = Path(temporary.name)
                    remaining = content_length
                    while remaining:
                        chunk = self.rfile.read(min(1024 * 1024, remaining))
                        if not chunk:
                            raise OSError("Upload CMDB terputus sebelum selesai")
                        temporary.write(chunk)
                        remaining -= len(chunk)

                updated_at = datetime.now(timezone.utc).isoformat(timespec="seconds")
                connection = sqlite3.connect(database_path, timeout=30)
                connection.execute("PRAGMA synchronous=FULL")
                connection.execute("BEGIN IMMEDIATE")
                existing = connection.execute(
                    "SELECT id FROM cmdb_files WHERE id = 1"
                ).fetchone()
                if action == "save" and existing is not None:
                    connection.rollback()
                    self.send_json(409, {"error": "CMDB sudah tersimpan; gunakan Update CMDB"})
                    return
                if action == "update" and existing is None:
                    connection.rollback()
                    self.send_json(404, {"error": "CMDB belum tersimpan; gunakan Simpan CMDB"})
                    return

                if existing is None:
                    connection.execute(
                        "INSERT INTO cmdb_files (id, name, content, size, updated_at) "
                        "VALUES (1, ?, zeroblob(?), ?, ?)",
                        (filename, content_length, content_length, updated_at),
                    )
                else:
                    connection.execute(
                        "UPDATE cmdb_files SET name = ?, content = zeroblob(?), size = ?, updated_at = ? "
                        "WHERE id = 1",
                        (filename, content_length, content_length, updated_at),
                    )
                with connection.blobopen("cmdb_files", "content", 1, readonly=False) as blob:
                    with temporary_path.open("rb") as source:
                        while chunk := source.read(1024 * 1024):
                            blob.write(chunk)
                connection.commit()
                self.send_json(200, {
                    "exists": True,
                    "name": filename,
                    "size": content_length,
                    "updated_at": updated_at,
                })
            except (OSError, sqlite3.Error) as error:
                if connection is not None:
                    connection.rollback()
                self.log_error("CMDB save failed: %s", error)
                self.send_json(500, {"error": "Gagal menyimpan CMDB ke database"})
            finally:
                if connection is not None:
                    connection.close()
                if temporary_path is not None:
                    temporary_path.unlink(missing_ok=True)

        def compile_excel(self):
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
    database_path = os.environ.get("CMDB_DATABASE_PATH", "/data/cmdb.sqlite3")
    port = int(os.environ.get("PORT", "8000"))
    staging_database_url = os.environ.get("DATABASE_URL")
    try:
        staging_ttl_seconds = int(os.environ.get("CMDB_STAGING_TTL_SECONDS", "86400"))
    except ValueError as error:
        raise RuntimeError("CMDB_STAGING_TTL_SECONDS harus berupa bilangan bulat") from error
    if not 60 <= staging_ttl_seconds <= 30 * 24 * 60 * 60:
        raise RuntimeError("CMDB_STAGING_TTL_SECONDS harus antara 60 dan 2592000")
    server = ThreadingHTTPServer(
        ("0.0.0.0", port),
        make_handler(database_path, staging_database_url, staging_ttl_seconds),
    )
    if staging_database_url:
        start_staging_cleanup_worker(staging_database_url)
    server.serve_forever()


if __name__ == "__main__":
    main()
