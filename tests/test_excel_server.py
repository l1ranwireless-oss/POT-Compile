import io
import json
import sqlite3
import tempfile
import threading
import unittest
from http.server import ThreadingHTTPServer
from urllib.error import HTTPError
from urllib.request import Request, urlopen
from pathlib import Path
from unittest.mock import MagicMock, patch

from openpyxl import Workbook, load_workbook

from excel_server import initialize_cmdb_database, make_handler, parse_staging_rows


def workbook_bytes(rows):
    workbook = Workbook()
    worksheet = workbook.active
    for row in rows:
        worksheet.append(row)
    output = io.BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


class ExcelServerTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.database_path = Path(self.directory.name) / "data" / "cmdb.sqlite3"
        self.server = ThreadingHTTPServer(
            ("127.0.0.1", 0), make_handler(self.database_path)
        )
        self.base_url = f"http://127.0.0.1:{self.server.server_port}"
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()

    def tearDown(self):
        self.server.shutdown()
        self.server.server_close()
        self.thread.join()
        self.directory.cleanup()

    def test_excel_compile_validates_request_without_authentication(self):
        request = Request(
            self.base_url + "/api/excel/compile",
            data=b"not multipart",
            headers={"Content-Type": "application/octet-stream"},
            method="POST",
        )
        with self.assertRaises(HTTPError) as raised:
            urlopen(request)
        self.assertEqual(raised.exception.code, 400)
        raised.exception.close()

    def test_cmdb_database_is_created_automatically(self):
        self.assertTrue(self.database_path.is_file())
        with sqlite3.connect(self.database_path) as connection:
            table = connection.execute(
                "SELECT name FROM sqlite_master WHERE type = 'table' AND name = 'cmdb_files'"
            ).fetchone()
        self.assertEqual(table, ("cmdb_files",))

    def test_api_health_reports_python_and_cmdb_database_ready(self):
        with urlopen(self.base_url + "/api/health") as response:
            self.assertEqual(response.status, 200)
            self.assertEqual(json.loads(response.read()), {
                "status": "ok",
                "cmdb_storage": "sqlite",
            })

    def test_news_feed_returns_attributed_https_headlines(self):
        from excel_server import fetch_news_headlines

        feed = b"""<?xml version="1.0"?>
        <rss><channel>
          <item><title>Telecom update</title><link>https://publisher.example/one</link>
            <source>Publisher One</source><pubDate>Wed, 11 Nov 2026 20:00:00 GMT</pubDate></item>
          <item><title>Skip insecure link</title><link>http://publisher.example/two</link>
            <source>Publisher Two</source></item>
          <item><title>Finance update</title><link>https://publisher.example/three</link>
            <source>Publisher Three</source></item>
        </channel></rss>"""
        with patch("excel_server._news_cache", {"expires_at": 0, "items": None}), patch(
            "excel_server.urllib.request.urlopen", return_value=io.BytesIO(feed)
        ) as open_feed:
            items = fetch_news_headlines()
        self.assertEqual([item["title"] for item in items], ["Telecom update", "Finance update"])
        self.assertEqual(items[0]["source"], "Publisher One")
        self.assertTrue(all(item["url"].startswith("https://") for item in items))
        self.assertEqual(open_feed.call_args.kwargs["timeout"], 8)

    def test_news_endpoint_reports_upstream_failure(self):
        from excel_server import make_handler

        with patch("excel_server.fetch_news_headlines", side_effect=RuntimeError("Feed offline")):
            handler = make_handler(self.database_path)
            server = ThreadingHTTPServer(("127.0.0.1", 0), handler)
            thread = threading.Thread(target=server.serve_forever, daemon=True)
            thread.start()
            try:
                with self.assertRaises(HTTPError) as raised:
                    urlopen(f"http://127.0.0.1:{server.server_port}/api/briefing/news")
                self.assertEqual(raised.exception.code, 502)
                self.assertEqual(json.loads(raised.exception.read()), {"error": "Feed offline"})
                raised.exception.close()
            finally:
                server.shutdown()
                server.server_close()
                thread.join()

    def test_api_health_confirms_postgres_connection(self):
        from excel_server import make_handler

        with patch("excel_server.initialize_postgres_cmdb"), patch(
            "excel_server.psycopg.connect"
        ) as connect:
            connect.return_value.__enter__.return_value = MagicMock()
            handler = make_handler(
                self.database_path,
                cmdb_database_url="postgresql://test",
            )
            server = ThreadingHTTPServer(("127.0.0.1", 0), handler)
            thread = threading.Thread(target=server.serve_forever, daemon=True)
            thread.start()
            try:
                with urlopen(f"http://127.0.0.1:{server.server_port}/api/health") as response:
                    result = json.loads(response.read())
                self.assertEqual(result, {
                    "status": "ok",
                    "cmdb_storage": "postgresql",
                    "postgres": "ok",
                })
                connect.assert_called_with("postgresql://test", connect_timeout=5)
            finally:
                server.shutdown()
                server.server_close()
                thread.join()

    def test_cmdb_save_in_postgres_uses_insert_without_overwriting_existing_data(self):
        from excel_server import store_cmdb_in_postgres

        connection = MagicMock()
        connection.execute.return_value.fetchone.return_value = (1,)
        with patch("excel_server.psycopg.connect") as connect:
            connect.return_value.__enter__.return_value = connection
            status, result = store_cmdb_in_postgres(
                "postgresql://test", "save", "cmdb.xlsx", b"CMDB", "2026-11-11T00:00:00+00:00"
            )
        self.assertEqual(status, 200)
        self.assertEqual(result["size"], 4)
        self.assertIn("ON CONFLICT (id) DO NOTHING", connection.execute.call_args.args[0])

    def test_cmdb_update_in_postgres_requires_existing_file(self):
        from excel_server import store_cmdb_in_postgres

        connection = MagicMock()
        connection.execute.return_value.fetchone.return_value = None
        with patch("excel_server.psycopg.connect") as connect:
            connect.return_value.__enter__.return_value = connection
            status, result = store_cmdb_in_postgres(
                "postgresql://test", "update", "cmdb.xlsx", b"CMDB", "2026-11-11T00:00:00+00:00"
            )
        self.assertEqual(status, 404)
        self.assertIn("gunakan Simpan CMDB", result["error"])

    def test_existing_sqlite_cmdb_migrates_to_empty_postgres_storage(self):
        from excel_server import initialize_postgres_cmdb

        content = b"legacy CMDB"
        with sqlite3.connect(self.database_path) as connection:
            connection.execute(
                "INSERT INTO cmdb_files (id, name, content, size, updated_at) "
                "VALUES (1, ?, zeroblob(?), ?, ?)",
                ("legacy.xlsx", len(content), len(content), "2026-11-11T00:00:00+00:00"),
            )
            with connection.blobopen("cmdb_files", "content", 1, readonly=False) as blob:
                blob.write(content)

        postgres = MagicMock()
        postgres.execute.return_value.fetchone.return_value = None
        with patch("excel_server.psycopg.connect") as connect:
            connect.return_value.__enter__.return_value = postgres
            initialize_postgres_cmdb("postgresql://test", self.database_path)
        migration = postgres.execute.call_args_list[-1]
        self.assertIn("INSERT INTO cmdb_source_files", migration.args[0])
        self.assertEqual(migration.args[1], ("legacy.xlsx", content, len(content), "2026-11-11T00:00:00+00:00"))

    def test_cmdb_status_reports_postgres_as_active_storage(self):
        from excel_server import make_handler

        postgres = MagicMock()
        postgres.execute.return_value.fetchone.return_value = (
            "cmdb.xlsx", 4, "2026-11-11T00:00:00+00:00"
        )
        with patch("excel_server.initialize_postgres_cmdb"), patch(
            "excel_server.psycopg.connect"
        ) as connect:
            connect.return_value.__enter__.return_value = postgres
            handler = make_handler(
                self.database_path,
                cmdb_database_url="postgresql://test",
            )
            server = ThreadingHTTPServer(("127.0.0.1", 0), handler)
            thread = threading.Thread(target=server.serve_forever, daemon=True)
            thread.start()
            try:
                with urlopen(f"http://127.0.0.1:{server.server_port}/api/cmdb") as response:
                    result = json.loads(response.read())
                self.assertEqual(result["storage"], "postgresql")
                self.assertEqual(result["name"], "cmdb.xlsx")
            finally:
                server.shutdown()
                server.server_close()
                thread.join()

    def test_postgres_staging_reports_disabled_without_database_url(self):
        with urlopen(self.base_url + "/api/cmdb/staging") as response:
            self.assertEqual(json.loads(response.read()), {
                "enabled": False,
                "retention_seconds": 86400,
            })

    def test_postgres_staging_creation_requires_database_url(self):
        request = Request(self.base_url + "/api/cmdb/staging", data=b"", method="POST")
        with self.assertRaises(HTTPError) as raised:
            urlopen(request)
        self.assertEqual(raised.exception.code, 503)
        self.assertEqual(
            json.loads(raised.exception.read()),
            {"error": "PostgreSQL staging belum dikonfigurasi"},
        )
        raised.exception.close()

    def test_postgres_staging_rows_validate_shape_and_cell_types(self):
        rows = [["SITE-1", ["Name", 1, None, True]]]
        self.assertEqual(parse_staging_rows(rows), rows)
        for invalid_rows in ([], [["", ["value"]]], [["SITE-1", [{"nested": "value"}]]]):
            with self.subTest(rows=invalid_rows), self.assertRaises(ValueError):
                parse_staging_rows(invalid_rows)

    def test_cmdb_database_migrates_existing_server_file(self):
        content = b"Existing server CMDB"
        generation = "a" * 32 + ".bin"
        (self.database_path.parent / generation).write_bytes(content)
        (self.database_path.parent / "metadata.json").write_text(json.dumps({
            "file": generation,
            "name": "legacy.xlsx",
            "size": len(content),
            "updated_at": "2026-11-11T00:00:00+00:00",
        }), encoding="utf-8")

        initialize_cmdb_database(self.database_path)
        with urlopen(self.base_url + "/api/cmdb/file") as response:
            self.assertEqual(response.headers.get_filename(), "legacy.xlsx")
            self.assertEqual(response.read(), content)

    def test_cmdb_save_update_and_download_without_token(self):
        original = b"CMDB source bytes"
        request = Request(
            self.base_url + "/api/cmdb/save",
            data=original,
            headers={"X-CMDB-Filename": "cmdb.xlsx"},
            method="POST",
        )
        with urlopen(request) as response:
            saved = json.loads(response.read())
        self.assertEqual(saved["size"], len(original))

        with urlopen(self.base_url + "/api/cmdb") as response:
            status = json.loads(response.read())
        self.assertTrue(status["exists"])
        self.assertEqual(status["name"], "cmdb.xlsx")
        self.assertEqual(status["size"], len(original))

        duplicate = Request(
            self.base_url + "/api/cmdb/save",
            data=original,
            headers={"X-CMDB-Filename": "cmdb.xlsx"},
            method="POST",
        )
        with self.assertRaises(HTTPError) as raised:
            urlopen(duplicate)
        self.assertEqual(raised.exception.code, 409)
        raised.exception.close()

        replacement = b"Updated CMDB source"
        update = Request(
            self.base_url + "/api/cmdb/update",
            data=replacement,
            headers={"X-CMDB-Filename": "cmdb-latest.xlsx"},
            method="PUT",
        )
        with urlopen(update) as response:
            updated = json.loads(response.read())
        self.assertEqual(updated["name"], "cmdb-latest.xlsx")
        self.assertEqual(updated["size"], len(replacement))

        with urlopen(self.base_url + "/api/cmdb/file") as response:
            self.assertEqual(response.headers.get_filename(), "cmdb-latest.xlsx")
            self.assertEqual(response.read(), replacement)

    def test_cmdb_update_requires_initial_save(self):
        request = Request(
            self.base_url + "/api/cmdb/update",
            data=b"CMDB",
            headers={"X-CMDB-Filename": "cmdb.csv"},
            method="PUT",
        )
        with self.assertRaises(HTTPError) as raised:
            urlopen(request)
        self.assertEqual(raised.exception.code, 404)
        raised.exception.close()

    def test_cmdb_source_survives_server_restart(self):
        content = b"Persistent CMDB\n" + b"x" * (2 * 1024 * 1024)
        request = Request(
            self.base_url + "/api/cmdb/save",
            data=content,
            headers={"X-CMDB-Filename": "cmdb.csv"},
            method="POST",
        )
        with urlopen(request) as response:
            response.close()
        self.server.shutdown()
        self.server.server_close()
        self.thread.join()
        self.server = ThreadingHTTPServer(
            ("127.0.0.1", 0), make_handler(self.database_path)
        )
        self.base_url = f"http://127.0.0.1:{self.server.server_port}"
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()

        with urlopen(self.base_url + "/api/cmdb/file") as response:
            self.assertEqual(response.read(), content)

    def test_cmdb_rejects_unsupported_extensions_and_oversized_uploads(self):
        invalid_file = Request(
            self.base_url + "/api/cmdb/save",
            data=b"not a CMDB",
            headers={"X-CMDB-Filename": "payload.exe"},
            method="POST",
        )
        with self.assertRaises(HTTPError) as raised:
            urlopen(invalid_file)
        self.assertEqual(raised.exception.code, 400)
        raised.exception.close()

        oversized = Request(
            self.base_url + "/api/cmdb/save",
            data=b"x",
            headers={
                "X-CMDB-Filename": "cmdb.csv",
                "Content-Length": str(100 * 1024 * 1024 + 1),
            },
            method="POST",
        )
        with self.assertRaises(HTTPError) as raised:
            urlopen(oversized, timeout=2)
        self.assertEqual(raised.exception.code, 413)
        raised.exception.close()

    def test_excel_compile_returns_downloadable_workbook(self):
        boundary = "excel-compiler-test-boundary"
        files = [
            ("template.xlsx", workbook_bytes([["Name", "Amount"], ["Alpha", 1]])),
            ("extra.csv", b"amount,name\n2,Beta\n"),
        ]
        body = b"".join(
            b"--" + boundary.encode() + b"\r\n"
            + f'Content-Disposition: form-data; name="files"; filename="{name}"\r\n'.encode()
            + b"Content-Type: application/octet-stream\r\n\r\n"
            + content + b"\r\n"
            for name, content in files
        ) + b"--" + boundary.encode() + b"--\r\n"
        request = Request(
            self.base_url + "/api/excel/compile",
            data=body,
            headers={
                "Content-Type": f"multipart/form-data; boundary={boundary}",
            },
            method="POST",
        )

        with urlopen(request) as response:
            self.assertEqual(
                response.headers.get_content_type(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            workbook = load_workbook(io.BytesIO(response.read()), read_only=True)
        self.assertEqual(workbook.sheetnames, ["Sheet"])
        self.assertEqual(list(workbook.active.values), [
            ("Name", "Amount"),
            ("Alpha", 1),
            ("Beta", "2"),
        ])
        workbook.close()

    def test_excel_compile_spools_uploads_larger_than_one_mib(self):
        boundary = "excel-compiler-large-test-boundary"
        template = workbook_bytes([["Name"], ["Template"]])
        csv_content = b"Name\n" + (b"Added-" + b"x" * 1024 + b"\n") * 1_100
        self.assertGreater(len(csv_content), 1024 * 1024)
        files = [("template.xlsx", template), ("large.csv", csv_content)]
        body = b"".join(
            b"--" + boundary.encode() + b"\r\n"
            + f'Content-Disposition: form-data; name="files"; filename="{name}"\r\n'.encode()
            + b"Content-Type: application/octet-stream\r\n\r\n"
            + content + b"\r\n"
            for name, content in files
        ) + b"--" + boundary.encode() + b"--\r\n"
        request = Request(
            self.base_url + "/api/excel/compile",
            data=body,
            headers={
                "Content-Type": f"multipart/form-data; boundary={boundary}",
            },
            method="POST",
        )

        with urlopen(request) as response:
            workbook = load_workbook(io.BytesIO(response.read()), read_only=True)
        self.assertEqual(sum(1 for _ in workbook.active.iter_rows()), 1_102)
        workbook.close()


if __name__ == "__main__":
    unittest.main()
