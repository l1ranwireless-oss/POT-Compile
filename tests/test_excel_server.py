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

from openpyxl import Workbook, load_workbook

from excel_server import initialize_cmdb_database, make_handler


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
