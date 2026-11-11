import io
import threading
import unittest
from http.server import ThreadingHTTPServer
from urllib.error import HTTPError
from urllib.request import Request, urlopen

from openpyxl import Workbook, load_workbook

from excel_server import make_handler


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
        self.token = "test-token-" + ("x" * 32)
        self.server = ThreadingHTTPServer(
            ("127.0.0.1", 0), make_handler(self.token)
        )
        self.base_url = f"http://127.0.0.1:{self.server.server_port}"
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()

    def tearDown(self):
        self.server.shutdown()
        self.server.server_close()
        self.thread.join()

    def test_requires_authentication(self):
        request = Request(
            self.base_url + "/api/excel/compile",
            data=b"not multipart",
            headers={"Content-Type": "application/octet-stream"},
            method="POST",
        )
        with self.assertRaises(HTTPError) as raised:
            urlopen(request)
        self.assertEqual(raised.exception.code, 401)
        raised.exception.close()

    def test_cmdb_server_endpoints_are_not_available(self):
        requests = [
            Request(self.base_url + "/api/cmdb"),
            Request(self.base_url + "/api/cmdb/save", data=b"source"),
            Request(self.base_url + "/api/cmdb/update", data=b"source", method="PUT"),
        ]
        for request in requests:
            with self.subTest(path=request.full_url, method=request.get_method()):
                with self.assertRaises(HTTPError) as raised:
                    urlopen(request)
                self.assertEqual(raised.exception.code, 404)
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
                "Authorization": f"Bearer {self.token}",
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
                "Authorization": f"Bearer {self.token}",
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
