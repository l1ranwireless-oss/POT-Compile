import json
import tempfile
import threading
import unittest
from http.server import ThreadingHTTPServer
from urllib.error import HTTPError
from urllib.request import Request, urlopen

from cmdb_server import make_handler


class CMDBServerTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.token = "test-token-" + ("x" * 32)
        self.server = ThreadingHTTPServer(
            ("127.0.0.1", 0), make_handler(self.directory.name, self.token)
        )
        self.base_url = f"http://127.0.0.1:{self.server.server_port}"
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()

    def tearDown(self):
        self.server.shutdown()
        self.server.server_close()
        self.thread.join()
        self.directory.cleanup()

    def request(self, path, method="GET", body=None, filename=None, authorized=True):
        headers = {}
        if authorized:
            headers["Authorization"] = f"Bearer {self.token}"
        if filename:
            headers["X-CMDB-Filename"] = filename
        request = Request(
            self.base_url + path,
            data=body,
            headers=headers,
            method=method,
        )
        return urlopen(request)

    def test_requires_authentication(self):
        with self.assertRaises(HTTPError) as raised:
            self.request("/api/cmdb", authorized=False)
        self.assertEqual(raised.exception.code, 401)
        raised.exception.close()

    def test_save_update_and_load_source(self):
        content = b"CMDB source bytes"
        response = self.request(
            "/api/cmdb/save",
            method="POST",
            body=content,
            filename="cmdb.xlsx",
        )
        self.assertEqual(json.loads(response.read())["size"], len(content))

        with self.assertRaises(HTTPError) as raised:
            self.request(
                "/api/cmdb/save",
                method="POST",
                body=content,
                filename="cmdb.xlsx",
            )
        self.assertEqual(raised.exception.code, 409)
        raised.exception.close()

        replacement = b"Updated CMDB source"
        self.request(
            "/api/cmdb/update",
            method="PUT",
            body=replacement,
            filename="cmdb-latest.xlsx",
        ).close()

        status = json.loads(self.request("/api/cmdb").read())
        self.assertEqual(status["name"], "cmdb-latest.xlsx")
        self.assertEqual(status["size"], len(replacement))
        self.assertEqual(self.request("/api/cmdb/file").read(), replacement)

    def test_rejects_unsupported_extension(self):
        with self.assertRaises(HTTPError) as raised:
            self.request(
                "/api/cmdb/save",
                method="POST",
                body=b"not a CMDB",
                filename="payload.exe",
            )
        self.assertEqual(raised.exception.code, 400)
        raised.exception.close()

    def test_saved_source_survives_server_restart(self):
        content = b"Persistent CMDB"
        self.request(
            "/api/cmdb/save",
            method="POST",
            body=content,
            filename="cmdb.csv",
        ).close()
        self.server.shutdown()
        self.server.server_close()
        self.thread.join()

        self.server = ThreadingHTTPServer(
            ("127.0.0.1", 0), make_handler(self.directory.name, self.token)
        )
        self.base_url = f"http://127.0.0.1:{self.server.server_port}"
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()

        status = json.loads(self.request("/api/cmdb").read())
        self.assertTrue(status["exists"])
        self.assertEqual(self.request("/api/cmdb/file").read(), content)


if __name__ == "__main__":
    unittest.main()
