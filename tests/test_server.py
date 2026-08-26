import importlib
import os
import tempfile
import unittest
import uuid
from pathlib import Path


class TelemetryServerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.TemporaryDirectory()
        os.environ["MULTIDOCSYNC_DB"] = str(Path(cls.temp_dir.name) / "telemetry.sqlite3")
        os.environ["MULTIDOCSYNC_TELEMETRY_SALT"] = "unit-test-salt-not-for-production"
        os.environ["MULTIDOCSYNC_ADMIN_TOKEN"] = "unit-test-admin-not-for-production"
        module = importlib.import_module("server.app")
        cls.client = module.app.test_client()

    @classmethod
    def tearDownClass(cls):
        cls.temp_dir.cleanup()

    def test_health(self):
        response = self.client.get("/healthz")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json["status"], "ok")

    def test_rejects_invalid_payload(self):
        response = self.client.post("/v1/events", json={"event": "session_start"})
        self.assertEqual(response.status_code, 400)

    def test_accepts_event_and_reports_summary(self):
        payload = {
            "event": "session_start",
            "install_id": str(uuid.uuid4()),
            "session_id": str(uuid.uuid4()),
            "version": "0.1.0",
            "platform": "windows",
            "architecture": "amd64",
            "elapsed_seconds": 0,
        }
        response = self.client.post("/v1/events", json=payload)
        self.assertEqual(response.status_code, 202)
        self.assertEqual(self.client.get("/v1/admin/summary").status_code, 401)
        summary = self.client.get(
            "/v1/admin/summary",
            headers={"Authorization": "Bearer unit-test-admin-not-for-production"},
        )
        self.assertEqual(summary.status_code, 200)
        self.assertGreaterEqual(summary.json["users"], 1)


if __name__ == "__main__":
    unittest.main()
