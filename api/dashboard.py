from __future__ import annotations

import json
import sys
import traceback
from http.server import BaseHTTPRequestHandler
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

IMPORT_ERROR = None
try:
    from app import DashboardHandler as DashboardBase
except Exception as exc:  # pragma: no cover - deployed diagnostics
    IMPORT_ERROR = exc
    DashboardBase = BaseHTTPRequestHandler


class handler(DashboardBase):
    def _send_deploy_error(self, exc: BaseException) -> None:
        body = json.dumps(
            {
                "ok": False,
                "error": str(exc),
                "traceback": traceback.format_exc(),
            },
            ensure_ascii=False,
        ).encode("utf-8")
        self.send_response(500)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:
        if IMPORT_ERROR:
            self._send_deploy_error(IMPORT_ERROR)
            return
        try:
            super().do_GET()
        except Exception as exc:  # pragma: no cover - deployed diagnostics
            self._send_deploy_error(exc)

    def do_POST(self) -> None:
        if IMPORT_ERROR:
            self._send_deploy_error(IMPORT_ERROR)
            return
        try:
            super().do_POST()
        except Exception as exc:  # pragma: no cover - deployed diagnostics
            self._send_deploy_error(exc)
