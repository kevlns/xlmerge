"""Local-only HTTP server for the sheet-aware XLSX resolver UI."""

from __future__ import annotations

import json
import sys
import threading
import webbrowser
from pathlib import Path
from wsgiref.simple_server import make_server

from bottle import Bottle, request, response

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from xlsx_conflict import ConflictError, _build_commit_message, apply_manifest, load_manifest


def _json(data, status: int = 200):
    response.status = status
    response.content_type = "application/json; charset=utf-8"
    return json.dumps(data, ensure_ascii=False)


def _public_manifest(manifest):
    return {
        "summary": manifest["summary"],
        "files": [
            {
                "id": item["id"],
                "path": item["path"],
                "oursAuthor": item.get("oursAuthor", ""),
                "theirsAuthor": item.get("theirsAuthor", ""),
                "oursDate": item.get("oursDate", ""),
                "theirsDate": item.get("theirsDate", ""),
                "macro": item.get("macro"),
                "diff": item["diff"],
            }
            for item in manifest["files"]
        ],
    }


def serve(manifest_path: Path | str, *, open_browser: bool = True) -> bool:
    manifest_path = Path(manifest_path).resolve()
    manifest = load_manifest(manifest_path)
    html_path = Path(__file__).with_name("merge_ui.html")
    app = Bottle()
    state = {"resolved": False, "server": None}

    @app.get("/")
    def index():
        response.content_type = "text/html; charset=utf-8"
        return html_path.read_text(encoding="utf-8")

    @app.get("/api/diff")
    def diff():
        return _json(_public_manifest(manifest))

    @app.post("/api/commit-preview")
    def commit_preview():
        """只读：按当前 decisions 生成结构化 commit 信息，不执行任何写操作。"""
        payload = request.json or {}
        try:
            message = _build_commit_message(manifest, payload.get("decisions") or {})
            return _json({"message": message})
        except (ConflictError, OSError, ValueError, KeyError) as exc:
            return _json({"success": False, "error": str(exc)}, 400)

    @app.post("/api/resolve")
    def resolve():
        payload = request.json or {}
        try:
            result = apply_manifest(
                manifest_path,
                payload.get("decisions") or {},
                commit=bool(payload.get("commit", True)),
                push=bool(payload.get("push", False)),
                message=(payload.get("message") or "").strip() or None,
            )
            if result["success"]:
                state["resolved"] = True
                threading.Timer(2.0, state["server"].shutdown).start()
            return _json(result, 200 if result["success"] else 500)
        except (ConflictError, OSError, ValueError, KeyError) as exc:
            return _json({"success": False, "error": str(exc)}, 400)

    @app.post("/api/shutdown")
    def shutdown():
        threading.Thread(target=state["server"].shutdown, daemon=True).start()
        return _json({"success": True})

    server = make_server("127.0.0.1", 0, app)
    state["server"] = server
    url = f"http://127.0.0.1:{server.server_port}"
    print(f"MERGE_SERVER_URL={url}", flush=True)
    if open_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    finally:
        server.server_close()
    return bool(state["resolved"])
