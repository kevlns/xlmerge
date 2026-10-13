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
from prepare_progress import preparation_context, PreparationCancelled, checkpoint


def _payload() -> dict:
    """Read raw JSON without Bottle's 100 KiB ``request.json`` form-memory limit."""
    if not request.content_length:
        return {}
    data = request.body.read()
    if not data:
        return {}
    parsed = json.loads(data.decode("utf-8"))
    return parsed if isinstance(parsed, dict) else {}


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


def serve(manifest_path: Path | str, *, open_browser: bool = True, port: int = 0,
          prepare_repo=None, relative_path=None) -> bool:
    manifest_path = Path(manifest_path).resolve()
    manifest = None if prepare_repo else load_manifest(manifest_path)
    if manifest is not None and not manifest["files"]:
        print(
            "Manifest has no conflicts requiring review; use resolve/apply for auto-theirs entries.",
            file=sys.stderr,
        )
        return False
    html_path = Path(__file__).with_name("merge_ui.html")
    app = Bottle()
    state = {"resolved": False, "server": None}
    lock = threading.Lock()
    cancelled = threading.Event()
    status = {'state': 'preparing' if prepare_repo else 'ready', 'phase': '开始解析',
              'file': '', 'fileDone': 0, 'fileTotal': 0, 'sheet': '', 'sheetDone': 0, 'sheetTotal': 0, 'side': ''}

    def report(progress):
        with lock:
            status.update(progress)

    def prepare():
        nonlocal manifest
        try:
            from xlsx_conflict import prepare_conflicts
            with preparation_context(report, cancelled):
                path = prepare_conflicts(prepare_repo, relative_path=relative_path, runtime_dir=manifest_path.parent)
                checkpoint()
                prepared = load_manifest(path)
                with lock:
                    if cancelled.is_set():
                        raise PreparationCancelled()
                    manifest = prepared
                    status.update(state='ready', phase='可开始审阅')
        except PreparationCancelled:
            report({'state': 'cancelled', 'phase': '已中断，工作树未修改'})
        except Exception as exc:
            report({'state': 'failed', 'error': str(exc), 'phase': '解析失败'})

    worker = threading.Thread(target=prepare, name='xlmerge-prepare') if prepare_repo else None

    @app.get("/")
    def index():
        response.content_type = "text/html; charset=utf-8"
        with lock:
            ready = status['state'] == 'ready'
        return (html_path if ready else html_path.with_name('loading_ui.html')).read_text(encoding="utf-8")

    @app.get('/api/status')
    def progress():
        with lock:
            return _json(dict(status))

    @app.get("/api/diff")
    def diff():
        if manifest is None:
            return _json({'error': '解析尚未完成'}, 409)
        return _json(_public_manifest(manifest))

    @app.post("/api/commit-preview")
    def commit_preview():
        """只读：按当前 decisions 生成结构化 commit 信息，不执行任何写操作。"""
        payload = _payload()
        if manifest is None:
            return _json({'error': '解析尚未完成'}, 409)
        try:
            message = _build_commit_message(manifest, payload.get("decisions") or {})
            return _json({"message": message})
        except (ConflictError, OSError, ValueError, KeyError) as exc:
            return _json({"success": False, "error": str(exc)}, 400)

    @app.post("/api/resolve")
    def resolve():
        payload = _payload()
        if manifest is None:
            return _json({'error': '解析尚未完成'}, 409)
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

    @app.error(413)
    def too_large(error):
        return _json({"success": False, "error": "request body too large"}, 413)

    @app.post("/api/shutdown")
    def shutdown():
        with lock:
            if status['state'] == 'preparing':
                cancelled.set()
                status['state'] = 'cancelling'
        def stop():
            if worker:
                worker.join()
            state['server'].shutdown()
        threading.Thread(target=stop, daemon=True).start()
        return _json({"success": True})

    server = make_server("127.0.0.1", port, app)
    state["server"] = server
    if worker:
        worker.start()
    url = f"http://127.0.0.1:{server.server_port}"
    print(f"MERGE_SERVER_URL={url}", flush=True)
    if open_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    finally:
        server.server_close()
    return bool(state["resolved"])
