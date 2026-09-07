"""CLI entry point for repository-local XLSX/XLSM conflict resolution."""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
import time
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from xlsx_conflict import (
    ConflictError,
    add_auto_theirs_filter,
    apply_manifest,
    auto_theirs_path,
    detect_conflicts,
    find_repo_root,
    load_manifest,
    load_auto_theirs_filters,
    prepare_conflicts,
)


SCRIPT_PATH = Path(__file__).resolve()
PACKAGE_ROOT = SCRIPT_DIR.parent


def _configure_stdio() -> None:
    """Keep structured CLI output UTF-8 even under legacy Windows code pages."""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            reconfigure(encoding="utf-8", errors="replace")


def _print(data) -> None:
    print(json.dumps(data, ensure_ascii=False, indent=2))


def _launch_background(repo: Path, manifest: Path, *, open_browser: bool) -> dict:
    """Start the resolver as a detached child and return its URL promptly."""
    runtime_dir = manifest.parent
    stamp = int(time.time() * 1000)
    stdout_path = runtime_dir / f"resolver-{stamp}.stdout.log"
    stderr_path = runtime_dir / f"resolver-{stamp}.stderr.log"
    command = [
        sys.executable,
        str(SCRIPT_PATH),
        "--repo",
        str(repo),
        "serve",
        "--manifest",
        str(manifest),
    ]
    if not open_browser:
        command.append("--no-browser")
    creationflags = (
        getattr(subprocess, "CREATE_NO_WINDOW", 0)
        | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    )
    with stdout_path.open("w", encoding="utf-8") as stdout_handle, stderr_path.open("w", encoding="utf-8") as stderr_handle:
        process = subprocess.Popen(
            command,
            cwd=str(PACKAGE_ROOT),
            stdout=stdout_handle,
            stderr=stderr_handle,
            stdin=subprocess.DEVNULL,
            creationflags=creationflags,
            close_fds=True,
        )

    deadline = time.monotonic() + 15
    url = ""
    while time.monotonic() < deadline:
        output = stdout_path.read_text(encoding="utf-8", errors="replace") if stdout_path.exists() else ""
        for line in output.splitlines():
            if line.startswith("MERGE_SERVER_URL="):
                url = line.split("=", 1)[1].strip()
                break
        if url or process.poll() is not None:
            break
        time.sleep(0.1)

    if not url:
        if process.poll() is None:
            process.terminate()
        stderr = stderr_path.read_text(encoding="utf-8", errors="replace") if stderr_path.exists() else ""
        raise ConflictError(f"Resolver failed to start: {stderr.strip() or 'URL was not reported within 15 seconds'}")
    return {
        "ok": True,
        "pid": process.pid,
        "url": url,
        "manifest": str(manifest),
        "stdout": str(stdout_path),
        "stderr": str(stderr_path),
    }


def _auto_only(manifest: Path) -> dict | None:
    """Resolve an all-auto-theirs manifest immediately instead of starting an empty UI."""
    if load_manifest(manifest)["files"]:
        return None
    result = apply_manifest(manifest, {"files": {}})
    return {
        "ok": result["success"],
        "url": None,
        "manifest": str(manifest),
        "autoResolved": result.get("autoTheirs", []),
        "steps": result["steps"],
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Resolve Git XLSX/XLSM conflicts in a Git repository.")
    parser.add_argument("--repo", default=".", help="Git repository (default: %(default)s; the current directory).")
    parser.add_argument("--runtime-dir", help="Optional directory for extracted Git stage workbooks and manifest.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    subparsers.add_parser("detect", help="List unresolved .xlsx/.xlsm files.")

    filter_parser = subparsers.add_parser("filter", help="Manage repository-scoped auto-theirs workbook filters.")
    filter_subparsers = filter_parser.add_subparsers(dest="filter_command", required=True)
    filter_add = filter_subparsers.add_parser("add", help="Add one case-insensitive workbook path-suffix filter.")
    filter_add.add_argument("path", help="Relative .xlsx/.xlsm path suffix to adopt from Git stage 3.")

    prepare = subparsers.add_parser("prepare", help="Extract base/ours/theirs and build a sheet-aware manifest.")
    prepare.add_argument("--path", help="Only prepare this repository-relative .xlsx/.xlsm path.")

    resolve = subparsers.add_parser("resolve", help="Prepare conflicts and launch the local visual resolver.")
    resolve.add_argument("--path", help="Only resolve this repository-relative .xlsx/.xlsm path.")
    resolve.add_argument("--no-browser", action="store_true", help="Print the URL without opening a browser.")

    launch = subparsers.add_parser("launch", help="Prepare conflicts, start the visual resolver in background, and return its URL.")
    launch.add_argument("--path", help="Only resolve this repository-relative .xlsx/.xlsm path.")
    launch.add_argument("--no-browser", action="store_true", help="Start the server without opening a browser.")

    serve_parser = subparsers.add_parser("serve", help=argparse.SUPPRESS)
    serve_parser.add_argument("--manifest", required=True)
    serve_parser.add_argument("--no-browser", action="store_true")
    serve_parser.add_argument("--port", type=int, default=0, help="Fixed listen port (0 = ephemeral).")

    apply_parser = subparsers.add_parser("apply", help="Apply a decisions JSON without launching the UI.")
    apply_parser.add_argument("--manifest", required=True)
    apply_parser.add_argument("--decisions", required=True)
    apply_parser.add_argument("--no-commit", action="store_true")
    apply_parser.add_argument("--push", action="store_true")
    apply_parser.add_argument("--message", default=None, help="Override the auto-generated structured commit message.")
    return parser


def main(argv: list[str] | None = None) -> int:
    _configure_stdio()
    args = build_parser().parse_args(argv)
    try:
        # serve/apply 仅依赖 manifest（其 repoRoot 在 apply_manifest 内部解析），
        # 可在非 git 目录下运行；其余公开命令需要前置解析仓库。
        if args.command in ("detect", "filter", "prepare", "resolve", "launch"):
            repo = find_repo_root(args.repo)
        if args.command == "filter":
            result = add_auto_theirs_filter(repo, args.path)
            _print({"ok": True, "repo": str(repo), **result})
            return 0
        if args.command == "detect":
            conflicts = detect_conflicts(repo)
            filters = load_auto_theirs_filters(repo)
            auto = [item["path"] for item in conflicts if auto_theirs_path(item["path"], filters)]
            _print({
                "ok": True,
                "repo": str(repo),
                "count": len(conflicts),
                "reviewCount": len(conflicts) - len(auto),
                "autoTheirs": auto,
                "conflicts": conflicts,
            })
            return 0
        if args.command == "prepare":
            manifest = prepare_conflicts(repo, relative_path=args.path, runtime_dir=args.runtime_dir)
            _print({"ok": True, "manifest": str(manifest)})
            return 0
        if args.command == "resolve":
            manifest = prepare_conflicts(repo, relative_path=args.path, runtime_dir=args.runtime_dir)
            auto = _auto_only(manifest)
            if auto is not None:
                _print(auto)
                return 0 if auto["ok"] else 1
            from merge_server import serve

            return 0 if serve(manifest, open_browser=not args.no_browser) else 1
        if args.command == "launch":
            manifest = prepare_conflicts(repo, relative_path=args.path, runtime_dir=args.runtime_dir)
            auto = _auto_only(manifest)
            if auto is not None:
                _print(auto)
                return 0 if auto["ok"] else 1
            _print(_launch_background(repo, manifest, open_browser=not args.no_browser))
            return 0
        if args.command == "serve":
            from merge_server import serve

            return 0 if serve(args.manifest, open_browser=not args.no_browser, port=args.port) else 1
        decisions = json.loads(Path(args.decisions).read_text(encoding="utf-8"))
        result = apply_manifest(
            args.manifest,
            decisions,
            commit=not args.no_commit,
            push=args.push,
            message=args.message,
        )
        _print(result)
        return 0 if result["success"] else 1
    except (ConflictError, OSError, ValueError, json.JSONDecodeError) as exc:
        _print({"ok": False, "error": str(exc)})
        return 1


if __name__ == "__main__":
    sys.exit(main())
