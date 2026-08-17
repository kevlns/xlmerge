#!/usr/bin/env node
/**
 * xlmerge 启动垫片：定位可用 Python 解释器后透传参数给包内 CLI。
 * - 解释器优先级：XLMERGE_PYTHON > python3 > python > py -3（需 >= 3.9）
 * - 系统解释器已有 openpyxl/bottle 时直接使用；否则在 ~/.xlmerge/venv
 *   创建用户级 venv 并安装固定版本依赖
 * - 所有诊断信息输出到 stderr，保证 stdout 只有 Python CLI 的 JSON 结果
 */

"use strict";

const { spawnSync, spawn } = require("child_process");
const fs = require("fs");
const os = require("os");
const path = require("path");

const PKG_ROOT = path.resolve(__dirname, "..");
const ENTRY = path.join(PKG_ROOT, "python", "xlsx_resolver", "resolve_xlsx_conflict.py");
const PY_ROOT = path.join(PKG_ROOT, "python");
const VENV_DIR = process.env.XLMERGE_VENV || path.join(os.homedir(), ".xlmerge", "venv");
const DEPS = ["openpyxl==3.1.5", "bottle==0.13.4"];

function log(msg) {
  process.stderr.write(`[xlmerge] ${msg}\n`);
}

/** 候选解释器命令（数组形式，兼容 "py -3"） */
function candidates() {
  const list = [];
  if (process.env.XLMERGE_PYTHON) list.push([process.env.XLMERGE_PYTHON]);
  if (process.platform !== "win32") list.push(["python3"]);
  list.push(["python"]);
  if (process.platform === "win32") list.push(["py", "-3"]);
  // 去重
  const seen = new Set();
  return list.filter((c) => {
    const key = c.join(" ");
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

/** 探测候选解释器，返回 { cmd, version } 或 null */
function probeInterpreter() {
  for (const cmd of candidates()) {
    let result;
    try {
      result = spawnSync(cmd[0], [...cmd.slice(1), "-c", "import sys; print(sys.version_info >= (3, 9))"], {
        encoding: "utf-8",
        timeout: 20000,
      });
    } catch {
      continue;
    }
    if (result.error || result.status !== 0) continue;
    if (result.stdout.trim() === "True") {
      return cmd;
    }
    if (result.stdout.trim() === "False") {
      log(`跳过 ${cmd.join(" ")}：Python 版本低于 3.9`);
    }
  }
  return null;
}

function venvPython(venvDir) {
  return process.platform === "win32"
    ? path.join(venvDir, "Scripts", "python.exe")
    : path.join(venvDir, "bin", "python");
}

function hasDeps(cmd) {
  const result = spawnSync(cmd[0], [...cmd.slice(1), "-c", "import openpyxl, bottle"], {
    encoding: "utf-8",
    timeout: 20000,
  });
  return result.status === 0 && !result.error;
}

/** 确保依赖可用：优先系统解释器，其次用户级 venv；返回解释器命令 */
function ensureInterpreter() {
  const system = probeInterpreter();
  if (!system) {
    log("未找到可用的 Python >= 3.9，请安装 Python 或设置 XLMERGE_PYTHON 指向解释器路径。");
    process.exit(1);
  }
  if (hasDeps(system)) return system;

  const vpy = venvPython(VENV_DIR);
  if (fs.existsSync(vpy)) {
    if (hasDeps([vpy])) return [vpy];
    log("已有 venv 缺少依赖，重新安装...");
  } else {
    log(`系统 Python 缺少 openpyxl/bottle，创建用户级 venv：${VENV_DIR}`);
    const created = spawnSync(system[0], [...system.slice(1), "-m", "venv", VENV_DIR], {
      stdio: ["ignore", "ignore", "inherit"],
      timeout: 180000,
    });
    if (created.status !== 0) {
      log("venv 创建失败，请手动执行：python -m venv \"" + VENV_DIR + '" 并 pip 安装 ' + DEPS.join(" "));
      process.exit(1);
    }
  }
  const installed = spawnSync(vpy, ["-m", "pip", "install", "--quiet", ...DEPS], {
    stdio: ["ignore", "ignore", "inherit"],
    timeout: 600000,
  });
  if (installed.status !== 0) {
    log(`pip 安装失败（${DEPS.join(" ")}），请检查网络或手动安装。`);
    process.exit(1);
  }
  return [vpy];
}

function main() {
  if (!fs.existsSync(ENTRY)) {
    log(`找不到入口脚本：${ENTRY}`);
    process.exit(1);
  }

  if (process.argv.includes("--selfcheck")) {
    const system = probeInterpreter();
    const cmd = system ? ensureInterpreter() : null;
    process.stdout.write(
      JSON.stringify(
        {
          ok: Boolean(cmd),
          entry: ENTRY,
          python: cmd ? cmd.join(" ") : null,
          venv: fs.existsSync(venvPython(VENV_DIR)) ? VENV_DIR : null,
        },
        null,
        2,
      ) + "\n",
    );
    process.exit(cmd ? 0 : 1);
  }

  const cmd = ensureInterpreter();
  const child = spawn(cmd[0], [...cmd.slice(1), ENTRY, ...process.argv.slice(2)], {
    stdio: "inherit",
    env: process.env,
  });
  child.on("error", (err) => {
    log(`启动失败：${err.message}`);
    process.exit(1);
  });
  child.on("exit", (code) => process.exit(code === null ? 1 : code));
}

main();
