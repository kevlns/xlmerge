<div align="center">

<img src="https://raw.githubusercontent.com/kevlns/xlmerge/main/logo.png" alt="xlmerge logo" width="320" />

# xlmerge

**解决 Git 仓库中 `.xlsx` / `.xlsm` 策划配置表未合并冲突（merge conflict）的可视化工具。**

[![npm version](https://img.shields.io/npm/v/%40kevlns%2Fxlmerge?style=flat-square&color=cb3837)](https://www.npmjs.com/package/@kevlns/xlmerge)
[![npm downloads](https://img.shields.io/npm/dm/%40kevlns%2Fxlmerge?style=flat-square&color=4c8bf5)](https://www.npmjs.com/package/@kevlns/xlmerge)
[![CI](https://img.shields.io/github/actions/workflow/status/kevlns/xlmerge/test.yml?branch=main&style=flat-square&label=CI)](https://github.com/kevlns/xlmerge/actions/workflows/test.yml)
[![license](https://img.shields.io/github/license/kevlns/xlmerge?style=flat-square&color=2e8b57)](./LICENSE)

[Getting started](#getting-started) · [Usage](#usage) · [Development](#development) · [Package family](#package-family)

</div>

---

## Why xlmerge?

Git 对二进制 `.xlsx`/`.xlsm` 只会整文件冲突，无法像文本一样逐行合并。xlmerge 用**三方结构化差异引擎 + 本地可视化 UI**把这件事变得可控：自动提取 Git index 的 base/ours/theirs 三方版本，按 Sheet、逻辑行、逻辑列、Cell 四级计算差异，浏览器中逐 Cell 选择本地/远端版本，最终写回工作簿并生成结构化 commit。

- **逻辑 ID 定位** - 冲突身份用逻辑 ID（非物理坐标），插行删行后选择不会漂移
- **保真写回** - 保留公式（注入缓存值，`data_only` 读取正确）、样式、批注、合并单元格、冻结窗格、行高列宽
- **VBA 安全** - `.xlsm` 全程 `keep_vba`，保存后逐字节校验 VBA/ActiveX 负载
- **Luban 表头识别** - 自动识别 `##var` / `##type` / `##group`，注释列、占位列参与对齐
- **默认不越权** - 服务只监听 `127.0.0.1`，不自动 push

## Getting started

### Install

```bash
npm install -g @kevlns/xlmerge

# 或直接从 GitHub 仓库安装开发版本
npm install -g git+https://github.com/kevlns/xlmerge.git
```

要求：Node.js >= 16，系统 Python >= 3.9。首次运行若缺少依赖，会在 `~/.xlmerge/venv` 自动创建用户级 venv 并安装 `openpyxl==3.1.5`、`bottle==0.13.4`。

### Quick start

```bash
# 检测未合并的 .xlsx/.xlsm（输出 JSON）
xlmerge --repo <仓库路径> detect

# 提取三方版本并后台启动可视化解析器，返回 url（自动打开浏览器）
xlmerge --repo <仓库路径> launch
```

环境变量：

| 变量 | 作用 |
|------|------|
| `XLMERGE_PYTHON` | 指定解释器路径（如嵌入式 Python），跳过自动探测 |
| `XLMERGE_VENV` | 覆盖默认 venv 目录 |

## Usage

```bash
# 检测未合并的 .xlsx/.xlsm（输出 JSON）
xlmerge --repo <仓库路径> detect

# 提取三方版本并后台启动可视化解析器，返回 url（自动打开浏览器）
xlmerge --repo <仓库路径> launch

# 只处理单个文件
xlmerge --repo <仓库路径> launch --path new_meta/Items.xlsx

# 非交互（自动化/CI）：先 prepare 再 apply
xlmerge --repo <仓库路径> prepare
xlmerge --repo <仓库路径> apply --manifest <manifest.json> --decisions <decisions.json> [--no-commit] [--push]
```

`--repo` 缺省为当前目录（向上查找 Git 根）。

decisions.json 格式（Cell 用 manifest 中的稳定逻辑 ID，如 `row:base:4|name`）：

```json
{
  "files": {
    "new_meta/Items.xlsx": {
      "sheets": {
        "Items": { "cells": { "row:base:4|name": "theirs" } }
      }
    }
  }
}
```

### Agent 调用规范

包根 `AGENTS.md` 是 AI Agent 调用规范正本，包含适用场景、快速流程、禁止事项与边界。首次通过 v-cli 调用前应运行 `v-cli agent docs xlmerge` 掌握该规范；独立安装时也可直接读取包内 `AGENTS.md`。

## Development

```bash
# 单元测试（72 个，覆盖引擎与解析器全流程）
python -m unittest discover -s python/tests -v

# 校验 v-cli.plugin.json 与 CLI 解析器未漂移（命令/选项/副作用覆盖）
npm run check:manifest

# 端到端（构造真实冲突仓库 -> detect -> launch -> API 决策 -> 写回提交）
python e2e_test.py bin/xlmerge.js node

# 打包
npm pack
```

### 包结构

```
bin/xlmerge.js                  # Node 启动垫片：探测 Python、按需建 venv、透传参数
AGENTS.md                       # AI Agent 调用规范正本（场景/规范/流程/边界）
v-cli.plugin.json               # Agent 插件清单（命令/选项/副作用），与 CLI 解析器做漂移校验
python/
├── xlsx_merge_engine/          # 三方结构化差异引擎（可独立 CLI 调试）
│   └── xlsx_git_merge_bridge.py
├── xlsx_resolver/              # 冲突解析器（detect/prepare/launch/apply + bottle UI）
└── tests/                      # 引擎与解析器测试（不随 npm 包发布）
```

> 注：为兼容嵌入式 Python 发行版（`._pth` 忽略 `PYTHONPATH`），模块内使用包内自举的 `sys.path` 注入而非纯包导入；这是有意为之。

## Package family

kevlns 工具家族共享同一套发布约定（tag 驱动、CI 护栏、MIT）。

| Package | Purpose | Status |
| --- | --- | --- |
| [`v-cli`](https://github.com/kevlns/v-cli) | 个人工具箱 CLI | v0.2.1 |
| [`xlmerge`](https://github.com/kevlns/xlmerge) | Git 中 .xlsx/.xlsm 冲突可视化解决工具（本仓库） | v1.2.2 |
| [`u-cli-mod`](https://github.com/kevlns/u-cli-mod) | Unity 精确版本路由 + CLI + pipeline 包（Windows-first） | v0.1.1 |

## Compatibility

| Runtime | Supported versions |
| --- | --- |
| Node.js | `16` and later |
| Python | `3.9` and later |

## Contributing

```bash
git clone https://github.com/kevlns/xlmerge.git
cd xlmerge
python -m unittest discover -s python/tests -v
npm run check:manifest
```

For bugs and feature requests, use [GitHub Issues](https://github.com/kevlns/xlmerge/issues).

## License

Released under the [MIT License](./LICENSE).

<div align="center">

Part of the **kevlns** tool family.

</div>
