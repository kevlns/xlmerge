# xlmerge

解决 Git 仓库中 `.xlsx` / `.xlsm` 策划配置表的未合并冲突（merge conflict）。

三方结构化差异引擎 + 本地可视化 UI：自动提取 Git index 的 base/ours/theirs 三方版本，按 Sheet、逻辑行、逻辑列、Cell 四级计算差异，浏览器中逐 Cell 选择本地/远端版本，最终写回工作簿并生成结构化 commit。冲突身份使用逻辑 ID（非物理坐标），插行删行后选择不会漂移。

- 保留公式（注入缓存值，`data_only` 读取正确）、样式、批注、合并单元格、冻结窗格、行高列宽
- `.xlsm` 全程 `keep_vba`，保存后逐字节校验 VBA/ActiveX 负载
- 自动识别 Luban 风格表头（`##var` / `##type` / `##group`），注释列、占位列参与对齐
- 服务只监听 `127.0.0.1`，不自动 push

## 安装

```bash
# 从 GitHub 仓库直接安装（当前推荐方式）
npm install -g git+https://github.com/kevlns/xlmerge.git

# 发布到 npm 后（计划中）
npm install -g xlmerge
```

要求：Node.js >= 16，系统 Python >= 3.9。首次运行若缺少依赖，会在 `~/.xlmerge/venv` 自动创建用户级 venv 并安装 `openpyxl==3.1.5`、`bottle==0.13.4`。

环境变量：

| 变量 | 作用 |
|------|------|
| `XLMERGE_PYTHON` | 指定解释器路径（如嵌入式 Python），跳过自动探测 |
| `XLMERGE_VENV` | 覆盖默认 venv 目录 |

## 使用

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

### 作为 Agent Skill 使用

`skill/SKILL.md` 为技能文档，可复制（或软链）到目标 Agent 的 skills 目录（如 `.claude/skills/xlmerge/`、`~/.pi/agent/skills/`）。Skill 要求模型只做命令路由：`detect` -> `launch` -> 等待用户在 UI 完成 Cell 选择，禁止模型自行读取或总结 diff。

## 开发与测试

```bash
# 单元测试（63 个，覆盖引擎与解析器全流程）
python -m unittest discover -s python/tests -v

# 端到端（构造真实冲突仓库 -> detect -> launch -> API 决策 -> 写回提交）
python e2e_test.py bin/xlmerge.js node

# 打包
npm pack
```

## 包结构

```
bin/xlmerge.js                  # Node 启动垫片：探测 Python、按需建 venv、透传参数
skill/SKILL.md                  # Agent Skill 文档
python/
├── xlsx_merge_engine/          # 三方结构化差异引擎（可独立 CLI 调试）
│   └── xlsx_git_merge_bridge.py
├── xlsx_resolver/              # 冲突解析器（detect/prepare/launch/apply + bottle UI）
└── tests/                      # 引擎与解析器测试
```

> 注：为兼容嵌入式 Python 发行版（`._pth` 忽略 `PYTHONPATH`），模块内使用包内自举的 `sys.path` 注入而非纯包导入；这是有意为之。

## License

[MIT](LICENSE)

## 版本

- 0.0.1：首个个人仓库版本（源自内部 Skill v1.2.0，2025-08-10）
