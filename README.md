# Vant 表格合并工具

`@kevlns/xlmerge`

## 功能

在浏览器中审阅 Git 冲突工作簿的三方差异，校验后写回并提交合并结果。

## 环境要求

Node.js 16 及以上、Python 3.9 及以上；包含公式的工作簿需要 Windows 桌面版 Excel 生成可靠的计算快照。

## 安装

```sh
npm install -g @kevlns/xlmerge
```

## 快速开始

```sh
xlmerge --repo <仓库目录> detect
xlmerge --repo <仓库目录> launch
```

## 使用说明

### 功能说明

Git 对二进制 `.xlsx`/`.xlsm` 只会整文件冲突，无法像文本一样逐行合并。xlmerge 用**公式最终值快照 + 三方结构化差异引擎 + 本地可视化 UI**把这件事变得可控：自动提取 Git index 的 base/ours/theirs 三方版本，按 Sheet、逻辑行、逻辑列、Cell 四级计算差异，浏览器中逐 Cell 选择本地/远端版本，最终原子写回工作簿并生成结构化 commit。

- **逻辑 ID 定位** - 冲突身份用逻辑 ID（非物理坐标），插行删行后选择不会漂移
- **公式感知** - 三方工作簿先各自完整计算；diff 使用最终值，受影响 Sheet 固化为静态值，未受影响 Sheet 保留原公式及缓存
- **安全写回** - apply 校验三方来源、值快照与工作树指纹；批次替换失败自动回滚，并保留所选来源的样式、批注、合并单元格、冻结窗格、行高列宽
- **VBA 安全** - `.xlsm` 全程 `keep_vba`，保存后逐字节校验 VBA/ActiveX 负载
- **Luban 表头识别** - 自动识别 `##var` / `##type` / `##group`，注释列、占位列参与对齐
- **可配置整表过滤** - 仓库可把生成表加入过滤列表；命中时直接采用 Git stage 3 的远端原始文件
- **默认不越权** - 服务只监听 `127.0.0.1`，不自动 push

### 操作流程

```bash
# 检测未合并的 .xlsx/.xlsm（输出 JSON）
xlmerge --repo <仓库路径> detect

# 添加整表过滤项（写入仓库根目录 .xlmerge.json）
xlmerge --repo <仓库路径> filter add <相对路径或路径后缀.xlsx>

# 提取三方版本并后台启动可视化解析器，返回 url（自动打开浏览器）
xlmerge --repo <仓库路径> launch

# 只处理单个文件
xlmerge --repo <仓库路径> launch --path new_meta/Items.xlsx

# 前台阻塞式运行（与 launch 等价，进程前台等待完成）
xlmerge --repo <仓库路径> resolve

# 非交互（自动化/CI）：先 prepare 再 apply
xlmerge --repo <仓库路径> prepare
xlmerge --repo <仓库路径> apply --manifest <manifest.json> --decisions <decisions.json> [--no-commit] [--message <text>] [--push]
```

`--repo` 缺省为当前目录（向上查找 Git 根）。

过滤列表默认为空，保存在仓库根目录的 `.xlmerge.json`：

```json
{
  "autoTheirs": [
    "locale/Key.xlsx"
  ]
}
```

过滤项按路径段做不区分大小写的后缀匹配。命中的冲突工作簿不做 Cell 分析，而是逐字节采用 Git stage 3（theirs）并随同本批次提交；如果整批冲突都命中过滤项，`launch`/`resolve` 会直接写回并提交，不启动空白 UI。命令可重复执行，已有项不会重复写入。

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

#### Agent 调用规范

包根 `AGENTS.md` 是 AI Agent 调用规范正本，包含适用场景、快速流程、禁止事项与边界。首次通过 v-cli 调用前应运行 `v-cli agent docs xlmerge` 掌握该规范；直接安装时也可直接读取包内 `AGENTS.md`。

## 开发与验证

```sh
npm ci
npm test
npm run test:python
```

命令返回、结构校验、模拟测试与外部平台实际运行分别核验。发布版本和依赖以随包 `package.json` 为准。

## 文档

[Agent 调用规范](AGENTS.md) · [文档规范](docs/documentation.md)

## 许可与版权

本工具采用 [MIT 许可](LICENSE)。Copyright © 2026 Vant。Vant 为工具集品牌；Figma、Unity、Steam、微信及其组件归各自权利人所有，第三方组件适用各自许可。
