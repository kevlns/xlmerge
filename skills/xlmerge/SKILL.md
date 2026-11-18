---
name: xlmerge
description: "处理 Git 中的 Excel 三方冲突，支持 Agent 根据项目规则裁决并写回，也支持浏览器逐项审阅。"
---

# 配表冲突

在冲突仓库用 `detect` 定位工作簿，`prepare` 生成三方差异。根据稳定 ID 和业务规则裁决歧义，通过 `resolve` 提交裁决、`apply --no-commit` 写回，检查 `status` 与项目验证结果后继续原合并。需要人工可视化审阅时使用 `launch`。

在 Vant 会话的 shell 中，示例里的 `v-cli` 代表插件入口：PowerShell 写作 `node "$env:DSH_VANT_CLI"`，sh 写作 `node "$DSH_VANT_CLI"`。 后接 `xlmerge` 和对应参数。完整参数与交付条件见[操作与边界](references/project.md)。
