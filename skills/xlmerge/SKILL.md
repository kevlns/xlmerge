---
name: xlmerge
description: "处理 Git 中的 Excel 三方冲突，支持 Agent 根据项目规则裁决并写回，也支持浏览器逐项审阅。"
---

# 配表冲突

在发生合并的工区 detect → prepare，读取 manifest 中 base/ours/theirs 的差异与稳定 ID，按项目约束及变更意图裁决。仅将无法判断的业务项交控制源。用 apply --no-commit 写回并验证 Git 索引，按原任务完成提交或交付。用户选择可视化审阅时使用 launch。参数与裁决格式见[操作规范](references/project.md)。
