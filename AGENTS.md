# Vant 表格合并开发规范

## 定位

本文件仅指导维护 xlmerge 仓库。运行时技能正本为 skills/xlmerge/SKILL.md；命令机器契约为 v-cli.plugin.json。

## 调用流程

修改前读取相关源文件与 package.json 中的检查命令，保留已有工作区改动。

## 调用规范

技能正文与 references 独立维护，不从本文件生成运行规范。修改命令时同步机器清单及生成的命令文档。宿主配置和状态使用 VANT_HOME（默认 ~/.vant），项目数据保留在工程 .vant；安装不扫描其他 Agent 目录。外部宿主仅显式适配。

## 结果验证

运行 package.json 中与改动相关的检查；区分单测、真实后端、界面和安装验证。不得将模拟结果报告为真实生图或发布成功。

## 参考资料

[使用说明](README.md) · [运行技能](skills/xlmerge/SKILL.md)

## 许可与版权

MIT。Copyright © 2026 Vant。
