# 配表冲突处理

## 定位

在工单或会话绑定的 Git 工区运行。原始三方、工作表、逻辑行列、公式计算值和稳定 ID 均由 xlmerge 生成。

## 调用流程

1. 用 `xlmerge --repo <repo> detect` 核对未合并的工作簿。没有冲突则结束。先核对 Git 合并双方，明确 ours 和 theirs 在本次操作中指向谁。
2. 用 `xlmerge --repo <repo> prepare` 获取 manifest 路径。仅处理指定文件时加 `--path <相对路径>`。按 files[].diff.sheets 读取三方变化与项目规则、字段约束、引用及两侧提交意图。保留可兼容的单边变化。命中仓库既有 autoTheirs 规则的文件由工具处理。
3. 对双边改动使用清单稳定 ID，记录选择和业务依据。只上报无法同时满足的具体取舍，保留当前清单和已判项。Agent 在获得裁决后续办原任务。
4. 构造 decisions.json。例如同一工作表内两个字段分别取不同来源：

```json
{"files":{"<文件ID>":{"sheets":{"<Sheet名>":{"cells":{"<逻辑单元格ID一>":"ours","<逻辑单元格ID二>":"theirs"}}}}}}
```

工作表冲突使用 sheet，重命名用 rename，合并区域用 mergedCells，均按清单选 ours/theirs。rows 和 columns 中的 delete_modify 按 ID 选 keep/delete。宏冲突使用文件级 macro。遇到需要工作簿两侧都没有的新值，单独安排普通改表与验证。

5. 用 `xlmerge --repo <repo> apply --manifest <manifest.json> --decisions <decisions.json> --no-commit` 写回并暂存。逐项核对 success、Git 未合并列表与暂存结果；按原授权继续合并或交 Reviewer。只在明确要求推送时推送。来源变动重新 prepare；写回结果未知时先检查实际工作簿与索引。
6. 用户想亲自审阅时可用 launch，status/cancel 使用同一会话的 manifest。launch 在所有文件都命中 autoTheirs 时会直接完成合并并提交。只有明确要求长期改变过滤策略才用 filter add。

## 调用规范

- 只处理 Git 未合并状态中的 `.xlsx` / `.xlsm`；不处理普通修改。
- 文件级 delete/modify 冲突缺少完整双方工作簿时，解析器停止并报告；不要猜测删除还是保留。
- 过滤项要求 Git stage 3 存在；缺少远端版本时停止并报告，不退回本地版本。
- 写回前必须收齐所有必需裁决；缺少任何选择时不改工作树。
- prepare 生成的三方来源、审阅值快照和当时工作树指纹会在 apply 前全部校验；快照篡改、Git stage 变化、计算未完成或不可靠公式语义都会明确中止，整批次不写回、不 `git add`。
- 三方计算错误保留原公式、错误与位置作为未知值进入同一批次 UI，必须选择有效来源；无法可靠确认记录身份、表头或行列结构时必须选择整个 Sheet。双方同一记录、同一字段都错误时，工具按一致的显式基础类型替换（数值 0、string 空串、bool false），保留审计信息；ID、引用、约束字段、未知类型、重复 ID 不自动替换。未经处理的错误不能写回，安全限制不变。业务歧义按任务的决策渠道上报。
- 所有人工表和自动表先生成临时输出并通过预检，再统一替换；替换中途失败会回滚已替换文件，若文件系统同时阻止回滚则保留恢复副本并报告路径。
- 替换失败（如 Excel 正在占用目标文件，替换整体回滚）时，应要求用户关闭该工作簿后重试。
- 不直接修改 XLSX 内部 XML，不使用未经本次完整计算验证的旧公式缓存冒充最终值。
- 运行时提取文件和 manifest 默认放系统临时目录，不写入业务仓库。

### 环境

npm 包通过 `bin/xlmerge.js` 启动：自动检测系统 Python（≥3.9，`XLMERGE_PYTHON` 可直接指定；否则按平台探测 `python3` / `python`，Windows 追加 `py -3`）；若缺 `openpyxl` / `bottle` 依赖，则在 `~/.xlmerge/venv` 创建用户级 venv 并安装 `openpyxl==3.1.5`、`bottle==0.13.4`。不含公式的工作簿保持跨平台；含公式时必须在 Windows 上安装桌面版 Excel，解析器只让 Excel 打开隔离的临时副本，并拒绝外链、UDF/VBA、volatile、循环引用、动态数组等不可安全复现的公式语义。`XLMERGE_EXCEL_TIMEOUT_SECONDS` 可覆盖每个工作簿默认 120 秒的重算硬超时。正常流程不检查环境；仅入口明确报环境错误时诊断。

### 验证

```bash
cd <npm 包目录>
python -m unittest discover -s python/tests -v
```

后台审阅用 `xlmerge launch`。`xlmerge status --manifest <file>` 查询原服务，`xlmerge cancel --manifest <file>` 请求停止并检查 stopConfirmed；身份不一致拒绝停止。通过 v-cli 结构化执行时，launch/resolve 返回 accepted 与 handle；status/cancel 使用原 handle，后台进程继续持有工程租约，停止前不重放启动。

## 结果验证

核对原始错误、任务身份、停止状态及实际产物。只报告已经完成的验证，模拟测试与真实执行分别说明。

## 参考资料

使用说明（工具包 `README.md`） · 文档规范（工具包 `docs/documentation.md`） · v-cli.plugin.json

## 许可与版权

MIT。Copyright © 2026 Vant。
