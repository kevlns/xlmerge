# xlmerge Agent 使用规范

> 三方工作簿会在结构对齐前分别计算并固化公式最终值；apply 对来源、快照和工作树做完整预检，并以可回滚批次执行写回（版本以 package.json / npm 包为准）。

> 本文件是 AI Agent 调用 xlmerge 的规范正本。完整参数以 `xlmerge --help` / 子命令 `--help` 与 `v-cli.plugin.json`（经 v-cli 路由时用 `v-cli agent describe xlmerge`）为准；本文件锁定使用场景、调用规范、快速流程与边界。

处理任意 Git 仓库中未合并的 `.xlsx` / `.xlsm` 二进制冲突。命令入口（npm 安装后全局可用，`<repo>` 为配置表 Git 仓库路径；经 v-cli 路由为 `v-cli xlmerge …`）：

```bash
xlmerge --repo <repo> detect
xlmerge --repo <repo> filter add <path.xlsx>
xlmerge --repo <repo> launch
```

`npx` 等价：`npx --package @kevlns/xlmerge@latest xlmerge --repo <repo> detect`。

## 何时使用

用户要求处理策划表冲突、配置表冲突、xlsx/xlsm merge conflict 或多 Sheet 冲突，且 Git 仓库存在未合并的 `.xlsx`/`.xlsm` 时立即使用。这是一个**命令路由任务，不是表格分析任务**。

## 使用规范（Agent 必须遵守）

1. 用户确认要处理冲突后，**只做快速连贯的检测与启动**：`detect` → `count > 0` 时 `launch`。`reviewCount > 0` 时把返回的 `url` 告诉用户，然后停止分析，等待用户在页面完成选择；`reviewCount == 0` 时 `launch` 会直接采用过滤项的远端版本并提交，报告其 JSON 结果；`count == 0` 时报告没有未解决冲突并结束。
2. **不额外审查冲突情况**：不检查工作簿 Cell、不自行做三方 diff、不总结冲突内容、不读取 base/ours/theirs；这些复杂度由工具内部流程处理。
3. **不默认走无头自动合并**：`launch --no-browser` + 程序化 decisions/apply 的无头链路仅限自动化测试或用户明确提供决策 JSON / 明确要求自动化的场景；默认流程必须走 UI `launch`，决策权始终在用户。
4. **不替用户选择 ours/theirs**；不因为冲突文件多、Sheet 多而进入计划模式。
5. `apply` 默认写回并 `git commit`；`--no-commit` 仅测试决策时使用；`--push` 仅用户明确要求推送时使用，正常流程不自动 push。
6. 只处理单个文件时使用 `launch --path <仓库相对路径>`；多文件不按文件循环调用 `launch --path`，一次性同批次页面展示。

## 最高优先级快速流程

用户确认要处理冲突后，只做以下动作：

1. 执行检测：

```bash
xlmerge --repo <repo> detect
```

2. `count > 0` 时立即启动解析器；不传 `--path`，一次性处理所有冲突工作簿：

```bash
xlmerge --repo <repo> launch
```

3. `reviewCount > 0` 时，把 `launch` 返回的 `url` 告诉用户，然后停止分析，等待用户在页面完成选择。`launch` 负责提取全部 Git 三方版本、计算多文件多 Sheet diff、后台启动服务并打开浏览器。`reviewCount == 0` 时不会启动 UI，命令会写回所有过滤项、提交并返回 `autoResolved`。

如果 `count` 为 `0`，报告没有未解决的 `.xlsx` / `.xlsm` 冲突并结束。

## 仓库过滤列表

过滤列表默认为空，仅在用户明确要求某类生成表始终采用远端整表时添加：

```bash
xlmerge --repo <repo> filter add locale/Key.xlsx
```

命令把规范化后的相对 `.xlsx`/`.xlsm` 路径后缀写入仓库根目录 `.xlmerge.json` 的 `autoTheirs` 数组；匹配不区分大小写，重复添加是幂等操作。工具不内置任何业务路径。

命中过滤项的冲突文件不进入结构化分析或 UI，而是逐字节写入 Git index stage 3（theirs）。若 stage 3 不存在则安全失败；混合批次会在用户确认人工冲突后，与 UI 处理结果一起写回并提交。

## 快速流程禁止事项

- 不调用通用 `xlsx` Skill、Excel API 或临时 Python 读取 base/ours/theirs。
- 不运行 `prepare` 后读取 manifest，也不枚举或总结 Cell/Sheet 差异。
- 不先搜索配置表结构、Luban schema、策划字段含义或项目代码。
- 不替用户判断某个 Cell 应选 ours 还是 theirs。
- 不因为冲突文件多、Sheet 多而进入计划模式；这些复杂度由工具内部流程处理。
- 不使用阻塞式 `resolve` 作为正常入口；正常调用始终使用会快速返回 URL 的 `launch`。
- 默认流程禁止按文件循环调用 `launch --path`；单个文件异常也必须由同一批次页面展示或明确报错。

只处理单个文件时，仍然直接使用 `launch --path <仓库相对路径>`。

## 程序负责的内部流程

以下内容由解析器自动完成，用于理解边界，不要由模型重复实现：

1. 从 Git index 提取 base/ours/theirs（stage 1/2/3）。含公式时，三方原始工作簿先分别在隔离副本中完整计算并固化最终值快照，再经内容锚点引擎生成 Sheet、行、列和 Cell 四级三方差异；不含公式时走原文件快路径。冲突身份是 `仓库相对路径 + 逻辑 Sheet + 逻辑行 ID + 逻辑列 ID`，物理坐标只用于展示；新增/删除行列后选择不会漂移。Luban 风格表头（`##var` / `##type` / `##group` / `##`）会被自动识别并保留，`##var` 为空的注释列、占位列同样参与对齐与写回。

2. 本地服务只监听 `127.0.0.1`，浏览器负责用户交互：

   - 顶部第一层标签切换冲突文件，第二层 `SHEET` 标签切换工作表；选择状态按文件和 Sheet 隔离。
   - 两侧修改同一 Cell 且结果不同：默认显示 `theirs`，单击可在 `ours` / `theirs` 间切换；只有一侧修改时默认采用修改侧，仍可单击覆盖。
   - 一侧删除 Sheet、另一侧修改：按整个 Sheet 选择保留 `ours` 或 `theirs`；只在一侧新增的 Sheet 自动保留。
   - `.xlsm` 的 VBA/ActiveX 单边变化自动采用变化侧；双边变化时在文件级选择本地宏/远端宏（默认本地）。
   - 相对本地可自动合入的远端新增行/列用浅绿色整行/列展示并锁定；已确定删除的行/列用灰色墓碑展示并锁定。
   - 单击行号/列头批量切换；“全部本地/全部远端”作用于当前 Sheet；“撤销”回退最近一次选择，“重置”恢复默认三方选择。
   - 当前 Sheet 提供上一个/下一个冲突定位；每个需人工审阅的 Sheet 必须独立勾选“已确认”，顶部统计跨文件显示已确认数量，未全部确认时提交会给出可点击的定位清单。
   - 悬停 Cell 时展示 base、双方值、双方文件作者和当前选择。

3. 用户点击“写回并提交”后：

   - 先请求只读 `/api/commit-preview` 生成 commit 信息，弹出居中的二次确认窗口；确认后才执行写回与 `git add` / `git commit`。
   - 以原始 `ours` 工作簿为宏包/样式基底重建受影响 Sheet；受影响 Sheet 的全部公式写为已验证的最终静态值，未受影响 Sheet 保留原公式及缓存。保留所选来源的样式、批注、超链接、数字格式、行高列宽、冻结拆分等视图状态。
   - 保存到仓库原路径并重新打开验证可读；`.xlsm` 全程 `keep_vba` 模式，保存后逐字节校验 VBA 负载。
   - 存在冲突 Cell 选择时自动生成结构化提交信息（每个冲突 Cell 一行，含表、Sheet、单元格、选择、双方值）；过滤项以“远端全量”记录，`--message` 可覆盖。
   - 可视化入口不自动 push。

4. 用户完成页面操作并返回后复核：

```bash
git -C <repo> diff --name-only --diff-filter=U
git -C <repo> status --short --branch
git -C <repo> log -1 --oneline
```

第一条必须无输出；随后报告处理的文件、Sheet、人工选择数量和 commit 结果。

## 非交互入口

仅在自动化测试或用户已提供明确决策 JSON 时使用 prepare/apply；不要用于查看 diff。决策必须同时包含文件、Sheet 和 Cell（Cell 优先使用 manifest 的稳定逻辑 ID）：

```bash
xlmerge --repo <repo> prepare
xlmerge --repo <repo> apply --manifest <manifest.json> --decisions <decisions.json>
```

测试决策但不 commit 时加 `--no-commit`；只有用户明确要求推送时加 `--push`。

## 边界与失败处理

- 只处理 Git 未合并状态中的 `.xlsx` / `.xlsm`；不处理普通修改。
- 文件级 delete/modify 冲突缺少完整双方工作簿时，解析器停止并报告；不要猜测删除还是保留。
- 过滤项要求 Git stage 3 存在；缺少远端版本时停止并报告，不退回本地版本。
- 写回前必须收齐所有人工选择；缺少任何选择时不改工作树。
- prepare 生成的三方来源、最终值快照和当时工作树指纹会在 apply 前全部校验；快照篡改、Git stage 变化、公式计算失败、公式错误或不可靠公式语义都会明确中止，整批次不写回、不 `git add`。
- 所有人工表和自动表先生成临时输出并通过预检，再统一替换；替换中途失败会回滚已替换文件，若文件系统同时阻止回滚则保留恢复副本并报告路径。
- 替换失败（如 Excel 正在占用目标文件，替换整体回滚）时，应要求用户关闭该工作簿后重试。
- 不直接修改 XLSX 内部 XML，不使用未经本次完整计算验证的旧公式缓存冒充最终值。
- 运行时提取文件和 manifest 默认放系统临时目录，不写入业务仓库。

## 环境

npm 包通过 `bin/xlmerge.js` 启动：自动检测系统 Python（≥3.9，`XLMERGE_PYTHON` 可直接指定；否则按平台探测 `python3` / `python`，Windows 追加 `py -3`）；若缺 `openpyxl` / `bottle` 依赖，则在 `~/.xlmerge/venv` 创建用户级 venv 并安装 `openpyxl==3.1.5`、`bottle==0.13.4`。不含公式的工作簿保持跨平台；含公式时必须在 Windows 上安装桌面版 Excel，解析器只让 Excel 打开隔离的临时副本，并拒绝外链、UDF/VBA、volatile、循环引用、动态数组等不可安全复现的公式语义。`XLMERGE_EXCEL_TIMEOUT_SECONDS` 可覆盖每个工作簿默认 120 秒的重算硬超时。正常流程不检查环境；仅入口明确报环境错误时诊断。

## 验证

```bash
cd <npm 包目录>
python -m unittest discover -s python/tests -v
```
