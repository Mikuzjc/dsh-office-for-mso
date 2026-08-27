---
name: office-bridge
description: DSH 的 Office 桥接技能 —— 通过本地桥接服务（localhost:3000）和 Office 加载项，操控用户当前打开的 Word/Excel/PowerPoint 文档（读/写/格式化/图表/公式/批注）。用于"读一下 Word 全文""Excel 生成图表""把选中的翻译成英文"等请求。
whenToUse: 用户要求操作 Word/Excel/PowerPoint 文档（读取、编辑、格式化、分析、生成图表/公式、批注等），且这些文档是用户当前已打开的。
---

# Office Bridge（DSH ↔ Microsoft Office）

## 架构

```
DSH 会话(AI) --POST /office/command--> 桥接服务 localhost:3000 --轮询--> Office 加载项(窗格) --Office.js--> Word/Excel/PPT
```

- **桥接服务**：`dsh-office-for-mso` 项目目录下的 `server.js`，通常由 Windows 计划任务「DSH Office Bridge」托管（登录自启、后台常驻）
- **加载项窗格**：用户 Office 文档侧边的「DSH Office 执行器」窗格，**使用时必须保持开启**（它承载 Office.js 执行；关闭后指令会立即返回 `addin_offline`）

## 使用流程

1. **查在线实例**：`GET http://127.0.0.1:3000/office/status` → `instances` 字段列出**每个在线窗格的实例**：`{instanceId, host, docUrl, docTitle}`——识别文档：**docUrl 优先**（文档路径，如 `file:///C:/.../测试Word.docx`）；**docUrl 为空（未保存的新文档）时用 docTitle**（文档标题/文件名）；两者都空（罕见）再发 `read_selection` 读光标上下文辅助判断
   - 目标文档不在线 → 提醒用户：打开目标文档，并在 **开始/开发人员 → 加载项 → 开发人员加载项** 中打开「DSH Office 执行器」窗格并保持开启，然后重试
2. **查可用能力**（可选）：`GET /office/capabilities` → action 注册表（hosts / 参数 / 是否破坏性）
3. **发指令**：`POST /office/command`，body：`{ "action": "...", "instance": "<instanceId>", "args": {...} }`
   - **`instance` 必填（强制）**：先按用户目标文档的 docUrl（docUrl 空则 docTitle）匹配 `instances` 里的 instanceId 带上；**不带 instance 会被拒绝**（`code: instance_required`）——多文档时未指定实例会导致指令被错误文档执行（曾发生：测试文件激活时正式论文收到审批）
   - `host` 可附带（Word/Excel/PowerPoint，校验用）；指令一次一条、串行等待结果（超时 90s）
4. **处理结果/错误**（完整错误码表见 `GET /office/errors`）：
   - `ok: true` → 用 `result` 完成用户任务
   - **无匹配不是错误**：`search` / `replace_all(dryRun)` 无命中返回 `ok:true + count:0`；`read_*`/`list_*` 空内容返回空数组/空字符串——**如实告知用户"未找到"，不要重试**
   - `code: instance_required` → 先查 status 拿 instanceId 再重发
   - `code: addin_offline` → 窗格未开启，提醒用户开窗格，**不要重试**
   - `code: busy` → 上一条指令仍在执行，稍等重发
   - `code: timeout` → 90s 无结果；查 status 确认窗格在线，在线可重试一次，仍超时提醒用户重开窗格
   - `code: unknown_action` → 先查 capabilities 用正确的 action 名
   - `code: bad_args` → 按 error 文本修正参数（缺必填参数/非法值）后重发
   - `code: not_found` → 定位/查找目标不存在（书签、文本、命名区域）；如实告知用户，建议换关键词或换定位方式，**不要重试同一查找**
   - `code: confirm_required` → ask 模式覆盖操作：把 `result.preview` 展示给用户，确认后带 `confirm:true` 重发
   - `code: rejected` → 用户在窗格拒绝（或审批超时）；停止操作、尊重用户决定，**不要重发**
   - `code: requirement / unsupported` → 该 API 在此环境不可用，如实告知用户（绝不假装成功）
   - `code: execution` → 把错误原样报告给用户

## 常用动作速查

- **读全文**：`read_document`（Word 全文 / Excel 工作表已用区域，`args.sheet` 指定表 / PPT 全文件逐页）
- **读选区**：`read_selection`（`withStyles: true` 附带样式）；**写选区**：`write_selection`（`{text}`）
- **读表格**（Word）：`read_tables`（结构化行列）；**读区域**（Excel）：`read_range {address: "Sheet1!A1:B10"}`
- **翻译/改写/润色**：先 `read_selection` 拿文本 → 处理 → `write_selection` 写回
- **数据→图表**（Excel）：`read_range` 读数据 → 分析 → `insert_chart {type, dataRange, title}`
- **公式求值**（Excel）：`evaluate_formula {formula: "SUM(A1:A10)"}`（白名单 SUM/AVERAGE/COUNT/MAX/MIN/PRODUCT）
- **工作表管理**：`list_sheets` / `add_sheet` / `rename_sheet` / `delete_sheet`
- **格式化**：`format_selection`（选区）/ `format_range`（Excel 区域）/ `set_font`（Word 全文）/ `apply_style`（Word 内置样式）
- **查找**：`search`（Word，支持通配符）；**全文替换**：`replace_all`（**先 dryRun 预览**）
- **改完自动选中**（内置约定，无需额外指令）：`replace_all`（Word）选中**最后一处**修改、`append_text` / `insert_paragraph` 选中插入内容、`write_range`（Excel）选中写入区域——结果里 `selected:true`。**多处无法同时选中**（Office.js 单选区），多处替换只选中最后一处示意
- **定位选中**：`locate_select`（零副作用：只做选中，不改文档内容/样式/撤销栈）——定位器三选一 `{text: "片段"}`（Word/Excel 首个匹配）、`{bookmark|anchor: "名"}`（Word 书签/Excel 命名区域）、`{range|address: "A1:B5"}`（Excel 区域，可带 `sheet`）；`blinks` 默认 0 = 只选中保持，`blinks: 2-3` 才闪烁；定位不到返回 `not_found`（如 `text not found`），如实转告、不要重试

## 安全与边界

- **🔒 覆盖类操作写前 re-read（机制级）**：以下属覆盖类——`write_selection` / `replace_all` / `remove_empty_paragraphs` / `delete_sheet` / `format_selection` / `write_range` / `format_range` / `rename_sheet` / `apply_sort` / `apply_filter` / `set_font` / `apply_style`。**无论何种模式，覆盖操作执行前都会自动读取当前状态（re-read），绝不凭记忆覆盖**。确认模式由执行器开关 `OFFICE_CONFIRM_MODE` 决定：
  - **`auto`（默认）**：自动 re-read 后**直接执行**，结果附 `previousState`（被操作前的状态快照）。AI 应核验 `previousState` 与预期一致（找到自己在改什么），异常立即告知用户
  - **`ask`**：re-read 后**必须用户确认**（`confirm: true`）才执行——发指令（无 confirm）拿预览 → 向用户展示 → 确认后带 confirm 重发。**审批时窗格会自动选中操作目标**帮用户看清改哪里：`replace_all` 选第一处命中、`insert_paragraph`/`append_text` 选中插入点（文末最后一段，`location: afterSelection` 时保持选区）、`write_range`/`format_range`（Excel）选目标区域
  - 切换：桥接服务环境变量 `OFFICE_CONFIRM_MODE=ask`（或 `auto`），改后重启服务
- **破坏性操作 dryRun 预览**：`replace_all` / `remove_empty_paragraphs` / `delete_sheet` 的预览本身就是 dryRun（返回影响范围），确认时一并展示
- **⚠️ 删空行有风险（实测踩坑）**：`remove_empty_paragraphs` 删除空段落可能**破坏文档的节/样式分隔**（Word 会因此自动重排小节、打乱格式）。执行前**必须 dryRun 预览 + 明确告知用户风险**；建议仅在"确实需要清理多余空行"且用户确认后执行，文档结构复杂（含分节符/多级标题）时宁可保守
- **⚠️ 中文标点规范（写入时，高频坑）**：向文档写入**中文内容**时，标点必须用中文全角标点，**尤其引号不要弄成英式**：
  - 引号：用中文弯引号 **“ ”**（嵌套内层用 ‘ ’），**严禁用英式直引号 " "**
  - 逗号/句号/分号/冒号/问号/感叹号：全角 ，。；：？！
  - 括号全角（）、省略号 ……、破折号 ——
  - 中文句中的英文/数字/符号保留半角（如 "p = 0.007"、85.71%、Token、cot_5 中的下划线）
  - 例外：纯英文句子/代码/公式内保持英文标点；中英混排以中文标点为主
  - **⚠️ 箭头分场景（正式文书禁、非正式文书可）**：**正式文书（论文/学术报告/正式文件）**：正文**禁用 → 表示数值变化或因果**（"507.73 → 607.04" 不严谨），应写"由 507.73 增至 607.04"、"由 55 题减至 48 题"；`→` 仅可用于定义映射、流程示意。**非正式文书（笔记/草稿/PPT/日常文档）**：箭头可用（"A → B" 表示变化更直观），不强制文字化
  - 写入后可用 `read_selection` 复查引号：若发现 `"` 直引号包裹中文，应修正为 “ ”
- **⚠️ 写入不要多加空行（实测踩坑）**：向 Word 插入/追加段落时，**不要**在段落之间插入多余的空段落——Word 段落间距应由段落格式（间距/段后）控制，多余空行会导致 Word 自动重排、填充小节、打乱格式。`insert_paragraph` 一次一个段落、连续段落间不加空行；需要视觉间距时说明用段落格式，而非空行段落
- **⚠️ Word 自动编号干扰（自动套用格式）**：Word 的「键入时自动套用格式」会在回车时自动创建编号/项目符号（尤其段落以"数字."开头、或上一段是编号列表时），导致意外自动编号、多余空行。AI 写入文本时**避免以"数字."/"数字）"等编号样式开头**；若用户反馈"回车变成编号/多了编号"，指导其在 **文件 → 选项 → 校对 → 自动更正选项 → 键入时自动套用格式** 中取消「自动编号列表」和「自动项目符号列表」（应用级设置，Office.js 无 API 可改）
- **写入后检查样式**：向 Word 写入内容（insert_paragraph / write_selection / append_text）后，用 `read_styles` / `read_document` 检查结果，样式不对时用 `format_selection` / `apply_style` 修正（用户常抱怨"写入后样式不对"）
- **删空段落保护**：自动跳过含图片的段落与文档结尾段（曾误删流程图，已修复）
- **环境边界**（本机实测）：Word 表格插入 / paragraphFormat / 批注 不可用（返回 `requirement`）；PPT 只能读（全文件文本 / 备注），不能新建幻灯片 / 改排版；Excel 基本全功能。换机器 / 更新 Office 可能不同——收到 `requirement`/`unsupported` 时如实降级，不要硬来
- **多文档精确路由（instanceId）**：每个打开文档的窗格有唯一实例 ID，`GET /office/status` 的 `instances` 字段列出**各实例及其文档路径（docUrl）与文档名（docTitle，docUrl 为空时用）**——多 Word 文档（如测试文件 + 正式论文）同时打开时，先查 status 按 docUrl（空则 docTitle）识别目标文档，指令带 `instance` 参数精确路由，**避免指令被错误文档执行**（曾发生：测试文件激活时指令被正式论文执行）。不带 instance 时由任一匹配 host 窗格执行（兼容旧行为；多文档务必带 instance）
- **窗格必须保持开启**：操作期间窗格关闭会立即得到 `addin_offline`，提醒用户重新打开

## 高级用法：题注与交叉引用一致性检查（学术论文常见需求）

无需新 API，用现有 `read_document` 即可完成：

1. `read_document` 读全文 → 提取所有**题注**（正则匹配"图 1""表 2""Figure 1""Table 2"等题注行）
2. 提取所有**交叉引用**（正文中"见图 1""如表 2"等引用）
3. 对比两侧编号集合：题注里有但没被引用的、引用了但题注不存在的、编号对不上的——输出差异清单
4. 若有题注语言混乱（"图 1" vs "Figure 1" 混用）一并指出；Word 题注自动变英文的根因通常是题注样式/插入标签语言跟随 Word 界面语言——指导用户在 **引用 → 插入题注 → 编号/标签** 中检查标签语言，或改用中文题注样式

## 编辑仓库文件前（guard 守卫）

在 dsh-office-for-mso 工作区编辑任何文件前，先运行：

```powershell
powershell -ExecutionPolicy Bypass -File guard.ps1 [文件...]
```

若报告未提交改动（可能来自外部进程/其他 AI/脚本，或 pwsh/node 直接改文件留下），**先 read 最新内容再编辑**——DSH 的 edit 工具自带 re-read 强制，但脚本直改文件不经过它，guard 补上这个缺口。

## 安装（用户侧，一次性）

```powershell
git clone https://github.com/Mikuzjc/dsh-office-for-mso.git
cd dsh-office-for-mso
npm run setup   # Windows：服务常驻（计划任务）+ 自动注册加载项（UAC 一次）
# 关闭并重新打开 Office 文档，侧边栏出现「DSH Office 执行器」窗格即就绪
```
