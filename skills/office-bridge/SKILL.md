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

1. **查在线文档**：`GET http://127.0.0.1:3000/office/status` → `hosts` 字段列出在线文档（Word/Excel/PowerPoint）
   - 目标文档不在线 → 提醒用户：打开目标文档，并在 **开始/开发人员 → 加载项 → 开发人员加载项** 中打开「DSH Office 执行器」窗格并保持开启，然后重试
2. **查可用能力**（可选）：`GET /office/capabilities` → action 注册表（hosts / 参数 / 是否破坏性）
3. **发指令**：`POST /office/command`，body：`{ "action": "...", "host": "Word|Excel|PowerPoint", "args": {...} }`
   - `host` 指定目标文档，避免多文档同时打开时抢指令
   - 指令一次一条、串行等待结果（超时 90s）
4. **处理结果/错误**：
   - `ok: true` → 用 `result` 完成用户任务
   - `code: addin_offline` → 窗格未开启，提醒用户开窗格，**不要重试**
   - `code: unknown_action` → 先查 capabilities 用正确的 action 名
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

## 安全与边界

- **🔒 覆盖类操作机制级确认（confirm，写前 re-read）**：以下属覆盖类——`write_selection` / `replace_all` / `remove_empty_paragraphs` / `delete_sheet` / `format_selection` / `write_range` / `format_range` / `rename_sheet` / `apply_sort` / `apply_filter` / `set_font` / `apply_style`。**不带 `confirm: true` 的指令会被桥接拒绝并返回当前状态预览**（`code: confirm_required`，`result.preview` 含被操作位置的当前内容）。正确流程：① 发指令（不带 confirm）拿预览 → ② 向用户展示"将改动什么" → ③ 用户确认后带 `confirm: true` 重发执行。**这保证任何覆盖都基于最新文档状态，不会凭记忆覆盖用户刚做的修改**
- **破坏性操作 dryRun 预览**：`replace_all` / `remove_empty_paragraphs` / `delete_sheet` 的预览本身就是 dryRun（返回影响范围），确认时一并展示
- **⚠️ 删空行有风险（实测踩坑）**：`remove_empty_paragraphs` 删除空段落可能**破坏文档的节/样式分隔**（Word 会因此自动重排小节、打乱格式）。执行前**必须 dryRun 预览 + 明确告知用户风险**；建议仅在"确实需要清理多余空行"且用户确认后执行，文档结构复杂（含分节符/多级标题）时宁可保守
- **⚠️ 写入不要多加空行（实测踩坑）**：向 Word 插入/追加段落时，**不要**在段落之间插入多余的空段落——Word 段落间距应由段落格式（间距/段后）控制，多余空行会导致 Word 自动重排、填充小节、打乱格式。`insert_paragraph` 一次一个段落、连续段落间不加空行；需要视觉间距时说明用段落格式，而非空行段落
- **写入后检查样式**：向 Word 写入内容（insert_paragraph / write_selection / append_text）后，用 `read_styles` / `read_document` 检查结果，样式不对时用 `format_selection` / `apply_style` 修正（用户常抱怨"写入后样式不对"）
- **删空段落保护**：自动跳过含图片的段落与文档结尾段（曾误删流程图，已修复）
- **环境边界**（本机实测）：Word 表格插入 / paragraphFormat / 批注 不可用（返回 `requirement`）；PPT 只能读（全文件文本 / 备注），不能新建幻灯片 / 改排版；Excel 基本全功能。换机器 / 更新 Office 可能不同——收到 `requirement`/`unsupported` 时如实降级，不要硬来
- **多文档并行**：同时开多个文档时务必带 `host` 指定目标，否则可能被任一文档执行
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
