# dsh-office-for-mso

> [English](README.en.md) · [日本語](README.ja.md) · 中文

**DSH ↔ Office 桥接执行器（v1.3）**：在 DSH 会话里发指令，通过 Office 加载项操控你**正在打开的** Word / Excel / PowerPoint 文档——读取、写入、格式化、结构操作、图表/公式/批注，接近 Copilot for Office 的常见工作流。

```
你 ──DSH 会话发指令──▶ AI(agent) ──POST──▶ 桥接服务 localhost:3000
                                                │ 指令入队（按 host 路由）
                           Office 加载项（文档内后台轮询，1s）
                                                │ Office.js 执行（33 个 action）
                                                ▼
                            结果回传 ──▶ AI 取回 ──▶ 汇报给你
```

- **不冲突**：加载项跑在 Office 进程内，操作内存文档，与你的编辑由 Office 统一串行处理（无文件锁、无"最后保存者赢"）
- **无侧边栏交互**：窗格只是状态显示，所有指令从 DSH 下发
- **热更新**：所有 action 实现在 `actions.js`，改它 → 重启桥接服务 → 窗格自动重载，**无需重开窗格**

---

## 一、部署（一次性）

### 1. 克隆仓库并启动桥接服务

```powershell
git clone https://github.com/Mikuzjc/dsh-office-for-mso.git
cd dsh-office-for-mso
node server.js   # 或：powershell -ExecutionPolicy Bypass -File start.ps1
```

看到 `listening on http://127.0.0.1:3000` 即可。
**Windows 下启动时会自动注册加载项（WEF 注册表）**——首次请**关闭并重新打开 Office 文档**，窗格即出现；macOS 需用 Office 菜单手动加载（见 1.2/1.3）。
（下文路径示例均以你的实际项目目录为准。）

### 2. 把加载项加载到 Office（sideload）

> **平台**：核心（server.js + 加载项）仅依赖 Node.js，Windows / macOS 均可运行；`install.ps1`/计划任务为 Windows 专属的**可选**自动托管，macOS 手动 `node server.js` 即可。

> **一键 sideload（推荐；Windows 专属——基于 WEF 注册表机制；macOS 请用 Office 菜单手动加载）**：
> ```powershell
> cd <你的项目目录>
> powershell -ExecutionPolicy Bypass -File sideload.ps1   # 注册（-Remove 移除）
> ```
> 写入 WEF 注册表后，**关闭并重新打开 Office 文档**即自动出现窗格。

也可以手动（新版 Office 微软已隐藏「上传我的加载项」入口，请用**开发人员加载项**）：

1. 启用开发人员选项卡：**文件 → 选项 → 自定义功能区 → 主选项卡勾选「开发人员」** → 确定
2. 打开任意 Word / Excel / PowerPoint 文档
3. **开发人员选项卡 → 加载项**（或 **插入 → 我的加载项**）→ 打开「Office 加载项」对话框
4. 对话框底部左侧「管理」下拉选择 **开发人员加载项**
5. 点 **+（添加）** → 选择 `<你的项目目录>\manifest.xml`
6. 侧边栏出现「DSH Office 执行器」窗格，显示 **已连接：等待 DSH 指令** 即成功

之后**保持文档打开**即可；窗格可缩小/拖角落，无需操作。

> 旧版 Office 若仍有「上传我的加载项」入口，也可直接使用。
> 换电脑：重复以上两步（桥接服务 + 上传 manifest）即可。

### 1.3 首次使用

安装后，需**手动打开一次加载项**，DSH 会话才能操作文档：

1. 打开 Word / Excel / PowerPoint 文档
2. **开始（或开发人员）标签页 → 加载项 → 开发人员加载项 → 「DSH Office 执行器」**
3. 出现窗格（显示 **已连接：等待 DSH 指令**）后，即可在 DSH 会话下发指令
4. 窗格可调小/拖到角落，无需持续关注（窗格只是状态显示）

之后保持文档打开即可；窗格关闭时重复步骤 2 重新打开。

## 二、架构

| 文件 | 职责 |
|---|---|
| `server.js` | 桥接服务：指令队列、host 路由、心跳、静态文件、能力发现 `/office/capabilities`、actions 版本 `/office/actions-version` |
| `taskpane.js` | 加载项外壳：轮询/心跳/分发/热更新加载（**尽量少改**，改动需重开窗格） |
| `actions.js` | 全部 action 实现 + 注册表（**热更新单元**：改它无需重开窗格） |
| `pako.min.js` | 本地 zip 解压库（PPT OOXML 读取用，离线） |
| `manifest.xml` | 加载项清单（权限 ReadWriteDocument，已到顶） |

**多文档模型**：Word / Excel / PowerPoint 各自运行一个加载项实例；指令带 `host`（Word/Excel/PowerPoint）精确路由，`GET /office/status` 返回在线文档列表（`hosts` 字段）与窗格启动记录（`hellos`）。

**热更新机制**：窗格每次轮询前 GET `/office/actions-version`，比对 actions.js 的 mtime，变化则动态重载脚本。**改 actions.js → 重启 server → 自动生效**。

## 三、能力矩阵（33 个 action）

> `destructive=true` 的操作支持 `args.dryRun` 预览影响（replace_all / remove_empty_paragraphs / delete_sheet 已实现；其余 AI 层先读后写）。W=Word，E=Excel，P=PowerPoint。

### 通用
| action | 平台 | 说明 |
|---|---|---|
| `read_selection` | W/E/P | 读取当前选区文本；`withStyles=true` 附带样式 |
| `write_selection` | W/E/P | 用 `{text}` 替换当前选区 |
| `read_document` | W/E/P | Word 全文；Excel 工作表已用区域（`sheet` 指定表名，5000 格截断）；PPT 全文件逐页文本 |
| `read_styles` | W/E | 选区样式：Word（字体/字号/加粗/斜体/颜色/下划线/高亮）；Excel（逐格，上限 10×10） |
| `replace_all` | W/E | 全文查找替换 `{search, replace, dryRun?}` |
| `append_text` | W | 文末追加段落 `{text}` |

### Word 组
| action | 说明 |
|---|---|
| `read_tables` | 结构化读取全部表格（逐单元格，getRange().text 按 \t/\r 切分） |
| `set_font` | 全文（含表格段落）设置字体 `{font}` |
| `remove_empty_paragraphs` | 删除空段落（**跳过含图片段落与文档结尾段**，`dryRun` 预览） |
| `insert_paragraph` | 插入段落 `{text, style?, location?}`（style 支持 标题1-3/正文/引用/强调） |
| `insert_table` | 插入表格 `{rows}` ⚠️ **此环境 Word.js 表格插入 API 全部不可用（见环境边界）** |
| `insert_image` | 选区插图 `{base64, width?, height?}` |
| `apply_style` | 应用内置样式 `{style, scope: selection\|all}` |
| `format_selection` | 选区格式化 `{font, size, bold, italic, color, highlight}` |
| `set_paragraph_format` | 段落格式 `{alignment, indent, lineSpacing, listType}` ⚠️ **此环境 paragraphFormat 不可用** |
| `search` | 查找 `{query, matchCase?, wildcard?}` 返回命中列表 |
| `add_comment` | 选区加批注 ⚠️ **此环境 Word comments API 不可用** |
| `read_comments` | 列出批注 ⚠️ 同上 |
| `read_properties` | 文档属性（标题/作者/字数等） |

### Excel 组
| action | 说明 |
|---|---|
| `list_sheets` | 列出工作表（名称/位置/可见性） |
| `read_range` | 读区域 `{address: "Sheet1!A1:B10", limit?}`（值/公式/数字格式） |
| `write_range` | 批量写 `{address, values?/formulas?}`（二维数组，>5000 格自动分块） |
| `format_range` | 区域格式化 `{address, font, size, bold, fill, numberFormat, autoFit, tableStyle}` |
| `insert_chart` | 数据→图表 `{type: Column/Line/Bar/Pie/Area/Scatter/…, dataRange, title?}` |
| `add_sheet` / `rename_sheet` / `delete_sheet` | 工作表管理（delete 支持 dryRun） |
| `apply_sort` | 排序 `{address, fields: [{column, ascending}]}` |
| `apply_filter` | AutoFilter `{address, columns?}` |
| `evaluate_formula` | 公式求值 `{formula: "SUM(A1:A10)"}`（白名单 SUM/AVERAGE/COUNT/MAX/MIN/PRODUCT，workbook.functions 类型化求值） |
| `add_comment` / `read_comments` | 单元格批注（`cell` 可带表名前缀自动剥离） |
| `read_properties` | 工作簿属性 |

### PPT 组
| action | 说明 |
|---|---|
| `read_slides` | 当前选中页列表（SlideRange：id + title） |
| `ppt_read_notes` | 全文件备注（OOXML 解析 notesSlides + rels 映射） |
| `read_document` | 全文件逐页文本（通用） |

### 环境诊断
| action | 说明 |
|---|---|
| `get_environment` | 宿主版本/平台/requirementSets 支持情况 + Word 对象模型深层探测（用于定位环境边界） |

## 四、能力边界（实测结论，2026-08）

**归因澄清**：以下"不可用"分两类——
- **平台级不可能**（Office.js 规范无此 API，所有版本/机器都做不到）：PPT OOXML 写入、Word 页面设置/目录/剪贴板移动、窗格程序化刷新
- **本机运行时缺失**（requirementSets 声明支持 WordApi 1.8/ImageCoercion 1.1，但运行时对象属性实际缺失；非 node.js 问题、非 CDN 缓存——已验证）：Word 批注（body.comments 属性不存在）、paragraphFormat（属性不存在）、表格插入（insertTable/insertOoxml 存在但调用失败）。**换机器/更新 Office 很可能可用**，本项目如实返回 `requirement` 错误码降级

| 平台 | 可用 | 不可用及类别 |
|---|---|---|
| Word | 段落/文本插入替换、全文/表格读取、字体设置、样式应用、选区格式化、查找、文档属性、空段清理 | 表格插入 / paragraphFormat / 批注（**本机运行时缺失**）；页面设置/目录/剪贴板（**平台级**） |
| Excel | 工作表管理、区域读写（批量）、格式化、图表、排序、筛选、公式求值、批注、属性 | `workbook.getRange`/`calculate` 不存在（已用替代 API 绕过）；Range.autoFilter 须用 worksheet 级 |
| PPT | 全文件文本/备注读取、SlideRange | OOXML 写入、新建幻灯片/排版（**平台级**） |

**设计原则**：环境不支持的 API 返回 `code: requirement | unsupported | execution`，AI 层据此降级或如实告知，绝不假装成功。

## 五、安全护栏

- **破坏性操作 dryRun**：replace_all / remove_empty_paragraphs / delete_sheet 支持 `dryRun` 返回影响预览，AI 层默认先预览后执行
- **图片段落保护**：删除空段落时跳过含 inlinePictures 的段落（曾误删流程图，已修复）
- **文档结尾段保护**：Word 最后一段（段落标记）不可删除
- **性能护栏**：Excel 批量写分块（≤5000 格/批）、`getUsedRange(true)` 避免全列格式爆炸、大表读截断

## 六、服务托管（生产环境）

**推荐：Windows 计划任务「DSH Office Bridge」（登录自启，静默运行）**
- **一键安装**：`powershell -ExecutionPolicy Bypass -File install.ps1`（自动按当前目录注册，跨机器通用，无需改路径）
- 手动注册：触发器=用户登录时启动；Settings=常驻无时限、StartWhenAvailable；启动命令=`powershell -NoProfile -WindowStyle Hidden -Command "& '<node完整路径>' '<项目目录>\server.js'"`（静默，无窗口）
- 手动管理：
  - 启动：`Start-ScheduledTask -TaskName 'DSH Office Bridge'`
  - 停止：按端口找进程 `netstat -ano | findstr :3000` → `Stop-Process -Id <pid>`
  - 重启（改代码后）：停进程 → `Start-ScheduledTask -TaskName 'DSH Office Bridge'`

**开发时**：`powershell -ExecutionPolicy Bypass -File start.ps1` 前台运行；`npm start` 亦可。

**自检**：`powershell -ExecutionPolicy Bypass -File smoke-test.ps1`（检查服务/端点/在线文档）。

## 七、排查

| 现象 | 处理 |
|---|---|
| 窗格显示「桥接服务未连接」 | 确认 `node server.js` 在跑（`/office/status` 有响应） |
| DSH 指令报 timeout | 加载项没在线：确认文档打开 + 窗格显示已连接 |
| 窗格没出现 | 重新上传 manifest；确认 Office 未以管理员运行（localhost 例外要求普通权限） |
| 修改没生效 | 确认执行位置（光标/选区）正确；查看窗格执行日志 |
| 改了 actions.js 没生效 | 重启桥接服务（窗格自动热更新，无需重开窗格） |
| 改了 taskpane.js 没生效 | 需手动重开窗格（外壳代码无法程序化刷新，Office 桌面版限制） |

## 八、能力发现（AI 侧）

- `GET /office/capabilities` → action 注册表（名称/平台/是否破坏性/参数说明）
- `GET /office/status` → 在线文档（hosts）+ 窗格启动记录（hellos）
- `GET /office/actions-version` → actions.js 版本（热更新比对）

---

*非微软官方产品。`DSH` 是社区 AI 工具箱生态；本项目是 DSH 与 Microsoft Office 之间的独立桥接。*
*Vibe-coded：本项目由人类与 AI 结对协作开发（AI 辅助编码），全部功能经真实文档实测验证。*
