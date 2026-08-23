# dsh-office-for-mso

> [中文](README.md) · [日本語](README.ja.md) · English

**DSH ↔ Office bridge executor (v1.3)**: Issue commands from a DSH session to control your **currently open** Word / Excel / PowerPoint documents through an Office add-in — reading, writing, formatting, structural operations, charts/formulas/comments — approaching common Microsoft Copilot for Office workflows.

```
You ──DSH session──▶ AI(agent) ──POST──▶ Bridge service localhost:3000
                                              │ Command queue (host-routed)
                         Office add-in (background polling, 1s)
                                              │ Office.js execution (33 actions)
                                              ▼
                        Result returned ──▶ AI reads ──▶ Report to you
```

- **No conflicts**: the add-in runs inside the Office process and operates on the in-memory document; Office serializes it with your own editing (no file locks, no "last-save-wins")
- **No taskpane interaction needed**: the pane is just a status display; all commands come from DSH
- **Hot-reload**: all action implementations live in `actions.js`; edit it → restart the bridge → panes reload automatically, **no need to reopen panes**

---

## 1. Setup (one-time)

### 1.1 Clone the repo and start the bridge service

```powershell
git clone https://github.com/Mikuzjc/dsh-office-for-mso.git
cd dsh-office-for-mso
node server.js   # or: powershell -ExecutionPolicy Bypass -File start.ps1
```

Wait for `listening on http://127.0.0.1:3000`.
(All path examples below refer to your actual project directory.)

### 1.2 Sideload the add-in into Office

> **Platforms**: the core (server.js + add-in) only requires Node.js and runs on Windows / macOS (Office desktop on macOS supports add-in sideload too); `install.ps1`/Scheduled Task is an **optional** Windows-only auto-hosting — on macOS just run `node server.js` manually.

On newer Office (2024+, Microsoft has hidden the "Upload My Add-in" entry) use the **Developer Add-ins** flow:

1. Enable the Developer tab: **File → Options → Customize Ribbon → check "Developer" under Main Tabs** → OK
2. Open any Word / Excel / PowerPoint document
3. **Developer tab → Add-ins** (or **Insert → My Add-ins**) → open the "Office Add-ins" dialog
4. At the bottom-left of the dialog, in the **Manage** dropdown select **Developer Add-ins**
5. Click **+ (Add)** → select `<your-project-dir>\manifest.xml`
6. The **DSH Office Executor** pane appears in the sidebar, showing **Connected: waiting for DSH commands**

Then just **keep the document open**; the pane can be minimized/dragged to a corner.

> Older Office versions with the "Upload My Add-in" entry can use it directly.
> Moving to another machine: repeat the two steps above (bridge service + upload manifest).

## 2. Architecture

| File | Responsibility |
|---|---|
| `server.js` | Bridge service: command queue, host routing, heartbeats, static files, capability discovery `/office/capabilities`, actions version `/office/actions-version` |
| `taskpane.js` | Add-in shell: polling/heartbeat/dispatch/hot-reload loading (**change rarely**; changes require reopening the pane) |
| `actions.js` | All action implementations + registry (**hot-reload unit**: editing it does not require reopening panes) |
| `pako.min.js` | Local zip decompression library (for PPT OOXML reading, offline) |
| `manifest.xml` | Add-in manifest (permission ReadWriteDocument, the maximum) |

**Multi-document model**: Word / Excel / PowerPoint each run their own add-in instance; commands carry a `host` (Word/Excel/PowerPoint) for precise routing; `GET /office/status` returns the online document list (`hosts` field) and pane startup records (`hellos`).

**Hot-reload mechanism**: before each poll, the pane GETs `/office/actions-version` and compares against the mtime of `actions.js`; on change it dynamically reloads the script. **Edit `actions.js` → restart server → takes effect automatically**.

## 3. Capability matrix (33 actions)

> Operations with `destructive=true` support `args.dryRun` to preview the impact (implemented for replace_all / remove_empty_paragraphs / delete_sheet; others follow a read-then-write convention at the AI layer). W=Word, E=Excel, P=PowerPoint.

### Common
| action | Platforms | Description |
|---|---|---|
| `read_selection` | W/E/P | Read current selection text; `withStyles=true` also returns styles |
| `write_selection` | W/E/P | Replace selection with `{text}` |
| `read_document` | W/E/P | Word full text; Excel used range of a worksheet (`sheet` to pick, 5000-cell cap); PPT full file per-slide text |
| `read_styles` | W/E | Selection styles: Word (font/size/bold/italic/color/underline/highlight); Excel (per cell, max 10×10) |
| `replace_all` | W/E | Find & replace across document `{search, replace, dryRun?}` |
| `append_text` | W | Append paragraph at end `{text}` |

### Word group
| action | Description |
|---|---|
| `read_tables` | Structurally read all tables (per cell, split by \t/\r from getRange().text) |
| `set_font` | Set font for whole document (incl. table paragraphs) `{font}` |
| `remove_empty_paragraphs` | Remove empty paragraphs (**skips image paragraphs and the final document paragraph**, `dryRun` preview) |
| `insert_paragraph` | Insert paragraph `{text, style?, location?}` (styles: Heading1-3/Body/Quote/Strong) |
| `insert_table` | Insert table `{rows}` ⚠️ **table-insertion APIs unavailable in this environment (see capability limits)** |
| `insert_image` | Insert image at selection `{base64, width?, height?}` |
| `apply_style` | Apply built-in style `{style, scope: selection\|all}` |
| `format_selection` | Format selection `{font, size, bold, italic, color, highlight}` |
| `set_paragraph_format` | Paragraph format `{alignment, indent, lineSpacing, listType}` ⚠️ **paragraphFormat unavailable in this environment** |
| `search` | Search `{query, matchCase?, wildcard?}` returning hits list |
| `add_comment` | Add comment to selection ⚠️ **Word comments API unavailable in this environment** |
| `read_comments` | List comments ⚠️ same as above |
| `read_properties` | Document properties (title/author/word count etc.) |

### Excel group
| action | Description |
|---|---|
| `list_sheets` | List worksheets (name/position/visibility) |
| `read_range` | Read range `{address: "Sheet1!A1:B10", limit?}` (values/formulas/number formats) |
| `write_range` | Batch write `{address, values?/formulas?}` (2D array, auto-chunked >5000 cells) |
| `format_range` | Format range `{address, font, size, bold, fill, numberFormat, autoFit, tableStyle}` |
| `insert_chart` | Data→chart `{type: Column/Line/Bar/Pie/Area/Scatter/…, dataRange, title?}` |
| `add_sheet` / `rename_sheet` / `delete_sheet` | Worksheet management (delete supports dryRun) |
| `apply_sort` | Sort `{address, fields: [{column, ascending}]}` |
| `apply_filter` | AutoFilter `{address, columns?}` |
| `evaluate_formula` | Evaluate formula `{formula: "SUM(A1:A10)"}` (whitelist SUM/AVERAGE/COUNT/MAX/MIN/PRODUCT, typed evaluation via workbook.functions) |
| `add_comment` / `read_comments` | Cell comments (`cell` may carry a sheet prefix, auto-stripped) |
| `read_properties` | Workbook properties |

### PPT group
| action | Description |
|---|---|
| `read_slides` | Currently selected slides (SlideRange: id + title) |
| `ppt_read_notes` | Full-file speaker notes (OOXML notesSlides parsing + rels mapping) |
| `read_document` | Full-file per-slide text (common) |

### Environment diagnostics
| action | Description |
|---|---|
| `get_environment` | Host version/platform/requirementSets support + deep probe of the Word object model (for locating capability limits) |

## 4. Capability limits (measured, 2026-08)

**Attribution**: the "unavailable" items fall into two categories —
- **Platform-impossible** (no such API in the Office.js spec; no version/machine can do it): PPT OOXML writing, Word page setup/TOC/clipboard moves, programmatic pane refresh
- **Runtime-missing on this machine** (requirementSets declare WordApi 1.8 / ImageCoercion 1.1 support, but the runtime object properties are actually absent; NOT a node.js issue, NOT a CDN cache issue — verified): Word comments (`body.comments` property absent), `paragraphFormat` (property absent), table insertion (`insertTable`/`insertOoxml` exist but calls fail). **Likely to work on other machines/newer Office**; this project honestly returns `requirement` error codes and degrades gracefully

| Platform | Available | Unavailable & category |
|---|---|---|
| Word | paragraph/text insert & replace, full text/table reading, font, styles, selection formatting, search, document properties, empty-paragraph cleanup | table insertion / paragraphFormat / comments (**runtime-missing**); page setup/TOC/clipboard (**platform-level**) |
| Excel | worksheet management, range read/write (batch), formatting, charts, sort, filter, formula evaluation, comments, properties | `workbook.getRange`/`calculate` absent (bypassed with alternatives); Range.autoFilter requires worksheet-level API |
| PPT | full-file text/notes reading, SlideRange | OOXML writing, adding slides/layout (**platform-level**) |

**Design principle**: unsupported APIs return `code: requirement | unsupported | execution`; the AI layer degrades or reports honestly — never fakes success.

## 5. Safety guardrails

- **dryRun for destructive ops**: replace_all / remove_empty_paragraphs / delete_sheet support `dryRun` impact preview; the AI layer previews before executing
- **Image paragraph protection**: empty-paragraph removal skips paragraphs containing inlinePictures (a flowchart was once deleted by accident — fixed)
- **Final paragraph protection**: Word's last paragraph (paragraph mark) cannot be deleted
- **Performance guardrails**: Excel batch writes chunked (≤5000 cells/batch), `getUsedRange(true)` avoids whole-column format blowups, large reads truncated

## 6. Service hosting (production)

**Recommended: Windows Scheduled Task `DSH Office Bridge` (auto-start at logon, silent)**
- **One-click install**: `powershell -ExecutionPolicy Bypass -File install.ps1` (registers automatically using the current directory — machine-agnostic, no path editing needed)
- Manual: Trigger = on user logon; Settings = no time limit (persistent), StartWhenAvailable; launch command = `wscript <project-dir>\run-hidden.vbs "node <project-dir>\server.js"` (silent, no window flash)
- Manual management:
  - Start: `Start-ScheduledTask -TaskName 'DSH Office Bridge'`
  - Stop: find process by port `netstat -ano | findstr :3000` → `Stop-Process -Id <pid>`
  - Restart (after code changes): stop process → `Start-ScheduledTask -TaskName 'DSH Office Bridge'`

**During development**: `powershell -ExecutionPolicy Bypass -File start.ps1` (foreground); or `npm start`.

**Self-check**: `powershell -ExecutionPolicy Bypass -File smoke-test.ps1` (checks service/endpoints/online documents).

## 7. Troubleshooting

| Symptom | Fix |
|---|---|
| Pane shows "Bridge service not connected" | Confirm `node server.js` is running (`/office/status` responds) |
| DSH command times out | Add-in offline: confirm document open + pane shows Connected |
| Pane does not appear | Re-upload manifest; confirm Office is not running as administrator (localhost exception requires normal privileges) |
| Changes not applied | Confirm cursor/selection position; check the pane execution log |
| actions.js changes not applied | Restart the bridge (panes hot-reload, no pane reopen needed) |
| taskpane.js changes not applied | Reopen the pane manually (shell code cannot be programmatically refreshed; Office desktop limitation) |

## 8. Capability discovery (AI side)

- `GET /office/capabilities` → action registry (name/platform/destructive/args description)
- `GET /office/status` → online documents (hosts) + pane startup records (hellos)
- `GET /office/actions-version` → actions.js version (hot-reload comparison)

---

*Not an official Microsoft product. `DSH` is a community AI-harness ecosystem; this project is an independent bridge between DSH and Microsoft Office.*
*Vibe-coded: developed collaboratively by humans and AI (AI-assisted coding), with every feature verified against real documents.*
