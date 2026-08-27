# dsh-office-for-mso

> [中文](README.md) · [日本語](README.ja.md) · English

**A plugin/skill for DeepSeek Harness (DSH) (v1.3)**: Issue commands from a DSH session to control your **currently open** Word / Excel / PowerPoint documents through a **Microsoft Office** add-in — reading, writing, formatting, structural operations, charts/formulas/comments — approaching common Microsoft Copilot for Office workflows. **(In this project, "Office" means Microsoft Office, not WPS or other compatible suites.)**

```
You ──DSH session──▶ AI(agent) ──POST──▶ Bridge service localhost:3000
                                              │ Command queue (host-routed)
                         Office add-in (background polling, 1s)
                                              │ Office.js execution (34 actions)
                                              ▼
                        Result returned ──▶ AI reads ──▶ Report to you
```

- **No conflicts**: the add-in runs inside the Office process and operates on the in-memory document; Office serializes it with your own editing (no file locks, no "last-save-wins")
- **No taskpane interaction needed**: the pane is just a status display; all commands come from DSH
- **🔥 Hot-reload (selling point)**: all action implementations live in `actions.js`; edit it → restart the bridge → panes **hot-reload automatically**, **no pane reopen needed** — customize actions to your needs and DSH picks them up instantly (re-read/approval/punctuation rules in this project were all shipped this way)

---

## 1. Setup (one-time)

### 0. Prerequisites

- **This is a plugin/skill for [DeepSeek Harness (DSH)](https://github.com/search?q=deepseek+harness)**: you must install DSH (DeepSeek Harness, a Node.js-based AI session environment) first, and use this skill from a DSH session
- **Node.js ≥ 18** (required by both DSH and the bridge service)
- **Microsoft Office desktop** (Word / Excel / PowerPoint, Windows or macOS)

### 1.1 Clone and install (recommended: one-click persistent service)

```powershell
git clone https://github.com/Mikuzjc/dsh-office-for-mso.git
cd dsh-office-for-mso
npm run setup   # one-click: register Scheduled Task (auto-start at logon, silent persistent) + start service + auto-register add-in (first run shows a UAC prompt — click Yes)
```

> **`npm run setup` is the official install**: the service is hosted by the Windows Scheduled Task — **auto-starts at logon, runs persistently in the background, closing the terminal does not affect it**.
> For a quick preview only: `node server.js` (foreground; closing the terminal stops the service).

**On Windows the add-in is auto-registered on startup (WEF registry)** — on first run **close and reopen the Office document** and the pane appears; on macOS load manually via the Office menu (see 1.2/1.3).
(All path examples below refer to your actual project directory.)

### 1.2 Sideload the add-in into Microsoft Office

> **Platforms**: the core (server.js + add-in) only requires Node.js and runs on Windows / macOS (Office desktop on macOS supports add-in sideload too); `install.ps1`/Scheduled Task is an **optional** Windows-only auto-hosting — on macOS just run `node server.js` manually.

> **Windows users: no manual sideload needed** — `node server.js` (or `npm run setup`) **auto-registers** the add-in on startup (WEF registry, normal user rights). On first run **close and reopen the Office document** and the pane appears.
> Only if you need manual management:
> ```powershell
> cd <your-project-dir>
> powershell -ExecutionPolicy Bypass -File sideload.ps1   # manual register (-Remove to unregister)
> ```

**macOS users**: no auto-registration; load manually via the Office menu (Developer Add-ins flow below):

1. Enable the Developer tab: **File → Options → Customize Ribbon → check "Developer" under Main Tabs** → OK
2. Open any Word / Excel / PowerPoint document
3. **Developer tab → Add-ins** (or **Insert → My Add-ins**) → open the "Office Add-ins" dialog
4. At the bottom-left of the dialog, in the **Manage** dropdown select **Developer Add-ins**
5. Click **+ (Add)** → select `<your-project-dir>\manifest.xml`
6. The **DSH Office Executor** pane appears in the sidebar, showing **Connected: waiting for DSH commands**

Then just **keep the document open**; the pane can be minimized/dragged to a corner.

> Older Office versions with the "Upload My Add-in" entry can use it directly.
> Moving to another machine: repeat the two steps above (bridge service + upload manifest).

### 1.3 First use

After installation, you need to **open the add-in once** before DSH can operate on documents:

1. Open a Word / Excel / PowerPoint document
2. **Home (or Developer) tab → Add-ins → Developer Add-ins → "DSH Office Executor"**
3. Once the pane appears (showing **Connected: waiting for DSH commands**), you can issue commands from the DSH session
4. The pane can be resized small / dragged to a corner, **but it must stay open while in use** — it hosts the Office.js execution; closing it makes the document unreachable by DSH

Keep the **document open + pane open** afterwards; if the pane is closed accidentally, DSH will report `addin_offline` — repeat step 2 to reopen it.

### 1.4 Make DSH actively use this skill (important)

DSH only auto-invokes skills installed in its **skill library** (`~/.agents/skills/`) — cloning the repo alone does not make DSH use it. Install the SKILL.md into the library:

```powershell
# Windows: copy into the DSH skill library
mkdir "$HOME\.agents\skills\office-bridge" -Force | Out-Null
copy "skills\office-bridge\SKILL.md" "$HOME\.agents\skills\office-bridge\"
```

Or create an `office-bridge` skill in the DSH settings panel's skill manager (content in `skills/office-bridge/SKILL.md`). Once installed, the main model sees the skill description and will automatically call the bridge for requests like "read the Word document" / "generate a chart in Excel" / "translate the selection".

## 2. Architecture

| File | Responsibility |
|---|---|
| `server.js` | Bridge service: command queue, host routing, heartbeats, static files, capability discovery `/office/capabilities`, actions version `/office/actions-version` |
| `taskpane.js` | Add-in shell: polling/heartbeat/dispatch/hot-reload loading (**change rarely**; changes require reopening the pane) |
| `actions.js` | All action implementations + registry (**hot-reload unit**: editing it does not require reopening panes) |
| `pako.min.js` | Local zip decompression library (for PPT OOXML reading, offline) |
| `manifest.xml` | Add-in manifest (permission ReadWriteDocument, the maximum) |

**Multi-document model**: Word / Excel / PowerPoint each run their own add-in instance; commands carry a `host` (Word/Excel/PowerPoint) for precise routing; `GET /office/status` returns the online document list (`hosts` field), **per-instance details (`instances` field: instanceId / host / docUrl document path / docTitle document name — used to identify the target document; fall back to docTitle when docUrl is empty)** and pane startup records (`hellos`).

**Hot-reload mechanism**: before each poll, the pane GETs `/office/actions-version` and compares against the mtime of `actions.js`; on change it dynamically reloads the script. **Edit `actions.js` → restart server → takes effect automatically**.

## 3. Capability matrix (34 actions)

> Operations with `destructive=true` support `args.dryRun` to preview the impact (implemented for replace_all / remove_empty_paragraphs / delete_sheet; others follow a read-then-write convention at the AI layer). W=Word, E=Excel, P=PowerPoint.
> **Auto-select after writes** (zero side effects, selection only): write actions select the changed content on success — `replace_all` selects the last change / `append_text`, `insert_paragraph` select the inserted content / `write_range` selects the written range / `write_selection` keeps selection via the host; multiple disjoint changes can only select one range (Office.js single-selection limit).

### Common
| action | Platforms | Description |
|---|---|---|
| `read_selection` | W/E/P | Read current selection text; `withStyles=true` also returns styles |
| `write_selection` | W/E/P | Replace selection with `{text}` |
| `read_document` | W/E/P | Word full text; Excel used range of a worksheet (`sheet` to pick, 5000-cell cap); PPT full file per-slide text |
| `read_styles` | W/E | Selection styles: Word (font/size/bold/italic/color/underline/highlight); Excel (per cell, max 10×10) |
| `replace_all` | W/E | Find & replace across document `{search, replace, dryRun?}` |
| `append_text` | W | Append paragraph at end `{text}` |
| `locate_select` | W/E | Locate and select (zero side effects, no content/style changes): `{text}` first match / `{bookmark\|anchor}` / Excel `{range\|address}` (optional `sheet`); `blinks>0` to blink, default selects and holds |

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
- Manual: Trigger = on user logon; Settings = no time limit (persistent), StartWhenAvailable; launch command = `powershell -NoProfile -WindowStyle Hidden -Command "& '<node-full-path>' '<project-dir>\server.js'"` (silent, no window)
- Manual management:
  - Start: `Start-ScheduledTask -TaskName 'DSH Office Bridge'`
  - Stop: find process by port `netstat -ano | findstr :3000` → `Stop-Process -Id <pid>`
  - Restart (after code changes): stop process → `Start-ScheduledTask -TaskName 'DSH Office Bridge'`

**During development**: `powershell -ExecutionPolicy Bypass -File start.ps1` (foreground); or `npm start`.

**Self-check**: `powershell -ExecutionPolicy Bypass -File smoke-test.ps1` (checks service/endpoints/online documents).

## 6.5 Updating (installed users)

Pull the latest version and apply it (one command):

```powershell
cd <your-project-dir>
powershell -ExecutionPolicy Bypass -File update.ps1   # git pull + auto-restart service
```

- **actions.js changes**: panes **hot-reload** after the service restart — no pane reopen needed
- **Shell (taskpane.js/html) changes**: reopen the pane once
- **server.js / install.ps1 changes**: update.ps1 restarts the service; if the Scheduled Task definition changed, rerun `install.ps1`

> No auto-push mechanism: an installed user runs `update.ps1` once to receive all updates (currently no other installers; evolves with the repo).

## 7. Troubleshooting

| Symptom | Fix |
|---|---|
| Pane shows "Bridge service not connected" | Confirm `node server.js` is running (`/office/status` responds) |
| DSH command times out | Add-in offline: confirm document open + pane shows Connected |
| DSH command returns `addin_offline` immediately | Pane not open: open the document → Home/Developer → Add-ins → Developer Add-ins → open the "DSH Office Executor" pane and keep it open |
| Pane does not appear | Re-upload manifest; confirm Office is not running as administrator (localhost exception requires normal privileges) |
| Changes not applied | Confirm cursor/selection position; check the pane execution log |
| actions.js changes not applied | Restart the bridge (panes hot-reload, no pane reopen needed) |
| taskpane.js changes not applied | Reopen the pane manually (shell code cannot be programmatically refreshed; Office desktop limitation) |

## 8. Capability discovery (AI side)

- `GET /office/capabilities` → action registry (name/platform/destructive/args description)
- `GET /office/status` → online documents (hosts) + pane startup records (hellos)
- `GET /office/actions-version` → actions.js version (hot-reload comparison)

## 9. AI usage conventions (for DSH and other AI callers)

- Before sending a command, `GET /office/status` first: confirm the target host is online (`hosts` contains the document with a fresh heartbeat) before dispatching
- On `code: addin_offline`, **do not retry** — remind the user: open the target document → Home/Developer → Add-ins → Developer Add-ins → open the "DSH Office Executor" pane and keep it open
- The pane must stay open for operations to run (it can be resized/dragged, but not closed)

## 10. Error codes at a glance (AI side)

Every error returns `{ok:false, code, error}` (`error` is a human-readable reason); the full enumeration — including per-code AI handling advice and whether retrying is allowed — is always available via `GET /office/errors`:

| code | meaning | AI handling |
|---|---|---|
| `instance_required` | missing `instance` (required for multi-document routing) | check status, get the instanceId, resend |
| `addin_offline` | pane not open / offline | ask the user to open the pane, **do not retry** |
| `busy` | previous command still running | wait briefly and resend |
| `timeout` | no result within 90s | check status; retry once if online, otherwise ask the user to reopen the pane |
| `bad_json` / `bad_args` | invalid body / invalid args | fix per `error` and resend |
| `unknown_action` | action name does not exist | check capabilities and use the correct name |
| `unsupported_host` | action not supported in this host | use a supported action/application |
| `confirm_required` | approval needed (ask mode) | show `result.preview` to the user, resend with `confirm:true` after approval |
| `rejected` | user rejected in the pane / approval timed out | stop, **do not resend** |
| `not_found` | locate/search target does not exist | tell the user honestly, **do not retry the same lookup** |
| `requirement` / `unsupported` | API missing in this environment | say so honestly, never pretend success |
| `execution` | runtime exception | report the `error` verbatim |

**"No match" is not an error**: `search` / `replace_all(dryRun)` with zero hits return `ok:true + count:0`, and empty `read_*` / `list_*` results return empty arrays/strings — all are successful responses; the AI should tell the user "not found" instead of retrying.

---

*Not an official Microsoft product. `DSH` is a community AI-harness ecosystem; this project is an independent bridge between DSH and Microsoft Office.*
*Vibe-coded: developed collaboratively by humans and AI (AI-assisted coding), with every feature verified against real documents.*
