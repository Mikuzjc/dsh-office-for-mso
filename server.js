// server.js — DSH ↔ Office 加载项桥接服务
// 职责：
//   1) 托管加载项静态页面（taskpane.html/js/css），供 Office 通过 http://localhost:3000 加载
//   2) /office/command —— DSH 侧（agent 经 pwsh）POST 一条 Office 操作指令，同步等待加载项执行结果（最长 waitMs）
//   3) /office/poll   —— 加载项轮询取指令（带 ?host= 精确路由，避免多文档抢指令）
//   4) /office/result —— 加载项回传执行结果
//   5) /office/heartbeat / /office/status —— 心跳（按 host 记录）与状态（含在线文档列表）
// 多文档模型：Word / Excel / PowerPoint 各自运行一个加载项实例，指令可指定 host 精确路由；
//             不指定 host 时任一在线文档可取（兼容旧窗格）。
// 无第三方依赖（Node >= 18 自带 fetch）。
// 启动：node server.js   （端口可用环境变量 PORT 覆盖）

const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');

const PORT = Number(process.env.PORT || 3000);
const ROOT = __dirname;
const COMMAND_TIMEOUT_MS = Number(process.env.COMMAND_TIMEOUT_MS || 90000);
const VALID_HOSTS = ['Word', 'Excel', 'PowerPoint'];
// 覆盖类操作确认模式：auto=自动 re-read 后执行（结果附 previousState，不打扰用户）；ask=窗格显示审批面板，用户点确认才执行
let CONFIRM_MODE = process.env.OFFICE_CONFIRM_MODE || 'auto';

// ---- 能力注册表（与 taskpane.js 的 ACTIONS 保持一致；供 AI 层 /office/capabilities 发现能力）----
// 一期全集：现有 10 个 + 新增（Word 组 / Excel 组 / PPT 组）。destructive=true 的操作须先 dryRun 预览。
const CAPABILITIES = [
  // 通用（多平台）
  { name: 'read_selection', hosts: ['Word', 'Excel', 'PowerPoint'], destructive: false, args: { withStyles: 'boolean 可选：同时返回样式' } },
  { name: 'write_selection', hosts: ['Word', 'Excel', 'PowerPoint'], destructive: true, args: { text: 'string 必填：替换选区的文本' } },
  { name: 'read_document', hosts: ['Word', 'Excel', 'PowerPoint'], destructive: false, args: { sheet: 'string 可选(Excel)：指定工作表名' } },
  { name: 'read_styles', hosts: ['Word', 'Excel'], destructive: false, args: {} },
  { name: 'replace_all', hosts: ['Word', 'Excel'], destructive: true, args: { search: 'string 必填', replace: 'string 必填', dryRun: 'boolean 可选：仅返回命中数不执行' } },
  { name: 'append_text', hosts: ['Word'], destructive: false, args: { text: 'string 必填：文末追加段落' } },
  // Word 组
  { name: 'read_tables', hosts: ['Word'], destructive: false, args: {} },
  { name: 'set_font', hosts: ['Word'], destructive: true, args: { font: 'string 必填：字体名' } },
  { name: 'remove_empty_paragraphs', hosts: ['Word'], destructive: true, args: { dryRun: 'boolean 可选' } },
  { name: 'insert_paragraph', hosts: ['Word'], destructive: false, args: { text: 'string 必填', style: 'string 可选：内置样式如 Heading1', location: 'string 可选：end|start|afterSelection' } },
  { name: 'insert_table', hosts: ['Word'], destructive: false, args: { rows: 'array 必填：二维数组', location: 'string 可选' } },
  { name: 'insert_image', hosts: ['Word'], destructive: false, args: { base64: 'string 必填：PNG/JPEG base64', width: 'number 可选(px)', height: 'number 可选(px)' } },
  { name: 'apply_style', hosts: ['Word'], destructive: true, args: { style: 'string 必填：Heading1/标题 1/正文等', scope: 'string 可选：selection|all' } },
  { name: 'format_selection', hosts: ['Word', 'Excel'], destructive: true, args: { font: 'string 可选', size: 'number 可选', bold: 'boolean 可选', italic: 'boolean 可选', color: 'string 可选(#RRGGBB)', highlight: 'string 可选(Word)' } },
  { name: 'set_paragraph_format', hosts: ['Word'], destructive: true, args: { alignment: 'string 可选：left|center|right|justify', indent: 'number 可选(cm)', lineSpacing: 'number 可选', listType: 'string 可选：bulleted|numbered|none' } },
  { name: 'search', hosts: ['Word'], destructive: false, args: { query: 'string 必填', matchCase: 'boolean 可选', wildcard: 'boolean 可选：通配符' } },
  { name: 'add_comment', hosts: ['Word', 'Excel'], destructive: false, args: { text: 'string 必填', selectionText: 'string 可选(Word)：给含该文本的选区加批注', cell: 'string 可选(Excel)：A1 单元格' } },
  { name: 'read_comments', hosts: ['Word', 'Excel'], destructive: false, args: {} },
  { name: 'read_properties', hosts: ['Word', 'Excel'], destructive: false, args: {} },
  // Excel 组
  { name: 'list_sheets', hosts: ['Excel'], destructive: false, args: {} },
  { name: 'read_range', hosts: ['Excel'], destructive: false, args: { address: 'string 必填：Sheet!A1:B10', limit: 'number 可选：最大格数' } },
  { name: 'write_range', hosts: ['Excel'], destructive: true, args: { address: 'string 必填', values: 'array 可选：二维数组', formulas: 'array 可选：二维数组(A1 公式)' } },
  { name: 'format_range', hosts: ['Excel'], destructive: true, args: { address: 'string 必填', font: 'string 可选', size: 'number 可选', bold: 'boolean 可选', fill: 'string 可选(#RRGGBB)', numberFormat: 'string 可选', autoFit: 'boolean 可选', tableStyle: 'string 可选(如 TableStyleMedium2)' } },
  { name: 'insert_chart', hosts: ['Excel'], destructive: false, args: { type: 'string 必填：Line|Column|Bar|Pie|Area|Scatter 等', dataRange: 'string 必填：Sheet!A1:B10', title: 'string 可选' } },
  { name: 'add_sheet', hosts: ['Excel'], destructive: false, args: { name: 'string 必填' } },
  { name: 'rename_sheet', hosts: ['Excel'], destructive: true, args: { oldName: 'string 必填', newName: 'string 必填' } },
  { name: 'delete_sheet', hosts: ['Excel'], destructive: true, args: { name: 'string 必填', dryRun: 'boolean 可选' } },
  { name: 'apply_sort', hosts: ['Excel'], destructive: true, args: { address: 'string 必填', fields: 'array 必填：[{column:1,ascending:true}]' } },
  { name: 'apply_filter', hosts: ['Excel'], destructive: true, args: { address: 'string 必填', columns: 'array 可选：[{column,filterType,criteria}]' } },
  { name: 'evaluate_formula', hosts: ['Excel'], destructive: false, args: { formula: 'string 必填：如 SUM(A1:A10)' } },
  // PPT 组
  { name: 'read_slides', hosts: ['PowerPoint'], destructive: false, args: {} },
  { name: 'ppt_read_notes', hosts: ['PowerPoint'], destructive: false, args: {} },
  // 环境诊断
  { name: 'get_environment', hosts: ['Word', 'Excel', 'PowerPoint'], destructive: false, args: {} },
  // 定位 + 选中（零副作用：只做选中/取消选中 UI 反馈，不改文档内容/样式）
  { name: 'locate_select', hosts: ['Word', 'Excel'], destructive: false, args: { bookmark: 'string 可选：Word 书签 / Excel 命名区域名', anchor: 'string 可选：同 bookmark', text: 'string 可选：定位首个匹配文本', range: 'string 可选(Excel)：A1:B5 或 Sheet!A1:B5', address: 'string 可选(Excel)：同 range', sheet: 'string 可选(Excel)：指定工作表', blinks: 'number 可选：闪烁次数(默认0=只选中保持,1-5才闪)', interval: 'number 可选：间隔ms(默认300,100-800)' } },
];

// ---- 指令队列（一次只处理一条，先到先得）----
let pending = null;        // { commandId, action, args, host, resolve, timer }  host: null = 任意文档可取
const heartbeats = {};     // host -> 最近心跳时间戳（用于感知哪些文档在线）
const hellos = {};         // host -> { ts, version } 窗格启动上报（诊断自动刷新是否生效）

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
};

function sendJson(res, code, obj) {
  const body = JSON.stringify(obj);
  res.writeHead(code, { 'Content-Type': 'application/json; charset=utf-8', 'Content-Length': Buffer.byteLength(body) });
  res.end(body);
}

// 代码版本：taskpane.js 的 mtime（毫秒）。窗格心跳比对发现不一致 → 自动 reload，无需手动重开窗格。
function getCodeVersion() {
  try {
    const st = fs.statSync(path.join(ROOT, 'taskpane.js'));
    return String(st.mtimeMs);
  } catch (e) { return '0'; }
}

// actions 版本：actions.js 的 mtime（毫秒），与 /office/actions-version 一致
function getActionsVersion() {
  try {
    const st = fs.statSync(path.join(ROOT, 'actions.js'));
    return String(st.mtimeMs);
  } catch (e) { return '0'; }
}

// ---- 启动时自动 sideload 加载项（Windows：写入 HKCU WEF Developer 注册表，普通权限即可）----
// 让 `node server.js` 一步完成：启动服务 + 注册加载项。macOS/其他平台跳过（Office 菜单手动加载）。
function ensureSideloaded() {
  if (process.platform !== 'win32') {
    console.log('[dsh-office] 非 Windows 平台，跳过加载项自动注册（请用 Office 菜单手动加载）');
    return;
  }
  const { spawnSync } = require('node:child_process');
  const reg = (args) => spawnSync('reg', args, { stdio: 'ignore', windowsHide: true }).status === 0;
  try {
    const manifest = path.join(ROOT, 'manifest.xml');
    const xml = fs.readFileSync(manifest, 'utf8');
    const m = xml.match(/<Id>\s*([0-9a-fA-F-]{36})\s*<\/Id>/);
    if (!m) { console.log('[dsh-office] manifest.xml 缺少 <Id>，跳过自动注册'); return; }
    const guid = m[1];
    const key = 'HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer';
    if (reg(['query', key, '/v', guid])) {
      console.log(`[dsh-office] 加载项已注册 (WEF: ${guid})`);
      return;
    }
    if (reg(['add', key, '/v', guid, '/t', 'REG_SZ', '/d', manifest, '/f'])) {
      console.log('[dsh-office] 加载项已自动注册 → 关闭并重新打开 Office 文档后，窗格将出现');
    } else {
      console.log('[dsh-office] 加载项自动注册失败（可手动运行 sideload.ps1）');
    }
  } catch (e) {
    console.log('[dsh-office] 加载项自动注册跳过:', e.message);
  }
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    let data = '';
    req.on('data', (c) => { data += c; if (data.length > 1e6) { reject(new Error('body too large')); req.destroy(); } });
    req.on('end', () => { try { resolve(data ? JSON.parse(data) : {}); } catch (e) { reject(e); } });
    req.on('error', reject);
  });
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
  const p = url.pathname;

  // ---- 静态文件（office-addin 目录内任意文件；API 路径 /office/* 已单独处理）----
  if (req.method === 'GET' && !p.startsWith('/office/')) {
    const file = p === '/' ? 'taskpane.html' : p.slice(1);
    const full = path.join(ROOT, file);
    if ((full.startsWith(ROOT + path.sep) || full === ROOT) && fs.existsSync(full) && fs.statSync(full).isFile()) {
      const ext = path.extname(full).toLowerCase();
      res.writeHead(200, { 'Content-Type': MIME[ext] || 'application/octet-stream', 'Cache-Control': 'no-store' });
      if (file === 'taskpane.js') {
        // 前缀注入代码版本标记（须在文件头部，taskpane.js 顶部读取 window.__CODE_VERSION__）
        const src = fs.readFileSync(full, 'utf8');
        res.end(`;window.__CODE_VERSION__ = ${JSON.stringify(getCodeVersion())};\n` + src);
        return;
      }
      if (file === 'actions.js') {
        // 前缀注入 actions 版本标记：窗格热更新后可验证新脚本确实执行（防止加载成功但 IIFE 抛错/旧值残留）
        const src = fs.readFileSync(full, 'utf8');
        res.end(`;window.__ACTIONS_VERSION__ = ${JSON.stringify(getActionsVersion())};\n` + src);
        return;
      }
      fs.createReadStream(full).pipe(res);
      return;
    }
    sendJson(res, 404, { error: 'not found' });
    return;
  }

  // ---- CORS（Office 加载项页面与桥接服务同源 localhost:3000，跨主机时用）----
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') { res.writeHead(204); res.end(); return; }

  // POST /office/command —— DSH 侧发指令，同步等待结果；body.host 可选（Word/Excel/PowerPoint），缺省=任意文档
  if (req.method === 'POST' && p === '/office/command') {
    let body;
    try { body = await readBody(req); } catch { sendJson(res, 400, { ok: false, error: 'bad json' }); return; }
    if (!body.action || typeof body.action !== 'string') { sendJson(res, 400, { ok: false, error: 'action required' }); return; }
    const host = body.host && typeof body.host === 'string' ? body.host : null;
    if (host && !VALID_HOSTS.includes(host)) { sendJson(res, 400, { ok: false, error: `invalid host: ${host} (valid: ${VALID_HOSTS.join('/')})` }); return; }
    if (pending) { sendJson(res, 409, { ok: false, error: 'a command is already pending' }); return; }
    // 快速失败：目标 host 窗格不在线时立即返回明确错误（而不是傻等 90s 超时），AI 层据此提醒用户正确用法
    const now = Date.now();
    const OFFLINE_MS = 15000; // 心跳每 5s 一次，15s 无心跳视为窗格离线
    if (host && (!heartbeats[host] || now - heartbeats[host] > OFFLINE_MS)) {
      sendJson(res, 200, { ok: false, code: 'addin_offline', error: `「${host}」加载项窗格未开启：请打开该文档，并在 开始/开发人员 → 加载项 → 开发人员加载项 中打开「DSH Office 执行器」窗格并保持开启，然后重试` });
      return;
    }
    if (!host && Object.keys(heartbeats).length === 0) {
      sendJson(res, 200, { ok: false, code: 'addin_offline', error: '没有在线的 Office 加载项：请打开 Word/Excel/PowerPoint 文档，并在 开始/开发人员 → 加载项 → 开发人员加载项 中打开「DSH Office 执行器」窗格并保持开启，然后重试' });
      return;
    }
    const commandId = crypto.randomUUID();
    const result = await new Promise((resolve) => {
      const timer = setTimeout(() => resolve({ ok: false, error: 'timeout: no add-in heartbeat or execution', timeout: true }), COMMAND_TIMEOUT_MS);
      pending = { commandId, action: body.action, args: body.args || {}, host, resolve: (r) => { clearTimeout(timer); resolve(r); } };
    });
    if (pending && pending.commandId === commandId) pending = null; // 清理（超时路径）
    sendJson(res, 200, { ok: result.ok !== false, commandId, host: host || 'any', ...result });
    return;
  }

  // GET /office/poll —— 加载项取指令；?host= 精确匹配：指令指定了 host 时只有同 host 的窗格能取走
  if (req.method === 'GET' && p === '/office/poll') {
    const host = url.searchParams.get('host') || '';
    if (pending && (!pending.host || pending.host === host)) {
      sendJson(res, 200, { commandId: pending.commandId, action: pending.action, args: pending.args });
    } else {
      sendJson(res, 200, { commandId: null });
    }
    return;
  }

  // POST /office/result —— 加载项回传结果
  if (req.method === 'POST' && p === '/office/result') {
    let body;
    try { body = await readBody(req); } catch { sendJson(res, 400, { ok: false, error: 'bad json' }); return; }
    if (pending && pending.commandId === body.commandId) {
      const r = pending.resolve;
      pending = null;
      // 错误也转发 code/result（如 confirm_required 的 result.preview），否则 AI 拿不到预览
      r(body.ok === false
        ? { ok: false, error: body.error || 'add-in error', code: body.code, result: body.result }
        : { ok: true, result: body.result });
      sendJson(res, 200, { ok: true });
    } else {
      sendJson(res, 404, { ok: false, error: 'unknown or expired commandId' });
    }
    return;
  }

  // POST /office/heartbeat —— body.host 记录该文档在线；响应带 codeVersion 供窗格自动刷新
  if (req.method === 'POST' && p === '/office/heartbeat') {
    let body;
    try { body = await readBody(req); } catch { body = {}; }
    const host = body.host && typeof body.host === 'string' && VALID_HOSTS.includes(body.host) ? body.host : null;
    if (host) heartbeats[host] = Date.now();
    sendJson(res, 200, { ok: true, host, codeVersion: getCodeVersion() });
    return;
  }

  // POST /office/hello —— 窗格启动上报（每次 JS 加载执行一次；用于验证自动刷新/reload 是否生效）
  if (req.method === 'POST' && p === '/office/hello') {
    let body;
    try { body = await readBody(req); } catch { body = {}; }
    const host = body.host && typeof body.host === 'string' && VALID_HOSTS.includes(body.host) ? body.host : null;
    if (host) hellos[host] = { ts: Date.now(), version: body.version || '' };
    sendJson(res, 200, { ok: true });
    return;
  }

  // GET /office/status —— 状态：pending 指令 + 各在线文档（按 host 心跳）+ 窗格启动记录
  if (req.method === 'GET' && p === '/office/status') {
    const hosts = {};
    for (const h of VALID_HOSTS) if (heartbeats[h]) hosts[h] = heartbeats[h];
    sendJson(res, 200, {
      pending: pending ? { commandId: pending.commandId, action: pending.action, host: pending.host || 'any' } : null,
      hosts,
      hellos,
      codeVersion: getCodeVersion(),
      lastHeartbeat: Object.keys(heartbeats).length ? Math.max(...Object.values(heartbeats)) : 0,
    });
    return;
  }

  // GET /office/actions-version —— actions.js 版本（mtime），窗格据此热更新实现
  if (req.method === 'GET' && p === '/office/actions-version') {
    let v = '0';
    try { v = String(fs.statSync(path.join(ROOT, 'actions.js')).mtimeMs); } catch (e) { /* 文件缺失 */ }
    sendJson(res, 200, { version: v });
    return;
  }

  // GET/POST /office/config —— 运行时配置（confirmMode: auto|ask，窗格开关同步用）
  if (req.method === 'GET' && p === '/office/config') {
    sendJson(res, 200, { confirmMode: CONFIRM_MODE });
    return;
  }
  if (req.method === 'POST' && p === '/office/config') {
    let body;
    try { body = await readBody(req); } catch { body = {}; }
    if (body.confirmMode === 'auto' || body.confirmMode === 'ask') {
      CONFIRM_MODE = body.confirmMode;
      sendJson(res, 200, { ok: true, confirmMode: CONFIRM_MODE });
    } else {
      sendJson(res, 400, { ok: false, error: 'confirmMode must be "auto" or "ask"' });
    }
    return;
  }

  // GET /office/capabilities —— 能力发现：返回 action 注册表（名称/平台/是否破坏性/参数说明）
  if (req.method === 'GET' && p === '/office/capabilities') {
    sendJson(res, 200, { version: '1.1.0', actions: CAPABILITIES });
    return;
  }

  sendJson(res, 404, { error: 'unknown endpoint' });
});

server.listen(PORT, '127.0.0.1', () => {
  console.log(`[dsh-office-bridge] listening on http://127.0.0.1:${PORT}`);
  console.log(`  add-in taskpane: http://127.0.0.1:${PORT}/taskpane.html`);
  console.log(`  command endpoint: POST /office/command  (timeout ${COMMAND_TIMEOUT_MS}ms, host-routed)`);
  ensureSideloaded();
});
