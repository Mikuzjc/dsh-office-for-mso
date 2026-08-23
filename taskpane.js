// taskpane.js — DSH Office 执行器外壳（v1.3，actions 热更新架构）
// 职责：窗格 UI、桥接轮询/心跳、execute 分发、actions.js 热更新加载。
// 所有 action 实现都在 actions.js（可热更新）：改 actions.js → 重启桥接服务 → 窗格自动重新加载，无需重开窗格。
// 本文件（外壳）应尽量少改；改动本文件仍需手动重开窗格（无法程序化刷新 Office 桌面版窗格）。

const BRIDGE = 'http://127.0.0.1:3000';
const POLL_MS = 1000;
const HEARTBEAT_MS = 5000;

const statusEl = document.getElementById('status');
const logEl = document.getElementById('log');
const hostEl = document.getElementById('host');

function setStatus(cls, text) {
  statusEl.className = 'status ' + cls;
  statusEl.textContent = text;
}
function log(line) {
  const t = new Date().toLocaleTimeString('zh-CN', { hour12: false });
  logEl.textContent = `[${t}] ${line}\n` + logEl.textContent;
  if (logEl.textContent.length > 4000) logEl.textContent = logEl.textContent.slice(0, 4000);
}

async function api(path, opts) {
  const r = await fetch(BRIDGE + path, opts);
  if (!r.ok) throw new Error(`HTTP ${r.status}`);
  return r.json();
}

function hostName() {
  return String(Office.context.host || '');
}

function okResult(result) { return { ok: true, result }; }
function errResult(code, error) { return { ok: false, code, error }; }

// ================= 公共辅助（暴露给 actions.js 使用） =================

function getSelectionTextAsync() {
  return new Promise((resolve) => {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
      (r) => resolve(r.status === Office.AsyncResultStatus.Succeeded
        ? okResult({ text: r.value || '' })
        : errResult('execution', r.error && r.error.message)));
  });
}

function setSelectionTextAsync(text) {
  return new Promise((resolve) => {
    Office.context.document.setSelectedDataAsync(String(text), { coercionType: Office.CoercionType.Text },
      (r) => resolve(r.status === Office.AsyncResultStatus.Succeeded
        ? okResult({ replaced: String(text).length })
        : errResult('execution', r.error && r.error.message)));
  });
}

// Word 内置样式名映射（小写去空格）
function mapBuiltInStyle(name) {
  const n = String(name || '').toLowerCase().replace(/[\s-]/g, '');
  const map = {
    '标题': 'title', title: 'title', '副标题': 'subtitle', subtitle: 'subtitle',
    '标题1': 'heading1', heading1: 'heading1', '标题2': 'heading2', heading2: 'heading2',
    '标题3': 'heading3', heading3: 'heading3', '正文': 'normal', normal: 'normal',
    '引用': 'quote', quote: 'quote', '强调': 'strong', strong: 'strong',
  };
  return map[n] || null;
}

function mapAlignment(a) {
  const m = { left: Word.Alignment.left, center: Word.Alignment.center, right: Word.Alignment.right, justify: Word.Alignment.justified };
  return m[String(a || '').toLowerCase()] || null;
}

// Excel：把 "Sheet1!A1:B10" 或 "A1:B10" 地址解析为 Range 对象（Workbook 无 getRange，需经 Worksheet）
function excelRange(ctx, address) {
  const addr = String(address || '');
  const bang = addr.lastIndexOf('!');
  if (bang >= 0) {
    const sheetName = addr.slice(0, bang);
    const rangePart = addr.slice(bang + 1);
    return ctx.workbook.worksheets.getItem(sheetName).getRange(rangePart);
  }
  return ctx.workbook.worksheets.getActiveWorksheet().getRange(addr);
}

// PowerPoint：getFileAsync 全文件拉取（分片 → 回调提取器）
function readPptFile(extractor) {
  return new Promise((resolve) => {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, (fileResult) => {
      if (fileResult.status !== Office.AsyncResultStatus.Succeeded) {
        return resolve(errResult('execution', fileResult.error && fileResult.error.message));
      }
      const file = fileResult.value;
      const chunks = [];
      let received = 0;
      const done = (out) => { try { file.closeAsync(() => resolve(out)); } catch (e) { resolve(out); } };
      const onSlice = (sliceResult) => {
        if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
          return done(errResult('execution', sliceResult.error && sliceResult.error.message));
        }
        chunks.push(new Uint8Array(sliceResult.value.data));
        received += sliceResult.value.size;
        if (received >= file.size) {
          try { done(extractor(chunks)); }
          catch (e) { done(errResult('execution', String(e && e.message || e))); }
        } else {
          file.getSliceAsync(received, onSlice);
        }
      };
      file.getSliceAsync(0, onSlice);
    });
  });
}

// 极简 zip 容器解析（local file header 遍历，支持 stored(0) / deflate(8)）
function unzip(bytes) {
  const entries = {};
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const decoder = new TextDecoder('utf-8');
  let i = 0;
  while (i + 30 <= bytes.length) {
    if (view.getUint32(i, true) !== 0x04034b50) break;
    const method = view.getUint16(i + 8, true);
    const compSize = view.getUint32(i + 18, true);
    const nameLen = view.getUint16(i + 26, true);
    const extraLen = view.getUint16(i + 28, true);
    const name = decoder.decode(bytes.subarray(i + 30, i + 30 + nameLen));
    const dataStart = i + 30 + nameLen + extraLen;
    entries[name] = { method, data: bytes.subarray(dataStart, dataStart + compSize) };
    i = dataStart + compSize;
  }
  return entries;
}

function inflateEntry(entry) {
  if (entry.method === 0) return entry.data;      // stored
  if (entry.method === 8) return pako.inflateRaw(entry.data); // deflate
  throw new Error('unsupported zip method: ' + entry.method);
}

window.__UTIL__ = {
  okResult, errResult, hostName,
  excelRange, getSelectionTextAsync, setSelectionTextAsync,
  mapBuiltInStyle, mapAlignment, readPptFile, unzip, inflateEntry,
};

// ================= actions.js 热更新 =================

let loadedActionsVersion = '';

function loadActionsScript() {
  return new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = '/actions.js?ts=' + Date.now();
    s.onload = () => resolve();
    s.onerror = () => reject(new Error('actions.js load failed'));
    document.head.appendChild(s);
  });
}

async function ensureActions() {
  try {
    const r = await api('/office/actions-version');
    if (r && r.version && r.version !== loadedActionsVersion) {
      await loadActionsScript();
      loadedActionsVersion = r.version;
      log('actions 已加载/热更新: ' + r.version.slice(0, 10));
    }
  } catch (e) { /* 忽略：桥接未就绪时下次再试 */ }
}

// ================= 分发器 =================

function execute(action, args) {
  const ACTIONS = window.__ACTIONS__ || {};
  const meta = ACTIONS[action];
  if (!meta) return errResult('unknown_action', `unknown action: ${action}`);
  const host = hostName();
  if (!meta.hosts.includes(host)) return errResult('unsupported_host', `${action} not supported in ${host}`);
  return Promise.resolve()
    .then(() => meta.impl(host, args || {}))
    .catch((e) => {
      const debug = (e && e.debugInfo) ? (' [debug] ' + JSON.stringify(e.debugInfo)) : '';
      return errResult('execution', String(e && e.message || e) + debug);
    });
}

// ================= 后台循环 =================

let busy = false;
async function poll() {
  if (busy) return;
  busy = true;
  try {
    await ensureActions();
    const host = hostName();
    const cmd = await api('/office/poll?host=' + encodeURIComponent(host));
    if (cmd && cmd.commandId) {
      log(`执行: ${cmd.action} ${JSON.stringify(cmd.args || {}).slice(0, 80)}`);
      const out = await execute(cmd.action, cmd.args || {});
      await api('/office/result', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ commandId: cmd.commandId, ok: out.ok, result: out.result, error: out.error, code: out.code }),
      });
      log(out.ok ? `完成: ${cmd.action}` : `失败: ${cmd.action} -> ${out.error}`);
    }
  } catch (e) {
    setStatus('err', '桥接服务未连接，等待重试…');
  } finally {
    busy = false;
  }
}

function heartbeat() {
  api('/office/heartbeat', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ host: hostName() }),
  }).catch(() => { /* 忽略，下次再试 */ });
}

Office.onReady((info) => {
  const hostNameStr = Office.HostType[info.host] || String(info.host);
  hostEl.textContent = `host: ${hostNameStr} (${Office.version})`;
  setStatus('ok', '已连接：等待 DSH 指令');
  log(`Office 就绪 (${hostNameStr})`);
  api('/office/hello', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ host: hostName(), version: 'shell-v1.3' }),
  }).catch(() => {});
  setInterval(poll, POLL_MS);
  setInterval(heartbeat, HEARTBEAT_MS);
  heartbeat();
});
