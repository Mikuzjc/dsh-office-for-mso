// actions.js — DSH Office 执行器的全部 action 实现与注册表（热更新单元）
// 依赖外壳（taskpane.js）提供的全局 window.__UTIL__（okResult/errResult/hostName/excelRange/readPptFile/unzip/inflateEntry 等）。
// 窗格每次轮询前检查本文件版本，变更时动态重新加载 —— 改本文件后重启桥接服务即可热更新，无需重开窗格。
(function () {
  const U = window.__UTIL__;
  const okResult = U.okResult, errResult = U.errResult;
  const hostName = U.hostName;
  const excelRange = U.excelRange;
  const getSelectionTextAsync = U.getSelectionTextAsync;
  const setSelectionTextAsync = U.setSelectionTextAsync;
  const mapBuiltInStyle = U.mapBuiltInStyle;
  const mapAlignment = U.mapAlignment;
  const readPptFile = U.readPptFile;
  const unzip = U.unzip;
  const inflateEntry = U.inflateEntry;

  // ================= 通用 =================

  async function readSelection(host, args) {
    if (args && args.withStyles) {
      if (host === 'Word') {
        try {
          return await Word.run(async (ctx) => {
            const sel = ctx.document.getSelection();
            sel.load('text, font/name, font/size, font/bold, font/italic, font/color, font/underline, font/highlightColor');
            await ctx.sync();
            const f = sel.font;
            return okResult({ host: 'Word', text: sel.text || '', styles: { font: { name: f.name, size: f.size, bold: f.bold, italic: f.italic, color: f.color, underline: f.underline, highlightColor: f.highlightColor } } });
          });
        } catch (e) { return errResult('execution', String(e && e.message || e)); }
      }
      if (host === 'Excel') {
        try {
          return await Excel.run(async (ctx) => {
            const range = ctx.workbook.getSelectedRange();
            range.load('rowCount, columnCount');
            await ctx.sync();
            const rows = Math.min(range.rowCount, 10);
            const cols = Math.min(range.columnCount, 10);
            const cellRefs = [];
            for (let r = 0; r < rows; r++) {
              for (let c = 0; c < cols; c++) {
                const cell = range.getCell(r, c);
                cell.load('text, format/fill/color, format/font/name, format/font/size, format/font/bold, format/font/italic, format/font/color, format/numberFormat');
                cellRefs.push(cell);
              }
            }
            await ctx.sync();
            const cells = [];
            for (let r = 0; r < rows; r++) {
              for (let c = 0; c < cols; c++) {
                const cell = cellRefs[r * cols + c];
                const f = cell.format.font;
                cells.push({ row: r + 1, col: c + 1, value: cell.text, font: { name: f.name, size: f.size, bold: f.bold, italic: f.italic, color: f.color }, fill: cell.format.fill.color, numberFormat: cell.format.numberFormat });
              }
            }
            return okResult({ host: 'Excel', rows, cols, total: range.rowCount * range.columnCount, cells });
          });
        } catch (e) { return errResult('execution', String(e && e.message || e)); }
      }
      return errResult('unsupported_host', 'read_selection withStyles not supported in PowerPoint');
    }
    // Word：有选中返回选中文本；未选中返回光标左右各 5 字（XXXXX | XXXXX，| = 光标位置）
    if (host === 'Word') {
      try {
        return await Word.run(async (ctx) => {
          const sel = ctx.document.getSelection();
          sel.load('text, start');
          await ctx.sync();
          const text = sel.text || '';
          if (text) return okResult({ text });
          const pos = sel.start;
          if (pos === undefined || pos === null) return okResult({ text: '' });
          const body = ctx.document.body;
          // 未选中：取光标所在段落，段内定位取左右各 5 字；定位失败则退回报整段
          const para = sel.paragraphs.getFirst();
          const pr = para.getRange();
          pr.load('start');
          para.load('text');
          await ctx.sync();
          const pt = para.text || '';
          if (!pt) return okResult({ text: '' });
          if (pr.start !== undefined) {
            const rel = Math.max(0, Math.min(pt.length, pos - pr.start));
            const s = Math.max(0, rel - 5);
            const e = Math.min(pt.length, rel + 5);
            return okResult({ text: pt.slice(s, rel) + ' | ' + pt.slice(rel, e) });
          }
          return okResult({ text: pt });
        });
      } catch (e) { return errResult('execution', String(e && e.message || e)); }
    }
    return getSelectionTextAsync();
  }

  async function writeSelection(host, args) {
    if (args.text === undefined || args.text === null) return errResult('bad_args', 'text required');
    // 段落规范化：\r\n / \r → \n（setSelectedDataAsync 对 \n 转 Word 段落标记）
    const text = String(args.text).replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    return setSelectionTextAsync(text);
  }

  async function readDocument(host, args) {
    if (host === 'Word') {
      try {
        return await Word.run(async (ctx) => {
          const body = ctx.document.body;
          body.load('text');
          await ctx.sync();
          return okResult({ text: body.text || '', host: 'Word', characterCount: (body.text || '').length });
        });
      } catch (e) { return errResult('execution', String(e && e.message || e)); }
    }
    if (host === 'Excel') {
      try {
        return await Excel.run(async (ctx) => {
          const sheetName = args.sheet ? String(args.sheet) : null;
          const ws = sheetName ? ctx.workbook.worksheets.getItem(sheetName) : ctx.workbook.worksheets.getActiveWorksheet();
          ws.load('name');
          const range = ws.getUsedRange(true);
          range.load('text, rowCount, columnCount');
          await ctx.sync();
          const rows = range.text || [];
          const MAX_CELLS = 5000;
          const out = [];
          let cellCount = 0;
          let truncated = false;
          for (let r = 0; r < rows.length; r++) {
            cellCount += rows[r].length;
            if (cellCount > MAX_CELLS) { truncated = true; break; }
            out.push(rows[r].join('\t'));
          }
          return okResult({ text: out.join('\n'), host: 'Excel', sheet: ws.name, rows: rows.length, cols: range.columnCount, truncated });
        });
      } catch (e) { return errResult('execution', String(e && e.message || e)); }
    }
    if (host === 'PowerPoint') {
      return readPptDocument();
    }
    return errResult('unsupported_host', 'read_document unsupported host');
  }

  async function readStyles(host, args) {
    if (host === 'Word') {
      try {
        return await Word.run(async (ctx) => {
          const sel = ctx.document.getSelection();
          sel.load('text, font/name, font/size, font/bold, font/italic, font/color, font/underline, font/highlightColor');
          await ctx.sync();
          const f = sel.font;
          return okResult({ host: 'Word', text: sel.text || '', styles: { font: { name: f.name, size: f.size, bold: f.bold, italic: f.italic, color: f.color, underline: f.underline, highlightColor: f.highlightColor } } });
        });
      } catch (e) { return errResult('execution', String(e && e.message || e)); }
    }
    if (host === 'Excel') {
      try {
        return await Excel.run(async (ctx) => {
          const range = ctx.workbook.getSelectedRange();
          range.load('rowCount, columnCount');
          await ctx.sync();
          const rows = Math.min(range.rowCount, 10);
          const cols = Math.min(range.columnCount, 10);
          const cellRefs = [];
          for (let r = 0; r < rows; r++) {
            for (let c = 0; c < cols; c++) {
              const cell = range.getCell(r, c);
              cell.load('text, format/fill/color, format/font/name, format/font/size, format/font/bold, format/font/italic, format/font/color, format/numberFormat');
              cellRefs.push(cell);
            }
          }
          await ctx.sync();
          const cells = [];
          for (let r = 0; r < rows; r++) {
            for (let c = 0; c < cols; c++) {
              const cell = cellRefs[r * cols + c];
              const f = cell.format.font;
              cells.push({ row: r + 1, col: c + 1, value: cell.text, font: { name: f.name, size: f.size, bold: f.bold, italic: f.italic, color: f.color }, fill: cell.format.fill.color, numberFormat: cell.format.numberFormat });
            }
          }
          return okResult({ host: 'Excel', rows, cols, total: range.rowCount * range.columnCount, cells });
        });
      } catch (e) { return errResult('execution', String(e && e.message || e)); }
    }
    return errResult('unsupported_host', 'read_styles not supported in PowerPoint');
  }

  async function replaceAll(host, args) {
    const search = String(args.search ?? '');
    const replace = String(args.replace ?? '');
    if (!search) return errResult('bad_args', 'search required');
    if (host === 'Word') {
      try {
        return await Word.run(async (ctx) => {
          const results = ctx.document.body.search(search, { matchCase: false, matchWholeWord: false, ignoreSpace: true });
          if (args.dryRun) {
            // 预览：返回命中数量 + 命中文本上下文（审批面板可看清将替换什么）
            results.load('items/text');
            await ctx.sync();
            const hits = results.items.map((r) => r.text).slice(0, 20);
            return okResult({ host: 'Word', dryRun: true, wouldReplace: results.items.length, hits });
          }
          results.load('length');
          await ctx.sync();
          let count = 0;
          let lastRange = null;
          for (let i = 0; i < results.items.length; i++) {
            lastRange = results.items[i];
            lastRange.insertText(replace, Word.InsertLocation.replace);
            count++;
          }
          await ctx.sync();
          // 改完自动选中最后一处修改（保持高亮反馈；Office.js 只能单选区，多处无法同时选中）
          if (lastRange) {
            try { lastRange.select(); await ctx.sync(); } catch (e) { /* 选中失败不影响结果 */ }
          }
          return okResult({ replaced: count, host: 'Word', selected: lastRange ? true : false });
        });
      } catch (e) { return errResult('execution', String(e && e.message || e)); }
    }
    if (host === 'Excel') {
      if (args.dryRun) return errResult('unsupported', 'dryRun not supported in Excel replace_all');
      try {
        return await Excel.run(async (ctx) => {
          const range = ctx.workbook.worksheets.getActiveWorksheet().getUsedRange(true);
          range.replaceAll(search, replace, { completeMatch: false, matchCase: false });
          await ctx.sync();
          return okResult({ replaced: -1, host: 'Excel' });
        });
      } catch (e) { return errResult('execution', String(e && e.message || e)); }
    }
    return errResult('unsupported_host', 'replace_all not supported in PowerPoint');
  }

  async function appendText(host, args) {
    const text = String(args.text ?? '');
    if (!text) return errResult('bad_args', 'text required');
    if (host === 'Word') {
      try {
        return await Word.run(async (ctx) => {
          const para = ctx.document.body.insertParagraph(text, Word.InsertLocation.end);
          await ctx.sync();
          // 改完自动选中追加的段落
          try { para.getRange().select(); await ctx.sync(); } catch (e) { /* 选中失败不影响结果 */ }
          return okResult({ appended: text.length, host: 'Word', selected: true });
        });
      } catch (e) { return errResult('execution', String(e && e.message || e)); }
    }
    return errResult('unsupported_host', 'append_text only supported in Word');
  }

  // ================= Word 组 =================

  async function readTables(host, args) {
    if (host !== 'Word') return errResult('unsupported_host', 'read_tables only supported in Word');
    try {
      return await Word.run(async (ctx) => {
        const tables = ctx.document.body.tables;
        tables.load('items/rowCount');
        await ctx.sync();
        const out = [];
        for (let ti = 0; ti < tables.items.length; ti++) {
          const t = tables.items[ti];
          const range = t.getRange();
          range.load('text');
          await ctx.sync();
          const raw = range.text || '';
          const lines = raw.split('\r').filter((l) => l.length > 0);
          const cells = lines.map((l) => l.split('\t').map((v) => v.trim()));
          out.push({ index: ti + 1, rowCount: t.rowCount, columnCount: cells.length ? cells[0].length : 0, cells });
        }
        return okResult({ host: 'Word', tables: out });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function setFont(host, args) {
    const font = String(args.font || '');
    if (!font) return errResult('bad_args', 'font required');
    if (host !== 'Word') return errResult('unsupported_host', 'set_font only supported in Word');
    try {
      return await Word.run(async (ctx) => {
        const paras = ctx.document.body.paragraphs;
        paras.load('items/font/name, items/text');
        await ctx.sync();
        let count = 0;
        for (const p of paras.items) {
          p.font.name = font;
          count++;
        }
        await ctx.sync();
        return okResult({ host: 'Word', paragraphs: count, font });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function removeEmptyParagraphs(host, args) {
    if (host !== 'Word') return errResult('unsupported_host', 'remove_empty_paragraphs only supported in Word');
    try {
      return await Word.run(async (ctx) => {
        const paras = ctx.document.body.paragraphs;
        paras.load('items/text, items/inlinePictures/items');
        await ctx.sync();
        const toDelete = [];
        const n = paras.items.length;
        let skippedImages = 0;
        for (let i = 0; i < n; i++) {
          if (i === n - 1) continue;
          const p = paras.items[i];
          if (p.inlinePictures && p.inlinePictures.items && p.inlinePictures.items.length > 0) {
            skippedImages++;
            continue;
          }
          if (p.text.trim() === '') toDelete.push(p);
        }
        if (args.dryRun) return okResult({ host: 'Word', dryRun: true, wouldRemove: toDelete.length, skippedImageParagraphs: skippedImages });
        for (const p of toDelete) p.delete();
        await ctx.sync();
        return okResult({ host: 'Word', removed: toDelete.length, skippedImageParagraphs: skippedImages });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function insertParagraph(host, args) {
    const text = String(args.text ?? '');
    if (!text) return errResult('bad_args', 'text required');
    if (host !== 'Word') return errResult('unsupported_host', 'insert_paragraph only supported in Word');
    try {
      return await Word.run(async (ctx) => {
        const loc = args.location === 'start' ? Word.InsertLocation.start : (args.location === 'afterSelection' ? Word.InsertLocation.after : Word.InsertLocation.end);
        let para;
        if (args.location === 'afterSelection') {
          const sel = ctx.document.getSelection();
          para = sel.insertParagraph(text, Word.InsertLocation.after);
        } else {
          para = ctx.document.body.insertParagraph(text, loc);
        }
        const style = mapBuiltInStyle(args.style);
        if (style) { para.styleBuiltIn = Word.BuiltInStyleName[style]; }
        await ctx.sync();
        // 改完自动选中插入的段落
        try { para.getRange().select(); await ctx.sync(); } catch (e) { /* 选中失败不影响结果 */ }
        return okResult({ host: 'Word', inserted: text.length, style: style || null, location: args.location || 'end', selected: true });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  // 表格 OOXML 构造（insertTable 系列在此环境不可用，用 insertOoxml 兜底）
  function escapeXml(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function buildTableOoxml(stringRows) {
    const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const colW = Math.max(800, Math.floor(9000 / (stringRows[0] ? stringRows[0].length : 1)));
    const rowsXml = stringRows.map((row) => {
      const cells = row.map((c) => `<w:tc><w:tcPr><w:tcW w:w="${colW}" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">${escapeXml(c)}</w:t></w:r></w:p></w:tc>`).join('');
      return `<w:tr>${cells}</w:tr>`;
    }).join('');
    const grid = (stringRows[0] || []).map(() => `<w:gridCol w:w="${colW}"/>`).join('');
    return `<w:tbl ${ns}><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${rowsXml}</w:tbl>`;
  }

  async function insertTable(host, args) {
    const rows = args.rows;
    if (!Array.isArray(rows) || rows.length === 0 || !Array.isArray(rows[0])) return errResult('bad_args', 'rows (2D array) required');
    if (host !== 'Word') return errResult('unsupported_host', 'insert_table only supported in Word');
    try {
      return await Word.run(async (ctx) => {
        const stringRows = rows.map((r) => r.map((v) => String(v ?? '')));
        const loc = args.location === 'start' ? Word.InsertLocation.start : Word.InsertLocation.end;
        const ooxml = buildTableOoxml(stringRows);
        ctx.document.body.insertOoxml(ooxml, loc);
        await ctx.sync();
        return okResult({ host: 'Word', rows: stringRows.length, cols: stringRows[0].length });
      });
    } catch (e) {
      const debug = (e && e.debugInfo) ? (' [debug] ' + JSON.stringify(e.debugInfo)) : '';
      return errResult('execution', `${e && e.message || e}${debug}`);
    }
  }

  async function insertImage(host, args) {
    const b64 = String(args.base64 || '').replace(/^data:image\/\w+;base64,/, '');
    if (!b64) return errResult('bad_args', 'base64 required');
    if (host !== 'Word') return errResult('unsupported_host', 'insert_image only supported in Word (PPT/Excel 二期)');
    try {
      return await Word.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        const pic = sel.insertInlinePictureFromBase64(b64, Word.InsertLocation.replace);
        if (args.width) pic.width = Number(args.width);
        if (args.height) pic.height = Number(args.height);
        await ctx.sync();
        return okResult({ host: 'Word', inserted: true });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function applyStyle(host, args) {
    const style = mapBuiltInStyle(args.style);
    if (!style) return errResult('bad_args', `unknown style: ${args.style}`);
    if (host !== 'Word') return errResult('unsupported_host', 'apply_style only supported in Word');
    try {
      return await Word.run(async (ctx) => {
        const bs = Word.BuiltInStyleName[style];
        let count = 0;
        if (args.scope === 'all') {
          const paras = ctx.document.body.paragraphs;
          paras.load('items');
          await ctx.sync();
          for (const p of paras.items) { p.styleBuiltIn = bs; count++; }
        } else {
          const sel = ctx.document.getSelection();
          sel.styleBuiltIn = bs;
          count = 1;
        }
        await ctx.sync();
        return okResult({ host: 'Word', applied: count, style, scope: args.scope || 'selection' });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function formatSelection(host, args) {
    if (host === 'Word') {
      try {
        return await Word.run(async (ctx) => {
          const sel = ctx.document.getSelection();
          const f = sel.font;
          if (args.font) f.name = String(args.font);
          if (args.size) f.size = Number(args.size);
          if (args.bold !== undefined) f.bold = !!args.bold;
          if (args.italic !== undefined) f.italic = !!args.italic;
          if (args.color) f.color = String(args.color);
          if (args.highlight) f.highlightColor = String(args.highlight);
          await ctx.sync();
          return okResult({ host: 'Word', formatted: true });
        });
      } catch (e) { return errResult('execution', String(e && e.message || e)); }
    }
    if (host === 'Excel') {
      try {
        return await Excel.run(async (ctx) => {
          const range = ctx.workbook.getSelectedRange();
          const fmt = range.format;
          if (args.font) fmt.font.name = String(args.font);
          if (args.size) fmt.font.size = Number(args.size);
          if (args.bold !== undefined) fmt.font.bold = !!args.bold;
          if (args.italic !== undefined) fmt.font.italic = !!args.italic;
          if (args.color) fmt.font.color = String(args.color);
          await ctx.sync();
          return okResult({ host: 'Excel', formatted: true });
        });
      } catch (e) { return errResult('execution', String(e && e.message || e)); }
    }
    return errResult('unsupported_host', 'format_selection not supported in PowerPoint');
  }

  async function setParagraphFormat(host, args) {
    if (host !== 'Word') return errResult('unsupported_host', 'set_paragraph_format only supported in Word');
    try {
      return await Word.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        const pf = sel.paragraphFormat;
        if (!pf) return errResult('requirement', 'Word paragraphFormat API not available in this environment');
        if (args.alignment) {
          const al = mapAlignment(args.alignment);
          if (al) pf.alignment = al;
        }
        if (args.indent !== undefined) pf.leftIndent = Number(args.indent) * 28.35;
        if (args.lineSpacing !== undefined) pf.lineSpacing = Number(args.lineSpacing);
        if (args.listType && args.listType !== 'none') {
          const paras = sel.paragraphs;
          paras.load('items');
          await ctx.sync();
          ctx.document.body.lists.create(paras.items, args.listType === 'numbered' ? Word.ListType.numbered : Word.ListType.bulleted);
        }
        await ctx.sync();
        return okResult({ host: 'Word', formatted: true, listType: args.listType || null });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function search(host, args) {
    const query = String(args.query ?? '');
    if (!query) return errResult('bad_args', 'query required');
    if (host !== 'Word') return errResult('unsupported_host', 'search only supported in Word');
    try {
      return await Word.run(async (ctx) => {
        const results = ctx.document.body.search(query, { matchCase: !!args.matchCase, matchWildcards: !!args.wildcard });
        results.load('items/text');
        await ctx.sync();
        const hits = results.items.map((r, i) => ({ index: i + 1, text: r.text }));
        return okResult({ host: 'Word', hits, count: hits.length });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function addComment(host, args) {
    const text = String(args.text ?? '');
    if (!text) return errResult('bad_args', 'text required');
    if (host === 'Word') {
      try {
        return await Word.run(async (ctx) => {
          const comments = ctx.document.body.comments;
          if (!comments) return errResult('requirement', 'Word comments API (WordApi 1.4+) not available');
          const sel = ctx.document.getSelection();
          comments.add(sel, text);
          await ctx.sync();
          return okResult({ host: 'Word', added: true });
        });
      } catch (e) { return errResult('execution', String(e && e.message || e)); }
    }
    if (host === 'Excel') {
      if (!args.cell) return errResult('bad_args', 'cell (e.g. A1 or Sheet1!A1) required for Excel');
      try {
        return await Excel.run(async (ctx) => {
          // cell 可带表名（Sheet1!A1）→ 加到目标表；否则加到激活表
          let cellAddr = String(args.cell);
          let ws;
          const bang = cellAddr.lastIndexOf('!');
          if (bang >= 0) {
            ws = ctx.workbook.worksheets.getItem(cellAddr.slice(0, bang));
            cellAddr = cellAddr.slice(bang + 1);
          } else {
            ws = ctx.workbook.worksheets.getActiveWorksheet();
          }
          ws.comments.add(cellAddr, text);
          await ctx.sync();
          return okResult({ host: 'Excel', added: true, cell: cellAddr, sheet: ws.name });
        });
      } catch (e) { return errResult('execution', String(e && e.message || e)); }
    }
    return errResult('unsupported_host', 'add_comment not supported in PowerPoint');
  }

  async function readComments(host, args) {
    if (host === 'Word') {
      try {
        return await Word.run(async (ctx) => {
          const comments = ctx.document.body.comments;
          if (!comments) return errResult('requirement', 'Word comments API (WordApi 1.4+) not available');
          comments.load('items/content, items/author, items/createdDate, items/resolved');
          await ctx.sync();
          const list = comments.items.map((c) => ({ content: c.content, author: c.author, createdDate: c.createdDate, resolved: c.resolved }));
          return okResult({ host: 'Word', comments: list });
        });
      } catch (e) { return errResult('execution', String(e && e.message || e)); }
    }
    if (host === 'Excel') {
      try {
        return await Excel.run(async (ctx) => {
          const ws = ctx.workbook.worksheets.getActiveWorksheet();
          const comments = ws.comments;
          // 注：此环境 Comment.getRange() 不存在、cellAddress 可能为空 → 尽力返回
          comments.load('items/author, items/content, items/resolved, items/cellAddress');
          await ctx.sync();
          const list = comments.items.map((c) => ({ cell: c.cellAddress || '', content: c.content, author: c.author, resolved: c.resolved }));
          return okResult({ host: 'Excel', comments: list });
        });
      } catch (e) { return errResult('execution', String(e && e.message || e)); }
    }
    return errResult('unsupported_host', 'read_comments not supported in PowerPoint');
  }

  async function readProperties(host, args) {
    if (host === 'Word') {
      try {
        return await Word.run(async (ctx) => {
          const p = ctx.document.properties;
          p.load('title, author, subject, keywords, comments, category, lastModifiedBy');
          await ctx.sync();
          return okResult({ host: 'Word', properties: { title: p.title, author: p.author, subject: p.subject, keywords: p.keywords, comments: p.comments, category: p.category, lastModifiedBy: p.lastModifiedBy } });
        });
      } catch (e) { return errResult('execution', String(e && e.message || e)); }
    }
    if (host === 'Excel') {
      try {
        return await Excel.run(async (ctx) => {
          const p = ctx.workbook.properties;
          p.load('title, author, subject, keywords, comments, category');
          await ctx.sync();
          return okResult({ host: 'Excel', properties: { title: p.title, author: p.author, subject: p.subject, keywords: p.keywords, comments: p.comments, category: p.category } });
        });
      } catch (e) { return errResult('execution', String(e && e.message || e)); }
    }
    return errResult('unsupported_host', 'read_properties not supported in PowerPoint');
  }

  // ================= Excel 组 =================

  async function listSheets(host, args) {
    if (host !== 'Excel') return errResult('unsupported_host', 'list_sheets only supported in Excel');
    try {
      return await Excel.run(async (ctx) => {
        const sheets = ctx.workbook.worksheets;
        sheets.load('items/name, items/position, items/visibility');
        await ctx.sync();
        const list = sheets.items.map((s) => ({ name: s.name, position: s.position, visibility: s.visibility }));
        return okResult({ host: 'Excel', sheets: list });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function readRange(host, args) {
    if (host !== 'Excel') return errResult('unsupported_host', 'read_range only supported in Excel');
    const address = String(args.address || '');
    if (!address) return errResult('bad_args', 'address (e.g. Sheet1!A1:B10) required');
    try {
      return await Excel.run(async (ctx) => {
        const range = excelRange(ctx, address);
        range.load('values, formulas, numberFormats, rowCount, columnCount');
        await ctx.sync();
        const limit = args.limit ? Number(args.limit) : 5000;
        const total = (range.rowCount || 0) * (range.columnCount || 0);
        const truncated = total > limit;
        const sliceRows = truncated ? Math.max(1, Math.floor(limit / (range.columnCount || 1))) : range.rowCount;
        const values = (range.values || []).slice(0, sliceRows);
        return okResult({ host: 'Excel', address, rows: values.length, cols: range.columnCount, truncated, values });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function writeRange(host, args) {
    if (host !== 'Excel') return errResult('unsupported_host', 'write_range only supported in Excel');
    const address = String(args.address || '');
    if (!address) return errResult('bad_args', 'address required');
    const data = args.values || args.formulas;
    if (!Array.isArray(data) || data.length === 0 || !Array.isArray(data[0])) return errResult('bad_args', 'values or formulas (2D array) required');
    try {
      return await Excel.run(async (ctx) => {
        const totalCells = data.length * data[0].length;
        const BATCH = 5000;
        if (totalCells <= BATCH) {
          const range = excelRange(ctx, address);
          if (args.values) range.values = args.values;
          if (args.formulas) range.formulas = args.formulas;
          await ctx.sync();
          // 改完自动选中写入区域
          try { range.select(); await ctx.sync(); } catch (e) { /* 选中失败不影响结果 */ }
          return okResult({ host: 'Excel', rows: data.length, cols: data[0].length, batched: 1, selected: true });
        }
        const perChunk = Math.max(1, Math.floor(BATCH / data[0].length));
        let chunks = 0;
        for (let i = 0; i < data.length; i += perChunk) {
          const chunk = data.slice(i, i + perChunk);
          const base = excelRange(ctx, address).getCell(i, 0);
          const target = base.getResizedRange(chunk.length - 1, data[0].length - 1);
          if (args.values) target.values = chunk;
          if (args.formulas) target.formulas = chunk;
          await ctx.sync();
          chunks++;
        }
        // 改完自动选中写入区域
        try { excelRange(ctx, address).select(); await ctx.sync(); } catch (e) { /* 选中失败不影响结果 */ }
        return okResult({ host: 'Excel', rows: data.length, cols: data[0].length, batched: chunks, selected: true });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function formatRange(host, args) {
    if (host !== 'Excel') return errResult('unsupported_host', 'format_range only supported in Excel');
    const address = String(args.address || '');
    if (!address) return errResult('bad_args', 'address required');
    try {
      return await Excel.run(async (ctx) => {
        const range = excelRange(ctx, address);
        const fmt = range.format;
        if (args.font) fmt.font.name = String(args.font);
        if (args.size) fmt.font.size = Number(args.size);
        if (args.bold !== undefined) fmt.font.bold = !!args.bold;
        if (args.fill) fmt.fill.color = String(args.fill);
        if (args.numberFormat) fmt.numberFormat = String(args.numberFormat);
        if (args.autoFit) fmt.autoFitColumns();
        if (args.tableStyle) {
          const table = range.getSurroundingTable();
          if (table) { table.style = String(args.tableStyle); }
        }
        await ctx.sync();
        return okResult({ host: 'Excel', formatted: true, address });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  const CHART_TYPES = {
    line: 'Line', column: 'ColumnClustered', bar: 'BarClustered', pie: 'Pie',
    area: 'Area', scatter: 'XYScatter', doughnut: 'Doughnut', radar: 'Radar',
  };

  async function insertChart(host, args) {
    if (host !== 'Excel') return errResult('unsupported_host', 'insert_chart only supported in Excel');
    const dataRange = String(args.dataRange || '');
    if (!dataRange) return errResult('bad_args', 'dataRange (e.g. Sheet1!A1:B10) required');
    try {
      return await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getActiveWorksheet();
        const typeName = CHART_TYPES[String(args.type || '').toLowerCase()] || 'ColumnClustered';
        // sourceData 传 Range 对象（字符串跨表引用在此环境报 sourceData 错误）
        const source = excelRange(ctx, dataRange);
        const chart = ws.charts.add(typeName, source, 'auto');
        if (args.title) {
          chart.title.text = args.title;
        }
        await ctx.sync();
        return okResult({ host: 'Excel', chart: { type: typeName, dataRange, title: args.title || null } });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function addSheet(host, args) {
    if (host !== 'Excel') return errResult('unsupported_host', 'add_sheet only supported in Excel');
    const name = String(args.name || '');
    if (!name) return errResult('bad_args', 'name required');
    try {
      return await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.add(name);
        ws.load('name');
        await ctx.sync();
        return okResult({ host: 'Excel', added: ws.name });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function renameSheet(host, args) {
    if (host !== 'Excel') return errResult('unsupported_host', 'rename_sheet only supported in Excel');
    const oldName = String(args.oldName || '');
    const newName = String(args.newName || '');
    if (!oldName || !newName) return errResult('bad_args', 'oldName and newName required');
    try {
      return await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(oldName);
        ws.name = newName;
        await ctx.sync();
        return okResult({ host: 'Excel', renamed: { from: oldName, to: newName } });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function deleteSheet(host, args) {
    if (host !== 'Excel') return errResult('unsupported_host', 'delete_sheet only supported in Excel');
    const name = String(args.name || '');
    if (!name) return errResult('bad_args', 'name required');
    try {
      return await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(name);
        ws.load('name, position, visibility');
        await ctx.sync();
        if (args.dryRun) {
          return okResult({ host: 'Excel', dryRun: true, wouldDelete: { name: ws.name, position: ws.position, visibility: ws.visibility } });
        }
        ws.delete();
        await ctx.sync();
        return okResult({ host: 'Excel', deleted: name });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function applySort(host, args) {
    if (host !== 'Excel') return errResult('unsupported_host', 'apply_sort only supported in Excel');
    const address = String(args.address || '');
    const fields = args.fields;
    if (!address) return errResult('bad_args', 'address required');
    if (!Array.isArray(fields) || fields.length === 0) return errResult('bad_args', 'fields [{column,ascending}] required');
    try {
      return await Excel.run(async (ctx) => {
        const range = excelRange(ctx, address);
        range.sort.apply(fields.map((f) => ({ key: Number(f.column) - 1, ascending: f.ascending !== false })), false);
        await ctx.sync();
        return okResult({ host: 'Excel', sorted: fields.length });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  async function applyFilter(host, args) {
    if (host !== 'Excel') return errResult('unsupported_host', 'apply_filter only supported in Excel');
    const address = String(args.address || '');
    if (!address) return errResult('bad_args', 'address required');
    try {
      return await Excel.run(async (ctx) => {
        // autoFilter 属于 worksheet，且 range 必须在同一工作表上 → 从 address 解析目标表
        const range = excelRange(ctx, address);
        const bang = address.lastIndexOf('!');
        const ws = bang >= 0 ? ctx.workbook.worksheets.getItem(address.slice(0, bang)) : ctx.workbook.worksheets.getActiveWorksheet();
        if (!ws.autoFilter) return errResult('requirement', 'Excel autoFilter API (ExcelApi 1.9+) not available');
        const criteria = Array.isArray(args.columns)
          ? args.columns.map((c) => ({ column: Number(c.column), filterType: c.filterType || 'Values', criteria1: c.criteria1, criteria2: c.criteria2 }))
          : undefined;
        ws.autoFilter.apply(range, criteria);
        await ctx.sync();
        return okResult({ host: 'Excel', filtered: true, columns: criteria ? criteria.length : 0 });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  // 单参数统计函数白名单：SUM/AVERAGE/COUNT/MAX/MIN/PRODUCT（用 workbook.functions 类型化求值，不写单元格）
  const EVAL_FUNCTIONS = { SUM: 'sum', AVERAGE: 'average', COUNT: 'count', MAX: 'max', MIN: 'min', PRODUCT: 'product' };

  async function evaluateFormula(host, args) {
    if (host !== 'Excel') return errResult('unsupported_host', 'evaluate_formula only supported in Excel');
    const formula = String(args.formula || '').trim();
    if (!formula) return errResult('bad_args', 'formula required (e.g. SUM(A1:A10))');
    const m = formula.match(/^([A-Z_]+)\((.*)\)$/i);
    const fnUpper = m ? m[1].toUpperCase() : '';
    const fnName = EVAL_FUNCTIONS[fnUpper];
    if (!fnName) return errResult('bad_args', `unsupported formula: ${formula}；支持 ${Object.keys(EVAL_FUNCTIONS).join('/')}(区域)`);
    try {
      return await Excel.run(async (ctx) => {
        const argStr = m[2].trim();
        if (!argStr) return errResult('bad_args', '函数缺少区域参数');
        const ranges = argStr.split(',').map((s) => excelRange(ctx, s.trim()));
        const funcs = ctx.workbook.functions;
        if (!funcs) return errResult('requirement', 'Excel Functions API not available');
        const result = funcs[fnName](ranges[0]);
        result.load('value');
        await ctx.sync();
        const value = result.value;
        return okResult({ host: 'Excel', formula, value, function: fnUpper });
      });
    } catch (e) { return errResult('execution', String(e && e.message || e)); }
  }

  // ================= PPT 组 =================

  function readSlides(host, args) {
    if (host !== 'PowerPoint') return Promise.resolve(errResult('unsupported_host', 'read_slides only supported in PowerPoint'));
    return new Promise((resolve) => {
      Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,
        (r) => resolve(r.status === Office.AsyncResultStatus.Succeeded
          ? okResult({ host: 'PowerPoint', slides: r.value && r.value.slides ? r.value.slides.map((s) => ({ id: s.id, title: s.title || '' })) : [] })
          : errResult('execution', r.error && r.error.message)));
    });
  }

  function readPptDocument() {
    return readPptFile(extractPptText);
  }

  function readPptNotes(host, args) {
    if (host !== 'PowerPoint') return Promise.resolve(errResult('unsupported_host', 'ppt_read_notes only supported in PowerPoint'));
    return readPptFile(extractPptNotes);
  }

  function extractPptText(chunks) {
    const entries = concatChunksToZip(chunks);
    const slideNames = Object.keys(entries)
      .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
      .sort((a, b) => parseInt(a.match(/slide(\d+)/)[1], 10) - parseInt(b.match(/slide(\d+)/)[1], 10));
    const decoder = new TextDecoder('utf-8');
    const parts = [];
    for (const name of slideNames) {
      const xml = decoder.decode(inflateEntry(entries[name]));
      const doc = new DOMParser().parseFromString(xml, 'application/xml');
      const aT = doc.getElementsByTagName('a:t');
      let slideText = '';
      for (const t of aT) slideText += t.textContent;
      const num = name.match(/slide(\d+)/)[1];
      parts.push(`--- 幻灯片 ${num} ---\n${slideText.trim()}`);
    }
    return okResult({ host: 'PowerPoint', slides: slideNames.length, text: parts.join('\n\n') });
  }

  function extractPptNotes(chunks) {
    const entries = concatChunksToZip(chunks);
    const decoder = new TextDecoder('utf-8');
    const slideNums = Object.keys(entries)
      .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
      .map((n) => parseInt(n.match(/slide(\d+)/)[1], 10));
    const notesMap = {};
    for (const sn of slideNums) {
      const relsName = `ppt/slides/_rels/slide${sn}.xml.rels`;
      const relsEntry = entries[relsName];
      if (!relsEntry) continue;
      const relsDoc = new DOMParser().parseFromString(decoder.decode(inflateEntry(relsEntry)), 'application/xml');
      const rels = relsDoc.getElementsByTagName('Relationship');
      let notesTarget = null;
      for (const rel of rels) {
        if (rel.getAttribute('Type') && rel.getAttribute('Type').indexOf('notesSlide') >= 0) {
          notesTarget = rel.getAttribute('Target');
          break;
        }
      }
      if (!notesTarget) continue;
      const notesName = 'ppt/notesSlides/' + notesTarget.split('/').pop();
      const notesEntry = entries[notesName];
      if (!notesEntry) continue;
      const notesDoc = new DOMParser().parseFromString(decoder.decode(inflateEntry(notesEntry)), 'application/xml');
      const aT = notesDoc.getElementsByTagName('a:t');
      let text = '';
      for (const t of aT) text += t.textContent;
      notesMap[sn] = text.trim();
    }
    const parts = Object.keys(notesMap).sort((a, b) => Number(a) - Number(b)).map((sn) => `--- 幻灯片 ${sn} 备注 ---\n${notesMap[sn] || '(无备注文本)'}`);
    return okResult({ host: 'PowerPoint', notes: parts.join('\n\n'), slidesWithNotes: parts.length });
  }

  function concatChunksToZip(chunks) {
    const total = chunks.reduce((a, c) => a + c.length, 0);
    const buf = new Uint8Array(total);
    let off = 0;
    for (const c of chunks) { buf.set(c, off); off += c.length; }
    return unzip(buf);
  }

  // ================= 环境诊断 =================

  function isSetSupported(name, version) {
    try {
      return !!Office.context.requirements.isSetSupported(name, version);
    } catch (e) { return false; }
  }

  async function getEnvironment(host, args) {
    const base = {
      host: String(Office.context.host || ''),
      platform: String(Office.context.platform || ''),
      officeVersion: String(Office.version || ''),
      diagnostics: (function () {
        try {
          const d = Office.context.diagnostics || {};
          return {
            officeContext: String(d.OfficeContext || ''),
            hostAppVersion: String(d.HostApplicationVersion || ''),
            hostPlatformVersion: String(d.HostPlatformVersion || ''),
            hostSession: String(d.HostSession || ''),
            hasDiagnostics: !!Office.context.diagnostics,
          };
        } catch (e) { return { error: String(e && e.message || e) }; }
      })(),
      requirementSets: {
        WordApi_1_1: isSetSupported('WordApi', '1.1'),
        WordApi_1_4: isSetSupported('WordApi', '1.4'),
        WordApi_1_5: isSetSupported('WordApi', '1.5'),
        WordApi_1_8: isSetSupported('WordApi', '1.8'),
        ExcelApi_1_9: isSetSupported('ExcelApi', '1.9'),
        ExcelApi_1_10: isSetSupported('ExcelApi', '1.10'),
        PowerPointApi_1_1: isSetSupported('PowerPointApi', '1.1'),
        ImageCoercion_1_1: isSetSupported('ImageCoercion', '1.1'),
      },
    };
    if (host === 'Word') {
      try {
        return await Word.run(async (ctx) => {
          const body = ctx.document.body;
          const sel = ctx.document.getSelection();
          // 探测对象属性是否真实存在（不 load，只检查存在性）——区分"API 不存在" vs "存在但调用失败"
          const probe = {
            bodyHasComments: !!body.comments,
            selHasParagraphFormat: !!sel.paragraphFormat,
            bodyInsertTableType: typeof body.insertTable,
            bodyInsertOoxmlType: typeof body.insertOoxml,
            selHasFont: !!sel.font,
          };
          return okResult({ ...base, probe });
        });
      } catch (e) { return errResult('execution', String(e && e.message || e)); }
    }
    return okResult(base);
  }

  // ================= 定位 + 选中（locate_select）=================
  // 零副作用：仅"选中/取消选中"的 UI 反馈，不修改文档内容/样式/撤销栈。
  // 定位器（三选一）：
  //   bookmark | anchor —— Word 书签 / Excel 命名区域名
  //   text               —— Word 正文首个匹配；Excel 已用区域首个包含该文本的单元格
  //   range | address    —— Excel 区域地址（A1:B5 或 Sheet!A1:B5）
  // 可选：sheet（Excel 工作表）、blinks（闪烁次数，默认 0=只选中保持，>0 才闪烁）、interval（间隔 ms，默认 300，100-800）
  function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }

  async function locateSelect(host, args) {
    const blinks = Math.max(0, Math.min(5, Number(args.blinks) || 0));
    const interval = Math.max(100, Math.min(800, Number(args.interval) || 300));
    const reps = Math.max(1, blinks);
    const bookmark = (args.bookmark !== undefined ? String(args.bookmark) : args.anchor !== undefined ? String(args.anchor) : null);
    const text = args.text !== undefined ? String(args.text) : null;
    const address = (args.range !== undefined ? String(args.range) : args.address !== undefined ? String(args.address) : null);
    if (!bookmark && !text && !address) return errResult('bad_args', 'locate_select requires one of: bookmark/anchor, text, range/address');

    if (host === 'Word') {
      try {
        let anySelected = false;
        for (let i = 0; i < reps; i++) {
          if (bookmark) {
            // 主路径：goToByIdAsync 滚动并选中书签内容（getBookmark 在部分宿主不可用，仅作增强）
            const navigated = await new Promise((resolve) => {
              Office.context.document.goToByIdAsync(bookmark, Office.GoToType.Bookmark, (res) => resolve(res.status === Office.AsyncResultStatus.Succeeded));
            });
            let precise = false;
            try {
              await Word.run(async (ctx) => {
                const bm = ctx.document.getBookmark(bookmark);
                bm.load('name');
                await ctx.sync();
                bm.getRange().select();
                await ctx.sync();
              });
              precise = true;
            } catch (e) { /* getBookmark 不可用，goToByIdAsync 已选中 */ }
            if (!navigated && !precise) throw new Error(`locate_select: bookmark not found: ${bookmark}`);
            if (navigated || precise) anySelected = true;
          } else {
            await Word.run(async (ctx) => {
              const results = ctx.document.body.search(text, { matchCase: false, ignoreSpace: true });
              results.load('length');
              await ctx.sync();
              if (!results.items || results.items.length === 0) throw new Error(`locate_select: text not found: ${text}`);
              results.items[0].select();
              await ctx.sync();
            });
            anySelected = true;
          }
          if (blinks > 0) {
            await sleep(interval);
            await Word.run(async (ctx) => {
              // 取消整块高亮：折叠为零宽选区（留在原位）。枚举可能不可用，用字面量 + 降级链。
              const sel = ctx.document.getSelection();
              try { sel.collapse('start'); }
              catch (e) {
                try { sel.getRange('start').select(); }
                catch (e2) { /* 保持选中 */ }
              }
              await ctx.sync();
            });
            await sleep(interval);
          }
        }
        return okResult({ host: 'Word', located: bookmark || text, selected: anySelected, blinks });
      } catch (e) { return errResult('execution', String(e && e.message || e)); }
    }

    if (host === 'Excel') {
      try {
        const sheetName = args.sheet !== undefined ? String(args.sheet) : null;
        // 先解析目标地址（一次），避免每轮重复扫描
        let targetAddress = null;
        if (bookmark) {
          await Excel.run(async (ctx) => {
            const nm = ctx.workbook.names.getItem(bookmark);
            const r = nm.getRange();
            r.load('address');
            await ctx.sync();
            targetAddress = r.address;
          });
        } else if (address) {
          targetAddress = address;
        } else {
          await Excel.run(async (ctx) => {
            const ws = sheetName ? ctx.workbook.worksheets.getItem(sheetName) : ctx.workbook.worksheets.getActiveWorksheet();
            const used = ws.getUsedRange(true);
            used.load('text, rowCount, columnCount');
            await ctx.sync();
            let fr = -1, fc = -1;
            for (let r = 0; r < used.rowCount && fr < 0; r++) {
              for (let c = 0; c < used.columnCount && fr < 0; c++) {
                if (String((used.text[r] && used.text[r][c]) || '').indexOf(text) >= 0) { fr = r; fc = c; }
              }
            }
            if (fr < 0) throw new Error(`locate_select: text not found: ${text}`);
            const cell = used.getCell(fr, fc);
            cell.load('address');
            await ctx.sync();
            targetAddress = cell.address;
          });
        }
        if (!targetAddress) throw new Error('locate_select: could not resolve target');
        const [sheetPart, addrPart] = targetAddress.indexOf('!') >= 0 ? targetAddress.split('!') : [sheetName, targetAddress];
        for (let i = 0; i < reps; i++) {
          await Excel.run(async (ctx) => {
            const ws = sheetPart ? ctx.workbook.worksheets.getItem(sheetPart) : ctx.workbook.worksheets.getActiveWorksheet();
            ws.getRange(addrPart).select(); // 整块选中（绿色边框）
            await ctx.sync();
          });
          if (blinks > 0) {
            await sleep(interval);
            await Excel.run(async (ctx) => {
              const ws = sheetPart ? ctx.workbook.worksheets.getItem(sheetPart) : ctx.workbook.worksheets.getActiveWorksheet();
              ws.getRange(addrPart.split(':')[0]).select(); // 选回首格：取消整块高亮、留在原位
              await ctx.sync();
            });
            await sleep(interval);
          }
        }
        return okResult({ host: 'Excel', located: bookmark || text || targetAddress, selected: true, blinks });
      } catch (e) { return errResult('execution', String(e && e.message || e)); }
    }

    return errResult('unsupported_host', 'locate_select not supported in PowerPoint');
  }

  // ===== 审批预选（selectTarget）：ask 确认模式下，预览时主动选中将要修改的内容 =====
  // 注意：Office.js 只能选中一个连续区域，"一次性多处"无法同时选中 —— 多处替换场景选第一处示意。
  async function selectReplaceFirst(host, args) {
    if (host !== 'Word' || !args.search) return;
    try {
      await Word.run(async (ctx) => {
        const results = ctx.document.body.search(String(args.search), { matchCase: false, ignoreSpace: true });
        results.load('length');
        await ctx.sync();
        if (results.items && results.items.length > 0) { results.items[0].select(); await ctx.sync(); }
      });
    } catch (e) { /* 预选失败不阻塞审批 */ }
  }
  async function selectExcelTargetRange(host, args) {
    if (host !== 'Excel' || !args.address) return;
    try {
      await Excel.run(async (ctx) => { excelRange(ctx, String(args.address)).select(); await ctx.sync(); });
    } catch (e) { /* 预选失败不阻塞审批 */ }
  }

  // ================= 注册表 =================

  // 机制级 re-read：覆盖类操作（destructive）必须先预览当前状态，确认后带 confirm:true 才执行。
  // 每个覆盖类 action 映射到"读当前状态"的函数，预览结果随 confirm_required 返回给 AI。
  const PREVIEW_READERS = {
    write_selection: (host, args) => readSelection(host, {}),
    replace_all: (host, args) => replaceAll(host, Object.assign({}, args, { dryRun: true })),
    remove_empty_paragraphs: (host, args) => removeEmptyParagraphs(host, Object.assign({}, args, { dryRun: true })),
    delete_sheet: (host, args) => deleteSheet(host, Object.assign({}, args, { dryRun: true })),
    format_selection: (host, args) => readSelection(host, { withStyles: true }),
    write_range: (host, args) => readRange(host, { address: args.address }),
    format_range: (host, args) => readRange(host, { address: args.address }),
    rename_sheet: (host, args) => listSheets(host, {}),
    apply_sort: (host, args) => readRange(host, { address: args.address }),
    apply_filter: (host, args) => readRange(host, { address: args.address }),
    set_font: (host, args) => okResult({ host, wouldSetFont: args.font, note: '影响全文所有段落（含表格内文字），属覆盖类操作' }),
    apply_style: (host, args) => readSelection(host, {}),
    insert_paragraph: (host, args) => readSelection(host, {}),
    append_text: (host, args) => readSelection(host, {}),
  };

  // 注册表（外壳 execute 按 meta.destructive + meta.preview 做机制级 re-read 确认）
  const ACTIONS_TABLE = {
    // 通用
    read_selection: { hosts: ['Word', 'Excel', 'PowerPoint'], destructive: false, impl: readSelection },
    write_selection: { hosts: ['Word', 'Excel', 'PowerPoint'], destructive: true, impl: writeSelection },
    read_document: { hosts: ['Word', 'Excel', 'PowerPoint'], destructive: false, impl: readDocument },
    read_styles: { hosts: ['Word', 'Excel'], destructive: false, impl: readStyles },
    replace_all: { hosts: ['Word', 'Excel'], destructive: true, impl: replaceAll, selectTarget: selectReplaceFirst },
    append_text: { hosts: ['Word'], destructive: true, impl: appendText },
    // Word
    read_tables: { hosts: ['Word'], destructive: false, impl: readTables },
    set_font: { hosts: ['Word'], destructive: true, impl: setFont },
    remove_empty_paragraphs: { hosts: ['Word'], destructive: true, impl: removeEmptyParagraphs },
    insert_paragraph: { hosts: ['Word'], destructive: true, impl: insertParagraph },
    insert_table: { hosts: ['Word'], destructive: false, impl: insertTable },
    insert_image: { hosts: ['Word'], destructive: false, impl: insertImage },
    apply_style: { hosts: ['Word'], destructive: true, impl: applyStyle },
    format_selection: { hosts: ['Word', 'Excel'], destructive: true, impl: formatSelection },
    set_paragraph_format: { hosts: ['Word'], destructive: true, impl: setParagraphFormat },
    search: { hosts: ['Word'], destructive: false, impl: search },
    add_comment: { hosts: ['Word', 'Excel'], destructive: false, impl: addComment },
    read_comments: { hosts: ['Word', 'Excel'], destructive: false, impl: readComments },
    read_properties: { hosts: ['Word', 'Excel'], destructive: false, impl: readProperties },
    // Excel
    list_sheets: { hosts: ['Excel'], destructive: false, impl: listSheets },
    read_range: { hosts: ['Excel'], destructive: false, impl: readRange },
    write_range: { hosts: ['Excel'], destructive: true, impl: writeRange, selectTarget: selectExcelTargetRange },
    format_range: { hosts: ['Excel'], destructive: true, impl: formatRange, selectTarget: selectExcelTargetRange },
    insert_chart: { hosts: ['Excel'], destructive: false, impl: insertChart },
    add_sheet: { hosts: ['Excel'], destructive: false, impl: addSheet },
    rename_sheet: { hosts: ['Excel'], destructive: true, impl: renameSheet },
    delete_sheet: { hosts: ['Excel'], destructive: true, impl: deleteSheet },
    apply_sort: { hosts: ['Excel'], destructive: true, impl: applySort },
    apply_filter: { hosts: ['Excel'], destructive: true, impl: applyFilter },
    evaluate_formula: { hosts: ['Excel'], destructive: false, impl: evaluateFormula },
    // PPT
    read_slides: { hosts: ['PowerPoint'], destructive: false, impl: readSlides },
    ppt_read_notes: { hosts: ['PowerPoint'], destructive: false, impl: readPptNotes },
    // 环境诊断
    get_environment: { hosts: ['Word', 'Excel', 'PowerPoint'], destructive: false, impl: getEnvironment },
    // 定位 + 选中（零副作用：只做选中/取消选中 UI 反馈，不改文档；blinks>0 才闪烁）
    locate_select: { hosts: ['Word', 'Excel'], destructive: false, impl: locateSelect },
  };
  // 注入 preview（机制级 re-read：覆盖类操作先读当前状态，AI 确认后带 confirm:true 才执行）
  for (const [name, meta] of Object.entries(ACTIONS_TABLE)) {
    if (meta.destructive && PREVIEW_READERS[name]) meta.preview = PREVIEW_READERS[name];
  }
  window.__ACTIONS__ = ACTIONS_TABLE;

  // 分发器（外壳 execute 转发到这里，便于热更新拦截逻辑）
  // confirmMode: auto（默认）= 覆盖操作自动 re-read 后执行，结果附 previousState 供 AI 核验，不打扰用户；
  //              ask = re-read 后必须用户确认（confirm:true）才执行。
  window.__EXECUTE__ = function (action, args) {
    // 用户定制优先（user-actions.js 可同名覆盖内置 action 或新增）
    const userActions = window.__USER_ACTIONS__ || {};
    const meta = userActions[action] || ACTIONS_TABLE[action];
    if (!meta) return errResult('unknown_action', `unknown action: ${action}`);
    const host = String(Office.context.host || '');
    if (!meta.hosts.includes(host)) return errResult('unsupported_host', `${action} not supported in ${host}`);
    const run = () => Promise.resolve().then(() => meta.impl(host, args || {})).catch((e) => {
      const debug = (e && e.debugInfo) ? (' [debug] ' + JSON.stringify(e.debugInfo)) : '';
      return errResult('execution', String(e && e.message || e) + debug);
    });
    // 写后验证：覆盖操作执行成功后，再读一次被操作区域（afterState）+ 附整个文档全文（避免盲人摸象）
    const withAfterState = (r) => {
      if (meta.destructive && r.ok && meta.preview && r.result && typeof r.result === 'object') {
        return Promise.resolve(meta.preview(host, args)).then((a) => {
          if (a.ok) r.result.afterState = a.result;
          // 附整个文档：Word 全文 / Excel 工作表已用区域；PPT 全文件读取慢，跳过
          if (host !== 'PowerPoint') {
            return Promise.resolve(readDocument(host, {})).then((d) => {
              if (d.ok) r.result.document = d.result.text;
              return r;
            });
          }
          return r;
        });
      }
      return Promise.resolve(r);
    };
    // 覆盖类操作：先取确认模式
    const modeP = window.__CONFIRM_MODE__ !== undefined
      ? Promise.resolve(window.__CONFIRM_MODE__)
      : fetch('/office/config').then((r) => r.json()).then((c) => { window.__CONFIRM_MODE__ = (c && c.confirmMode) || 'auto'; return window.__CONFIRM_MODE__; }).catch(() => { window.__CONFIRM_MODE__ = 'auto'; return 'auto'; });
    if (meta.destructive && args && args.confirm !== true) {
      if (!meta.preview) return errResult('confirm_required', `覆盖类操作 ${action} 需先预览确认（confirm:true）`);
      return modeP.then((mode) => Promise.resolve(meta.preview(host, args)).then((p) => {
        if (!p.ok) return p;
        if (mode === 'auto') {
          // 自动模式：改前 re-read（previousState）→ 执行 → 改后验证（afterState）；标记 approvedBy=auto
          return run().then((r) => {
            if (r.ok && r.result && typeof r.result === 'object') {
              r.result.previousState = p.result;
              r.result.approvedBy = 'auto';
            }
            return withAfterState(r);
          });
        }
        // ask 模式：必须用户确认；先选中将要修改的目标（审批可视化，失败不阻塞）
        const preSelect = meta.selectTarget ? Promise.resolve(meta.selectTarget(host, args)).catch(() => {}) : Promise.resolve();
        return preSelect.then(() => ({ ok: false, code: 'confirm_required', error: '覆盖类操作需确认：请向用户展示当前状态预览（result.preview），确认后带 confirm:true 重发', result: { preview: p.result, action, args, selectedTarget: true } }));
      }));
    }
    return withAfterState(run());
  };
})();
