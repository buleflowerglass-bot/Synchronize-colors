const PROP_KEY = 'DOC_COLOR_SYNC_SETTINGS_V2';

const DEFAULTS = {
  docUrl: '',              // 谷歌文档链接或 ID
  docColumn: 2,            // 文档表格中英文所在的列（从左数第几列）
  sheetUrl: '',            // 谷歌表格链接或 ID
  sheetName: '',           // 空 = 按链接里的 gid，没有 gid 则用第一个标签
  column: 'E',             // 表格中英文所在列
  startRow: 2,
  mode: 'background',      // background = 高亮色；text = 文字颜色
  ignoreColors: '',        // 逗号分隔，如 #cccccc,#d9d9d9
  cellBg: true,            // 整格填充底色
  clearUnmatched: false,   // 无颜色的单元格清除底色（仅在整格底色开启时有效）
  markItalic: true,        // 高亮文字改斜体
  markUnderline: false,    // 高亮文字加下划线
  markTextColor: false,    // 高亮文字改成对应颜色
  clearOldMarks: false,    // 清除已匹配单元格里的斜体/下划线
  maxSkip: 5000
};

/***** 网页入口 *****/
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('文档颜色同步')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/***** 设置读写（保存在你自己的账号下）*****/
function getSettings() {
  const raw = PropertiesService.getUserProperties().getProperty(PROP_KEY);
  const saved = raw ? JSON.parse(raw) : {};
  return Object.assign({}, DEFAULTS, saved);
}

function saveSettings(s) {
  PropertiesService.getUserProperties().setProperty(PROP_KEY, JSON.stringify(s));
}

function getInitData() {
  const settings = getSettings();
  let sheets = [];
  if (settings.sheetUrl) {
    try { sheets = getSheetNames(settings.sheetUrl); } catch (e) {}
  }
  return { settings: settings, sheets: sheets };
}

// 读取表格里所有标签名
function getSheetNames(sheetUrl) {
  const ss = openSpreadsheet_(sheetUrl);
  return ss.getSheets().map(sh => sh.getName());
}

/***** 页面按钮调用 *****/
function saveAndSync(s) {
  saveSettings(s);
  return syncColors(s);
}

// 列出文档英文列里出现的颜色及字数
function detectColors(docInput, docColumn, mode) {
  const doc = DocumentApp.openById(extractId_(docInput));
  const { colors } = extractDocChars_(doc.getBody(), mode, [], Number(docColumn) || 2);
  const count = {};
  colors.forEach(c => { if (c) count[c] = (count[c] || 0) + 1; });
  return Object.keys(count)
    .map(c => ({ color: c, count: count[c] }))
    .sort((a, b) => b.count - a.count);
}

/***** 核心同步 *****/
function syncColors(cfg) {
  if (!cfg.docUrl) return '请先填写文档链接或 ID。';
  if (!cfg.sheetUrl) return '请先填写表格链接或 ID。';

  let ss;
  try {
    ss = openSpreadsheet_(cfg.sheetUrl);
  } catch (e) {
    return '无法打开表格：' + e.message;
  }
  const sheet = resolveSheet_(ss, cfg);
  if (!sheet) return '找不到工作表：' + cfg.sheetName;

  const colLetter = String(cfg.column || 'E').trim().toUpperCase();
  if (!/^[A-Z]{1,3}$/.test(colLetter)) return '列格式不正确，请填写字母，例如 E。';
  const col = sheet.getRange(colLetter + '1').getColumn();
  const startRow = Number(cfg.startRow) || 1;

  if (!cfg.cellBg && !cfg.markItalic && !cfg.markUnderline && !cfg.markTextColor && !cfg.clearOldMarks) {
    return '请至少勾选一种标记方式。';
  }

  let doc;
  try {
    doc = DocumentApp.openById(extractId_(cfg.docUrl));
  } catch (e) {
    return '无法打开文档：' + e.message;
  }

  const ignore = String(cfg.ignoreColors || '')
    .split(',').map(s => s.trim().toLowerCase()).filter(Boolean);

  const { text: docStr, colors: docColors } =
    extractDocChars_(doc.getBody(), cfg.mode, ignore, Number(cfg.docColumn) || 2);

  if (!docStr) return '文档里没有读取到英文内容，请检查“文档表格中英文所在列”。';

  const lastRow = sheet.getLastRow();
  const numRows = lastRow - startRow + 1;
  if (numRows <= 0) return '该列没有可处理的数据。';

  const range = sheet.getRange(startRow, col, numRows, 1);
  const values = range.getValues();
  const formulas = range.getFormulas();
  const bgs = range.getBackgrounds();
  const rich = range.getRichTextValues();
  const newRich = rich.map(r => r[0]);
  const changed = new Array(numRows).fill(false);

  const doRich = cfg.markItalic || cfg.markUnderline || cfg.markTextColor || cfg.clearOldMarks;
  const maxSkip = Number(cfg.maxSkip) || 5000;

  let pointer = 0, bgCount = 0, markedCells = 0;
  const unmatched = [];

  for (let i = 0; i < numRows; i++) {
    const { norm, map } = normalizeWithMap_(String(values[i][0]));
    if (!norm) continue;

    const m = findFrom_(docStr, norm, pointer, maxSkip);
    if (!m) { unmatched.push(startRow + i); continue; }
    pointer = m.pos + m.len;

    // 1) 整格底色：取该单元格范围内字数最多的颜色
    if (cfg.cellBg) {
      const count = {};
      for (let k = m.pos; k < m.pos + m.len; k++) {
        const c = docColors[k];
        if (c) count[c] = (count[c] || 0) + 1;
      }
      let best = null, bestN = 0;
      for (const c in count) if (count[c] > bestN) { best = c; bestN = count[c]; }
      if (best) { bgs[i][0] = best; bgCount++; }
      else if (cfg.clearUnmatched) { bgs[i][0] = '#ffffff'; }
    }

    // 2) 精确标记高亮的那几个词
    if (doRich && !formulas[i][0]) {
      const marks = [];
      let k = 0;
      while (k < m.len) {
        const c = docColors[m.pos + k];
        if (!c) { k++; continue; }
        let j = k;
        while (j + 1 < m.len && docColors[m.pos + j + 1] === c) j++;
        marks.push({ s: map[k], e: map[j] + 1, color: c });
        k = j + 1;
      }
      if (marks.length || cfg.clearOldMarks) {
        newRich[i] = applyMarks_(rich[i][0], marks, cfg);
        changed[i] = true;
        if (marks.length) markedCells++;
      }
    }
  }

  if (cfg.cellBg) range.setBackgrounds(bgs);

  // 只把有改动的连续区块写回
  let i = 0;
  while (i < numRows) {
    if (!changed[i]) { i++; continue; }
    let j = i;
    while (j + 1 < numRows && changed[j + 1]) j++;
    sheet.getRange(startRow + i, col, j - i + 1, 1)
      .setRichTextValues(newRich.slice(i, j + 1).map(v => [v]));
    i = j + 1;
  }

  let msg = '完成：工作表「' + sheet.getName() + '」' + colLetter + ' 列。';
  if (cfg.cellBg) msg += '\n整格标色：' + bgCount + ' 个单元格。';
  if (doRich) msg += '\n文字标记：' + markedCells + ' 个单元格含高亮文字。';
  if (unmatched.length) {
    msg += '\n未能匹配的行（共 ' + unmatched.length + ' 个）：' +
      unmatched.slice(0, 30).join(', ') + (unmatched.length > 30 ? ' ...' : '');
  }
  return msg;
}

/***** 富文本标记 *****/
function applyMarks_(rv, marks, cfg) {
  let base = rv;

  if (cfg.clearOldMarks) {
    const b = rv.copy();
    rv.getRuns().forEach(run => {
      const st = run.getTextStyle().copy().setItalic(false).setUnderline(false).build();
      b.setTextStyle(run.getStartIndex(), run.getEndIndex(), st);
    });
    base = b.build();
  }
  if (!marks.length) return base;

  const builder = base.copy();
  const runs = base.getRuns();

  marks.forEach(mk => {
    runs.forEach(run => {
      const s = Math.max(mk.s, run.getStartIndex());
      const e = Math.min(mk.e, run.getEndIndex());
      if (s >= e) return;
      const sb = run.getTextStyle().copy();
      if (cfg.markItalic) sb.setItalic(true);
      if (cfg.markUnderline) sb.setUnderline(true);
      if (cfg.markTextColor) sb.setForegroundColor(mk.color);
      builder.setTextStyle(s, e, sb.build());
    });
  });
  return builder.build();
}

/***** 表格与文档定位 *****/
function extractId_(s) {
  const m = String(s).match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : String(s).trim();
}

function extractGid_(s) {
  const m = String(s).match(/[#&?]gid=(\d+)/);
  return m ? Number(m[1]) : null;
}

function openSpreadsheet_(urlOrId) {
  return SpreadsheetApp.openById(extractId_(urlOrId));
}

// 优先用选择的标签；没选则用链接里的 gid；都没有就用第一个标签
function resolveSheet_(ss, cfg) {
  if (cfg.sheetName) return ss.getSheetByName(cfg.sheetName);
  const gid = extractGid_(cfg.sheetUrl);
  if (gid !== null) {
    const hit = ss.getSheets().filter(sh => sh.getSheetId() === gid)[0];
    if (hit) return hit;
  }
  return ss.getSheets()[0];
}

/***** 文档读取 *****/
// 只读取文档中所有表格里指定列（英文列）的文字，按行顺序拼接
function extractDocChars_(body, mode, ignore, docColumn) {
  const textEls = [];
  const n = body.getNumChildren();
  for (let i = 0; i < n; i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.TABLE) continue;
    const table = child.asTable();
    for (let r = 0; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      if (row.getNumCells() >= docColumn) {
        collectTexts_(row.getCell(docColumn - 1), textEls);
      }
    }
  }

  let str = '';
  const colors = [];

  textEls.forEach(t => {
    const txt = t.getText();
    if (!txt) return;
    const idxs = t.getTextAttributeIndices();
    for (let k = 0; k < idxs.length; k++) {
      const start = idxs[k];
      const end = k + 1 < idxs.length ? idxs[k + 1] : txt.length;
      let color = mode === 'text' ? t.getForegroundColor(start) : t.getBackgroundColor(start);
      if (color) color = color.toLowerCase();
      if (mode === 'text' && color === '#000000') color = null;
      if (color && ignore.indexOf(color) !== -1) color = null;

      for (let i = start; i < end; i++) {
        const ch = txt.charAt(i);
        if (/\s/.test(ch)) continue;
        str += normChar_(ch);
        colors.push(color);
      }
    }
  });
  return { text: str, colors: colors };
}

function collectTexts_(el, out) {
  if (el.getType() === DocumentApp.ElementType.TEXT) { out.push(el.asText()); return; }
  if (el.getNumChildren) {
    for (let i = 0; i < el.getNumChildren(); i++) collectTexts_(el.getChild(i), out);
  }
}

/***** 文本规范化与匹配 *****/
function normChar_(ch) {
  if (ch === '\u2019' || ch === '\u2018') return "'";
  if (ch === '\u201C' || ch === '\u201D') return '"';
  if (ch === '\u2013' || ch === '\u2014' || ch === '\u2015') return '-';
  return ch;
}

// 去掉空白并统一符号，同时记录每个字符在原文中的位置
function normalizeWithMap_(s) {
  let norm = '';
  const map = [];
  for (let i = 0; i < s.length; i++) {
    const ch = s.charAt(i);
    if (/\s/.test(ch)) continue;
    norm += normChar_(ch);
    map.push(i);
  }
  return { norm: norm, map: map };
}

// 从 pointer 起向后查找，返回 {pos, len}；整句找不到时用前 20 个字符兜底
function findFrom_(hay, needle, pointer, maxSkip) {
  const limit = Math.min(hay.length, pointer + maxSkip + needle.length);
  const win = hay.substring(pointer, limit);
  let p = win.indexOf(needle);
  if (p >= 0) return { pos: pointer + p, len: needle.length };
  if (needle.length > 20) {
    p = win.indexOf(needle.substring(0, 20));
    if (p >= 0) return { pos: pointer + p, len: 20 };
  }
  return null;
}