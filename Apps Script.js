const PROP_KEY = 'DOC_COLOR_SYNC_SETTINGS';

const DEFAULTS = {
  docId: '',
  docColumn: 2,           // 文档表格中英文所在的列（从左数第几列）
  sheetName: '',          // 空 = 当前工作表
  column: 'E',            // 表格中英文所在列
  startRow: 2,
  mode: 'background',     // background = 高亮色；text = 文字颜色
  ignoreColors: '',       // 逗号分隔，如 #cccccc,#d9d9d9
  clearUnmatched: false,
  maxSkip: 5000
};

/***** 菜单与侧边栏 *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('文档颜色同步')
    .addItem('打开设置面板', 'showSidebar')
    .addItem('按上次设置直接同步', 'runWithSavedSettings')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('文档颜色同步')
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

/***** 设置读写（供侧边栏调用）*****/
function getSettings() {
  const raw = PropertiesService.getDocumentProperties().getProperty(PROP_KEY);
  const saved = raw ? JSON.parse(raw) : {};
  return Object.assign({}, DEFAULTS, saved);
}

function saveSettings(s) {
  PropertiesService.getDocumentProperties().setProperty(PROP_KEY, JSON.stringify(s));
}

function getInitData() {
  return {
    settings: getSettings(),
    sheets: SpreadsheetApp.getActive().getSheets().map(sh => sh.getName())
  };
}

/***** 入口 *****/
function runWithSavedSettings() {
  const msg = syncColors(getSettings());
  SpreadsheetApp.getUi().alert(msg);
}

// 侧边栏点击“同步”时调用：先保存再执行，返回结果文字
function saveAndSync(s) {
  saveSettings(s);
  return syncColors(s);
}

// 侧边栏“检测颜色”按钮：列出文档英文列里出现的颜色及字数
function detectColors(docInput, docColumn, mode) {
  const docId = extractDocId_(docInput);
  const doc = DocumentApp.openById(docId);
  const { colors } = extractDocChars_(doc.getBody(), mode, [], Number(docColumn) || 2);
  const count = {};
  colors.forEach(c => { if (c) count[c] = (count[c] || 0) + 1; });
  return Object.keys(count)
    .map(c => ({ color: c, count: count[c] }))
    .sort((a, b) => b.count - a.count);
}

/***** 核心同步 *****/
function syncColors(cfg) {
  if (!cfg.docId) return '请先填写文档链接或 ID。';

  const docId = extractDocId_(cfg.docId);
  const ss = SpreadsheetApp.getActive();
  const sheet = cfg.sheetName ? ss.getSheetByName(cfg.sheetName) : ss.getActiveSheet();
  if (!sheet) return '找不到工作表：' + cfg.sheetName;

  const colLetter = String(cfg.column || 'E').trim().toUpperCase();
  if (!/^[A-Z]{1,3}$/.test(colLetter)) return '列格式不正确，请填写字母，例如 E。';
  const col = sheet.getRange(colLetter + '1').getColumn();
  const startRow = Number(cfg.startRow) || 1;

  let doc;
  try {
    doc = DocumentApp.openById(docId);
  } catch (e) {
    return '无法打开文档，请检查链接/ID 和访问权限。';
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
  const bgs = range.getBackgrounds();

  let pointer = 0, colored = 0;
  const unmatched = [];
  const maxSkip = Number(cfg.maxSkip) || 5000;

  for (let i = 0; i < values.length; i++) {
    const norm = normalize_(String(values[i][0]));
    if (!norm) continue;

    const pos = findFrom_(docStr, norm, pointer, maxSkip);
    if (pos < 0) { unmatched.push(startRow + i); continue; }
    pointer = pos + norm.length;

    // 统计该范围内字数最多的颜色
    const count = {};
    for (let k = pos; k < pos + norm.length; k++) {
      const c = docColors[k];
      if (c) count[c] = (count[c] || 0) + 1;
    }
    let best = null, bestN = 0;
    for (const c in count) if (count[c] > bestN) { best = c; bestN = count[c]; }

    if (best) { bgs[i][0] = best; colored++; }
    else if (cfg.clearUnmatched) { bgs[i][0] = '#ffffff'; }
  }

  range.setBackgrounds(bgs);

  let msg = '完成：工作表「' + sheet.getName() + '」' + colLetter + ' 列，已标色 ' + colored + ' 个单元格。';
  if (unmatched.length) {
    msg += '\n未能匹配的行（共 ' + unmatched.length + ' 个）：' +
      unmatched.slice(0, 30).join(', ') + (unmatched.length > 30 ? ' ...' : '');
  }
  return msg;
}

/***** 工具函数 *****/
function extractDocId_(s) {
  const m = String(s).match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : String(s).trim();
}

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

function normChar_(ch) {
  if (ch === '\u2019' || ch === '\u2018') return "'";
  if (ch === '\u201C' || ch === '\u201D') return '"';
  if (ch === '\u2013' || ch === '\u2014' || ch === '\u2015') return '-';
  return ch;
}

function normalize_(s) {
  let out = '';
  for (let i = 0; i < s.length; i++) {
    const ch = s.charAt(i);
    if (/\s/.test(ch)) continue;
    out += normChar_(ch);
  }
  return out;
}

// 从 pointer 起向后查找；找不到时用前 20 个字符做兜底匹配
function findFrom_(hay, needle, pointer, maxSkip) {
  const limit = Math.min(hay.length, pointer + maxSkip + needle.length);
  const win = hay.substring(pointer, limit);
  let p = win.indexOf(needle);
  if (p >= 0) return pointer + p;
  if (needle.length > 20) {
    p = win.indexOf(needle.substring(0, 20));
    if (p >= 0) return pointer + p;
  }
  return -1;
}