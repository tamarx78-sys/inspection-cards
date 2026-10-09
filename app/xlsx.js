// xlsx の読み取り・レイアウト推定・セル書き戻し。
// 書き戻しはシート XML の該当セルだけを文字列レベルで差し替えるので、
// 罫線・結合・書式・その他の要素は元のまま残る。
import { readZip, writeZip } from './zip.js';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export const colName = (c) => {
  let s = '';
  for (; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s;
  return s;
};
export const colNum = (s) => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
export const parseRef = (ref) => {
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(ref);
  return m ? { c: colNum(m[1]), r: +m[2] } : null;
};
const key = (r, c) => `${r},${c}`;

const parseXml = (text) => new DOMParser().parseFromString(text, 'application/xml');
const all = (node, name) => [...node.getElementsByTagNameNS('*', name)];
const first = (node, name) => node.getElementsByTagNameNS('*', name)[0] || null;

function resolvePath(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

// ---- 日付・時刻の書式判定 ----
// 組み込み書式番号は日本語ロケールの解釈 (55, 56 は「yyyy年m月」「m月d日」)
const BUILTIN_DATE = new Set([14, 15, 16, 17, 22, 27, 28, 29, 30, 31, 34, 35, 36, 50, 51, 52, 53, 54, 55, 56, 57, 58]);
const BUILTIN_TIME = new Set([18, 19, 20, 21, 32, 33, 45, 46, 47]);
function classifyFormat(id, code) {
  if (BUILTIN_DATE.has(id)) return 'date';
  if (BUILTIN_TIME.has(id)) return 'time';
  if (!code) return null;
  const s = code.split(';')[0].replace(/"[^"]*"|\[[^\]]*\]|\\./g, '').toLowerCase();
  if (/[yd]/.test(s)) return 'date';
  if (/[hs]/.test(s)) return 'time';
  if (/m/.test(s) && !/[0#?]/.test(s)) return 'date';
  return null;
}

export function serialToDate(v) {
  const ms = Math.round((v - 25569) * 86400000);
  const d = new Date(ms);
  return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate() };
}
export function serialToTime(v) {
  const mins = Math.round((v % 1) * 1440);
  return { h: Math.floor(mins / 60) % 24, mi: mins % 60 };
}
export const dateToSerial = (y, m, d) => (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;
export const timeToSerial = (h, mi) => (h * 60 + mi) / 1440;
const pad = (n) => String(n).padStart(2, '0');

/** xlsx バイト列を読み込んでワークブックオブジェクトを返す */
export async function openWorkbook(buffer) {
  const zip = readZip(buffer);
  const wbPath = 'xl/workbook.xml';
  const wbXml = parseXml(await zip.readText(wbPath));
  const relsXml = parseXml(await zip.readText('xl/_rels/workbook.xml.rels'));
  const rels = new Map(all(relsXml, 'Relationship').map((r) => [r.getAttribute('Id'), r.getAttribute('Target')]));
  const sheets = all(wbXml, 'sheet').map((s) => ({
    name: s.getAttribute('name'),
    path: resolvePath(wbPath, rels.get(s.getAttributeNS(NS_REL, 'id') || s.getAttribute('r:id'))),
  }));

  let shared = [];
  const sstText = await zip.readText('xl/sharedStrings.xml');
  if (sstText) {
    shared = all(parseXml(sstText), 'si').map((si) => richText(si));
  }

  const formats = []; // cellXfs index -> 'date' | 'time' | null
  const stText = await zip.readText('xl/styles.xml');
  if (stText) {
    const st = parseXml(stText);
    const custom = new Map(all(st, 'numFmt').map((n) => [+n.getAttribute('numFmtId'), n.getAttribute('formatCode')]));
    const cellXfs = first(st, 'cellXfs');
    if (cellXfs) {
      for (const xf of [...cellXfs.children].filter((x) => x.localName === 'xf')) {
        const id = +(xf.getAttribute('numFmtId') || 0);
        formats.push(classifyFormat(id, custom.get(id)));
      }
    }
  }

  return { zip, sheets, shared, formats };
}

function richText(node) {
  // <si>/<is> 内の <t> を連結 (ふりがな <rPh> は除外)
  return all(node, 't')
    .filter((t) => t.parentNode.localName !== 'rPh')
    .map((t) => t.textContent)
    .join('');
}

/** シートを読み込み、セル・結合情報を返す */
export async function loadSheet(wb, sheet) {
  const text = await wb.zip.readText(sheet.path);
  const doc = parseXml(text);
  const cells = new Map();
  let maxR = 0, maxC = 0;
  for (const c of all(doc, 'c')) {
    const pos = parseRef(c.getAttribute('r'));
    if (!pos) continue;
    const t = c.getAttribute('t') || 'n';
    const s = +(c.getAttribute('s') || 0);
    const vNode = first(c, 'v');
    const hasF = !!first(c, 'f');
    let value = null;
    if (t === 's' && vNode) value = wb.shared[+vNode.textContent] ?? '';
    else if (t === 'inlineStr') value = first(c, 'is') ? richText(first(c, 'is')) : '';
    else if (t === 'b' && vNode) value = vNode.textContent === '1';
    else if ((t === 'str' || t === 'e') && vNode) value = vNode.textContent;
    else if (vNode && vNode.textContent !== '') value = +vNode.textContent;
    const fmt = wb.formats[s] || null;
    const empty = (value === null || value === '') && !hasF;
    cells.set(key(pos.r, pos.c), { ...pos, t, s, value, fmt, hasF, empty });
    if (!empty) { maxR = Math.max(maxR, pos.r); maxC = Math.max(maxC, pos.c); }
  }
  const merges = all(doc, 'mergeCell').map((m) => {
    const [a, b] = m.getAttribute('ref').split(':').map(parseRef);
    return { r1: a.r, c1: a.c, r2: (b || a).r, c2: (b || a).c };
  });
  const parseRanges = (sqref) => (sqref || '').split(/\s+/).filter(Boolean).map((part) => {
    const [a, b] = part.split(':').map(parseRef);
    return a ? { r1: a.r, c1: a.c, r2: (b || a).r, c2: (b || a).c } : null;
  }).filter(Boolean);
  // 入力規則 (プルダウン) の選択肢。同じシート内の範囲か "a,b,c" 形式のみ対応
  const validations = all(doc, 'dataValidation')
    .filter((v) => v.getAttribute('type') === 'list')
    .map((v) => ({
      ranges: parseRanges(v.getAttribute('sqref')),
      formula: (first(v, 'formula1')?.textContent || '').trim(),
    }));
  const sheetObj = {
    ...sheet, xml: text, cells, merges, maxR, maxC,
    get(r, c) { return cells.get(key(r, c)) || null; },
    validationOptions(r, c) {
      const v = validations.find((x) => x.ranges.some((m) => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2));
      if (!v) return null;
      if (v.formula.startsWith('"')) return v.formula.slice(1, -1).split(',').map((s) => s.trim()).filter(Boolean);
      if (v.formula.includes('!')) return null; // 他シート参照は未対応
      const [a, b] = v.formula.split(':').map(parseRef);
      if (!a) return null;
      const out = [];
      for (let rr = a.r; rr <= (b || a).r; rr++) {
        for (let cc = a.c; cc <= (b || a).c; cc++) {
          const s = displayValue(this.get(rr, cc)).trim();
          if (s) out.push(s);
        }
      }
      return out.length ? out : null;
    },
    /** 結合セルを考慮した値 (結合範囲内なら左上の値) */
    valueAt(r, c) {
      const m = merges.find((m) => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2);
      const cell = m ? this.get(m.r1, m.c1) : this.get(r, c);
      return cell && !cell.empty ? cell : null;
    },
  };
  return sheetObj;
}

export function displayValue(cell) {
  if (!cell || cell.empty) return '';
  const v = cell.value;
  if (typeof v === 'number') {
    if (cell.fmt === 'date') { const d = serialToDate(v); return `${d.y}/${d.m}/${d.d}`; }
    if (cell.fmt === 'time') { const t = serialToTime(v); return `${t.h}:${pad(t.mi)}`; }
    return String(+v.toPrecision(12));
  }
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (cell.hasF && (v === null || v === '')) return '(数式)';
  return String(v);
}

// ---- レイアウト推定 (横に記録を追加していく点検表) ----
const META_LABELS = /^(日付|点検日|実施日|日時|年月日)$/;
const ROLE_PATTERNS = [
  ['name', /機器|名称|項目|点検箇所|設備/],
  ['code', /記号|コード|ID|タグ/i],
  ['range', /管理値|基準|規格|範囲|許容/],
  ['contact', /連絡|担当|部署/],
];

const text = (cell) => (cell && typeof cell.value === 'string' ? cell.value.trim() : '');

/** 管理値テキストから {min, max} を推定 */
export function parseRange(s) {
  if (!s) return null;
  const n = s.normalize('NFKC').replace(/\s/g, '');
  const num = '(-?\\d+(?:\\.\\d+)?)';
  let m = new RegExp(`^${num}[~〜～\\-]${num}$`).exec(n);
  if (m) return { min: +m[1], max: +m[2] };
  m = new RegExp(`^(?:≦|<=|≤)${num}$|^${num}以下$`).exec(n);
  if (m) return { max: +(m[1] ?? m[2]) };
  m = new RegExp(`^(?:≧|>=|≥)${num}$|^${num}以上$`).exec(n);
  if (m) return { min: +(m[1] ?? m[2]) };
  m = new RegExp(`^${num}±${num}$`).exec(n);
  if (m) return { min: +m[1] - +m[2], max: +m[1] + +m[2] };
  return null;
}

export function detectLayout(sheet) {
  // 1. 「日付」ラベルのある列 = ラベル列。その右が記録列
  let labelCol = 0, dateRow = 0;
  for (let r = 1; r <= Math.min(15, sheet.maxR) && !labelCol; r++) {
    for (let c = 1; c <= sheet.maxC; c++) {
      if (META_LABELS.test(text(sheet.get(r, c)))) { labelCol = c; dateRow = r; break; }
    }
  }
  if (!labelCol) throw new Error('「日付」ラベルが見つからないため、表の構造を判定できませんでした');

  // 2. ヘッダ行: ラベル列より左に文字列が 2 つ以上ある最初の行
  let headerRow = 0;
  for (let r = 1; r <= Math.min(20, sheet.maxR); r++) {
    let n = 0;
    for (let c = 1; c < labelCol; c++) if (text(sheet.get(r, c))) n++;
    if (n >= 2) { headerRow = r; break; }
  }
  if (!headerRow) headerRow = dateRow;

  // 3. 列の役割
  const columns = [];
  for (let c = 1; c < labelCol; c++) {
    const h = text(sheet.get(headerRow, c));
    const role = ROLE_PATTERNS.find(([, re]) => re.test(h))?.[0] || (h ? 'other' : 'section');
    columns.push({ c, header: h, role });
  }
  // 名称列が見つからなければ、section 以外の最初の列を名称とする
  if (!columns.some((x) => x.role === 'name')) {
    const f = columns.find((x) => x.role !== 'section');
    if (f) f.role = 'name';
  }

  // 4. メタ行 (日付・時間・名前など): ヘッダ行までにラベル列に文字がある行
  const metaRows = [];
  for (let r = 1; r <= headerRow; r++) {
    let label = text(sheet.get(r, labelCol));
    if (!label) continue;
    if (label.includes('|')) label = label.split('|').pop().trim();
    let type = 'text';
    if (META_LABELS.test(label)) type = 'date';
    else if (/時間|時刻/.test(label)) type = 'time';
    metaRows.push({ row: r, label, type });
  }

  // 5. 項目行: ヘッダ行より下で、左側の列に何か書いてある行
  const itemRows = [];
  for (let r = headerRow + 1; r <= sheet.maxR; r++) {
    const has = columns.some((x) => x.role !== 'section' && text(sheet.get(r, x.c)));
    if (has) itemRows.push(r);
  }

  const allRows = [...metaRows.map((m) => m.row), ...itemRows];
  const recordCols = [];
  for (let c = labelCol + 1; c <= sheet.maxC; c++) {
    if (allRows.some((r) => sheet.get(r, c) && !sheet.get(r, c).empty)) recordCols.push(c);
  }
  const lastCol = recordCols.length ? recordCols[recordCols.length - 1] : labelCol;

  return {
    sheetName: sheet.name,
    labelCol, headerRow, columns, metaRows, itemRows, recordCols,
    targetCol: lastCol + 1,
  };
}

/** レイアウトと対象列から、カード定義を作る */
export function buildCards(sheet, layout, targetCol) {
  const prevCols = layout.recordCols.filter((c) => c < targetCol);
  const prevCol = prevCols.length ? prevCols[prevCols.length - 1] : null;
  const role = (r, name) => {
    const col = layout.columns.find((x) => x.role === name);
    return col ? displayValue(sheet.valueAt(r, col.c)) : '';
  };
  const typeFromCell = (cell, fallback) => {
    if (!cell || cell.empty) return fallback;
    if (typeof cell.value === 'number') return cell.fmt || 'number';
    return 'text';
  };
  const mk = (r, base) => {
    const cur = sheet.get(r, targetCol);
    const prev = prevCol ? sheet.get(r, prevCol) : null;
    return {
      ...base,
      row: r,
      ref: colName(targetCol) + r,
      locked: !!(cur && !cur.empty),
      existing: cur && !cur.empty ? displayValue(cur) : '',
      prev: prev && !prev.empty ? displayValue(prev) : '',
      style: (prev && prev.s) || (cur && cur.s) || 0,
    };
  };

  const meta = layout.metaRows.map((m) => mk(m.row, {
    kind: 'meta', label: m.label, type: m.type === 'text' && /名前|氏名|点検者|担当/.test(m.label) ? 'name' : m.type, chips: [],
  }));
  const items = layout.itemRows.map((r) => {
    const rangeText = role(r, 'range');
    const range = parseRange(rangeText);
    const prevCell = prevCol ? sheet.get(r, prevCol) : null;
    const extra = layout.columns
      .filter((x) => x.role === 'other')
      .map((x) => [x.header, displayValue(sheet.valueAt(r, x.c))])
      .filter(([, v]) => v)
      .map(([h, v]) => `${h} ${v}`);
    const options = sheet.validationOptions(r, targetCol);
    const contact = role(r, 'contact');
    return mk(r, {
      kind: 'item',
      section: role(r, 'section'),
      label: role(r, 'name') || `${r}行目`,
      sub: role(r, 'code'),
      rangeText: range && rangeText !== '-' ? rangeText : '',
      range,
      rangeKind: range ? 'range' : null,
      effect: contact ? `${contact}へ連絡` : '',
      chips: [contact, ...extra].filter(Boolean),
      options,
      type: options ? 'choice' : typeFromCell(prevCell, range ? 'number' : 'text'),
      cond: {},
    });
  });
  return { meta, items, conditions: [], warnings: [], prevCol, targetCol, ngMessage: '管理値外' };
}

// ---- 書き戻し ----
const escXml = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/**
 * 入力値を Excel セル表現に変換する。
 * 返り値: { kind:'n', v:number } | { kind:'s', v:string } | null(空)
 */
export function toCellValue(type, raw) {
  if (raw === undefined || raw === null || String(raw).trim() === '') return null;
  const s = String(raw).trim();
  if (type === 'date') {
    const m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s);
    if (m) return { kind: 'n', v: dateToSerial(+m[1], +m[2], +m[3]) };
  } else if (type === 'time') {
    const m = /^(\d{1,2}):(\d{2})/.exec(s);
    if (m) return { kind: 'n', v: timeToSerial(+m[1], +m[2]) };
  } else if (type === 'number') {
    const n = s.normalize('NFKC').replace(/,/g, '');
    if (/^[-+]?(\d+\.?\d*|\.\d+)$/.test(n)) return { kind: 'n', v: +n };
  }
  return { kind: 's', v: s };
}

function cellXml(p, ref, style, val) {
  const s = style ? ` s="${style}"` : '';
  if (val.kind === 'n') return `<${p}c r="${ref}"${s}><${p}v>${val.v}</${p}v></${p}c>`;
  const sp = /^\s|\s$/.test(val.v) ? ' xml:space="preserve"' : '';
  return `<${p}c r="${ref}"${s} t="inlineStr"><${p}is><${p}t${sp}>${escXml(val.v)}</${p}t></${p}is></${p}c>`;
}

/** 開始タグ位置 start から要素の終端 (排他) を返す。要素は同名でネストしない前提 */
function elementEnd(xml, start, closeTag) {
  const gt = xml.indexOf('>', start);
  if (xml[gt - 1] === '/') return gt + 1;
  const close = xml.indexOf(closeTag, gt);
  return close + closeTag.length;
}

/**
 * シート XML に値を書き込む。空セルにのみ書き込み、値・数式のあるセルは触らない。
 * updates: [{ row, col, style, val }]
 */
export function patchSheetXml(xml, updates) {
  const pm = /<(\w+:)?sheetData\b/.exec(xml);
  if (!pm) throw new Error('sheetData が見つかりません');
  const p = pm[1] || '';
  const sdOpen = pm.index;
  const sdOpenEnd = xml.indexOf('>', sdOpen) + 1;
  let head = xml.slice(0, sdOpenEnd);
  let body, tail;
  if (xml[sdOpenEnd - 2] === '/') {
    // <sheetData/> (空シート)
    head = head.slice(0, -2) + '>';
    body = '';
    tail = `</${p}sheetData>` + xml.slice(sdOpenEnd);
  } else {
    const sdClose = xml.indexOf(`</${p}sheetData>`, sdOpenEnd);
    body = xml.slice(sdOpenEnd, sdClose);
    tail = xml.slice(sdClose);
  }

  // 行に分解
  const rowRe = new RegExp(`<${p}row[\\s>/]`, 'g');
  const rows = [];
  let m;
  while ((m = rowRe.exec(body))) {
    const end = elementEnd(body, m.index, `</${p}row>`);
    const chunk = body.slice(m.index, end);
    const r = +/\sr="(\d+)"/.exec(chunk.slice(0, chunk.indexOf('>')))[1];
    rows.push({ r, chunk });
    rowRe.lastIndex = end;
  }

  const byRow = new Map();
  for (const u of updates) {
    if (!byRow.has(u.row)) byRow.set(u.row, []);
    byRow.get(u.row).push(u);
  }

  const written = [];
  for (const [r, ups] of byRow) {
    let row = rows.find((x) => x.r === r);
    if (!row) {
      row = { r, chunk: `<${p}row r="${r}"></${p}row>` };
      const idx = rows.findIndex((x) => x.r > r);
      rows.splice(idx < 0 ? rows.length : idx, 0, row);
    }
    let chunk = row.chunk;
    if (chunk.endsWith('/>')) chunk = chunk.slice(0, -2) + `></${p}row>`;
    const openEnd = chunk.indexOf('>') + 1;
    let open = chunk.slice(0, openEnd);
    const inner = chunk.slice(openEnd, chunk.length - `</${p}row>`.length);

    const cellRe = new RegExp(`<${p}c[\\s>/]`, 'g');
    const cells = [];
    let cm;
    while ((cm = cellRe.exec(inner))) {
      const end = elementEnd(inner, cm.index, `</${p}c>`);
      const cx = inner.slice(cm.index, end);
      const tag = cx.slice(0, cx.indexOf('>'));
      const ref = /\sr="([A-Z]+)(\d+)"/.exec(tag);
      cells.push({ c: ref ? colNum(ref[1]) : -1, start: cm.index, end, xml: cx });
      cellRe.lastIndex = end;
    }

    let out = inner;
    // 後ろから差し込む/置き換えるため、列の降順で処理
    for (const u of [...ups].sort((a, b) => b.col - a.col)) {
      const ref = colName(u.col) + r;
      const ex = cells.find((x) => x.c === u.col);
      if (ex) {
        if (new RegExp(`<${p}(v|f|is)[\\s>]`).test(ex.xml)) continue; // 空でないセルは上書きしない
        const st = /\ss="(\d+)"/.exec(ex.xml.slice(0, ex.xml.indexOf('>')));
        const style = st ? +st[1] : u.style;
        out = out.slice(0, ex.start) + cellXml(p, ref, style, u.val) + out.slice(ex.end);
      } else {
        const next = cells.find((x) => x.c > u.col);
        const pos = next ? next.start : inner.length;
        out = out.slice(0, pos) + cellXml(p, ref, u.style, u.val) + out.slice(pos);
      }
      written.push(ref);
    }
    // spans="1:7" を対象列まで広げる
    const maxCol = Math.max(...ups.map((u) => u.col));
    open = open.replace(/\sspans="(\d+):(\d+)"/, (s, a, b) => ` spans="${a}:${Math.max(+b, maxCol)}"`);
    row.chunk = open + out + `</${p}row>`;
  }

  let result = head + rows.map((x) => x.chunk).join('') + tail;
  // dimension を広げる
  const maxR = Math.max(...updates.map((u) => u.row));
  const maxC = Math.max(...updates.map((u) => u.col));
  result = result.replace(new RegExp(`(<${p}dimension\\s+ref=")([^"]+)(")`), (s, a, ref, z) => {
    const [r1, r2] = ref.split(':');
    const e = parseRef(r2 || r1);
    if (!e) return s;
    return `${a}${r1.split(':')[0]}:${colName(Math.max(e.c, maxC))}${Math.max(e.r, maxR)}${z}`;
  });
  return { xml: result, written };
}

/** 入力値を書き込んだ xlsx の Blob を作る (cardList: 書き込む対象のカード) */
export async function exportWorkbook(wb, sheet, targetCol, cardList, values) {
  const updates = [];
  for (const card of cardList) {
    if (card.locked) continue;
    const val = toCellValue(card.type, values[card.row]);
    if (!val) continue;
    updates.push({ row: card.row, col: targetCol, style: card.style, val });
  }
  if (!updates.length) throw new Error('書き込む値がありません');
  const { xml, written } = patchSheetXml(sheet.xml, updates);
  const blob = await writeZip(wb.zip, new Map([[sheet.path, new TextEncoder().encode(xml)]]));
  return { blob, written };
}
