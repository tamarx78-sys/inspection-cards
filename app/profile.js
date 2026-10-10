// 点検表ごとの「設定ファイル」(JSON) に従ってカードを組み立てる。
// 点検表に固有の名前・行・見出しは設定ファイル側だけに書き、このコードには持たない。
import { colNum, colName, parseRef, displayValue } from './xlsx.js';

export const PROFILE_TYPE = 'inspection-cards-profile';

const norm = (s) => String(s ?? '').normalize('NFKC').replace(/\s+/g, '').trim();
const text = (cell) => (cell && !cell.empty ? displayValue(cell).trim() : '');
const num = (s) => {
  const n = String(s ?? '').normalize('NFKC').replace(/,/g, '').trim();
  return /^[-+]?(\d+\.?\d*|\.\d+)$/.test(n) ? +n : null;
};

/** 設定ファイルの読み込みと最低限の検証 */
export function parseProfile(json) {
  const p = typeof json === 'string' ? JSON.parse(json) : json;
  if (p?.type !== PROFILE_TYPE) throw new Error('点検カードの設定ファイルではありません');
  for (const k of ['name', 'sheet', 'columns', 'items', 'records', 'meta']) {
    if (!p[k]) throw new Error(`設定ファイルに「${k}」がありません`);
  }
  return p;
}

/** このブックに使える設定か */
export function profileMatches(profile, wb, fileName) {
  if (!wb.sheets.some((s) => s.name === profile.sheet)) return false;
  return !profile.fileMatch || fileName.includes(profile.fileMatch);
}

/**
 * 設定に従ってカード一式を作る。
 * 返り値: { meta, items, conditions, warnings, recordCols, prevCol, targetCol }
 */
export function buildProfileCards(sheet, profile, targetCol) {
  const warnings = [];
  const C = Object.fromEntries(Object.entries(profile.columns).map(([k, v]) => [k, colNum(v)]));
  const at = (r, key) => (C[key] ? sheet.get(r, C[key]) : null);

  // 見出しの確認 (表の形が変わっていないか)
  for (const [ref, expect] of Object.entries(profile.headerCheck || {})) {
    const p = parseRef(ref);
    const got = p ? text(sheet.valueAt(p.r, p.c)) : '';
    if (norm(got) !== norm(expect)) warnings.push(`見出し ${ref} が「${expect}」ではありません (実際: 「${got}」)`);
  }

  // 基本情報 (日付・時間・担当など)
  const labelCol = colNum(profile.meta.labelCol);
  const metaRows = [];
  for (const m of profile.meta.rows) {
    let found = 0;
    for (let r = 1; r < profile.items.startRow; r++) {
      if (norm(text(sheet.get(r, labelCol))) === norm(m.label)) { found = r; break; }
    }
    if (found) metaRows.push({ ...m, row: found });
    else warnings.push(`基本情報「${m.label}」の行が見つかりません`);
  }

  // 項目行: startRow から stopWhen に当たる行の手前まで
  const keyCol = colNum(profile.items.keyCol);
  const stop = profile.items.stopWhen;
  const rows = [];
  for (let r = profile.items.startRow; r <= sheet.maxR; r++) {
    if (stop && norm(text(sheet.get(r, colNum(stop.col)))) === norm(stop.equals)) break;
    const hasAny = ['name', 'sub', 'unit', 'set', 'min', 'max'].some((k) => text(at(r, k)));
    if (hasAny) rows.push(r);
  }

  // 記録列: startCol 以降で何か入っている列
  const recStart = colNum(profile.records.startCol);
  const recRows = [...metaRows.map((m) => m.row), ...rows];
  const recordCols = [];
  for (let c = recStart; c <= Math.max(sheet.maxC, recStart); c++) {
    if (recRows.some((r) => sheet.get(r, c) && !sheet.get(r, c).empty)) recordCols.push(c);
  }
  const nextCol = recordCols.length ? recordCols[recordCols.length - 1] + 1 : recStart;
  targetCol ??= nextCol;
  const prevCols = recordCols.filter((c) => c < targetCol);
  const prevCol = prevCols.length ? prevCols[prevCols.length - 1] : null;

  const base = (r) => {
    const cur = sheet.get(r, targetCol);
    const prev = prevCol ? sheet.get(r, prevCol) : null;
    return {
      row: r,
      ref: colName(targetCol) + r,
      locked: !!(cur && !cur.empty),
      existing: cur && !cur.empty ? displayValue(cur) : '',
      prev: prev && !prev.empty ? displayValue(prev) : '',
      style: (prev && prev.s) || (cur && cur.s) || 0,
    };
  };

  const meta = metaRows.map((m) => ({
    ...base(m.row), kind: 'meta', label: m.label.replace(/[:：]\s*$/, ''), type: m.type, chips: [],
  }));

  // 行を特定するキー (設定の items.keyCol の値) → 行
  const byKey = new Map();
  for (const r of rows) {
    const k = text(sheet.get(r, keyCol));
    if (k) byKey.set(norm(k), r);
  }
  let lastSection = '';
  let lastName = '';
  const kinds = profile.kinds || {};
  const items = rows.map((r) => {
    const section = text(at(r, 'section')) || lastSection;
    lastSection = section;
    // 名前が空の行は直前の項目の続き。全角スペースの連続は 1 つにまとめる
    const name = text(at(r, 'name')).replace(/[\s　]+/g, ' ') || lastName;
    lastName = name;
    const sub = text(at(r, 'sub'));
    const setText = text(at(r, 'set'));
    const kindText = text(at(r, 'kind'));
    const min = num(text(at(r, 'min')));
    const max = num(text(at(r, 'max')));
    let rangeKind = null;
    if (kinds.range?.some((k) => norm(k) === norm(kindText))) rangeKind = 'range';
    else if (kinds.reference?.some((k) => norm(k) === norm(kindText))) rangeKind = 'reference';
    const hasRange = rangeKind && (min !== null || max !== null);
    const digitsText = text(at(r, 'digits'));
    const b = base(r);

    // 選択肢: 入力規則 > 設定値が属する選択肢セット
    let options = sheet.validationOptions?.(r, targetCol) || null;
    if (!options) {
      const probe = [setText, b.prev].filter(Boolean).map(norm);
      options = (profile.choiceSets || []).find((set) => set.some((o) => probe.includes(norm(o)))) || null;
    }
    // 単位がある行・数値の設定値や範囲がある行は数値入力 (iPad で数字キーボードを出す)
    const unit = text(at(r, 'unit'));
    // 手がかりがない行は設定の defaultType に従う (前回値が文字の行は文字入力のまま)
    const numeric = unit || num(setText) !== null || min !== null || max !== null || num(b.prev) !== null;
    const fallback = profile.defaultType === 'number' && !(b.prev && num(b.prev) === null) ? 'number' : 'text';
    const type = options ? 'choice' : numeric ? 'number' : fallback;

    return {
      ...b,
      kind: 'item',
      key: text(sheet.get(r, keyCol)),
      section,
      label: name || `${r}行目`,
      sub,
      unit,
      watch: !!text(at(r, 'watch')),
      chips: [text(at(r, 'where'))].filter(Boolean),
      setText,
      rangeKind: hasRange ? rangeKind : null,
      range: hasRange ? { min: min ?? undefined, max: max ?? undefined } : null,
      rangeText: hasRange ? `${min ?? ''}～${max ?? ''}` : '',
      kindText,
      effect: text(at(r, 'effect')),
      digits: num(digitsText),
      options,
      type,
      cond: {}, // 条件による振り分け (下で設定)
    };
  });

  // 設定の「行の指定」を解決する ({no, name} → 行)。名前が違えば警告して使わない
  const itemByRow = new Map(items.map((c) => [c.row, c]));
  const resolve = (ref, what) => {
    const r = byKey.get(norm(ref.no));
    if (!r) { warnings.push(`${what}: キー ${ref.no} の行が見つかりません`); return null; }
    const card = itemByRow.get(r);
    if (ref.name && card && !norm(card.label + card.sub).includes(norm(ref.name)) && norm(card.label) !== norm(ref.name)) {
      warnings.push(`${what}: キー ${ref.no} の項目名が「${ref.name}」ではありません (実際: 「${card.label}」)`);
      return null;
    }
    return card;
  };

  // 一覧表示 (項目名と入力欄を 1 行にして縦に並べる) にする行
  for (const ref of profile.compactRows || []) {
    const c = resolve(ref, '一覧表示の行');
    if (c) c.compact = true;
  }

  // 任意入力の行 (空欄でも未入力として数えない)
  for (const ref of profile.optionalRows || []) {
    const c = resolve(ref, '任意入力の行');
    if (c) c.optional = true;
  }

  // 対象外の行
  const excluded = new Set();
  for (const ref of profile.exclude || []) {
    const c = resolve(ref, '対象外の行');
    if (c) excluded.add(c.row);
  }

  // 条件
  const conditions = [];
  for (const cd of profile.conditions || []) {
    const cond = { key: cd.key, label: cd.label, options: cd.options || null, kind: cd.match?.type || 'select' };
    if (cd.source) {
      const src = resolve(cd.source, `条件「${cd.label}」`);
      if (!src) continue;
      cond.sourceRow = src.row;
      if (cd.options && !src.options) { src.options = cd.options; src.type = 'choice'; }
      src.conditionSource = cd.key;
    }
    if (cd.match?.type === 'token') {
      const opts = cd.options.map(norm);
      for (const c of items) {
        // 英数字のかたまりとそれ以外を分けて単語にする (例:「AB幅」→ AB / 幅、「ABC部」→ ABC / 部)
        const tokens = cd.match.columns.flatMap((k) => String(k === 'name' ? c.label : c[k] || '')
          .normalize('NFKC').match(/[A-Za-z0-9.]+|[^\sA-Za-z0-9.()（）/、,]+/g) || []).map(norm).filter(Boolean);
        const hit = cd.options.find((o, i) => tokens.includes(opts[i]));
        if (hit && c.row !== cond.sourceRow) c.cond[cd.key] = { option: hit };
      }
    } else if (cd.match?.type === 'band') {
      const re = new RegExp(cd.match.pattern);
      for (const c of items) {
        for (const k of cd.match.columns) {
          const m = re.exec(String(k === 'name' ? c.label : c[k] || '').normalize('NFKC'));
          if (m) { c.cond[cd.key] = { min: +m[1], max: m[2] !== undefined ? +m[2] : undefined, text: m[0] }; break; }
        }
      }
      cond.bandMins = [...new Set(items.map((c) => c.cond[cd.key]?.min).filter((v) => v !== undefined))].sort((a, b) => a - b);
    } else if (cd.rows) {
      // 選択式の条件: 選んだ選択肢の行だけを使う
      for (const [opt, refs] of Object.entries(cd.rows)) {
        for (const ref of refs) {
          const c = resolve(ref, `条件「${cd.label}」の${opt}`);
          if (c) c.cond[cd.key] = { option: opt };
        }
      }
      if (cd.placeBefore) cond.placeBefore = resolve(cd.placeBefore, `条件「${cd.label}」の表示位置`)?.row ?? null;
      cond.writes = {};
      for (const [opt, ref] of Object.entries(cd.writes || {})) {
        const c = resolve(ref, `条件「${cd.label}」の記入先`);
        if (c) cond.writes[opt] = c.row;
      }
      // 既定値: 前回の記録で値が入っていた側
      const score = Object.fromEntries(cd.options.map((o) => [o, 0]));
      for (const c of items) if (c.cond[cd.key] && c.prev) score[c.cond[cd.key].option]++;
      const best = cd.options.reduce((a, o) => (score[o] > score[a] ? o : a), cd.options[0]);
      cond.defaultOption = score[best] > 0 ? best : null;
    }
    conditions.push(cond);
  }

  return {
    meta,
    items: items.filter((c) => !excluded.has(c.row)),
    conditions,
    warnings,
    recordCols,
    prevCol,
    targetCol,
    nextCol,
    ngMessage: profile.ngMessage || '管理値外',
  };
}

/** 条件の現在値 (選択式はセッションに保存した値、入力連動は元の項目の入力値) */
export function conditionValues(cards, session) {
  const out = {};
  for (const cd of cards.conditions) {
    if (cd.sourceRow) out[cd.key] = String(session.values[cd.sourceRow] ?? '').trim();
    else out[cd.key] = session.cond?.[cd.key] ?? cd.defaultOption ?? '';
  }
  return out;
}

/** この項目は現在の条件で入力対象か */
export function isApplicable(card, cards, values) {
  for (const cd of cards.conditions) {
    const c = card.cond?.[cd.key];
    if (!c) continue;
    const v = values[cd.key];
    if (cd.kind === 'band') {
      const cs = num(v);
      if (cs === null) continue; // 未入力のうちは全部出す
      const band = cd.bandMins.filter((m) => m <= cs).pop();
      if (band === undefined || c.min !== band) return false;
    } else if (v && norm(v) !== norm(c.option)) {
      return false;
    } else if (!v && cd.kind === 'select') {
      return false; // 選択式は選ぶまで出さない
    }
  }
  return true;
}
