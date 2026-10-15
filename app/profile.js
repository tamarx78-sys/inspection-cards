// 点検表ごとの「設定ファイル」(JSON) に従ってカードを組み立てる。
// 点検表に固有の名前・行・見出しは設定ファイル側だけに書き、このコードには持たない。
import { colNum, colName, parseRef, displayValue, shiftFormula, baseCard } from './xlsx.js';

export const PROFILE_TYPE = 'inspection-cards-profile';

/** 文字列の比較用 (全角・半角をそろえ、空白を除く) */
export const norm = (s) => String(s ?? '').normalize('NFKC').replace(/\s+/g, '').trim();
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

/**
 * このブックに使える設定か。fileMatch があればファイル名で判断する (シート名は改訂で変わることがあるため見ない)。
 * fileMatch がなければ、設定のシート名のシートがあるか
 */
export function profileMatches(profile, wb, fileName) {
  if (profile.fileMatch) return String(fileName).normalize('NFKC').includes(String(profile.fileMatch).normalize('NFKC'));
  return wb.sheets.some((s) => s.name === profile.sheet);
}

/** 列の見出しの確認 (表の形が変わっていないか)。合わない箇所の説明の配列を返す */
export function checkHeaders(sheet, profile) {
  const out = [];
  for (const [ref, expect] of Object.entries(profile.headerCheck || {})) {
    const p = parseRef(ref);
    const got = p ? text(sheet.valueAt(p.r, p.c)) : '';
    if (norm(got) !== norm(expect)) out.push(`見出し ${ref} が「${expect}」ではありません (実際: 「${got}」)`);
  }
  return out;
}

/**
 * この設定で使えるシートの候補。列の見出し (headerCheck) がすべて合うシート。
 * headerCheck がない設定は、設定のシート名のシートだけ。
 * 返り値: { candidates: [シート名], named: { name, problems } | null (設定のシート名のシートの確認結果) }
 */
export async function profileSheets(profile, wb, loadSheet) {
  const hasCheck = Object.keys(profile.headerCheck || {}).length > 0;
  const candidates = [];
  let named = null;
  for (const meta of wb.sheets) {
    if (!hasCheck && meta.name !== profile.sheet) continue;
    const sheet = await loadSheet(wb, meta);
    const problems = checkHeaders(sheet, profile);
    if (meta.name === profile.sheet) named = { name: meta.name, problems };
    if (!problems.length) candidates.push(meta.name);
  }
  return { candidates, named };
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
  warnings.push(...checkHeaders(sheet, profile));

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

  const base = (r) => baseCard(sheet, r, targetCol, prevCol);
  const dateRow = metaRows.find((m) => m.type === 'date')?.row; // 過去のメモに日付を添えるため

  const meta = metaRows.map((m) => ({
    ...base(m.row), kind: 'meta', label: m.label.replace(/[:：]\s*$/, ''), type: m.type, chips: [],
  }));

  // 行を特定するキー (設定の items.keyCol の値) → 行。同じキーの行や、キーが空の行もある (項目名で区別する)
  const byKey = new Map();
  for (const r of rows) {
    const k = norm(text(sheet.get(r, keyCol)));
    byKey.set(k, [...(byKey.get(k) || []), r]);
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
    let rangeKind = null;
    if (kinds.range?.some((k) => norm(k) === norm(kindText))) rangeKind = 'range';
    else if (kinds.reference?.some((k) => norm(k) === norm(kindText))) rangeKind = 'reference';
    let min = num(text(at(r, 'min')));
    let max = num(text(at(r, 'max')));
    // 「設定±2」のような書き方: 設定値 ± 幅を管理幅にする (kinds.tolerance に書いた文字で始まるもの)
    const tol = (kinds.tolerance || []).map(norm).find((k) => norm(kindText).startsWith(k));
    if (tol && num(setText) !== null) {
      const w = num(norm(kindText).slice(tol.length));
      if (w !== null) {
        const round = (x) => +x.toFixed(6);
        min = round(num(setText) - w); max = round(num(setText) + w); rangeKind = 'range';
      }
    }
    const hasRange = rangeKind && (min !== null || max !== null);
    const digitsText = text(at(r, 'digits'));
    const b = base(r);
    // コメント (メモ): 項目の列 (記録列より左) に付いたものと、前回の記録に付いたもの
    const comments = [];
    for (let c = 1; c < recStart; c++) { const t = sheet.comment?.(r, c); if (t) comments.push(t); }
    // 過去の記録に付いたコメントのうち、いちばん新しいもの (日付の行があれば日付を添える)
    let lastMemo = null;
    for (const col of [...prevCols].reverse()) {
      const t = sheet.comment?.(r, col);
      if (t) {
        lastMemo = { text: t, prev: col === prevCol, date: dateRow ? text(sheet.get(dateRow, col)) : '' };
        break;
      }
    }

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
      note: text(at(r, 'note')), // 備考の列
      comments,
      lastMemo,
      digits: num(digitsText),
      options,
      type,
      cond: {}, // 条件による振り分け (下で設定)
    };
  });

  // 設定の「行の指定」を解決する ({no, name} → 行)。名前が違えば警告して使わない
  const itemByRow = new Map(items.map((c) => [c.row, c]));
  // 設定に控えてある行番号 (_row) と実際の行が違う = 行が増減した。キーと項目名が合えば使うが、
  // 行番号で書いた式 (formula) などがずれるおそれがあるので知らせる
  const moved = new Map();
  const nameOk = (card, name) => !name || norm(card.label + card.sub).includes(norm(name)) || norm(card.label) === norm(name);
  const resolve = (ref, what) => {
    const rs = byKey.get(norm(ref.no ?? ''));
    if (!rs) { warnings.push(`${what}: キー ${ref.no} の行が見つかりません`); return null; }
    // 同じキーの行が複数あるときは、項目名が合う行 → 控えの行番号 (_row) が合う行の順に選ぶ
    const named = rs.filter((r) => itemByRow.get(r) && nameOk(itemByRow.get(r), ref.name));
    const pick = named.find((r) => r === ref._row) ?? named[0] ?? rs[0];
    const card = itemByRow.get(pick);
    if (ref.name && card && !nameOk(card, ref.name)) {
      warnings.push(`${what}: キー ${ref.no} の項目名が「${ref.name}」ではありません (実際: 「${card.label}」)`);
      return null;
    }
    if (card && Number.isInteger(ref._row) && ref._row !== card.row) moved.set(card.row, `キー ${ref.no} ${ref.name || ''}: ${ref._row} → ${card.row} 行目`);
    return card;
  };

  // 入力方法の指定 (自動判定より優先)。options があれば選択式。"A|B" の文字列でも配列でもよい
  for (const ref of profile.rowTypes || []) {
    const c = resolve(ref, '入力方法の指定');
    if (!c) continue;
    const options = Array.isArray(ref.options) ? ref.options
      : typeof ref.options === 'string' ? ref.options.split('|').map((o) => o.trim()).filter(Boolean) : null;
    if (options?.length) { c.options = options; c.type = 'choice'; }
    else if (['number', 'text', 'choice'].includes(ref.type)) c.type = ref.type;
    else warnings.push(`入力方法の指定: キー ${ref.no} の type / options が正しくありません`);
    // 選べない選択肢 (表示はするが押せない)。"A|B" の文字列でも配列でもよい
    const disabled = Array.isArray(ref.disabled) ? ref.disabled
      : typeof ref.disabled === 'string' ? ref.disabled.split('|').map((o) => o.trim()).filter(Boolean) : [];
    if (disabled.length) c.disabledOptions = disabled;
    if (c.type === 'choice' && !c.options?.length) {
      warnings.push(`入力方法の指定: キー ${ref.no} は選択式ですが選択肢 (options) がありません`);
      c.type = 'text';
    }
  }

  // 見出し (区分) の指定: 区分の列の値の代わりに、指定した見出しにまとめる
  for (const g of profile.sectionRows || []) {
    for (const ref of g.rows || []) {
      const c = resolve(ref, `見出し「${g.section}」の行`);
      if (c) c.section = g.section;
    }
  }

  // 自動計算の行: 式を今回の列に書き、値は入力から計算して表示する。
  // 式は設定の formula ({col} が今回の列) か、前回までの記録で最後に式が入っている列からずらして作る
  for (const ref of profile.formulaRows || []) {
    const c = resolve(ref, '自動計算の行');
    if (!c || c.locked) continue;
    let f = ref.formula ? ref.formula.replace(/^=/, '').replaceAll('{col}', colName(targetCol)) : '';
    if (!f) {
      const src = [...prevCols].reverse().find((col) => sheet.get(c.row, col)?.formula);
      if (src) f = shiftFormula(sheet.get(c.row, src).formula, 0, targetCol - src);
    }
    if (!f) { warnings.push(`自動計算の行: キー ${ref.no} の式が前回までの記録に見つかりません (設定の formula で指定できます)`); continue; }
    c.type = 'formula';
    c.formula = f;
  }

  // 一覧表示 (項目名と入力欄を 1 行にして縦に並べる) にする行
  for (const ref of profile.compactRows || []) {
    const c = resolve(ref, '一覧表示の行');
    if (c) c.compact = true;
  }

  // 列分け表示: 指定した列の値 (例: 左右・系統) ごとに分け、それぞれを一覧表示で縦に並べる。
  // 入力の順番は「1 つ目の値の行を上から下へ → 2 つ目の値の行…」になるよう並べ替える
  const splitOrder = new Map(); // 行 → 並べ替え後の順番キー
  (profile.splitGroups || []).forEach((g, gi) => {
    const members = (g.rows || []).map((ref) => resolve(ref, '列分け表示の行')).filter(Boolean);
    // 分ける値: by の欄の値。pattern があれば、その欄 (既定は項目名) から正規表現で取り出す (最初の ( ) の中)
    const by = g.by || (g.pattern ? 'label' : 'sub');
    const re = g.pattern ? new RegExp(g.pattern) : null;
    const keyOf = (c) => {
      const v = String(c[by] ?? '').normalize('NFKC');
      if (!re) return v;
      const m = re.exec(v);
      return m ? (m[1] ?? m[0]).trim() : '';
    };
    const keys = (g.order || [...new Set(members.map(keyOf))]).map((k) => String(k).normalize('NFKC'));
    for (const c of members) {
      const key = keyOf(c);
      c.compact = true;
      c.split = { id: gi, key };
      if (by === 'sub') c.hideSub = true; // 見出しに出すので行の表示からは省く
      // 項目名から取り出したときは、その部分を項目名の表示から省く
      if (re && by === 'label') c.label = c.label.normalize('NFKC').replace(re, '').trim() || c.label;
      const ki = keys.indexOf(key);
      splitOrder.set(c.row, [gi, ki < 0 ? keys.length : ki]);
    }
  });

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

  if (moved.size) {
    const ex = [...moved.values()];
    warnings.push(`行の位置が設定と変わっています (${moved.size} か所。例: ${ex.slice(0, 3).join(' / ')}${ex.length > 3 ? ' ほか' : ''})`);
  }

  // 列分け表示のかたまりは、最初の行の位置にまとめて置き、値ごとの順に並べ替える
  const ordered = [];
  const placed = new Set();
  for (const c of items) {
    const so = splitOrder.get(c.row);
    if (!so) { ordered.push(c); continue; }
    if (placed.has(so[0])) continue;
    placed.add(so[0]);
    const group = items.filter((x) => splitOrder.get(x.row)?.[0] === so[0]);
    group.sort((a, b) => splitOrder.get(a.row)[1] - splitOrder.get(b.row)[1] || a.row - b.row);
    ordered.push(...group);
  }

  return {
    meta,
    items: ordered.filter((c) => !excluded.has(c.row)),
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
