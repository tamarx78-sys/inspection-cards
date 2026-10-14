import {
  openWorkbook, loadSheet, detectLayout, buildCards, exportWorkbook, toCellValue, colName, displayValue,
} from './xlsx.js';
import {
  parseProfile, profileMatches, profileSheets, buildProfileCards, conditionValues, isApplicable,
} from './profile.js';
import {
  putSession, getSession, deleteSession, listSessions, requestPersist, putProfile, deleteProfile, listProfiles,
} from './store.js';
import { evaluate } from './formula.js';

const $ = (sel, root = document) => root.querySelector(sel);
const main = $('#main');
const el = (tag, attrs = {}, ...children) => {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (v === undefined || v === null || v === false) continue;
    if (k === 'class') e.className = v;
    else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
    else if (k === 'dataset') Object.assign(e.dataset, v);
    else e.setAttribute(k, v === true ? '' : v);
  }
  return add(e, children);
};
/** 配列をネストごと展開し、null/false を飛ばして追加する */
const add = (parent, children) => {
  for (const c of [children].flat(Infinity)) if (c !== null && c !== undefined && c !== false) parent.append(c);
  return parent;
};
const mount = (...children) => { keypad.hide(); main.replaceChildren(); add(main, children); };
const pad = (n) => String(n).padStart(2, '0');
const today = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
const nowTime = () => { const d = new Date(); return `${pad(d.getHours())}:${pad(d.getMinutes())}`; };
const norm = (s) => String(s ?? '').normalize('NFKC').replace(/\s+/g, '').trim();

// 現在の作業状態
let S = null; // { session, wb, sheet, cards, filter }

function toast(msg, kind = '') {
  const t = el('div', { class: `toast ${kind}` }, msg);
  document.body.append(t);
  setTimeout(() => t.classList.add('show'), 10);
  setTimeout(() => { t.classList.remove('show'); setTimeout(() => t.remove(), 300); }, 2600);
}

function setHeader(title, actions = []) {
  $('#title').textContent = title;
  $('#actions').replaceChildren(...actions);
}

/** セッション (設定ファイルの有無) に応じてカード一式を作る */
async function makeCards(wb, sheetName, profile, targetCol) {
  const meta = wb.sheets.find((s) => s.name === sheetName);
  if (!meta) throw new Error(`シート「${sheetName}」が見つかりません`);
  const sheet = await loadSheet(wb, meta);
  if (profile) return { sheet, cards: buildProfileCards(sheet, profile, targetCol) };
  const layout = detectLayout(sheet);
  return { sheet, cards: buildCards(sheet, layout, targetCol ?? layout.targetCol), layout };
}

// 点検表が設定と合わない (行の増減など) ときの注意の文言 (利用者の指定)
const REVISED_MSG = '点検表が改定されたため正しく入力できない可能性があります。Excelによる入力をお勧めします。';

// ---------------- ホーム ----------------
async function showHome() {
  S = null;
  setHeader('点検カード');
  const [sessions, profiles] = await Promise.all([listSessions(), listProfiles()]);
  const fileInput = el('input', {
    type: 'file',
    accept: '.xlsx,.xlsm,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel.sheet.macroEnabled.12',
    class: 'hidden', onchange: (e) => e.target.files[0] && openFile(e.target.files[0]),
  });
  mount(
    el('section', { class: 'home' },
      el('button', { class: 'btn primary big', onclick: () => fileInput.click() }, '📂 点検表 (Excel) を開く'),
      fileInput,
      el('p', { class: 'hint' }, '開いた点検表と入力内容はこの端末内に保存されます。Wi-Fi が無くても入力・書き出しできます。'),
      sessions.length ? el('h2', {}, '作業中の点検表') : null,
      el('div', { class: 'session-list' },
        sessions.map((s) => {
          const filled = Object.values(s.values || {}).filter((v) => String(v ?? '').trim()).length;
          return el('div', { class: 'session' },
            el('div', { class: 'session-main', onclick: () => resume(s.id) },
              el('div', { class: 'session-name' }, s.fileName),
              el('div', { class: 'session-sub' },
                `${s.sheetName} / ${colName(s.targetCol)}列 · ${filled}件入力 · ${new Date(s.updatedAt).toLocaleString('ja-JP')}`,
                s.exportedAt ? el('span', { class: 'badge ok' }, '書き出し済') : null),
            ),
            el('button', {
              class: 'btn ghost danger',
              onclick: async () => {
                if (!confirm(`「${s.fileName}」の作業データを削除しますか?\n(元の Excel ファイルには影響しません)`)) return;
                await deleteSession(s.id);
                showHome();
              },
            }, '削除'),
          );
        }),
      ),
      el('h2', {}, '設定ファイル'),
      el('p', { class: 'hint' }, '点検表と同じフォルダにある設定ファイル (.json) を読み込むと、その点検表専用のカードになります。'),
      el('div', { class: 'session-list' },
        profiles.map((p) => el('div', { class: 'session' },
          el('div', { class: 'session-main' },
            el('div', { class: 'session-name' }, p.name),
            el('div', { class: 'session-sub' }, `読み込み ${new Date(p.importedAt).toLocaleString('ja-JP')}`)),
          el('button', {
            class: 'btn ghost danger',
            onclick: async () => {
              if (!confirm(`設定「${p.name}」をこの端末から削除しますか?`)) return;
              await deleteProfile(p.id);
              showHome();
            },
          }, '削除')))),
      profileImportButton(() => showHome()),
      versionLine,
    ),
  );
  checkUpdate();
}

/** 設定ファイルを取り込む。then: 取り込み後の画面 (省略時はホーム) */
async function importProfile(file, then = showHome) {
  try {
    const p = parseProfile(await file.text());
    await putProfile(p);
    requestPersist();
    toast(`設定「${p.name}」を読み込みました`, 'ok');
    await then(p);
  } catch (e) {
    console.error(e);
    toast(`設定ファイルを読み込めませんでした: ${e.message}`, 'error');
  }
}

/** 「設定ファイルを読み込む」ボタン (ファイル選択つき) */
function profileImportButton(then, label = '⚙ 設定ファイルを読み込む') {
  const input = el('input', {
    type: 'file', accept: '.json,application/json', class: 'hidden',
    onchange: (e) => e.target.files[0] && importProfile(e.target.files[0], then),
  });
  return [el('button', { class: 'btn', onclick: () => input.click() }, label), input];
}

async function openFile(file) {
  try {
    const bytes = await file.arrayBuffer();
    const wb = await openWorkbook(bytes);
    const profiles = (await listProfiles()).filter((p) => profileMatches(p, wb, file.name));
    showSetup({ fileName: file.name, bytes, wb, profiles });
  } catch (e) {
    console.error(e);
    toast(`読み込めませんでした: ${e.message}`, 'error');
  }
}

// ---------------- 設定・書き込む列の選択 ----------------
async function showSetup({ fileName, bytes, wb, profiles }) {
  setHeader(fileName, [el('button', { class: 'btn ghost', onclick: showHome }, '戻る')]);
  const profileSel = el('select', { class: 'field' },
    profiles.map((p, i) => el('option', { value: i }, p.name)),
    el('option', { value: '' }, '設定なし (自動判定)'));
  const sheetSel = el('select', { class: 'field' });
  const sheetLbl = el('label', { class: 'lbl' }, 'シート', sheetSel);
  let sheetCheck = null; // 設定を使うときのシートの候補 ({ candidates, named })
  // 候補が複数あるときに選んだシートを、設定ごとに覚えておく。候補の組み合わせが変わったら (シートが増えたなど) 選び直してもらう
  const choiceKey = (profile) => `sheetChoice:${profile.name}`;
  const loadChoice = (profile, names) => {
    try {
      const c = JSON.parse(localStorage.getItem(choiceKey(profile)) || 'null');
      return c && JSON.stringify(c.candidates) === JSON.stringify(names) && names.includes(c.chosen) ? c.chosen : null;
    } catch { return null; }
  };
  let rememberedSheet = null;

  /** シートの選択肢を作る。設定を使うときは列の見出しが合うシートだけ (シート名は改訂で変わるため名前では決めない) */
  async function fillSheets() {
    const profile = profileSel.value === '' ? null : profiles[+profileSel.value];
    sheetCheck = profile ? await profileSheets(profile, wb, loadSheet) : null;
    const names = sheetCheck ? sheetCheck.candidates : wb.sheets.map((x) => x.name);
    rememberedSheet = profile && names.length > 1 ? loadChoice(profile, names) : null;
    sheetSel.replaceChildren(
      // 候補が複数あるときは選ぶまで始められないようにする (知らないうちに古いシートへ書かないため)
      sheetCheck && names.length > 1 && !rememberedSheet ? el('option', { value: '', selected: true }, '— シートを選んでください —') : null,
      ...names.map((n) => el('option', { value: n, selected: n === rememberedSheet }, n)));
  }
  const colSel = el('select', { class: 'field' });
  const info = el('div', { class: 'setup-info' });
  const startBtn = el('button', { class: 'btn primary big' }, '入力をはじめる');
  let current = null; // { profile, sheetName, sheet, cards }

  const excelHint = 'アプリでは入力できないので、この点検表は Excel で入力してください (設定ファイルの更新が必要です)。';
  async function refresh() {
    const profile = profileSel.value === '' ? null : profiles[+profileSel.value];
    const candidates = sheetCheck?.candidates || [];
    sheetLbl.hidden = profile ? candidates.length < 2 : wb.sheets.length < 2;
    const stop = (...msg) => {
      current = null;
      info.replaceChildren(el('div', { class: 'attention' }, msg));
      colSel.replaceChildren();
      startBtn.disabled = true;
    };
    if (profile && !candidates.length) {
      // 列の見出しが合うシートがない = 表の形が変わった
      stop(el('p', {}, `この点検表は、設定「${profile.name}」と表の形 (列の見出し) が合いません。表が改訂された可能性があります。`),
        el('p', {}, excelHint),
        sheetCheck?.named?.problems.length ? el('ul', {}, sheetCheck.named.problems.map((w) => el('li', {}, w))) : null);
      return;
    }
    if (profile && candidates.length > 1 && !sheetSel.value) {
      stop(el('p', {}, `この設定に合うシートが ${candidates.length} つあります: ${candidates.map((n) => `「${n}」`).join('、')}`),
        el('p', {}, '入力するシートを選んでください。表が改訂されて新しいシートが追加された可能性があります。'));
      return;
    }
    const sheetName = sheetSel.value || candidates[0] || wb.sheets[0]?.name;
    try {
      const { sheet, cards } = await makeCards(wb, sheetName, profile);
      current = { profile, sheetName, sheet, cards };
    } catch (e) {
      current = null;
      info.replaceChildren(el('p', { class: 'error' }, e.message));
      if (!profile) info.append(el('p', { class: 'hint' }, 'この点検表用の設定ファイルを読み込むと開けるようになります。'));
      colSel.replaceChildren();
      startBtn.disabled = true;
      return;
    }
    const { sheet, cards } = current;
    startBtn.disabled = false;
    const dateRow = cards.meta.find((m) => m.type === 'date')?.row;
    const recordCols = cards.recordCols || [];
    const nextCol = cards.nextCol ?? cards.targetCol;
    const opts = recordCols.map((c) => {
      const d = dateRow ? displayValue(sheet.get(dateRow, c)) : '';
      return el('option', { value: c }, `${colName(c)}列 ${d ? `(${d})` : ''} - 空欄のみ追記`);
    });
    opts.push(el('option', { value: nextCol, selected: true }, `${colName(nextCol)}列 (新しい記録)`));
    colSel.replaceChildren(...opts);
    const sections = new Set(cards.items.map((c) => c.section).filter(Boolean));
    info.replaceChildren();
    add(info, [
      el('dl', {},
        el('dt', {}, 'シート'), el('dd', {}, sheetName),
        el('dt', {}, '基本情報'), el('dd', {}, cards.meta.map((m) => m.label).join('・') || '-'),
        el('dt', {}, '点検項目'), el('dd', {}, `${cards.items.length}件 (${sections.size}区分)`),
        cards.conditions.length ? [el('dt', {}, '条件'), el('dd', {}, cards.conditions.map((c) => c.label).join('・'))] : null,
        el('dt', {}, '前回の記録'), el('dd', {}, cards.prevCol ? `${colName(cards.prevCol)}列` : 'なし'),
      ),
      profile && sheetName !== profile.sheet ? el('p', { class: 'hint' },
        `シート名が設定 (「${profile.sheet}」) と違いますが、列の見出しが合うのでこのシートを使います。`) : null,
      profile && candidates.length > 1 ? el('p', { class: rememberedSheet === sheetName ? 'hint' : 'warn' },
        rememberedSheet === sheetName
          ? `前回選んだシートです (ほかに ${candidates.filter((n) => n !== sheetName).map((n) => `「${n}」`).join('、')} も設定に合います)。`
          : `この設定に合うシートがほかにもあります (${candidates.filter((n) => n !== sheetName).map((n) => `「${n}」`).join('、')})。入力するシートで合っているか確認してください。`) : null,
      // 行が設定と合わない (行の増減など) ときは注意にとどめ、入力は止めない (利用者の指定)
      cards.warnings.length ? el('div', { class: 'attention' },
        el('p', {}, REVISED_MSG),
        el('details', {}, el('summary', {}, `設定と合わない箇所 (${cards.warnings.length} 件)`),
          el('ul', {}, cards.warnings.map((w) => el('li', {}, w))))) : null,
    ]);
  }

  profileSel.addEventListener('change', async () => { await fillSheets(); refresh(); });
  sheetSel.addEventListener('change', refresh);
  startBtn.addEventListener('click', async () => {
    const targetCol = +colSel.value;
    const { profile, sheetName } = current;
    if (profile && sheetCheck?.candidates.length > 1) {
      try { localStorage.setItem(choiceKey(profile), JSON.stringify({ candidates: sheetCheck.candidates, chosen: sheetName })); } catch {}
    }
    const { sheet, cards } = await makeCards(wb, sheetName, profile, targetCol);
    const values = {};
    // 基本情報は今日の日付・現在時刻・前回の名前で埋めておく
    for (const m of cards.meta) {
      if (m.locked) continue;
      if (m.type === 'date') values[m.row] = today();
      else if (m.type === 'time') values[m.row] = nowTime();
      else if (m.type === 'name') values[m.row] = localStorage.getItem('lastName') || '';
    }
    const session = {
      id: `${Date.now()}`, fileName, bytes, sheetName, targetCol, values, profile, cond: {}, auto: {}, fromPrev: {},
      createdAt: Date.now(),
    };
    S = { session, wb, sheet, cards, filter: 'all' };
    applyConditionWrites();
    await putSession(session);
    requestPersist();
    showCards();
  });

  // 設定を取り込んだら、同じ点検表で設定を選び直して開き直す
  const reopenWithProfile = async (p) => {
    const all = (await listProfiles()).filter((x) => profileMatches(x, wb, fileName));
    if (!all.some((x) => x.name === p.name)) toast(`設定「${p.name}」はこの点検表には合いません`, 'error');
    showSetup({ fileName, bytes, wb, profiles: all });
  };

  mount(
    el('section', { class: 'setup' },
      profiles.length ? el('label', { class: 'lbl' }, '設定', profileSel) : el('div', { class: 'warn-box' },
        el('p', { class: 'warn' }, 'この点検表用の設定ファイルが、この端末にはまだ読み込まれていません。'),
        el('p', {}, '点検表と同じフォルダにある設定ファイル (.json) を読み込んでください。'),
        profileImportButton(reopenWithProfile)),
      sheetLbl,
      el('label', { class: 'lbl' }, '書き込む列', colSel),
      info,
      startBtn,
    ),
  );
  if (!profiles.length) profileSel.value = '';
  await fillSheets();
  await refresh();
}

async function resume(id) {
  try {
    const session = await getSession(id);
    const wb = await openWorkbook(session.bytes);
    const { sheet, cards } = await makeCards(wb, session.sheetName, session.profile || null, session.targetCol);
    session.cond ??= {};
    session.auto ??= {};
    session.fromPrev ??= {};
    S = { session, wb, sheet, cards, filter: 'all' };
    showCards();
  } catch (e) {
    console.error(e);
    toast(`再開できませんでした: ${e.message}`, 'error');
  }
}

// ---------------- 保存 ----------------
let saveTimer = null;
function scheduleSave() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => putSession(S.session).catch((e) => toast(`保存失敗: ${e.message}`, 'error')), 300);
}
// ページを離れる/バックグラウンドに回る前に確実に保存
const flushSave = () => { if (S && saveTimer) { clearTimeout(saveTimer); saveTimer = null; putSession(S.session); } };
document.addEventListener('visibilitychange', () => document.visibilityState === 'hidden' && flushSave());
window.addEventListener('pagehide', flushSave);

// ---------------- 判定 ----------------
function judge(card, raw) {
  if (card.locked) return { state: 'locked' };
  const s = String(raw ?? '').trim();
  if (!s) return { state: 'empty' };
  if (card.type === 'choice') {
    const set = card.setText;
    if (set && card.options?.some((o) => norm(o) === norm(set)) && norm(s) !== norm(set)) {
      return { state: 'ng', msg: `設定 (${set}) と違います${card.effect ? ` / 影響: ${card.effect}` : ''}` };
    }
    return { state: 'ok' };
  }
  if (card.type === 'number' || card.type === 'formula') {
    const v = toCellValue('number', s);
    // 自動計算は表示と同じ桁 (桁数の列) に丸めてから判定する
    if (card.type === 'formula' && v.kind === 'n' && Number.isInteger(card.digits)) v.v = +v.v.toFixed(card.digits);
    if (v.kind !== 'n') return { state: 'warn', msg: '数値ではありません (文字として書き込みます)' };
    const { range } = card;
    if (range && ((range.min !== undefined && v.v < range.min) || (range.max !== undefined && v.v > range.max))) {
      if (card.rangeKind === 'reference') return { state: 'warn', msg: '参考範囲から外れています' };
      return { state: 'ng', msg: `${S.cards.ngMessage}${card.effect ? ` / 影響: ${card.effect}` : ''}` };
    }
    const decimals = (s.normalize('NFKC').split('.')[1] || '').length;
    if (card.type !== 'formula' && card.digits !== null && card.digits !== undefined && decimals > card.digits) {
      return { state: 'warn', msg: `小数点以下は ${card.digits} 桁です` };
    }
  }
  return { state: 'ok' };
}

// ---------------- 条件による振り分け ----------------
const condValues = () => conditionValues(S.cards, S.session);
const applicable = (card) => card.kind === 'meta' || isApplicable(card, S.cards, condValues());

/** 選択式の条件の「記入先」セルに、選んだ選択肢を書く */
function applyConditionWrites() {
  const { session, cards } = S;
  const cv = condValues();
  for (const cd of cards.conditions) {
    for (const [opt, row] of Object.entries(cd.writes || {})) {
      if (cv[cd.key] === opt) {
        if (!String(session.values[row] ?? '').trim()) { session.values[row] = opt; session.auto[row] = true; }
      } else if (session.auto[row]) {
        delete session.values[row];
        delete session.auto[row];
      }
    }
  }
}

/** 書き込み対象のカード (任意入力を含む) */
function targetCards() {
  return [...S.cards.meta, ...S.cards.items].filter((c) => !c.locked && applicable(c));
}
/** 入力が必要なカード (任意入力を除く)。進み具合と未入力の数え方に使う */
const requiredCards = () => targetCards().filter((c) => !c.optional && c.type !== 'formula');
const isEmptyRequired = (c) => !c.optional && c.type !== 'formula' && judge(c, S.session.values[c.row]).state === 'empty';

/** 自動計算の行を、今の入力値から計算し直す (計算できなければ空) */
function recalcFormulas() {
  const { cards, sheet, session } = S;
  const fcards = cards.items.filter((c) => c.type === 'formula');
  if (!fcards.length) return;
  const get = (r, c) => {
    if (c === cards.targetCol) {
      const v = toCellValue('number', session.values[r]);
      if (v) return v.kind === 'n' ? v.v : null;
    }
    const cell = sheet.get(r, c);
    return cell && typeof cell.value === 'number' ? cell.value : null;
  };
  for (const card of fcards) {
    let v = null;
    try { v = evaluate(card.formula, get); } catch (e) { console.warn('式を計算できません', card.formula, e); }
    session.values[card.row] = v === null ? '' : String(+v.toPrecision(12));
  }
  for (const n of main.querySelectorAll('.card.is-formula')) n._refresh?.();
}

// テンキーで入力中 (まだ確定していない) の行。判定・自動計算は確定してから行う
// (打っている途中の値で「管理値外」を出さないため)。確定 = Enter・戻る・進む・閉じる・別の欄へ移る
let editingRow = null;
/** 画面に出す判定 (入力中の行は判定しない) */
function shownJudge(card) {
  const j = judge(card, S.session.values[card.row]);
  return editingRow === card.row && j.state !== 'empty' ? { state: 'editing' } : j;
}
function beginEditing(row) { editingRow = row; }
function endEditing() {
  if (editingRow === null) return;
  const row = editingRow;
  editingRow = null;
  recalcFormulas();
  main.querySelector(`.card[data-row="${row}"]`)?._refresh?.();
  updateProgress();
}

function updateProgress() {
  const list = requiredCards();
  const done = list.filter((c) => judge(c, S.session.values[c.row]).state !== 'empty').length;
  const ng = targetCards().filter((c) => shownJudge(c).state === 'ng').length;
  $('#progress-text').textContent = `${done} / ${list.length}`;
  $('#progress-bar').style.width = `${list.length ? (done / list.length) * 100 : 0}%`;
  $('#ng-count').textContent = ng ? `管理値外 ${ng}` : '';
}

function applyFilter() {
  for (const node of main.querySelectorAll('.card[data-row]')) {
    const card = node._card;
    node._refresh?.();
    const j = shownJudge(card);
    const show = applicable(card) && (
      S.filter === 'all' || (S.filter === 'empty' && isEmptyRequired(card)) || (S.filter === 'ng' && j.state === 'ng'));
    node.hidden = !show;
  }
  for (const node of main.querySelectorAll('.card.cond')) node.hidden = S.filter !== 'all';
  for (const sec of main.querySelectorAll('.section')) {
    sec.hidden = ![...sec.querySelectorAll('.card')].some((c) => !c.hidden);
  }
}

// ---------------- カード ----------------
function setValue(card, value, { fromPrev = false } = {}) {
  const { session } = S;
  session.values[card.row] = value;
  if (fromPrev) session.fromPrev[card.row] = true;
  else delete session.fromPrev[card.row];
  delete session.auto[card.row];
  if (card.type === 'name') localStorage.setItem('lastName', value);
  if (editingRow !== card.row) recalcFormulas(); // 入力中は確定してから計算する
  scheduleSave();
  if (card.conditionSource) applyFilter(); // 条件の元になる項目が変わると、対象の行が変わる
  updateProgress();
}

function renderInput(card, refresh) {
  const value = S.session.values[card.row] ?? '';
  if (card.locked) return el('div', { class: 'locked-value' }, card.existing, el('span', { class: 'badge' }, '記入済'));
  if (card.type === 'choice') {
    const group = el('div', { class: 'choices' });
    const paint = () => {
      const cur = norm(S.session.values[card.row]);
      for (const b of group.children) b.classList.toggle('on', norm(b.dataset.v) === cur);
    };
    for (const o of card.options) {
      group.append(el('button', {
        class: `choice ${norm(o) === norm(card.prev) ? 'prev' : ''}`, dataset: { v: o }, type: 'button',
        onclick: () => {
          const same = norm(S.session.values[card.row]) === norm(o);
          setValue(card, same ? '' : o);
          paint();
          refresh();
          if (!same) focusNext(card, { scroll: false });
        },
      }, o));
    }
    paint();
    group._paint = paint;
    return group;
  }
  if (card.type === 'formula') {
    // 自動計算: 入力欄の代わりに計算結果を表示する
    const out = el('div', { class: 'formula-value' });
    out._paint = () => {
      const v = String(S.session.values[card.row] ?? '');
      const shown = v && card.digits !== null && card.digits !== undefined ? (+v).toFixed(card.digits) : v;
      out.replaceChildren(el('span', {}, shown || '—'), el('span', { class: 'badge' }, '自動計算'));
    };
    out._paint();
    return out;
  }
  const onInput = (e) => { setValue(card, e.target.value); refresh(); };
  const common = { class: 'field big', value, enterkeyhint: 'next', oninput: onInput, onchange: onInput };
  if (card.type === 'date') return el('input', { ...common, type: 'date' });
  if (card.type === 'time') return el('input', { ...common, type: 'time' });
  // 管理値はチップに出しているので、入力欄には重ねて出さない
  const placeholder = card.type === 'number' && !card.rangeText && !card.setText ? '数値' : '';
  if (card.type === 'number') {
    // 画面キーボードは出さず、アプリのテンキーで入力する。フォーカスは当てない (readonly)。
    // iPad の Safari はフォーカス時に自動でスクロールし、その間のタップが効かなくなるため
    return el('input', { ...common, type: 'text', inputmode: 'none', readonly: true, autocomplete: 'off', placeholder, dataset: { keypad: '' } });
  }
  return el('input', { ...common, type: 'text', autocomplete: 'off', placeholder, list: card.type === 'name' ? 'names' : undefined });
}

const canUsePrev = (card) => card.kind === 'item' && !card.locked && card.prev !== '' && card.type !== 'choice';

/** 前回値の表示 (桁数の列がある数値は、その桁に揃える。書き込む値はそのまま) */
const prevText = (card) => {
  const v = toCellValue('number', card.prev);
  return v?.kind === 'n' && Number.isInteger(card.digits) && card.type !== 'choice' ? v.v.toFixed(card.digits) : card.prev;
};

function cardChips(card) {
  if (card.kind !== 'item') return null;
  const chips = card.chips.map((c) => el('span', { class: 'chip where' }, c));
  if (card.rangeKind === 'range') chips.push(el('span', { class: 'chip range' }, `管理 ${card.rangeText}`));
  else if (card.rangeKind === 'reference') chips.push(el('span', { class: 'chip ref' }, `参考 ${card.rangeText}`));
  if (card.setText && card.type !== 'choice') chips.push(el('span', { class: 'chip' }, `設定 ${card.setText}`));
  else if (card.setText && card.type === 'choice') chips.push(el('span', { class: 'chip range' }, `設定 ${card.setText}`));
  if (!card.rangeKind && card.kindText && card.type !== 'choice' && !/[～~]/.test(card.kindText)) {
    chips.push(el('span', { class: 'chip' }, card.kindText));
  }
  if (card.prev) chips.push(el('span', { class: 'chip prev' }, `前回 ${prevText(card)}`));
  return el('div', { class: 'chips' }, chips);
}

/** 備考 (備考の列・セルのコメント) と、前回の記録に付いたコメント */
function cardNotes(card) {
  const notes = [card.note, ...(card.comments || [])].filter(Boolean);
  const uniq = [...new Set(notes)];
  const memo = card.lastMemo;
  if (!uniq.length && !memo) return null;
  const memoTag = memo && (memo.prev ? '前回メモ' : `メモ${memo.date ? ` (${memo.date})` : ''}`);
  return el('div', { class: 'notes' },
    uniq.map((t) => el('div', { class: 'note' }, el('span', { class: 'note-tag' }, '備考'), t)),
    memo ? el('div', { class: 'note prev' }, el('span', { class: 'note-tag' }, memoTag), memo.text) : null);
}

function renderCard(card) {
  const status = el('div', { class: 'status' });
  const node = el('div', {
    class: `card ${card.type === 'choice' ? 'is-choice' : ''} ${card.type === 'formula' ? 'is-formula' : ''} ${card.compact ? 'compact' : ''}`, dataset: { row: card.row },
  });
  node._card = card;
  const prevBadge = el('span', { class: 'badge prev-badge' }, '前回値');
  const autoBadge = el('span', { class: 'badge prev-badge' }, '自動');
  let input;
  const refresh = () => {
    const j = shownJudge(card);
    node.dataset.state = j.state;
    status.textContent = j.msg || '';
    prevBadge.hidden = !S.session.fromPrev[card.row];
    autoBadge.hidden = !S.session.auto[card.row];
    if (input?._paint) input._paint();
    else if (input?.tagName === 'INPUT' && document.activeElement !== input) input.value = S.session.values[card.row] ?? '';
  };
  node._refresh = refresh;
  input = renderInput(card, refresh);
  if (input.tagName === 'INPUT') {
    input.addEventListener('keydown', (e) => {
      if (e.key !== 'Enter' || e.isComposing) return;
      e.preventDefault();
      // 空のまま Enter = 前回値で確定 (ありがちな「前回と同じ」を 1 打で)
      if (!input.value.trim() && canUsePrev(card)) {
        input.value = card.prev;
        setValue(card, card.prev, { fromPrev: true });
        refresh();
      }
      focusNext(card);
    });
    input.addEventListener('blur', applyFilterLater);
  }
  add(node, [
    el('div', { class: 'card-head' },
      el('div', { class: 'card-title' },
        card.watch ? el('span', { class: 'watch', title: '監視項目' }, '●') : null,
        card.label,
        card.sub && !card.hideSub ? el('span', { class: 'code' }, card.sub) : null,
        card.unit ? el('span', { class: 'code' }, `[${card.unit}]`) : null),
      card.optional ? el('span', { class: 'badge' }, '任意') : null,
      prevBadge, autoBadge,
      el('div', { class: 'card-ref' }, card.ref),
    ),
    cardChips(card),
    input,
    status,
    cardNotes(card),
  ]);
  refresh();
  return node;
}

/** 選択式の条件カード (どちらかの系統だけを入力する場合など) */
function renderConditionCard(cd) {
  const group = el('div', { class: 'choices' });
  const paint = () => {
    const cur = condValues()[cd.key];
    for (const b of group.children) b.classList.toggle('on', b.dataset.v === cur);
  };
  for (const o of cd.options) {
    group.append(el('button', {
      class: 'choice', dataset: { v: o }, type: 'button',
      onclick: () => {
        S.session.cond[cd.key] = o;
        applyConditionWrites();
        scheduleSave();
        paint();
        applyFilter();
        updateProgress();
      },
    }, o));
  }
  paint();
  return el('div', { class: 'card cond' },
    el('div', { class: 'card-head' }, el('div', { class: 'card-title' }, `${cd.label} を選択`)),
    el('div', { class: 'hint' }, '選んだ側の項目だけを入力します'),
    group);
}

// 入力中にカードが消えないよう、フィルタ反映はフォーカスが外れてから
let filterTimer;
function applyFilterLater() {
  clearTimeout(filterTimer);
  filterTimer = setTimeout(() => {
    if (!main.contains(document.activeElement) || document.activeElement === document.body) applyFilter();
  }, 400);
}

function focusNext(card, { scroll = true, back = false } = {}) {
  const nodes = [...main.querySelectorAll('.card[data-row]')].filter((n) => !n.hidden);
  const i = nodes.findIndex((n) => n._card === card);
  const next = (back ? nodes.slice(0, i).reverse() : nodes.slice(i + 1)).find((n) => !n._card.locked && n._card.type !== 'formula');
  if (!next) {
    if (back) { toast('最初の項目です'); return; }
    document.activeElement?.blur?.();
    keypad.hide();
    toast('最後の項目です');
    return;
  }
  const inp = next.querySelector('input');
  if (inp) activateInput(inp);
  else { document.activeElement?.blur?.(); keypad.hide(); } // 選択式はキーボードを閉じてボタンで選ぶ
  // テンキー使用中はスクロールを一瞬で済ませる (スクロール中のタップは iPad で効かないため)
  if (scroll || !inp) next.scrollIntoView({ block: 'center', behavior: keypad.isOpen() ? 'auto' : 'smooth' });
  if (keypad.isOpen()) keypad.reveal();
}

/** 入力欄を入力中にする。数値欄はフォーカスせずテンキーの対象にする */
function activateInput(inp) {
  if (inp.matches('[data-keypad]')) {
    if (document.activeElement !== document.body) document.activeElement?.blur?.();
    keypad.show(inp);
  } else {
    inp.focus({ preventScroll: true });
  }
}

function jumpToFirstEmpty() {
  const node = [...main.querySelectorAll('.card[data-row]')].find((n) => !n.hidden && isEmptyRequired(n._card));
  if (!node) { toast('未入力の項目はありません', 'ok'); return; }
  node.scrollIntoView({ block: 'center', behavior: keypad.isOpen() ? 'auto' : 'smooth' });
  const inp = node.querySelector('input');
  if (inp) activateInput(inp);
  if (keypad.isOpen()) keypad.reveal();
}

function showCards() {
  const { session, cards } = S;
  recalcFormulas();
  setHeader(session.fileName, [
    el('button', { class: 'btn ghost', onclick: () => { flushSave(); showHome(); } }, '一覧'),
    el('button', { class: 'btn primary', onclick: showExport }, '書き出し'),
  ]);
  const names = el('datalist', { id: 'names' },
    [localStorage.getItem('lastName')].filter(Boolean).map((n) => el('option', { value: n })));

  const filterBtns = [['all', 'すべて'], ['empty', '未入力'], ['ng', '管理値外']].map(([k, label]) =>
    el('button', {
      class: `seg ${S.filter === k ? 'on' : ''}`,
      onclick: (e) => {
        S.filter = k;
        for (const b of e.target.parentNode.children) b.classList.toggle('on', b === e.target);
        applyFilter();
      },
    }, label));

  // 項目と条件カードを並べ、工程 (区分) ごとにまとめる
  const selectConds = cards.conditions.filter((c) => c.kind === 'select');
  const groups = [];
  const push = (name, node, own = false, compact = false) => {
    const last = groups[groups.length - 1];
    if (!own && last && !last.own && !last.split && last.name === name && last.compact === compact) last.nodes.push(node);
    else groups.push({ name, nodes: [node], own, compact });
  };
  // 列分け表示: 値ごとの一覧を横に並べる (幅が足りなければ縦に積む)
  const pushSplit = (name, card, node) => {
    const last = groups[groups.length - 1];
    let g = last?.split?.id === card.split.id ? last : null;
    if (!g) { g = { name, split: { id: card.split.id, lists: new Map() } }; groups.push(g); }
    if (!g.split.lists.has(card.split.key)) g.split.lists.set(card.split.key, []);
    g.split.lists.get(card.split.key).push(node);
  };
  for (const cd of selectConds.filter((c) => !c.placeBefore)) push(cd.label, renderConditionCard(cd), true);
  for (const c of cards.items) {
    for (const cd of selectConds.filter((x) => x.placeBefore === c.row)) push(cd.label, renderConditionCard(cd), true);
    if (c.split) pushSplit(c.section, c, renderCard(c));
    else push(c.section, renderCard(c), false, !!c.compact);
  }
  const renderGroupBody = (g) => {
    if (g.split) {
      return el('div', { class: 'split' }, [...g.split.lists].map(([key, nodes]) => el('div', { class: 'split-col' },
        el('div', { class: 'split-head' }, key || '-'),
        el('div', { class: 'compact-list' }, nodes))));
    }
    return el('div', { class: g.compact ? 'compact-list' : 'cards' }, g.nodes);
  };

  mount(
    names,
    el('div', { class: 'toolbar' },
      el('div', { class: 'progress' },
        el('div', { class: 'progress-label' },
          el('span', { id: 'progress-text' }), el('span', { id: 'ng-count', class: 'ng' }),
          el('span', { class: 'target' }, `${colName(cards.targetCol)}列に記入`)),
        el('div', { class: 'progress-track' }, el('div', { id: 'progress-bar' }))),
      el('div', { class: 'toolbar-row' },
        el('div', { class: 'segs' }, filterBtns),
        el('button', { class: 'btn ghost', onclick: jumpToFirstEmpty }, '未入力へ ↓')),
    ),
    cards.warnings.length ? el('div', { class: 'attention' }, el('p', {}, REVISED_MSG)) : null,
    el('div', { class: 'section' },
      el('h2', { class: 'section-title' }, '基本情報'),
      el('div', { class: 'cards' }, cards.meta.map(renderCard))),
    // 直前と同じ見出しが続くとき (一覧表示とカードの切り替わりなど) は、見出しを繰り返さない
    groups.map((g, i) => el('div', { class: 'section' },
      i > 0 && groups[i - 1].name === g.name ? null : el('h2', { class: 'section-title' }, g.name || '項目'),
      renderGroupBody(g))),
    el('div', { class: 'bottom-space' }),
  );
  updateProgress();
  applyFilter();
}

// ---------------- 書き出し ----------------
async function showExport() {
  flushSave();
  const { session, wb, sheet, cards } = S;
  const list = targetCards();
  const empty = list.filter(isEmptyRequired);
  const ng = list.filter((c) => judge(c, session.values[c.row]).state === 'ng');
  const fromPrev = list.filter((c) => session.fromPrev?.[c.row] && session.values[c.row] === c.prev);
  const skipped = cards.items.filter((c) => !c.locked && !applicable(c) && String(session.values[c.row] ?? '').trim());
  let result;
  try {
    result = await exportWorkbook(wb, sheet, cards.targetCol, list, session.values);
  } catch (e) {
    toast(e.message, 'error');
    return;
  }
  // 書き込み済みファイルは元と同じ名前、読み込み時点の元ファイルは _backup_日時 を付けて一緒に出す。
  // (Safari からは元ファイルを直接上書きできないため、保存時に「置き換え」てもらう)
  const d = new Date();
  const stamp = `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}`;
  const name = session.fileName;
  const m = /^(.*?)(\.[^.]+)?$/.exec(name);
  const backupName = `${m[1]}_backup_${stamp}${m[2] || '.xlsx'}`;
  const type = /\.xlsm$/i.test(name) ? 'application/vnd.ms-excel.sheet.macroEnabled.12' : result.blob.type;
  const file = new File([result.blob], name, { type });
  const backup = new File([session.bytes], backupName, { type });
  const files = [file, backup];

  const markDone = async () => {
    session.exportedAt = Date.now();
    await putSession(session);
  };
  const download = async () => {
    for (const f of files) {
      const url = URL.createObjectURL(f);
      const a = el('a', { href: url, download: f.name });
      document.body.append(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 10000);
      await new Promise((r) => setTimeout(r, 300));
    }
    await markDone();
    close();
  };
  const share = async () => {
    try {
      await navigator.share({ files });
      await markDone();
      close();
    } catch (e) {
      if (e.name !== 'AbortError') toast(`共有できませんでした: ${e.message}`, 'error');
    }
  };
  const canShare = !!navigator.canShare?.({ files });
  const isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
  const where = (c) => `${c.section ? `${c.section} / ` : ''}${c.label}${c.sub ? ` ${c.sub}` : ''}`;

  const dlg = el('div', { class: 'modal-bg' },
    el('div', { class: 'modal' },
      el('h2', {}, 'Excel に書き出し'),
      el('p', {}, `${colName(cards.targetCol)}列に ${result.written.length} セル書き込みます。`),
      empty.length ? el('p', { class: 'warn' }, `未入力が ${empty.length} 件あります (空欄のまま書き出します)。`) : null,
      fromPrev.length ? el('p', {}, `前回値のまま確定: ${fromPrev.length} 件`) : null,
      skipped.length ? el('p', {}, `条件の対象外になった入力 ${skipped.length} 件は書き込みません。`) : null,
      ng.length ? el('div', { class: 'ng-list' },
        el('p', { class: 'error' }, `管理値外が ${ng.length} 件あります:`),
        el('ul', {}, ng.map((c) => el('li', {}, `${where(c)}: ${session.values[c.row]}${c.rangeText ? ` (管理 ${c.rangeText})` : c.setText ? ` (設定 ${c.setText})` : ''}`)))) : null,
      el('div', { class: 'files' },
        el('div', {}, el('span', { class: 'badge ok' }, '記入済'), ' ', name),
        el('div', {}, el('span', { class: 'badge' }, '元の状態'), ' ', backupName)),
      el('p', { class: 'hint' }, '元のファイルと同じフォルダに保存し、同じ名前のファイルは「置き換え」てください。元の状態は backup のファイルに残ります。'),
      el('div', { class: 'modal-actions' },
        canShare && isIOS ? el('button', { class: 'btn primary big', onclick: share }, '共有 / 「ファイル」に保存') : null,
        el('button', { class: `btn ${canShare && isIOS ? '' : 'primary'} big`, onclick: download }, 'ダウンロード'),
        el('button', { class: 'btn ghost', onclick: () => close() }, 'キャンセル'),
      ),
    ),
  );
  const close = () => dlg.remove();
  document.body.append(dlg);
}

// ---------------- テンキー ----------------
// iPad の Safari は inputmode="decimal" でもテンキーにならない (通常のキーボードの数字段が出るだけ) ため、
// 数値入力欄 (data-keypad) では画面下にアプリのテンキーを出す。
// 数値欄にはフォーカスを当てず (readonly)、入力中の欄 (target) はここで管理する。
// iPad ではフォーカス時の自動スクロールや、なめらかなスクロールの最中のタップが効かなくなるため。
// 閉じるのは「閉じる」・ほかの入力欄へ移ったとき・選択式を押したとき・数値欄でない項目へ移ったときだけ。
const keypad = (() => {
  let target = null;
  let fresh = false; // 入力中の欄に移った直後。最初のキーで値を置き換える (打ち直しやすくするため)
  const label = el('div', { class: 'keypad-label' });
  const setValue = (v) => {
    fresh = false;
    beginEditing(target.closest('.card')?._card?.row ?? null);
    target.classList.remove('keypad-fresh');
    target.value = v;
    target.dispatchEvent(new Event('input', { bubbles: true }));
  };
  const actions = {
    back: () => setValue(fresh ? '' : target.value.slice(0, -1)),
    // Enter と同じ (空欄なら前回値で確定して次へ)
    enter: () => target.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true })),
    // 戻る / 進む: 値には触れずに前後の項目へ移るだけ
    prev: () => { const card = target.closest('.card')?._card; if (card) focusNext(card, { back: true }); },
    next: () => { const card = target.closest('.card')?._card; if (card) focusNext(card); },
    close: () => hide(),
  };
  const press = (act) => {
    if (!target) return;
    // 入力欄が消えた・隠れた (絞り込みなど) ときは閉じる
    if (!target.isConnected || target.closest('.card')?.hidden) { hide(); return; }
    if (['enter', 'prev', 'next', 'close'].includes(act)) endEditing(); // 確定 (移れなかったときも判定を出す)
    if (actions[act]) actions[act]();
    else setValue((fresh ? '' : target.value) + act);
  };
  const key = (text, act, cls = '') => el('button', { type: 'button', class: `key ${cls}`, dataset: { act } }, text);
  // 大きさを変えるつまみ: 上端で高さ、テンキー部分の左端で横幅
  const hGrip = el('div', { class: 'keypad-hgrip', title: '上下にドラッグで高さを変更 (ダブルタップで元に戻す)' });
  const wGrip = el('div', { class: 'keypad-wgrip', title: '左右にドラッグで横幅を変更 (ダブルタップで元に戻す)' });
  const pad = el('div', { class: 'keypad', hidden: true },
    hGrip,
    el('div', { class: 'keypad-side' }, label,
      // 狭い画面 (iPhone) では記号だけにして高さを詰める (k-label を CSS で隠す)
      el('div', { class: 'keypad-nav' },
        key([el('span', {}, '▲'), el('span', { class: 'k-label' }, ' 戻る')], 'prev', 'nav'),
        key([el('span', {}, '▼'), el('span', { class: 'k-label' }, ' 進む')], 'next', 'nav')),
      key([el('span', { class: 'k-narrow' }, '×'), el('span', { class: 'k-label' }, '閉じる')], 'close', 'close')),
    wGrip,
    el('div', { class: 'keypad-keys' },
      key('7', '7'), key('8', '8'), key('9', '9'), key('-', '-'),
      key('4', '4'), key('5', '5'), key('6', '6'), key('BS', 'back', 'fn'),
      key('1', '1'), key('2', '2'), key('3', '3'), key('Enter', 'enter', 'next'),
      key('0', '0', 'zero'), key('.', '.')));
  // pointerdown で処理し既定動作を止める (ほかの要素にフォーカスを移さない・反応を速くするため)
  let swallowUntil = 0;
  pad.addEventListener('pointerdown', (e) => {
    e.preventDefault();
    const act = e.target.closest('.key')?.dataset.act;
    if (act) press(act);
    // キーでテンキーが閉じたら、続く click が下の数値欄に届いて開き直さないよう、少しの間 click を無視する
    if (pad.hidden) swallowUntil = Date.now() + 400;
  });
  document.body.append(pad);

  // ---- 大きさの調整 (端末ごと・縦向き/横向きごとに記憶。最小は標準の大きさ) ----
  const sizeKey = () => `keypadSize:${window.innerWidth > window.innerHeight ? 'landscape' : 'portrait'}`;
  let size = { h: 1, w: 1 }; // 記憶している倍率
  const applied = { h: 1, w: 1 }; // 今かかっている倍率 (画面に収まるよう size を抑えたもの)
  const rem = () => parseFloat(getComputedStyle(document.documentElement).fontSize) || 16;
  const BASE_KEY_H = 3.25; // rem (CSS の .key と同じ)
  const BASE_KEYS_W = 22; // rem (CSS の .keypad-keys と同じ)
  const SIDE_MIN = 11; // rem: 左側 (戻る・進むなど) に最低限残す幅
  /** 今の画面で取れる最大の倍率 */
  function limits() {
    const r = rem();
    const cs = getComputedStyle(pad);
    const innerW = pad.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight);
    const fixedW = wGrip.offsetWidth + 2 * (parseFloat(cs.columnGap) || 0);
    const maxW = Math.max(1, (innerW - fixedW - SIDE_MIN * r) / (BASE_KEYS_W * r));
    // 高さ: テンキー全体が画面の半分を超えない。倍率 1 のときの高さから、キー 4 段ぶんの伸びで割る
    const step = 4 * BASE_KEY_H * r; // 倍率が 1 増えたときの高さの増え方
    const baseH = pad.offsetHeight - step * (applied.h - 1);
    const maxH = Math.max(1, 1 + (window.innerHeight * 0.5 - baseH) / step);
    return { maxW, maxH };
  }
  function applySize({ save = false } = {}) {
    const { maxW, maxH } = limits();
    const h = Math.min(Math.max(size.h, 1), maxH);
    const w = Math.min(Math.max(size.w, 1), maxW);
    applied.h = h; applied.w = w;
    pad.style.setProperty('--kp-h', h.toFixed(3));
    pad.style.setProperty('--kp-w', w.toFixed(3));
    pad.style.setProperty('--kp-f', Math.min(h, w).toFixed(3)); // 文字はキーの短い辺に合わせる
    pad.classList.toggle('keypad-narrow-side', w > 1 && (maxW - w) * BASE_KEYS_W < 4); // 左側が狭いときは縦に並べる
    if (save) { try { localStorage.setItem(sizeKey(), JSON.stringify(size)); } catch {} }
    if (!pad.hidden) {
      document.documentElement.style.setProperty('--keypad-h', `${pad.offsetHeight}px`);
    }
  }
  function loadSize() {
    try { size = { h: 1, w: 1, ...JSON.parse(localStorage.getItem(sizeKey()) || '{}') }; } catch { size = { h: 1, w: 1 }; }
  }
  /** つまみのドラッグ。axis: 'h' (上下) / 'w' (左右)。ダブルタップで標準に戻す */
  function dragGrip(grip, axis) {
    let start = null;
    let lastTap = 0;
    grip.addEventListener('pointerdown', (e) => {
      e.preventDefault();
      e.stopPropagation();
      if (Date.now() - lastTap < 350) { size[axis] = 1; applySize({ save: true }); reveal(); lastTap = 0; return; }
      lastTap = Date.now();
      grip.setPointerCapture(e.pointerId);
      start = { x: e.clientX, y: e.clientY, v: Math.min(size[axis], axis === 'h' ? limits().maxH : limits().maxW) };
      grip.classList.add('dragging');
    });
    grip.addEventListener('pointermove', (e) => {
      if (!start) return;
      const r = rem();
      // 上へ・左へ動かすと大きくなる
      const d = axis === 'h' ? (start.y - e.clientY) / (4 * BASE_KEY_H * r) : (start.x - e.clientX) / (BASE_KEYS_W * r);
      if (Math.abs(d) > 0.02) lastTap = 0; // 動かしたらダブルタップとみなさない
      const { maxW, maxH } = limits();
      size[axis] = Math.min(Math.max(start.v + d, 1), axis === 'h' ? maxH : maxW);
      applySize();
    });
    const end = () => {
      if (!start) return;
      start = null;
      grip.classList.remove('dragging');
      applySize({ save: true });
      reveal();
    };
    grip.addEventListener('pointerup', end);
    grip.addEventListener('pointercancel', end);
  }
  dragGrip(hGrip, 'h');
  dragGrip(wGrip, 'w');
  loadSize();

  /** 入力中の項目がテンキーに隠れていれば、見える位置まで一瞬でずらす */
  function reveal() {
    if (!target) return;
    const r = (target.closest('.card') || target).getBoundingClientRect();
    const limit = window.innerHeight - pad.offsetHeight - 12;
    const top = (document.querySelector('.toolbar')?.getBoundingClientRect().bottom ?? 0) + 8;
    if (r.bottom > limit) window.scrollBy(0, r.bottom - limit);
    else if (r.top < top) window.scrollBy(0, r.top - top);
  }
  function show(input) {
    endEditing(); // 別の欄へ移る = 確定
    if (target && target !== input) applyFilterLater(); // 入力を終えた欄の絞り込みを反映 (フォーカス外れの代わり)
    target?.classList.remove('keypad-target', 'keypad-fresh');
    target = input;
    fresh = true;
    shownCount++;
    input.classList.add('keypad-target', 'keypad-fresh');
    const card = input.closest('.card')?._card;
    label.replaceChildren(
      el('div', { class: 'keypad-title' }, card?.label ?? '', card?.split?.key ? ` ${card.split.key}` : '',
        card?.sub && !card.hideSub ? el('span', { class: 'code' }, card.sub) : null,
        card?.unit ? el('span', { class: 'code' }, `[${card.unit}]`) : null),
      card?.prev ? el('div', { class: 'keypad-prev' }, `前回 ${prevText(card)}`) : null);
    pad.hidden = false;
    document.documentElement.classList.add('keypad-open');
    applySize();
  }
  function hide() {
    endEditing();
    if (target) applyFilterLater();
    target?.classList.remove('keypad-target', 'keypad-fresh');
    target = null;
    pad.hidden = true;
    document.documentElement.classList.remove('keypad-open');
  }
  // 数値欄のタップ: フォーカスさせずにテンキーの対象にする (スクロールの指の動きでは反応しないよう click で)
  document.addEventListener('pointerdown', (e) => { if (e.target.matches?.('input[data-keypad]')) e.preventDefault(); });
  let shownCount = 0; // show() の回数。click の間に入力中の欄が変わったかを見る
  let shownAtClick = 0;
  document.addEventListener('click', (e) => {
    shownAtClick = shownCount;
    if (Date.now() < swallowUntil) { swallowUntil = 0; e.preventDefault(); e.stopPropagation(); }
  }, true);
  document.addEventListener('click', (e) => {
    if (e.target.matches?.('input[data-keypad]')) { activateInput(e.target); reveal(); }
    // 選択式を押したら閉じる。ただし選んだ結果、次の数値欄へ移ったときは閉じない
    else if (target && e.target.closest?.('.choice') && shownCount === shownAtClick) hide();
  });
  // Tab キーなどでフォーカスが来たとき・ほかの入力欄へ移ったとき
  document.addEventListener('focusin', (e) => {
    if (e.target.matches?.('input[data-keypad]')) show(e.target);
    else if (target && !pad.contains(e.target)) hide();
  });
  // 外付けキーボード
  document.addEventListener('keydown', (e) => {
    // e.isTrusted: テンキーの Enter が送る合成イベントは対象外 (拾うと次々に進んでしまう)
    if (!target || !e.isTrusted || e.isComposing || e.metaKey || e.ctrlKey || e.altKey) return;
    const a = document.activeElement;
    if (a && a !== target && a !== document.body && a.matches('input, textarea, select')) return;
    const map = { Backspace: 'back', Enter: 'enter', ArrowUp: 'prev', ArrowDown: 'next', Escape: 'close' };
    const act = map[e.key] ?? (/^[0-9.-]$/.test(e.key) ? e.key : null);
    if (!act || (act === 'enter' && e.target.matches?.('input[data-keypad]'))) return; // 数値欄自身の Enter は既存の処理に任せる
    e.preventDefault();
    press(act);
  });
  // 回転などで高さが変わったら、入力中の項目を見える位置に戻す
  window.addEventListener('resize', () => {
    loadSize(); // 縦向き・横向きで別々に記憶している
    if (pad.hidden) return;
    applySize();
    target?.closest('.card')?.scrollIntoView({ block: 'center' });
  });
  return { show, hide, reveal, isOpen: () => !pad.hidden };
})();

// ---------------- 文字サイズ・画面の向き ----------------
const SIZES = [['', '標準'], ['l', '大'], ['xl', '特大'], ['s', '小']]; // 押すたびに 標準 → 大 → 特大 → 小 → 標準
function applySize(key) {
  if (key) document.documentElement.dataset.size = key;
  else delete document.documentElement.dataset.size;
  try { localStorage.setItem('textSize', key); } catch {}
  $('#size').replaceChildren('A', el('small', {}, SIZES.find(([k]) => k === key)[1]));
  updateTopbarH();
}
$('#size').addEventListener('click', () => {
  const cur = document.documentElement.dataset.size || '';
  const i = SIZES.findIndex(([k]) => k === cur);
  const [key, label] = SIZES[(i + 1) % SIZES.length];
  // 拡大後も、今見ているカードが画面内に残るようにする
  const focused = document.activeElement?.closest?.('.card');
  applySize(key);
  focused?.scrollIntoView({ block: 'center' });
  toast(`文字サイズ: ${label}`);
});

// ヘッダーの高さ (文字サイズ・回転で変わる) に合わせて、ツールバーの固定位置をずらす
function updateTopbarH() {
  document.documentElement.style.setProperty('--topbar-h', `${$('.topbar').getBoundingClientRect().height}px`);
}
new ResizeObserver(updateTopbarH).observe($('.topbar'));
window.addEventListener('resize', updateTopbarH);
applySize(document.documentElement.dataset.size || '');

// ---------------- 起動 ----------------
function updateOnline() {
  $('#net').textContent = navigator.onLine ? '' : 'オフライン';
}
window.addEventListener('online', updateOnline);
window.addEventListener('offline', updateOnline);
updateOnline();

// ---------------- 版の確認 ----------------
// 開いたとき・アプリに戻ってきたときに、オンラインなら新しい版がないか確認する。
// 新しい版は Service Worker が裏で読み込み、すぐ有効になる (sw.js の skipWaiting)。
// 画面は古いまま動いているので「再読み込み」を案内する (入力途中の内容は保存してから読み込み直す)。
const versionLine = el('p', { class: 'hint version-line' });
let swReg = null;
let currentVersion = '';
let lastCheck = 0;
const setVersionStatus = (msg) => {
  versionLine.textContent = `アプリの版: ${currentVersion || '不明'}${msg ? ` · ${msg}` : ''}`;
};
/** Service Worker に版を問い合わせる */
function askVersion(sw) {
  return new Promise((resolve) => {
    if (!sw) { resolve(''); return; }
    const ch = new MessageChannel();
    ch.port1.onmessage = (e) => resolve(e.data?.version || '');
    sw.postMessage('version', [ch.port2]);
    setTimeout(() => resolve(''), 2000);
  });
}
async function checkUpdate({ force = false } = {}) {
  if (!swReg) return;
  if (!navigator.onLine) { setVersionStatus('オフラインのため最新版か確認できません'); return; }
  if (!force && Date.now() - lastCheck < 60 * 1000) return; // 連続して確認しない
  lastCheck = Date.now();
  setVersionStatus('最新版か確認中…');
  try {
    await swReg.update();
  } catch {
    setVersionStatus('最新版か確認できませんでした');
    return;
  }
  if (swReg.installing || swReg.waiting) setVersionStatus('新しい版を読み込み中…');
  else setVersionStatus('最新版です');
}
function showUpdateBanner(version) {
  document.querySelector('.update-banner')?.remove();
  const reload = () => { flushSave(); location.reload(); };
  document.body.append(el('div', { class: 'update-banner' },
    el('span', {}, `新しい版${version ? ` (${version})` : ''}を読み込みました。`,
      S ? '入力内容は保存されます。' : ''),
    el('button', { class: 'btn primary', onclick: reload }, '再読み込み')));
}

if ('serviceWorker' in navigator) {
  const hadController = !!navigator.serviceWorker.controller; // 初回インストールでは案内しない
  navigator.serviceWorker.addEventListener('controllerchange', async () => {
    if (!hadController) return;
    const v = await askVersion(navigator.serviceWorker.controller);
    setVersionStatus(v && v !== currentVersion ? `新しい版 ${v} があります (再読み込みで切り替わります)` : '');
    showUpdateBanner(v);
  });
  navigator.serviceWorker.register('./sw.js').then(async (reg) => {
    swReg = reg;
    const ready = await navigator.serviceWorker.ready;
    currentVersion = await askVersion(navigator.serviceWorker.controller || ready.active);
    setVersionStatus('');
    checkUpdate({ force: true });
  }).catch((e) => console.warn('SW 登録失敗', e));
  // ホーム画面から戻ってきたとき (iPad では起動し直さずに再開することが多い)
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') checkUpdate(); });
  window.addEventListener('online', () => checkUpdate());
}
if (typeof DecompressionStream === 'undefined') {
  mount(el('p', { class: 'error' }, 'このブラウザは対応していません (iPadOS 16.4 以降の Safari が必要です)'));
} else {
  showHome();
}

// 開発用フック (自動テストからファイルを読み込ませる)
window.__app = { openFile, importProfile, showHome, checkUpdate, get state() { return S; } };
