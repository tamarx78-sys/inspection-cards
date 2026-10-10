import {
  openWorkbook, loadSheet, detectLayout, buildCards, exportWorkbook, toCellValue, colName, displayValue,
} from './xlsx.js';
import {
  parseProfile, profileMatches, buildProfileCards, conditionValues, isApplicable,
} from './profile.js';
import {
  putSession, getSession, deleteSession, listSessions, requestPersist, putProfile, deleteProfile, listProfiles,
} from './store.js';

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
    ),
  );
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
  const sheetSel = el('select', { class: 'field' }, wb.sheets.map((s, i) => el('option', { value: i }, s.name)));
  const sheetLbl = el('label', { class: 'lbl' }, 'シート', sheetSel);
  const colSel = el('select', { class: 'field' });
  const info = el('div', { class: 'setup-info' });
  const startBtn = el('button', { class: 'btn primary big' }, '入力をはじめる');
  let current = null; // { profile, sheetName, sheet, cards }

  async function refresh() {
    const profile = profileSel.value === '' ? null : profiles[+profileSel.value];
    sheetLbl.hidden = !!profile || wb.sheets.length < 2;
    const sheetName = profile ? profile.sheet : wb.sheets[+sheetSel.value].name;
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
      cards.warnings.length ? el('div', { class: 'warn-box' },
        el('p', { class: 'warn' }, '設定ファイルと表が合わない箇所があります。該当する振り分けは使わず、行をすべて表示します:'),
        el('ul', {}, cards.warnings.map((w) => el('li', {}, w)))) : null,
    ]);
  }

  profileSel.addEventListener('change', refresh);
  sheetSel.addEventListener('change', refresh);
  startBtn.addEventListener('click', async () => {
    const targetCol = +colSel.value;
    const { profile, sheetName } = current;
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
  if (card.type === 'number') {
    const v = toCellValue('number', s);
    if (v.kind !== 'n') return { state: 'warn', msg: '数値ではありません (文字として書き込みます)' };
    const { range } = card;
    if (range && ((range.min !== undefined && v.v < range.min) || (range.max !== undefined && v.v > range.max))) {
      if (card.rangeKind === 'reference') return { state: 'warn', msg: '参考範囲から外れています' };
      return { state: 'ng', msg: `${S.cards.ngMessage}${card.effect ? ` / 影響: ${card.effect}` : ''}` };
    }
    const decimals = (s.normalize('NFKC').split('.')[1] || '').length;
    if (card.digits !== null && card.digits !== undefined && decimals > card.digits) {
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
const requiredCards = () => targetCards().filter((c) => !c.optional);
const isEmptyRequired = (c) => !c.optional && judge(c, S.session.values[c.row]).state === 'empty';

function updateProgress() {
  const list = requiredCards();
  const done = list.filter((c) => judge(c, S.session.values[c.row]).state !== 'empty').length;
  const ng = targetCards().filter((c) => judge(c, S.session.values[c.row]).state === 'ng').length;
  $('#progress-text').textContent = `${done} / ${list.length}`;
  $('#progress-bar').style.width = `${list.length ? (done / list.length) * 100 : 0}%`;
  $('#ng-count').textContent = ng ? `管理値外 ${ng}` : '';
}

function applyFilter() {
  for (const node of main.querySelectorAll('.card[data-row]')) {
    const card = node._card;
    node._refresh?.();
    const j = judge(card, S.session.values[card.row]);
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
  const onInput = (e) => { setValue(card, e.target.value); refresh(); };
  const common = { class: 'field big', value, enterkeyhint: 'next', oninput: onInput, onchange: onInput };
  if (card.type === 'date') return el('input', { ...common, type: 'date' });
  if (card.type === 'time') return el('input', { ...common, type: 'time' });
  // 管理値はチップに出しているので、入力欄には重ねて出さない
  const placeholder = card.type === 'number' && !card.rangeText && !card.setText ? '数値' : '';
  if (card.type === 'number') {
    // 画面キーボードは出さず、アプリのテンキーで入力する (外付けキーボードはそのまま使える)
    return el('input', { ...common, type: 'text', inputmode: 'none', autocomplete: 'off', placeholder, dataset: { keypad: '' } });
  }
  return el('input', { ...common, type: 'text', autocomplete: 'off', placeholder, list: card.type === 'name' ? 'names' : undefined });
}

const canUsePrev = (card) => card.kind === 'item' && !card.locked && card.prev !== '' && card.type !== 'choice';

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
  if (card.prev) chips.push(el('span', { class: 'chip prev' }, `前回 ${card.prev}`));
  return el('div', { class: 'chips' }, chips);
}

function renderCard(card) {
  const status = el('div', { class: 'status' });
  const node = el('div', {
    class: `card ${card.type === 'choice' ? 'is-choice' : ''} ${card.compact ? 'compact' : ''}`, dataset: { row: card.row },
  });
  node._card = card;
  const prevBadge = el('span', { class: 'badge prev-badge' }, '前回値');
  const autoBadge = el('span', { class: 'badge prev-badge' }, '自動');
  let input;
  const refresh = () => {
    const j = judge(card, S.session.values[card.row]);
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

function focusNext(card, { scroll = true } = {}) {
  const nodes = [...main.querySelectorAll('.card[data-row]')].filter((n) => !n.hidden);
  const i = nodes.findIndex((n) => n._card === card);
  const next = nodes.slice(i + 1).find((n) => !n._card.locked);
  if (!next) {
    document.activeElement?.blur?.();
    toast('最後の項目です');
    return;
  }
  const inp = next.querySelector('input');
  if (inp) inp.focus({ preventScroll: true });
  else document.activeElement?.blur?.(); // 選択式はキーボードを閉じてボタンで選ぶ
  if (scroll || !inp) next.scrollIntoView({ block: 'center', behavior: 'smooth' });
}

function jumpToFirstEmpty() {
  const node = [...main.querySelectorAll('.card[data-row]')].find((n) => !n.hidden && isEmptyRequired(n._card));
  if (!node) { toast('未入力の項目はありません', 'ok'); return; }
  node.scrollIntoView({ block: 'center', behavior: 'smooth' });
  node.querySelector('input')?.focus({ preventScroll: true });
}

function showCards() {
  const { session, cards } = S;
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
    cards.warnings.length ? el('p', { class: 'warn' }, `設定ファイルと合わない箇所が ${cards.warnings.length} 件あります (開くときの画面で確認できます)`) : null,
    el('div', { class: 'section' },
      el('h2', { class: 'section-title' }, '基本情報'),
      el('div', { class: 'cards' }, cards.meta.map(renderCard))),
    groups.map((g) => el('div', { class: 'section' },
      el('h2', { class: 'section-title' }, g.name || '項目'),
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
// 数値入力欄 (data-keypad) にフォーカスしている間は画面下にアプリのテンキーを出す。
const keypad = (() => {
  let target = null;
  let hideTimer;
  let fixTimer;
  const label = el('div', { class: 'keypad-label' });
  const edit = (fn) => {
    const v = target.value;
    const a = target.selectionStart ?? v.length;
    const b = target.selectionEnd ?? a;
    const [next, pos] = fn(v, a, b);
    target.value = next;
    target.setSelectionRange(pos, pos);
    target.dispatchEvent(new Event('input', { bubbles: true }));
  };
  const insert = (s) => edit((v, a, b) => [v.slice(0, a) + s + v.slice(b), a + s.length]);
  const actions = {
    back: () => edit((v, a, b) => (a !== b ? [v.slice(0, a) + v.slice(b), a] : [v.slice(0, Math.max(0, a - 1)) + v.slice(a), Math.max(0, a - 1)])),
    clear: () => edit(() => ['', 0]),
    sign: () => edit((v, a) => (v.startsWith('-') ? [v.slice(1), Math.max(0, a - 1)] : [`-${v}`, a + 1])),
    // Enter と同じ (空欄なら前回値で確定して次へ)
    next: () => target.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true })),
    close: () => target.blur(),
  };
  const key = (text, act, cls = '') => el('button', { type: 'button', class: `key ${cls}`, dataset: { act } }, text);
  const pad = el('div', { class: 'keypad', hidden: true },
    el('div', { class: 'keypad-side' }, label, key('閉じる', 'close', 'close')),
    el('div', { class: 'keypad-keys' },
      key('7', '7'), key('8', '8'), key('9', '9'), key('⌫', 'back', 'fn'),
      key('4', '4'), key('5', '5'), key('6', '6'), key('クリア', 'clear', 'fn'),
      key('1', '1'), key('2', '2'), key('3', '3'), key('±', 'sign', 'fn'),
      key('0', '0'), key('.', '.'), key('次へ', 'next', 'next')));
  // pointerdown で処理し既定動作を止める (入力欄からフォーカスを外さないため)
  pad.addEventListener('pointerdown', (e) => {
    e.preventDefault();
    const act = e.target.closest('.key')?.dataset.act;
    if (!act || !target) return;
    if (actions[act]) actions[act]();
    else insert(act);
  });
  document.body.append(pad);

  const show = (input) => {
    clearTimeout(hideTimer);
    target = input;
    const card = input.closest('.card')?._card;
    label.replaceChildren(
      el('div', { class: 'keypad-title' }, card?.label ?? '', card?.unit ? el('span', { class: 'code' }, `[${card.unit}]`) : null),
      card?.prev ? el('div', { class: 'keypad-prev' }, `前回 ${card.prev}`) : null);
    pad.hidden = false;
    document.documentElement.classList.add('keypad-open');
    document.documentElement.style.setProperty('--keypad-h', `${pad.offsetHeight}px`);
    // カードへのスクロールが終わってもテンキーに隠れていれば、見える位置までずらす
    clearTimeout(fixTimer);
    fixTimer = setTimeout(() => {
      if (target !== input) return;
      const limit = window.innerHeight - pad.offsetHeight - 12;
      const r = (input.closest('.card') || input).getBoundingClientRect();
      if (r.bottom > limit) window.scrollBy({ top: r.bottom - limit, behavior: 'smooth' });
    }, 450);
  };
  const hide = () => {
    target = null;
    pad.hidden = true;
    document.documentElement.classList.remove('keypad-open');
  };
  document.addEventListener('focusin', (e) => { if (e.target.matches?.('input[data-keypad]')) show(e.target); });
  // 次の数値欄へ移るときは閉じずにそのまま使う
  document.addEventListener('focusout', () => {
    clearTimeout(hideTimer);
    hideTimer = setTimeout(() => { if (!document.activeElement?.matches?.('input[data-keypad]')) hide(); }, 50);
  });
  // 回転などで高さが変わったら、入力中の項目を見える位置に戻す
  window.addEventListener('resize', () => {
    if (pad.hidden) return;
    document.documentElement.style.setProperty('--keypad-h', `${pad.offsetHeight}px`);
    target?.closest('.card')?.scrollIntoView({ block: 'center' });
  });
  return { hide };
})();

// ---------------- 文字サイズ・画面の向き ----------------
const SIZES = [['', '標準'], ['l', '大'], ['xl', '特大']];
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

if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('./sw.js').catch((e) => console.warn('SW 登録失敗', e));
}
if (typeof DecompressionStream === 'undefined') {
  mount(el('p', { class: 'error' }, 'このブラウザは対応していません (iPadOS 16.4 以降の Safari が必要です)'));
} else {
  showHome();
}

// 開発用フック (自動テストからファイルを読み込ませる)
window.__app = { openFile, importProfile, showHome, get state() { return S; } };
