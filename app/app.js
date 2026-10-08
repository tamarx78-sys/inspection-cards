import {
  openWorkbook, loadSheet, detectLayout, buildCards, exportWorkbook, toCellValue, colName, displayValue,
} from './xlsx.js';
import { putSession, getSession, deleteSession, listSessions, requestPersist } from './store.js';

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
const mount = (...children) => { main.replaceChildren(); add(main, children); };
const pad = (n) => String(n).padStart(2, '0');
const today = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
const nowTime = () => { const d = new Date(); return `${pad(d.getHours())}:${pad(d.getMinutes())}`; };

// 現在の作業状態
let S = null; // { session, wb, sheet, layout, cards, filter }

function toast(msg, kind = '') {
  const t = el('div', { class: `toast ${kind}` }, msg);
  document.body.append(t);
  setTimeout(() => t.classList.add('show'), 10);
  setTimeout(() => { t.classList.remove('show'); setTimeout(() => t.remove(), 300); }, 2600);
}

function setHeader(title, actions = []) {
  $('#title').textContent = title;
  const a = $('#actions');
  a.replaceChildren(...actions);
}

// ---------------- ホーム ----------------
async function showHome() {
  S = null;
  setHeader('点検カード');
  const sessions = await listSessions();
  const fileInput = el('input', {
    type: 'file', accept: '.xlsx,.xlsm,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
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
              class: 'btn ghost danger', title: '削除',
              onclick: async () => {
                if (!confirm(`「${s.fileName}」の作業データを削除しますか?\n(元の Excel ファイルには影響しません)`)) return;
                await deleteSession(s.id);
                showHome();
              },
            }, '削除'),
          );
        }),
      ),
    ),
  );
}

async function openFile(file) {
  try {
    const bytes = await file.arrayBuffer();
    const wb = await openWorkbook(bytes);
    showSetup({ fileName: file.name, bytes, wb });
  } catch (e) {
    console.error(e);
    toast(`読み込めませんでした: ${e.message}`, 'error');
  }
}

// ---------------- 対象シート・列の選択 ----------------
async function showSetup({ fileName, bytes, wb }) {
  setHeader(fileName, [el('button', { class: 'btn ghost', onclick: showHome }, '戻る')]);
  const sheetSel = el('select', { class: 'field' }, wb.sheets.map((s, i) => el('option', { value: i }, s.name)));
  const colSel = el('select', { class: 'field' });
  const info = el('div', { class: 'setup-info' });
  const startBtn = el('button', { class: 'btn primary big' }, '入力をはじめる');
  let sheet, layout;

  async function refreshSheet() {
    sheet = await loadSheet(wb, wb.sheets[+sheetSel.value]);
    try {
      layout = detectLayout(sheet);
    } catch (e) {
      layout = null;
      info.replaceChildren(el('p', { class: 'error' }, e.message));
      colSel.replaceChildren();
      startBtn.disabled = true;
      return;
    }
    startBtn.disabled = false;
    const dateRow = layout.metaRows.find((m) => m.type === 'date')?.row;
    const opts = layout.recordCols.map((c) => {
      const d = dateRow ? displayValue(sheet.get(dateRow, c)) : '';
      return el('option', { value: c }, `${colName(c)}列 ${d ? `(${d})` : ''} - 空欄のみ追記`);
    });
    opts.push(el('option', { value: layout.targetCol, selected: true }, `${colName(layout.targetCol)}列 (新しい記録)`));
    colSel.replaceChildren(...opts);
    const cards = buildCards(sheet, layout, layout.targetCol);
    const sections = new Set(cards.items.map((c) => c.section).filter(Boolean));
    info.replaceChildren(
      el('dl', {},
        el('dt', {}, '基本情報'), el('dd', {}, layout.metaRows.map((m) => m.label).join('・') || '-'),
        el('dt', {}, '点検項目'), el('dd', {}, `${cards.items.length}件 (${sections.size}区分)`),
        el('dt', {}, '列の判定'), el('dd', {}, layout.columns.map((c) => `${colName(c.c)}:${c.header || '区分'}`).join(' / ')),
        el('dt', {}, '前回の記録'), el('dd', {}, cards.prevCol ? `${colName(cards.prevCol)}列` : 'なし'),
      ),
    );
  }

  sheetSel.addEventListener('change', refreshSheet);
  startBtn.addEventListener('click', async () => {
    const targetCol = +colSel.value;
    const cards = buildCards(sheet, layout, targetCol);
    const values = {};
    // 基本情報は今日の日付・現在時刻・前回の名前で埋めておく
    for (const m of cards.meta) {
      if (m.locked) continue;
      if (m.type === 'date') values[m.row] = today();
      else if (m.type === 'time') values[m.row] = nowTime();
      else if (/名前|氏名|点検者|担当/.test(m.label)) values[m.row] = localStorage.getItem('lastName') || '';
    }
    const session = {
      id: `${Date.now()}`, fileName, bytes, sheetName: sheet.name, targetCol, values, createdAt: Date.now(),
    };
    await putSession(session);
    requestPersist();
    S = { session, wb, sheet, layout, cards, filter: 'all' };
    showCards();
  });

  mount(
    el('section', { class: 'setup' },
      wb.sheets.length > 1 ? el('label', { class: 'lbl' }, 'シート', sheetSel) : null,
      el('label', { class: 'lbl' }, '書き込む列', colSel),
      info,
      startBtn,
    ),
  );
  await refreshSheet();
}

async function resume(id) {
  try {
    const session = await getSession(id);
    const wb = await openWorkbook(session.bytes);
    const meta = wb.sheets.find((s) => s.name === session.sheetName) || wb.sheets[0];
    const sheet = await loadSheet(wb, meta);
    const layout = detectLayout(sheet);
    const cards = buildCards(sheet, layout, session.targetCol);
    S = { session, wb, sheet, layout, cards, filter: 'all' };
    showCards();
  } catch (e) {
    console.error(e);
    toast(`再開できませんでした: ${e.message}`, 'error');
  }
}

// ---------------- カード入力 ----------------
let saveTimer = null;
function scheduleSave() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => putSession(S.session).catch((e) => toast(`保存失敗: ${e.message}`, 'error')), 300);
}
// ページを離れる/バックグラウンドに回る前に確実に保存
const flushSave = () => { if (S && saveTimer) { clearTimeout(saveTimer); saveTimer = null; putSession(S.session); } };
document.addEventListener('visibilitychange', () => document.visibilityState === 'hidden' && flushSave());
window.addEventListener('pagehide', flushSave);

function judge(card, raw) {
  if (card.locked) return { state: 'locked' };
  const s = String(raw ?? '').trim();
  if (!s) return { state: 'empty' };
  if (card.type === 'number') {
    const v = toCellValue('number', s);
    if (v.kind !== 'n') return { state: 'warn', msg: '数値ではありません (文字として書き込みます)' };
    const { range } = card;
    if (range && ((range.min !== undefined && v.v < range.min) || (range.max !== undefined && v.v > range.max))) {
      return { state: 'ng', msg: `管理値外${card.contact ? ` → ${card.contact}へ連絡` : ''}` };
    }
  }
  return { state: 'ok' };
}

function editableCards() {
  return [...S.cards.meta, ...S.cards.items].filter((c) => !c.locked);
}

function updateProgress() {
  const list = editableCards();
  const done = list.filter((c) => judge(c, S.session.values[c.row]).state !== 'empty').length;
  const ng = S.cards.items.filter((c) => judge(c, S.session.values[c.row]).state === 'ng').length;
  $('#progress-text').textContent = `${done} / ${list.length}`;
  $('#progress-bar').style.width = `${list.length ? (done / list.length) * 100 : 0}%`;
  $('#ng-count').textContent = ng ? `管理値外 ${ng}` : '';
}

function applyFilter() {
  for (const node of main.querySelectorAll('.card[data-row]')) {
    const card = node._card;
    const j = judge(card, S.session.values[card.row]);
    const show = S.filter === 'all' || (S.filter === 'empty' && j.state === 'empty') || (S.filter === 'ng' && j.state === 'ng');
    node.hidden = !show;
  }
  for (const sec of main.querySelectorAll('.section')) {
    sec.hidden = ![...sec.querySelectorAll('.card')].some((c) => !c.hidden);
  }
}

function renderInput(card, onInput) {
  const value = S.session.values[card.row] ?? '';
  if (card.locked) return el('div', { class: 'locked-value' }, card.existing, el('span', { class: 'badge' }, '記入済'));
  const common = { class: 'field big', value, enterkeyhint: 'next', oninput: onInput, onchange: onInput };
  if (card.type === 'date') return el('input', { ...common, type: 'date' });
  if (card.type === 'time') return el('input', { ...common, type: 'time' });
  // 前回値がある項目は、空のまま Enter で前回値を入れられることを案内する
  const placeholder = canUsePrev(card) ? `Enter で前回値 ${card.prev}` : card.rangeText || (card.type === 'number' ? '数値' : '');
  if (card.type === 'number') {
    return el('input', { ...common, type: 'text', inputmode: 'decimal', autocomplete: 'off', placeholder });
  }
  return el('input', { ...common, type: 'text', autocomplete: 'off', placeholder, list: card.kind === 'meta' ? 'names' : undefined });
}

const canUsePrev = (card) => card.kind === 'item' && !card.locked && card.prev !== '';

function renderCard(card) {
  const status = el('div', { class: 'status' });
  const node = el('div', { class: 'card', dataset: { row: card.row } });
  node._card = card;
  const prevBadge = el('span', { class: 'badge prev-badge' }, '前回値');
  const refresh = () => {
    const j = judge(card, S.session.values[card.row]);
    node.dataset.state = j.state;
    status.textContent = j.msg || '';
    prevBadge.hidden = !S.session.fromPrev[card.row];
  };
  const onInput = (e) => {
    S.session.values[card.row] = e.target.value;
    delete S.session.fromPrev[card.row];
    if (card.kind === 'meta' && /名前|氏名|点検者|担当/.test(card.label)) localStorage.setItem('lastName', e.target.value);
    refresh();
    updateProgress();
    scheduleSave();
  };
  const input = renderInput(card, onInput);
  if (input.tagName === 'INPUT') {
    input.addEventListener('keydown', (e) => {
      if (e.key !== 'Enter' || e.isComposing) return;
      e.preventDefault();
      // 空のまま Enter = 前回値で確定 (ありがちな「前回と同じ」を 1 打で)
      if (!input.value.trim() && canUsePrev(card)) {
        input.value = card.prev;
        S.session.values[card.row] = card.prev;
        S.session.fromPrev[card.row] = true;
        refresh();
        updateProgress();
        scheduleSave();
      }
      focusNext(card);
    });
    input.addEventListener('blur', applyFilterLater);
  }
  add(node, [
    el('div', { class: 'card-head' },
      el('div', { class: 'card-title' }, card.label, card.code ? el('span', { class: 'code' }, card.code) : null),
      prevBadge,
      el('div', { class: 'card-ref' }, card.ref),
    ),
    card.kind === 'item'
      ? el('div', { class: 'chips' },
        card.rangeText ? el('span', { class: 'chip range' }, `管理値 ${card.rangeText}`) : null,
        card.contact ? el('span', { class: 'chip' }, card.contact) : null,
        card.extra.map((x) => el('span', { class: 'chip' }, `${x.label} ${x.value}`)),
        card.prev ? el('span', { class: 'chip prev' }, `前回 ${card.prev}`) : null,
      )
      : null,
    input,
    status,
  ]);
  refresh();
  return node;
}

// 入力中にカードが消えないよう、フィルタ反映はフォーカスが外れてから
let filterTimer;
function applyFilterLater() {
  clearTimeout(filterTimer);
  filterTimer = setTimeout(() => {
    if (!main.contains(document.activeElement) || document.activeElement === document.body) applyFilter();
  }, 400);
}

function focusNext(card) {
  const nodes = [...main.querySelectorAll('.card[data-row]')].filter((n) => !n.hidden);
  const i = nodes.findIndex((n) => n._card === card);
  for (const n of nodes.slice(i + 1)) {
    const inp = n.querySelector('input');
    if (inp) { inp.focus(); n.scrollIntoView({ block: 'center', behavior: 'smooth' }); return; }
  }
  document.activeElement.blur();
  toast('最後の項目です');
}

function jumpToFirstEmpty() {
  const node = [...main.querySelectorAll('.card[data-row]')].find((n) => !n.hidden && n.dataset.state === 'empty');
  if (!node) { toast('未入力の項目はありません', 'ok'); return; }
  node.scrollIntoView({ block: 'center', behavior: 'smooth' });
  node.querySelector('input')?.focus({ preventScroll: true });
}

function showCards() {
  const { session, cards } = S;
  session.fromPrev ??= {}; // 前回値で確定した行 (旧バージョンの保存データには無い)
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

  const groups = [];
  for (const c of cards.items) {
    const last = groups[groups.length - 1];
    if (last && last.name === c.section) last.cards.push(c);
    else groups.push({ name: c.section, cards: [c] });
  }

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
    el('div', { class: 'section' },
      el('h2', { class: 'section-title' }, '基本情報'),
      el('div', { class: 'cards' }, cards.meta.map(renderCard))),
    groups.map((g) => el('div', { class: 'section' },
      el('h2', { class: 'section-title' }, g.name || '項目'),
      el('div', { class: 'cards' }, g.cards.map(renderCard)))),
    el('div', { class: 'bottom-space' }),
  );
  updateProgress();
  applyFilter();
}

// ---------------- 書き出し ----------------
async function showExport() {
  flushSave();
  const { session, wb, sheet, cards } = S;
  const empty = editableCards().filter((c) => judge(c, session.values[c.row]).state === 'empty');
  const ng = cards.items.filter((c) => judge(c, session.values[c.row]).state === 'ng');
  const fromPrev = cards.items.filter((c) => session.fromPrev?.[c.row] && session.values[c.row] === c.prev);
  let result;
  try {
    result = await exportWorkbook(wb, sheet, cards, session.values);
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

  const dlg = el('div', { class: 'modal-bg' },
    el('div', { class: 'modal' },
      el('h2', {}, 'Excel に書き出し'),
      el('p', {}, `${colName(cards.targetCol)}列に ${result.written.length} セル書き込みます。`),
      empty.length ? el('p', { class: 'warn' }, `未入力が ${empty.length} 件あります (空欄のまま書き出します)。`) : null,
      fromPrev.length ? el('p', {}, `前回値のまま確定: ${fromPrev.length} 件`) : null,
      ng.length ? el('div', { class: 'ng-list' },
        el('p', { class: 'error' }, `管理値外が ${ng.length} 件あります:`),
        el('ul', {}, ng.map((c) => el('li', {}, `${c.section ? c.section + ' / ' : ''}${c.label}: ${session.values[c.row]} (管理値 ${c.rangeText})`)))) : null,
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
applySize(document.documentElement.dataset.size || '');

// ヘッダーの高さ (文字サイズ・回転で変わる) に合わせて、ツールバーの固定位置をずらす
function updateTopbarH() {
  document.documentElement.style.setProperty('--topbar-h', `${$('.topbar').getBoundingClientRect().height}px`);
}
new ResizeObserver(updateTopbarH).observe($('.topbar'));
window.addEventListener('resize', updateTopbarH);
updateTopbarH();

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
window.__app = { openFile, showHome, get state() { return S; } };
