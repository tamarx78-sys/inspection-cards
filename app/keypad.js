// アプリ内のテンキー (数値入力用)。
// iPad の Safari は inputmode="decimal" でもテンキーにならない (通常のキーボードの数字段が出るだけ) ため、
// 数値入力欄 (data-keypad) では画面下にアプリのテンキーを出す。
// 数値欄にはフォーカスを当てず (readonly)、入力中の欄 (target) はここで管理する。
// iPad ではフォーカス時の自動スクロールや、なめらかなスクロールの最中のタップが効かなくなるため。
// 閉じるのは「閉じる」・ほかの入力欄へ移ったとき・選択式を押したとき・数値欄でない項目へ移ったときだけ。
//
// hooks (画面側の処理を受け取る):
//   onType(row)          キーで値を変えた (入力中の行。判定は確定まで保留する)
//   onCommit()           入力を確定した (Enter・戻る・進む・閉じる・別の欄へ移る)
//   onMove(card, back)   前後の項目へ移る
//   onActivate(input)    数値欄がタップされた (入力中にする)
//   onLeave()            入力中だった欄から離れた (絞り込みの反映など)
//   prevText(card)       前回値の表示
import { el } from './dom.js';

export function createKeypad(hooks) {
  let target = null;
  let fresh = false; // 入力中の欄に移った直後。最初のキーで値を置き換える (打ち直しやすくするため)
  let shownCount = 0; // show() の回数。click の間に入力中の欄が変わったかを見る
  let shownAtClick = 0;
  const label = el('div', { class: 'keypad-label' });
  const typeValue = (v) => {
    fresh = false;
    hooks.onType(target.closest('.card')?._card?.row ?? null);
    target.classList.remove('keypad-fresh');
    target.value = v;
    target.dispatchEvent(new Event('input', { bubbles: true }));
  };
  const actions = {
    back: () => typeValue(fresh ? '' : target.value.slice(0, -1)),
    // Enter と同じ (空欄なら前回値で確定して次へ)
    enter: () => target.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true })),
    // 戻る / 進む: 値には触れずに前後の項目へ移るだけ
    prev: () => { const card = target.closest('.card')?._card; if (card) hooks.onMove(card, true); },
    next: () => { const card = target.closest('.card')?._card; if (card) hooks.onMove(card, false); },
    close: () => hide(),
  };
  const press = (act) => {
    if (!target) return;
    // 入力欄が消えた・隠れた (絞り込みなど) ときは閉じる
    if (!target.isConnected || target.closest('.card')?.hidden) { hide(); return; }
    if (['enter', 'prev', 'next', 'close'].includes(act)) hooks.onCommit(); // 確定 (移れなかったときも判定を出す)
    if (actions[act]) actions[act]();
    else typeValue((fresh ? '' : target.value) + act);
  };
  const key = (text, act, cls = '') => el('button', { type: 'button', class: `key ${cls}`, dataset: { act } }, text);
  // 大きさを変えるつまみ: 上端で高さ、テンキー部分の左端で横幅
  const hGrip = el('div', { class: 'keypad-hgrip', title: '上下にドラッグで高さを変更 (ダブルタップで元に戻す)' });
  const wGrip = el('div', { class: 'keypad-wgrip', title: '左右にドラッグで横幅を変更 (ダブルタップで元に戻す)' });
  const panel = el('div', { class: 'keypad', hidden: true },
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
  panel.addEventListener('pointerdown', (e) => {
    e.preventDefault();
    const act = e.target.closest('.key')?.dataset.act;
    if (act) press(act);
    // キーでテンキーが閉じたら、続く click が下の数値欄に届いて開き直さないよう、少しの間 click を無視する
    if (panel.hidden) swallowUntil = Date.now() + 400;
  });
  document.body.append(panel);

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
    const cs = getComputedStyle(panel);
    const innerW = panel.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight);
    const fixedW = wGrip.offsetWidth + 2 * (parseFloat(cs.columnGap) || 0);
    const maxW = Math.max(1, (innerW - fixedW - SIDE_MIN * r) / (BASE_KEYS_W * r));
    // 高さ: テンキー全体が画面の半分を超えない。倍率 1 のときの高さから、キー 4 段ぶんの伸びで割る
    const step = 4 * BASE_KEY_H * r; // 倍率が 1 増えたときの高さの増え方
    const baseH = panel.offsetHeight - step * (applied.h - 1);
    const maxH = Math.max(1, 1 + (window.innerHeight * 0.5 - baseH) / step);
    return { maxW, maxH };
  }
  function applyKeypadSize({ save = false } = {}) {
    const { maxW, maxH } = limits();
    const h = Math.min(Math.max(size.h, 1), maxH);
    const w = Math.min(Math.max(size.w, 1), maxW);
    applied.h = h; applied.w = w;
    panel.style.setProperty('--kp-h', h.toFixed(3));
    panel.style.setProperty('--kp-w', w.toFixed(3));
    panel.style.setProperty('--kp-f', Math.min(h, w).toFixed(3)); // 文字はキーの短い辺に合わせる
    panel.classList.toggle('keypad-narrow-side', w > 1 && (maxW - w) * BASE_KEYS_W < 4); // 左側が狭いときは縦に並べる
    if (save) { try { localStorage.setItem(sizeKey(), JSON.stringify(size)); } catch {} }
    if (!panel.hidden) {
      document.documentElement.style.setProperty('--keypad-h', `${panel.offsetHeight}px`);
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
      if (Date.now() - lastTap < 350) { size[axis] = 1; applyKeypadSize({ save: true }); reveal(); lastTap = 0; return; }
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
      applyKeypadSize();
    });
    const end = () => {
      if (!start) return;
      start = null;
      grip.classList.remove('dragging');
      applyKeypadSize({ save: true });
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
    const limit = window.innerHeight - panel.offsetHeight - 12;
    const top = (document.querySelector('.toolbar')?.getBoundingClientRect().bottom ?? 0) + 8;
    if (r.bottom > limit) window.scrollBy(0, r.bottom - limit);
    else if (r.top < top) window.scrollBy(0, r.top - top);
  }
  function show(input) {
    hooks.onCommit(); // 別の欄へ移る = 確定
    if (target && target !== input) hooks.onLeave(); // 入力を終えた欄の絞り込みを反映 (フォーカス外れの代わり)
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
      card?.prev ? el('div', { class: 'keypad-prev' }, `前回 ${hooks.prevText(card)}`) : null);
    panel.hidden = false;
    document.documentElement.classList.add('keypad-open');
    applyKeypadSize();
  }
  function hide() {
    hooks.onCommit();
    if (target) hooks.onLeave();
    target?.classList.remove('keypad-target', 'keypad-fresh');
    target = null;
    panel.hidden = true;
    document.documentElement.classList.remove('keypad-open');
  }
  // 数値欄のタップ: フォーカスさせずにテンキーの対象にする (スクロールの指の動きでは反応しないよう click で)
  document.addEventListener('pointerdown', (e) => { if (e.target.matches?.('input[data-keypad]')) e.preventDefault(); });
  document.addEventListener('click', (e) => {
    shownAtClick = shownCount;
    if (Date.now() < swallowUntil) { swallowUntil = 0; e.preventDefault(); e.stopPropagation(); }
  }, true);
  document.addEventListener('click', (e) => {
    if (e.target.matches?.('input[data-keypad]')) { hooks.onActivate(e.target); reveal(); }
    // 選択式を押したら閉じる。ただし選んだ結果、次の数値欄へ移ったときは閉じない
    else if (target && e.target.closest?.('.choice') && shownCount === shownAtClick) hide();
  });
  // Tab キーなどでフォーカスが来たとき・ほかの入力欄へ移ったとき
  document.addEventListener('focusin', (e) => {
    if (e.target.matches?.('input[data-keypad]')) show(e.target);
    else if (target && !panel.contains(e.target)) hide();
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
    if (panel.hidden) return;
    applyKeypadSize();
    target?.closest('.card')?.scrollIntoView({ block: 'center' });
  });
  return { show, hide, reveal, isOpen: () => !panel.hidden, get target() { return target; } };
}
