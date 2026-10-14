// 画面の部品を作る小さなヘルパー (外部ライブラリを使わないため)
export const $ = (sel, root = document) => root.querySelector(sel);

/** 要素を作る。attrs: class / dataset / on〜 (イベント) / その他の属性。children は配列のネスト・null を含んでよい */
export const el = (tag, attrs = {}, ...children) => {
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
export const add = (parent, children) => {
  for (const c of [children].flat(Infinity)) if (c !== null && c !== undefined && c !== false) parent.append(c);
  return parent;
};

/** 画面下に短いお知らせを出す。kind: '' / 'ok' / 'error' */
export function toast(msg, kind = '') {
  const t = el('div', { class: `toast ${kind}` }, msg);
  document.body.append(t);
  setTimeout(() => t.classList.add('show'), 10);
  setTimeout(() => { t.classList.remove('show'); setTimeout(() => t.remove(), 300); }, 2600);
}
