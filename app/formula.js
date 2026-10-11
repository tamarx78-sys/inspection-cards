// 自動計算の行で使う、簡単な Excel 式の計算 (eval は使わない)。
// 対応: 数値、セル参照 (A1)、範囲 (A1:B2)、+ - * / ^、括弧、MAX MIN SUM AVERAGE ABS ROUND
import { colNum } from './xlsx.js';

const FUNCS = {
  MAX: (xs) => (xs.length ? Math.max(...xs) : 0),
  MIN: (xs) => (xs.length ? Math.min(...xs) : 0),
  SUM: (xs) => xs.reduce((a, b) => a + b, 0),
  AVERAGE: (xs) => (xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : NaN),
  ABS: (xs) => Math.abs(xs[0]),
  ROUND: (xs) => { const k = 10 ** (xs[1] || 0); return Math.round(xs[0] * k) / k; },
};

function tokenize(f) {
  const re = /\s*(?:(\d+\.?\d*(?:[eE][-+]?\d+)?|\.\d+)|\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?|([A-Z][A-Z0-9.]*)\s*\(|([-+*/^(),]))/gy;
  const out = [];
  let m;
  let pos = 0;
  const src = f.replace(/^=/, '').trim();
  while (pos < src.length) {
    re.lastIndex = pos;
    m = re.exec(src);
    if (!m) throw new Error(`式を読めません: ${src.slice(pos)}`);
    pos = re.lastIndex;
    if (m[1] !== undefined) out.push({ t: 'num', v: +m[1] });
    else if (m[2] !== undefined) {
      out.push({ t: 'ref', r1: +m[3], c1: colNum(m[2]), r2: +(m[5] ?? m[3]), c2: colNum(m[4] ?? m[2]) });
    } else if (m[6] !== undefined) out.push({ t: 'fn', v: m[6] });
    else out.push({ t: 'op', v: m[7] });
  }
  return out;
}

/**
 * 式を計算する。get(r, c) はセルの数値 (空・数値でなければ null) を返す。
 * 必要なセルが空なら null、計算できない式なら例外
 */
export function evaluate(formula, get) {
  const toks = tokenize(formula);
  let i = 0;
  let missing = false;
  const peek = () => toks[i];
  const take = (v) => { const t = toks[i]; if (v && t?.v !== v) throw new Error('式の形が正しくありません'); i++; return t; };
  const cellsOf = (t) => {
    const xs = [];
    for (let r = Math.min(t.r1, t.r2); r <= Math.max(t.r1, t.r2); r++) {
      for (let c = Math.min(t.c1, t.c2); c <= Math.max(t.c1, t.c2); c++) {
        const v = get(r, c);
        if (v === null || v === undefined) { if (t.r1 === t.r2 && t.c1 === t.c2) missing = true; continue; }
        xs.push(v);
      }
    }
    return xs;
  };
  // 関数の引数: 範囲は中の数値すべて
  const arg = () => {
    const t = peek();
    if (t?.t === 'ref' && (t.r1 !== t.r2 || t.c1 !== t.c2)) { i++; const xs = cellsOf(t); if (!xs.length) missing = true; return xs; }
    return [expr()];
  };
  function primary() {
    const t = take();
    if (!t) throw new Error('式が途中で終わっています');
    if (t.t === 'num') return t.v;
    if (t.t === 'ref') { const xs = cellsOf(t); return xs[0] ?? 0; }
    if (t.t === 'fn') {
      const fn = FUNCS[t.v];
      if (!fn) throw new Error(`未対応の関数: ${t.v}`);
      const xs = [];
      if (peek()?.v !== ')') { xs.push(...arg()); while (peek()?.v === ',') { take(','); xs.push(...arg()); } }
      take(')');
      return fn(xs);
    }
    if (t.v === '(') { const v = expr(); take(')'); return v; }
    if (t.v === '-') return -unary();
    if (t.v === '+') return unary();
    throw new Error('式の形が正しくありません');
  }
  function unary() { return primary(); }
  function power() { let v = unary(); while (peek()?.v === '^') { take(); v **= unary(); } return v; }
  function term() {
    let v = power();
    while (peek()?.v === '*' || peek()?.v === '/') { const op = take().v; const w = power(); v = op === '*' ? v * w : v / w; }
    return v;
  }
  function expr() {
    let v = term();
    while (peek()?.v === '+' || peek()?.v === '-') { const op = take().v; const w = term(); v = op === '+' ? v + w : v - w; }
    return v;
  }
  const v = expr();
  if (i < toks.length) throw new Error('式の形が正しくありません');
  if (missing || !Number.isFinite(v)) return null;
  return v;
}
