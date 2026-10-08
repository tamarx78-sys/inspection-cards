// 最小限の ZIP 読み書き (xlsx 用)。外部ライブラリなし。
// 展開/圧縮はブラウザ標準の (De)CompressionStream('deflate-raw') を使う (Safari 16.4+)。
// 書き戻し時、変更していないエントリは元の圧縮データをそのままコピーするので劣化しない。

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(bytes) {
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

async function pipe(bytes, stream) {
  const out = new Response(new Blob([bytes]).stream().pipeThrough(stream));
  return new Uint8Array(await out.arrayBuffer());
}
const inflateRaw = (b) => pipe(b, new DecompressionStream('deflate-raw'));
const deflateRaw = (b) => pipe(b, new CompressionStream('deflate-raw'));

const utf8 = new TextDecoder('utf-8');

/** ZIP を読み込み、エントリ一覧を返す。中身は read() で遅延展開。 */
export function readZip(buffer) {
  const u8 = new Uint8Array(buffer);
  const dv = new DataView(buffer);
  // End of central directory を末尾から探す
  let eocd = -1;
  for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) {
    if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('ZIP 形式ではありません (xlsx ではない可能性)');
  const count = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  const entries = new Map();
  for (let i = 0; i < count; i++) {
    if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('ZIP の中央ディレクトリが壊れています');
    const flags = dv.getUint16(p + 8, true);
    const method = dv.getUint16(p + 10, true);
    const time = dv.getUint16(p + 12, true);
    const date = dv.getUint16(p + 14, true);
    const crc = dv.getUint32(p + 16, true);
    const compSize = dv.getUint32(p + 20, true);
    const size = dv.getUint32(p + 24, true);
    const nameLen = dv.getUint16(p + 28, true);
    const extraLen = dv.getUint16(p + 30, true);
    const commentLen = dv.getUint16(p + 32, true);
    const localOff = dv.getUint32(p + 42, true);
    const nameBytes = u8.subarray(p + 46, p + 46 + nameLen);
    const name = utf8.decode(nameBytes);
    // ローカルヘッダから実データ位置を求める
    const lNameLen = dv.getUint16(localOff + 26, true);
    const lExtraLen = dv.getUint16(localOff + 28, true);
    const dataStart = localOff + 30 + lNameLen + lExtraLen;
    const comp = u8.subarray(dataStart, dataStart + compSize);
    entries.set(name, { name, nameBytes, flags, method, time, date, crc, compSize, size, comp });
    p += 46 + nameLen + extraLen + commentLen;
  }
  return {
    entries,
    has: (name) => entries.has(name),
    async read(name) {
      const e = entries.get(name);
      if (!e) return null;
      if (e.method === 0) return e.comp.slice();
      if (e.method === 8) return inflateRaw(e.comp);
      throw new Error(`未対応の圧縮方式です: ${e.method}`);
    },
    async readText(name) {
      const b = await this.read(name);
      return b ? utf8.decode(b) : null;
    },
  };
}

/**
 * 元の ZIP をベースに、replaced (Map<name, Uint8Array>) の内容を差し替えた ZIP を作る。
 * 差し替えないエントリは圧縮データを再利用する。
 */
export async function writeZip(zip, replaced) {
  const enc = new TextEncoder();
  const parts = [];
  const central = [];
  let offset = 0;
  const now = new Date();
  const dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
  const dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();

  const names = [...zip.entries.keys()];
  for (const n of replaced.keys()) if (!zip.entries.has(n)) names.push(n);

  for (const name of names) {
    let e;
    if (replaced.has(name)) {
      const raw = replaced.get(name);
      const comp = await deflateRaw(raw);
      e = {
        nameBytes: enc.encode(name), flags: 0x0800, method: 8,
        time: dosTime, date: dosDate, crc: crc32(raw),
        compSize: comp.length, size: raw.length, comp,
      };
    } else {
      // データディスクリプタ(bit3)は使わず、サイズをヘッダに直接書く
      e = { ...zip.entries.get(name) };
      e.flags &= ~0x0008;
    }
    const local = new Uint8Array(30 + e.nameBytes.length);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, 20, true);
    lv.setUint16(6, e.flags, true);
    lv.setUint16(8, e.method, true);
    lv.setUint16(10, e.time, true);
    lv.setUint16(12, e.date, true);
    lv.setUint32(14, e.crc, true);
    lv.setUint32(18, e.compSize, true);
    lv.setUint32(22, e.size, true);
    lv.setUint16(26, e.nameBytes.length, true);
    lv.setUint16(28, 0, true);
    local.set(e.nameBytes, 30);
    parts.push(local, e.comp);

    const cd = new Uint8Array(46 + e.nameBytes.length);
    const cv = new DataView(cd.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint16(4, 20, true);
    cv.setUint16(6, 20, true);
    cv.setUint16(8, e.flags, true);
    cv.setUint16(10, e.method, true);
    cv.setUint16(12, e.time, true);
    cv.setUint16(14, e.date, true);
    cv.setUint32(16, e.crc, true);
    cv.setUint32(20, e.compSize, true);
    cv.setUint32(24, e.size, true);
    cv.setUint16(28, e.nameBytes.length, true);
    cv.setUint32(42, offset, true);
    cd.set(e.nameBytes, 46);
    central.push(cd);
    offset += local.length + e.comp.length;
  }
  const cdSize = central.reduce((s, c) => s + c.length, 0);
  const end = new Uint8Array(22);
  const ev = new DataView(end.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, central.length, true);
  ev.setUint16(10, central.length, true);
  ev.setUint32(12, cdSize, true);
  ev.setUint32(16, offset, true);
  return new Blob([...parts, ...central, end], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}
