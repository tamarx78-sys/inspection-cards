// 入力途中のデータを IndexedDB に保存する (オフラインでも消えないように)
const DB_NAME = 'inspection-cards';
const STORE = 'sessions';

let dbPromise;
function db() {
  dbPromise ??= new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE, { keyPath: 'id' });
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return dbPromise;
}

async function tx(mode, fn) {
  const d = await db();
  return new Promise((resolve, reject) => {
    const t = d.transaction(STORE, mode);
    const req = fn(t.objectStore(STORE));
    t.oncomplete = () => resolve(req?.result);
    t.onerror = () => reject(t.error);
    t.onabort = () => reject(t.error);
  });
}

export const putSession = (s) => tx('readwrite', (st) => st.put({ ...s, updatedAt: Date.now() }));
export const getSession = (id) => tx('readonly', (st) => st.get(id));
export const deleteSession = (id) => tx('readwrite', (st) => st.delete(id));
export async function listSessions() {
  const all = await tx('readonly', (st) => st.getAll());
  return (all || []).sort((a, b) => b.updatedAt - a.updatedAt);
}

/** ブラウザに「消さないで」と依頼する (iPad の容量逼迫時の自動削除対策) */
export async function requestPersist() {
  try {
    if (navigator.storage?.persist && !(await navigator.storage.persisted())) {
      return await navigator.storage.persist();
    }
    return true;
  } catch {
    return false;
  }
}
