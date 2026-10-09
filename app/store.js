// 入力途中のデータと設定ファイルを IndexedDB に保存する (オフラインでも消えないように)
const DB_NAME = 'inspection-cards';
const SESSIONS = 'sessions';
const PROFILES = 'profiles';

let dbPromise;
function db() {
  dbPromise ??= new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 2);
    req.onupgradeneeded = () => {
      const d = req.result;
      if (!d.objectStoreNames.contains(SESSIONS)) d.createObjectStore(SESSIONS, { keyPath: 'id' });
      if (!d.objectStoreNames.contains(PROFILES)) d.createObjectStore(PROFILES, { keyPath: 'id' });
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return dbPromise;
}

async function tx(store, mode, fn) {
  const d = await db();
  return new Promise((resolve, reject) => {
    const t = d.transaction(store, mode);
    const req = fn(t.objectStore(store));
    t.oncomplete = () => resolve(req?.result);
    t.onerror = () => reject(t.error);
    t.onabort = () => reject(t.error);
  });
}

export const putSession = (s) => tx(SESSIONS, 'readwrite', (st) => st.put({ ...s, updatedAt: Date.now() }));
export const getSession = (id) => tx(SESSIONS, 'readonly', (st) => st.get(id));
export const deleteSession = (id) => tx(SESSIONS, 'readwrite', (st) => st.delete(id));
export async function listSessions() {
  const all = await tx(SESSIONS, 'readonly', (st) => st.getAll());
  return (all || []).sort((a, b) => b.updatedAt - a.updatedAt);
}

// 設定ファイル: 同じ名前のものは上書き
export const putProfile = (p) => tx(PROFILES, 'readwrite', (st) => st.put({ ...p, id: p.name, importedAt: Date.now() }));
export const deleteProfile = (id) => tx(PROFILES, 'readwrite', (st) => st.delete(id));
export async function listProfiles() {
  return (await tx(PROFILES, 'readonly', (st) => st.getAll())) || [];
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
