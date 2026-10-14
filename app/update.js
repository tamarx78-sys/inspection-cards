// 最新版の確認。開いたとき・アプリに戻ってきたときに、オンラインなら新しい版がないか確認する。
// 新しい版は Service Worker が裏で読み込み、すぐ有効になる (sw.js の skipWaiting)。
// 画面は古いまま動いているので「再読み込み」を案内する (入力途中の内容は保存してから読み込み直す)。
// ホームの「アプリの版」表示 (versionLine) と確認 (checkUpdate) を返す。
//   beforeReload()  再読み込みの前に呼ぶ (入力内容の保存)
//   hasWork()       入力中の作業があるか (案内の文言を変える)
import { el } from './dom.js';

export function initUpdateCheck({ beforeReload, hasWork }) {
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
    const reload = () => { beforeReload(); location.reload(); };
    document.body.append(el('div', { class: 'update-banner' },
      el('span', {}, `新しい版${version ? ` (${version})` : ''}を読み込みました。`,
        hasWork() ? '入力内容は保存されます。' : ''),
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
  return { versionLine, checkUpdate };
}
