(function () {
  // ---------- helpers ----------
  function toast(msg) {
    const el = document.createElement('div');
    el.textContent = msg;
    el.style.cssText = 'position:fixed;right:14px;bottom:14px;background:#111827;color:#fff;padding:10px 12px;border-radius:10px;z-index:99999;max-width:70vw;';
    document.body.appendChild(el);
    setTimeout(() => el.remove(), 2500);
  }

  function removeThemeToggleIfPresent() {
    try {
      document.querySelectorAll('button.btn').forEach((btn) => {
        const style = (btn.getAttribute('style') || '').toLowerCase().replace(/\s+/g, '');
        if (style.includes('position:fixed') && style.includes('right:14px') && style.includes('bottom:14px') && style.includes('z-index:99997')) {
          btn.remove();
        }
      });
    } catch (_) {}
  }

  // ---------- (7) Dark mode ----------
  function initTheme() {
    // Dark mode disabled by request: keep default light UI and do not inject toggle button.
    try {
      document.body.classList.remove('pwa-dark');
      localStorage.setItem('swiftfin_theme', 'light');
      const meta = document.querySelector('meta[name="theme-color"]');
      if (meta) meta.setAttribute('content', '#1a73e8');
    } catch (_) {}
  }

  // ---------- (8) Offline indicator ----------
  function showOfflineBar() {
    if (document.getElementById('pwa-offline')) return;
    const bar = document.createElement('div');
    bar.id = 'pwa-offline';
    bar.className = 'pwa-offline';
    bar.textContent = 'Offline: changes will sync when internet returns.';
    document.body.appendChild(bar);
    document.body.classList.add('pwa-has-offline');
  }
  function hideOfflineBar() {
    const bar = document.getElementById('pwa-offline');
    if (bar) bar.remove();
    document.body.classList.remove('pwa-has-offline');
  }
  function initOfflineUI() {
    if (!navigator.onLine) showOfflineBar();
    window.addEventListener('offline', showOfflineBar);
    window.addEventListener('online', () => {
      hideOfflineBar();
      replayQueue(); // try sync immediately when back online
    });
  }

  // ---------- (4) Lazy load images/backgrounds ----------
  function initLazyLoad() {
    // Use IntersectionObserver where available. [web:155]
    const supportsIO = ('IntersectionObserver' in window);

    document.querySelectorAll('img[data-src]').forEach((img) => {
      const load = () => { img.src = img.dataset.src; img.removeAttribute('data-src'); };
      if (!supportsIO) return load();

      const io = new IntersectionObserver((entries, obs) => {
        entries.forEach((e) => {
          if (!e.isIntersecting) return;
          load();
          obs.unobserve(img);
        });
      }, { rootMargin: '200px 0px' });

      io.observe(img);
    });

    document.querySelectorAll('[data-bg]').forEach((el) => {
      const load = () => { el.style.backgroundImage = `url(${el.dataset.bg})`; el.removeAttribute('data-bg'); };
      if (!supportsIO) return load();

      const io = new IntersectionObserver((entries, obs) => {
        entries.forEach((e) => {
          if (!e.isIntersecting) return;
          load();
          obs.unobserve(el);
        });
      }, { rootMargin: '200px 0px' });

      io.observe(el);
    });
  }

  // ---------- (6) Performance logs (no backend) ----------
  function initPerf() {
    window.addEventListener('load', () => {
      const t = performance.timing;
      const loadMs = t.loadEventEnd - t.navigationStart;
      const domMs = t.domContentLoadedEventEnd - t.navigationStart;
      console.log('SwiftFin perf(ms): load=', loadMs, 'dom=', domMs);
      const el = document.getElementById('perf-load-time');
      if (el && Number.isFinite(loadMs)) el.textContent = (loadMs / 1000).toFixed(2) + 's';
    });
  }

  // ---------- (1) Auto-logout (front-end only) ----------
  function initAutoLogout() {
    const LOGOUT_URL = '/logout'; // change if your logout route differs
    const INACTIVITY_MS = 5 * 60 * 1000;
    const WARNING_MS = 4 * 60 * 1000;
    let tWarn, tLogout;

    function removeWarn() {
      const el = document.getElementById('pwa-warn');
      if (el) el.remove();
    }

    function showWarn() {
      if (document.getElementById('pwa-warn')) return;
      const wrap = document.createElement('div');
      wrap.id = 'pwa-warn';
      wrap.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.65);z-index:99997;display:flex;align-items:center;justify-content:center;padding:16px;';
      wrap.innerHTML = `
        <div style="background:#fff;color:#111827;border-radius:12px;max-width:420px;width:100%;padding:16px">
          <div style="font-weight:900;margin-bottom:6px">Session expiring</div>
          <div style="color:#374151;margin-bottom:12px">No activity detected. Continue to stay logged in.</div>
          <button id="pwa-stay" class="pwa-btn" style="background:#16a34a">Stay logged in</button>
        </div>`;
      document.body.appendChild(wrap);
      wrap.querySelector('#pwa-stay').addEventListener('click', reset);
    }

    function logout() {
      removeWarn();
      window.location.href = LOGOUT_URL;
    }

    function reset() {
      removeWarn();
      clearTimeout(tWarn);
      clearTimeout(tLogout);
      tWarn = setTimeout(showWarn, WARNING_MS);
      tLogout = setTimeout(logout, INACTIVITY_MS);
    }

    ['mousemove','mousedown','keypress','scroll','touchstart','click'].forEach((ev) =>
      document.addEventListener(ev, reset, true)
    );
    reset();
  }

  // ---------- Offline form queue (no app.py changes) ----------
  // Mark any form with: data-offline="1"
  const DB_NAME = 'swiftfin_pwa';
  const DB_VERSION = 1;
  const STORE = 'request_queue';

  function openDB() {
    return new Promise((resolve, reject) => {
      const r = indexedDB.open(DB_NAME, DB_VERSION);
      r.onupgradeneeded = () => {
        const db = r.result;
        if (!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE, { keyPath: 'id', autoIncrement: true });
      };
      r.onsuccess = () => resolve(r.result);
      r.onerror = () => reject(r.error);
    });
  }

  async function addQueue(item) {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      const req = tx.objectStore(STORE).add(item);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  async function getAllQueue() {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readonly');
      const req = tx.objectStore(STORE).getAll();
      req.onsuccess = () => resolve(req.result || []);
      req.onerror = () => reject(req.error);
    });
  }

  async function deleteQueue(id) {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).delete(id);
      tx.oncomplete = () => resolve(true);
      tx.onerror = () => reject(tx.error);
    });
  }

  async function replayQueue() {
    if (!navigator.onLine) return;
    const items = await getAllQueue();
    if (!items.length) return;

    let synced = 0;
    for (const item of items) {
      try {
        const res = await fetch(item.url, {
          method: item.method,
          headers: item.headers,
          body: item.body
        });
        if (res.ok) {
          await deleteQueue(item.id);
          synced += 1;
        }
      } catch {
        // keep for next time
      }
    }
    if (synced) toast(`Synced ${synced} offline request(s).`);
  }

  function serializeForm(form) {
    const fd = new FormData(form);
    return new URLSearchParams(fd).toString();
  }

  function bindOfflineForms() {
    document.querySelectorAll('form[data-offline="1"]').forEach((form) => {
      if (form.dataset.offlineBound === '1') return;
      form.dataset.offlineBound = '1';

      form.addEventListener('submit', async (e) => {
        e.preventDefault();

        const url = form.getAttribute('action') || window.location.pathname;
        const method = (form.getAttribute('method') || 'POST').toUpperCase();
        const body = serializeForm(form);

        // If online, try normal submit via fetch
        if (navigator.onLine) {
          try {
            const res = await fetch(url, {
              method,
              headers: { 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' },
              body
            });
            if (res.ok) {
              toast('Saved.');
              setTimeout(() => window.location.reload(), 250);
              return;
            }
          } catch {
            // fall through to queue
          }
        }

        // Offline or failed network: queue
        await addQueue({
          url,
          method,
          headers: { 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' },
          body,
          created_at: Date.now()
        });
        toast('Offline: saved, will sync when online.');
      });
    });
  }

  // ---------- Init ----------
  document.addEventListener('DOMContentLoaded', () => {
    initTheme();
    removeThemeToggleIfPresent();
    initOfflineUI();
    initLazyLoad();
    initPerf();
    bindOfflineForms();
  });

  window.addEventListener('online', replayQueue);
  window.swiftfinReplayQueue = replayQueue; // optional manual call
})();
