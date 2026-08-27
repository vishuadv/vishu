const UX = (() => {
  function applySavedTheme() {
    try {
      document.body.classList.remove('dark');
      localStorage.setItem('swiftfin_theme', 'light');
    } catch (_) {}
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

  function addThemeToggle() {
    return;
  }

  function showOfflineBanner() {
    if (document.getElementById('swiftfin-offline')) return;
    const bar = document.createElement('div');
    bar.id = 'swiftfin-offline';
    bar.className = 'offline-banner';
    bar.textContent = 'Offline: changes will sync when internet returns.';
    document.body.appendChild(bar);
    document.body.classList.add('has-offline-banner');
  }

  function hideOfflineBanner() {
    const bar = document.getElementById('swiftfin-offline');
    if (bar) bar.remove();
    document.body.classList.remove('has-offline-banner');
  }

  function initOfflineIndicator() {
    if (!navigator.onLine) showOfflineBanner();
    window.addEventListener('offline', showOfflineBanner);
    window.addEventListener('online', () => {
      hideOfflineBanner();
      if (window.OfflineForms && typeof window.OfflineForms?.replayQueued === 'function') {
        window.OfflineForms.replayQueued();
      }
    });
  }

  function renderStaticDashboard() {
    const host = document.getElementById('quick-dashboard');
    if (!host) return;

    // Static dashboard (no app.py endpoint required)
    const data = {
      todayCollections: 0,
      pendingLoans: 0,
      overduePayments: 0,
      totalOutstanding: 0
    };

    host.innerHTML = `
      <div class="grid">
        <div class="stat"><div class="label">Today's collections</div><div class="value">₹${fmt(data.todayCollections)}</div></div>
        <div class="stat"><div class="label">Pending loans</div><div class="value">${data.pendingLoans}</div></div>
        <div class="stat"><div class="label">Overdue payments</div><div class="value">${data.overduePayments}</div></div>
        <div class="stat"><div class="label">Total outstanding</div><div class="value">₹${fmt(data.totalOutstanding)}</div></div>
      </div>
    `;
  }

  function fmt(n) {
    try { return new Intl.NumberFormat('en-IN').format(Number(n || 0)); }
    catch { return String(n || 0); }
  }

  function init() {
    applySavedTheme();
    removeThemeToggleIfPresent();
    addThemeToggle();
    initOfflineIndicator();
    renderStaticDashboard();
  }

  return { init };
})();

document.addEventListener('DOMContentLoaded', () => UX.init());
