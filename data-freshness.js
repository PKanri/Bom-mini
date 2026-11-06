/* Page 4: shows the latest day of data uploaded by the central admin, next to the branch label. */
(function () {
  'use strict';
  const TYPES = ['SALES', 'STOCK_IN', 'STOCK_OUT', 'EXTRA_IN'];
  const fmt = value => String(value || '').replace(/^(\d{4})-(\d{2})-(\d{2})$/, '$3/$2/$1');
  const later = (a, b) => (!a || (b && b > a) ? b : a);
  function addDays(value, n) { const d = new Date(value + 'T00:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }

  // Latest uploaded day of one source for one branch (null when nothing is uploaded).
  function latestUpload(type, branch) {
    let best = null;
    for (const b of Object.values(workflow.batches || {})) {
      if (b.type !== type) continue;
      const excluded = new Set(b.excludedDates || []);
      if (b.coverage) {
        for (const key of b.coverage) {
          let pair; try { pair = JSON.parse(key); } catch { continue; }
          if (pair[0] === branch && !excluded.has(pair[1])) best = later(best, pair[1]);
        }
      } else if ((b.branches || []).includes(branch) && b.end) {
        let d = b.end;
        for (let i = 0; i < 62 && excluded.has(d) && d > (b.start || d); i++) d = addDays(d, -1);
        if (!excluded.has(d)) best = later(best, d);
      }
    }
    // Branch accounts cannot read other branches' movement rows; central coverage lists the uploaded days.
    if (type !== 'SALES' && typeof bomCloud !== 'undefined' && bomCloud.member?.role === 'branch') {
      for (const c of bomCloud.coverage || []) if (c.kind === type) best = later(best, c.end_date);
    }
    return best;
  }

  window.renderDataFreshness = function renderDataFreshness(branch) {
    const el = document.getElementById('adminDataDate');
    if (!el || typeof workflow === 'undefined') return;
    const codes = branch && branch !== 'ALL' ? [branch] : Object.keys(workflow.branches || {});
    let latest = null;
    for (const code of codes) for (const type of TYPES) latest = later(latest, latestUpload(type, code));
    el.textContent = latest ? 'Admin อัปเดตล่าสุดคือ ' + fmt(latest) : 'Admin ยังไม่ได้อัปเดตข้อมูล';
  };
})();
