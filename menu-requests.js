/* Branch -> admin notifications for sales menus that have no BOM yet.
   Branch accounts press "แจ้ง Admin" on page 5; admins see the requests on page 2 and add the BOM or remove the request.
   Data lives in public.bom_menu_requests (see supabase-menu-requests.sql), never in the shared master document,
   so notifying cannot overwrite recipes or collide with other branches. */
(function () {
  'use strict';
  const $ = id => document.getElementById(id);
  const E = v => (typeof esc === 'function' ? esc(v) : String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])));
  const state = window.menuRequests = { list: [], loaded: false, loading: null, error: '' };
  const role = () => (typeof bomCloud !== 'undefined' && bomCloud.member ? bomCloud.member.role : null);
  const HINT = 'ยังไม่ได้ติดตั้งระบบแจ้งสินค้า ให้ผู้ดูแลรันไฟล์ supabase-menu-requests.sql ใน Supabase ก่อน';
  const friendly = e => {
    const m = String((e && e.message) || e);
    return /bom_menu_requests|bom_report_menus|bom_clear_menu_requests|PGRST|42P01|schema cache|Could not find/i.test(m) ? HINT : m;
  };
  const dmy = iso => { const d = new Date(iso); if (isNaN(d)) return ''; const z = n => String(n).padStart(2, '0'); return z(d.getDate()) + '/' + z(d.getMonth() + 1) + '/' + d.getFullYear() + ' ' + z(d.getHours()) + ':' + z(d.getMinutes()); };
  const repaint = () => {
    try { if (typeof renderMissingImportBom === 'function') renderMissingImportBom(); } catch (e) { /* informational */ }
    try { if (typeof window.renderBranchBomUsage === 'function') window.renderBranchBomUsage(true); } catch (e) { /* informational */ }
  };

  async function load() {
    if (typeof bomCloud === 'undefined' || !bomCloud.member) return;
    if (state.loading) return state.loading;
    state.loading = (async () => {
      try {
        state.list = await cloudRequest('/rest/v1/bom_menu_requests?select=*&order=requested_at.desc&limit=2000');
        state.error = ''; state.loaded = true;
      } catch (e) { state.error = friendly(e); }
      finally { state.loading = null; }
    })();
    await state.loading;
    repaint();
  }

  // Branch accounts: send menus to the admin. items = [{code,name,qty,reason}]
  async function report(items, context) {
    if (role() !== 'branch') throw Error('เฉพาะบัญชีสาขาเท่านั้นที่แจ้ง Admin ได้');
    const clean = items.filter(i => i && i.code).slice(0, 500).map(i => ({
      code: String(i.code), name: String(i.name || ''), qty: Number(i.qty) || 0, reason: String(i.reason || ''),
      file: String((context && context.file) || ''), date: String((context && context.date) || '')
    }));
    if (!clean.length) return 0;
    try { await cloudRequest('/rest/v1/rpc/bom_report_menus', { method: 'POST', body: JSON.stringify({ items: clean }) }); }
    catch (e) { throw Error(friendly(e)); }
    const now = new Date().toISOString(), branch = bomCloud.member.branch;
    for (const i of clean) {
      state.list = state.list.filter(r => !(r.code === i.code && r.branch === branch));
      state.list.push({ code: i.code, branch, name: i.name, qty: i.qty, source_file: i.file, report_date: i.date, reason: i.reason, requested_at: now });
    }
    return clean.length;
  }

  // Admin: remove requests (all branches) for the given menu codes.
  async function clear(codes) {
    if (role() !== 'admin') throw Error('เฉพาะ Admin เท่านั้น');
    try { await cloudRequest('/rest/v1/rpc/bom_clear_menu_requests', { method: 'POST', body: JSON.stringify({ codes }) }); }
    catch (e) { throw Error(friendly(e)); }
    const set = new Set(codes);
    state.list = state.list.filter(r => !set.has(r.code));
  }

  state.report = report; state.clear = clear; state.load = load; state.reload = load;
  state.isReported = code => role() === 'branch' && state.list.some(r => r.code === code && r.branch === bomCloud.member.branch);

  // Admin view on page 2 (replaces the old "all unmatched sales rows" list).
  window.renderMissingImportBom = function renderMissingImportBom() {
    const panel = $('missingImportBomPanel');
    if (!panel) return;
    const isAdmin = role() === 'admin';
    panel.hidden = !isAdmin;
    if (!isAdmin) return;
    const recipes = appState.recipes || [];
    const byCode = new Map();
    for (const r of state.list) {
      let g = byCode.get(r.code);
      if (!g) { g = { code: r.code, name: r.name || r.code, qty: 0, at: '', branches: [], reasons: new Set() }; byCode.set(r.code, g); }
      g.qty += Number(r.qty) || 0;
      if (!g.at || r.requested_at > g.at) g.at = r.requested_at;
      if (r.name) g.name = r.name;
      g.branches.push({ id: r.branch, qty: Number(r.qty) || 0 });
      if (r.reason) g.reasons.add(r.reason);
    }
    const all = [...byCode.values()];
    const NO_RECIPE = /^(ยังไม่มีเมนูในสูตร|มีเมนูแล้ว แต่ยังไม่มีวัตถุดิบ)/;
    for (const g of all) g.problem = [...g.reasons].some(t => !NO_RECIPE.test(t));
    // A request disappears once the menu has ingredients, except "recipe not ready" reports that the admin must clear.
    const done = all.filter(g => !g.problem && (recipes.find(x => x.menuCode === g.code) || {}).items?.length);
    const pending = all.filter(g => !done.includes(g)).sort((a, b) => b.qty - a.qty || a.code.localeCompare(b.code, undefined, { numeric: true }));
    const query = (($('missingImportBomSearch') || {}).value || '').trim().toLocaleLowerCase();
    const shown = pending.filter(g => (g.code + ' ' + g.name).toLocaleLowerCase().includes(query));
    const summary = $('missingImportBomSummary');
    summary.textContent = state.error
      ? state.error
      : !state.loaded ? 'กำลังโหลดรายการที่สาขาแจ้ง…'
        : `สาขาแจ้งมา ${pending.length} รายการที่ยังไม่มีสูตร` + (done.length ? ` · ทำสูตรแล้ว ${done.length} รายการ (ซ่อนไว้)` : '') + (query ? ` · แสดง ${shown.length} รายการ` : '');
    const bn = id => (workflow.branches && workflow.branches[id]) || id;
    const extra = g => [...g.reasons].filter(t => !/^(ยังไม่มีเมนูในสูตร|มีเมนูแล้ว แต่ยังไม่มีวัตถุดิบ)/.test(t)).join(' / ');
    const noRecipe = code => { const r = recipes.find(x => x.menuCode === code); return r ? (r.items && r.items.length ? 'มีสูตรแล้ว แต่สาขาแจ้งว่ายังคิดตามสูตรไม่ได้' : 'มีเมนูแล้ว แต่ยังไม่มีวัตถุดิบในสูตร') : 'ยังไม่มีเมนูในสูตร BOM'; };
    $('missingImportBomBody').innerHTML = shown.map(g => `<tr>
      <td class="p-3">${E(g.code)}</td>
      <td class="p-3"><span class="font-semibold">${E(g.name)}</span>
        <div class="text-xs text-slate-500 mt-1">${E(noRecipe(g.code))}${extra(g) ? ' · ' + E(extra(g)) : ''}</div>
        <div class="text-xs text-slate-500 mt-1">แจ้งโดย: ${g.branches.map(b => E(bn(b.id)) + ' (' + E(b.id) + ')').join(', ')}</div></td>
      <td class="p-3 whitespace-nowrap"><b>${g.qty.toLocaleString('th-TH', { maximumFractionDigits: 2 })}</b><div class="text-xs text-slate-500">แจ้งล่าสุด ${E(dmy(g.at))}</div></td>
      <td class="p-3 whitespace-nowrap"><button type="button" data-code="${E(g.code)}" data-name="${E(g.name)}" onclick="addBomFromImport(this.dataset.code,this.dataset.name,this)" class="border border-emerald-300 text-emerald-700 rounded-lg px-3 py-2 whitespace-nowrap">เพิ่มสูตร</button>
        <button type="button" data-code="${E(g.code)}" onclick="removeMenuRequest(this.dataset.code,this)" class="ml-2 border border-rose-200 text-rose-600 rounded-lg px-3 py-2 whitespace-nowrap" title="ลบรายการที่สาขาแจ้งนี้ออก (ไม่กระทบสูตรและยอดขาย)">ลบ</button></td></tr>`).join('')
      || `<tr><td colspan="4" class="p-3">${state.error ? E(state.error) : !state.loaded ? 'กำลังโหลด…' : pending.length ? 'ไม่พบรหัสหรือชื่อที่ค้นหา' : 'ยังไม่มีสาขาแจ้งสินค้าที่ไม่มีสูตร BOM (สาขากดแจ้งได้ที่หน้า 5 หลังอัปโหลดไฟล์ยอดขาย)'}</td></tr>`;
  };

  window.removeMenuRequest = async function removeMenuRequest(code, button) {
    if (role() !== 'admin' || !code) return;
    if (typeof appConfirm === 'function' && !await appConfirm('ลบรายการแจ้ง ' + code + ' ออกจากรายการหรือไม่?\nสาขาที่แจ้งไว้จะกดแจ้งใหม่ได้อีกครั้ง การลบนี้ไม่กระทบสูตรและยอดขาย')) return;
    if (button) button.disabled = true;
    try { await clear([code]); }
    catch (e) { if (typeof appAlert === 'function') appAlert('ลบไม่สำเร็จ: ' + e.message); }
    finally { if (button) button.disabled = false; }
    window.renderMissingImportBom();
  };

  // Load after every successful cloud load (fire and forget so sign-in is not delayed).
  const originalLoad = window.cloudLoad;
  if (typeof originalLoad === 'function') {
    window.cloudLoad = async function () {
      const result = await originalLoad.apply(this, arguments);
      state.loaded = false;
      load().catch(() => {});
      return result;
    };
  }
})();
