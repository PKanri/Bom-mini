/* Branch -> admin notifications for sales menus that have no BOM yet.
   Branch accounts press "แจ้ง Admin" on page 5; admins see the requests on page 2 and add the BOM or remove the request.
   Data lives in public.bom_menu_requests (see supabase-menu-requests.sql), never in the shared master document,
   so notifying cannot overwrite recipes or collide with other branches. */
(function () {
  'use strict';
  const $ = id => document.getElementById(id);
  const E = v => (typeof esc === 'function' ? esc(v) : String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])));
  const MAX_IMAGES = 3;
  const state = window.menuRequests = { list: [], loaded: false, loading: null, error: '', drafts: new Map(), maxImages: MAX_IMAGES };
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

  // Codes that ANY branch already attached photos to (codes only, no images). Needs bom_menu_codes_with_images() from the SQL file.
  state.imagedCodes = new Set();
  async function loadImagedCodes() {
    if (role() !== 'branch') { state.imagedCodes = new Set(); return; }
    try {
      const codes = await cloudRequest('/rest/v1/rpc/bom_menu_codes_with_images', { method: 'POST', body: '{}' });
      state.imagedCodes = new Set((Array.isArray(codes) ? codes : []).map(c => typeof c === 'string' ? c : Object.values(c || {})[0]).filter(Boolean));
    } catch (e) { /* SQL not updated yet: behave as before */ }
  }
  state.othersHaveImages = code => state.imagedCodes.has(code) && state.imageCount(code) === 0;
  async function load() {
    if (typeof bomCloud === 'undefined' || !bomCloud.member) return;
    if (state.loading) return state.loading;
    state.loading = (async () => {
      try {
        state.list = await cloudRequest('/rest/v1/bom_menu_requests?select=code,branch,name,qty,source_file,report_date,reason,requested_at,image_count&order=requested_at.desc&limit=2000');
        state.error = ''; state.loaded = true;
        await loadImagedCodes();
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
      file: String((context && context.file) || ''), date: String((context && context.date) || ''),
      ...(Array.isArray(i.images) ? { images: i.images.slice(0, MAX_IMAGES) } : {})
    }));
    if (!clean.length) return 0;
    try { await cloudRequest('/rest/v1/rpc/bom_report_menus', { method: 'POST', body: JSON.stringify({ items: clean }) }); }
    catch (e) { throw Error(friendly(e)); }
    const now = new Date().toISOString(), branch = bomCloud.member.branch;
    for (const i of clean) {
      const old = state.list.find(r => r.code === i.code && r.branch === branch);
      state.list = state.list.filter(r => r !== old);
      state.list.push({ code: i.code, branch, name: i.name, qty: i.qty, source_file: i.file, report_date: i.date, reason: i.reason, requested_at: now,
        image_count: Array.isArray(i.images) ? i.images.length : (old ? old.image_count || 0 : 0) });
      if (Array.isArray(i.images)) state.drafts.delete(i.code);
    }
    await loadImagedCodes();
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


  // ---- Recipe photos -------------------------------------------------------
  // Shrink a photo in the browser (max 1600 px, JPEG) so a recipe sheet stays readable but small.
  async function compressImage(file) {
    if (!file || !/^image\/(jpeg|png|webp|heic|heif)$/i.test(file.type || '') && !/\.(jpe?g|png|webp)$/i.test(file.name || '')) throw Error('รองรับเฉพาะไฟล์รูป JPG, PNG, WEBP (ไฟล์ PDF ให้แคปเป็นรูปก่อน)');
    if (file.size > 25 * 1024 * 1024) throw Error('ไฟล์รูปใหญ่เกิน 25 MB');
    let bitmap;
    try { bitmap = await createImageBitmap(file); } catch (e) { throw Error('เปิดไฟล์รูปไม่ได้ ลองบันทึกเป็น JPG แล้วแนบใหม่'); }
    let scale = Math.min(1, 1600 / Math.max(bitmap.width, bitmap.height)), quality = 0.82, out = '';
    for (let i = 0; i < 8; i++) {
      const canvas = document.createElement('canvas');
      canvas.width = Math.max(1, Math.round(bitmap.width * scale)); canvas.height = Math.max(1, Math.round(bitmap.height * scale));
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
      out = canvas.toDataURL('image/jpeg', quality);
      if (out.length <= 450000) break;
      quality = Math.max(0.6, quality - 0.07); scale *= 0.88;
    }
    if (bitmap.close) bitmap.close();
    if (out.length > 690000) throw Error('รูปใหญ่เกินไปแม้ย่อแล้ว ลองถ่ายใหม่ให้เห็นเฉพาะใบสูตร');
    return out;
  }

  // Branch: replace the photos of an already reported menu ([] clears).
  async function setImages(code, images) {
    if (role() !== 'branch') throw Error('เฉพาะบัญชีสาขา');
    try { await cloudRequest('/rest/v1/rpc/bom_set_menu_images', { method: 'POST', body: JSON.stringify({ menu_code: code, new_images: images }) }); }
    catch (e) { throw Error(friendly(e)); }
    const row = state.list.find(r => r.code === code && r.branch === bomCloud.member.branch);
    if (row) row.image_count = images.length;
    await loadImagedCodes();
  }

  // Photos are loaded only when someone opens them. Admin gets every branch that attached some; a branch gets its own.
  async function fetchImages(code) {
    const rows = await cloudRequest('/rest/v1/bom_menu_requests?select=branch,name,images&image_count=gt.0&code=eq.' + encodeURIComponent(code));
    return rows.map(r => ({ branch: r.branch, name: r.name, images: Array.isArray(r.images) ? r.images : [] }));
  }

  // Viewer: click an image to switch between "fit" and "actual size" for reading small print.
  function showImages(title, groups) {
    const old = document.getElementById('menuImageViewer'); if (old) old.remove();
    const dlg = document.createElement('dialog');
    dlg.id = 'menuImageViewer'; dlg.className = 'ponn-dialog';
    dlg.style.cssText = 'width:min(96vw,1100px);max-width:96vw;max-height:92vh;padding:16px;overflow:auto';
    const body = groups.map(g => `<section style="margin-top:12px">${g.label ? '<div style="font-weight:700;margin-bottom:6px">' + E(g.label) + '</div>' : ''}${g.images.map((src, i) => `<img src="${E(src)}" alt="${E(title)} รูปที่ ${i + 1}" data-zoom="0" style="display:block;max-width:100%;margin:0 0 10px;border:1px solid var(--b3,#c6d1e3);border-radius:8px;background:#fff;cursor:zoom-in">`).join('')}</section>`).join('') || '<p>ไม่มีรูปแนบ</p>';
    dlg.innerHTML = `<div style="display:flex;justify-content:space-between;align-items:center;gap:12px"><h2 style="margin:0;font-size:17px;font-weight:700">🖼️ ${E(title)}</h2><button type="button" id="menuImageClose" style="border:1px solid var(--b3,#c6d1e3);border-radius:8px;padding:6px 14px">ปิด</button></div><p style="font-size:12px;opacity:.7;margin:6px 0 0">กดที่รูปเพื่อดูขนาดจริง / ย่อให้พอดีหน้าจอ</p>${body}`;
    document.body.appendChild(dlg);
    dlg.addEventListener('click', e => {
      if (e.target.id === 'menuImageClose' || e.target === dlg) { dlg.close(); return; }
      const img = e.target.closest && e.target.closest('img[data-zoom]');
      if (img) { const on = img.dataset.zoom === '1'; img.dataset.zoom = on ? '0' : '1'; img.style.maxWidth = on ? '100%' : 'none'; img.style.cursor = on ? 'zoom-in' : 'zoom-out'; }
    });
    dlg.addEventListener('close', () => dlg.remove());
    dlg.showModal();
  }
  window.viewMenuImages = async function viewMenuImages(code, name, button) {
    if (button) button.disabled = true;
    try {
      const isAdmin = role() === 'admin';
      const bn = id => (typeof workflow !== 'undefined' && workflow.branches && workflow.branches[id]) || id;
      const groups = (await fetchImages(code)).filter(g => g.images.length).map(g => ({ label: isAdmin ? 'สาขา ' + bn(g.branch) + ' (' + g.branch + ')' : '', images: g.images }));
      showImages((code + ' ' + (name || '')).trim(), groups);
    } catch (e) { if (typeof appAlert === 'function') appAlert('เปิดรูปไม่สำเร็จ: ' + friendly(e)); }
    finally { if (button) button.disabled = false; }
  };
  state.compressImage = compressImage; state.setImages = setImages; state.fetchImages = fetchImages; state.showImages = showImages;
  state.imageCount = code => { const r = state.list.find(x => x.code === code && role() === 'branch' && x.branch === bomCloud.member.branch); return r ? Number(r.image_count) || 0 : 0; };

  // Once a menu has ingredients in the BOM its photos are no longer needed: the admin's page clears them automatically.
  async function purgeImages(codes) {
    if (state.purging || state.purgeFailed || !codes.length) return;
    state.purging = true;
    try {
      await cloudRequest('/rest/v1/rpc/bom_purge_menu_images', { method: 'POST', body: JSON.stringify({ codes }) });
      const set = new Set(codes);
      for (const r of state.list) if (set.has(r.code)) r.image_count = 0;
      state.lastPurged = codes.length;
    } catch (e) { state.purgeFailed = true; state.purgeError = friendly(e); }
    finally { state.purging = false; }
    try { window.renderMissingImportBom(); } catch (e) { /* informational */ }
  }
  state.purgeImages = purgeImages;
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
      g.images = (g.images || 0) + (Number(r.image_count) || 0);
      if (r.reason) g.reasons.add(r.reason);
    }
    const all = [...byCode.values()];
    const NO_RECIPE = /^(ยังไม่มีเมนูในสูตร|มีเมนูแล้ว แต่ยังไม่มีวัตถุดิบ)/;
    for (const g of all) g.problem = [...g.reasons].some(t => !NO_RECIPE.test(t));
    // A request disappears once the menu has ingredients, except "recipe not ready" reports that the admin must clear.
    const done = all.filter(g => !g.problem && (recipes.find(x => x.menuCode === g.code) || {}).items?.length);
    const purgeCodes = done.filter(g => g.images > 0).map(g => g.code);
    if (purgeCodes.length && state.loaded && recipes.length) purgeImages(purgeCodes);
    const pending = all.filter(g => !done.includes(g)).sort((a, b) => b.qty - a.qty || a.code.localeCompare(b.code, undefined, { numeric: true }));
    const query = (($('missingImportBomSearch') || {}).value || '').trim().toLocaleLowerCase();
    const shown = pending.filter(g => (g.code + ' ' + g.name).toLocaleLowerCase().includes(query));
    const summary = $('missingImportBomSummary');
    summary.textContent = state.error
      ? state.error
      : !state.loaded ? 'กำลังโหลดรายการที่สาขาแจ้ง…'
        : `สาขาแจ้งมา ${pending.length} รายการที่ยังไม่มีสูตร` + (done.length ? ` · ทำสูตรแล้ว ${done.length} รายการ (ซ่อนไว้ และลบรูปที่สาขาส่งมาแล้ว)` : '') + (state.purgeError ? ' · ลบรูปอัตโนมัติไม่สำเร็จ: ' + state.purgeError : '') + (query ? ` · แสดง ${shown.length} รายการ` : '');
    const bn = id => (workflow.branches && workflow.branches[id]) || id;
    const extra = g => [...g.reasons].filter(t => !/^(ยังไม่มีเมนูในสูตร|มีเมนูแล้ว แต่ยังไม่มีวัตถุดิบ)/.test(t)).join(' / ');
    const noRecipe = code => { const r = recipes.find(x => x.menuCode === code); return r ? (r.items && r.items.length ? 'มีสูตรแล้ว แต่สาขาแจ้งว่ายังคิดตามสูตรไม่ได้' : 'มีเมนูแล้ว แต่ยังไม่มีวัตถุดิบในสูตร') : 'ยังไม่มีเมนูในสูตร BOM'; };
    $('missingImportBomBody').innerHTML = shown.map(g => `<tr>
      <td class="p-3">${E(g.code)}</td>
      <td class="p-3"><span class="font-semibold">${E(g.name)}</span>
        <div class="text-xs text-slate-500 mt-1">${E(noRecipe(g.code))}${extra(g) ? ' · ' + E(extra(g)) : ''}</div>
        <div class="text-xs text-slate-500 mt-1">แจ้งโดย: ${g.branches.map(b => E(bn(b.id)) + ' (' + E(b.id) + ')').join(', ')}</div></td>
      <td class="p-3 whitespace-nowrap"><b>${g.qty.toLocaleString('th-TH', { maximumFractionDigits: 2 })}</b><div class="text-xs text-slate-500">แจ้งล่าสุด ${E(dmy(g.at))}</div></td>
      <td class="p-3 whitespace-nowrap">${g.images ? `<button type="button" data-code="${E(g.code)}" data-name="${E(g.name)}" onclick="viewMenuImages(this.dataset.code,this.dataset.name,this)" class="border rounded-lg px-3 py-2 whitespace-nowrap mr-2" style="border-color:var(--acc-b300,#93b8fb);color:var(--acc-t700,#1d4ed8)">🖼️ ดูรูปสูตร (${g.images})</button>` : '<span class="text-xs text-slate-400 mr-2">ไม่มีรูปแนบ</span>'}<button type="button" data-code="${E(g.code)}" data-name="${E(g.name)}" onclick="addBomFromImport(this.dataset.code,this.dataset.name,this)" class="border border-emerald-300 text-emerald-700 rounded-lg px-3 py-2 whitespace-nowrap">เพิ่มสูตร</button>
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
