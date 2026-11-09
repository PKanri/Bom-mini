/* Page 5: after the branch uploads its sales file, show which menus were deducted by the BOM
   (sorted by quantity sold, high to low) and which sold menus were NOT counted, with the reason. */
(function () {
  'use strict';
  let tab = 'used';
  let lastModel = null;
  const $ = id => document.getElementById(id);
  const num = v => Number(v).toLocaleString('th-TH', { maximumFractionDigits: 2 });
  const amt = v => Number(v).toLocaleString('th-TH', { maximumFractionDigits: 3 });
  const E = v => (typeof esc === 'function' ? esc(v) : String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])));

  const style = document.createElement('style');
  style.textContent = `
    #branchBomUsagePanel{border:1px solid var(--b2,#dde4ef);border-radius:16px;padding:16px;background:var(--surface-2,#f6f8fc)}
    #branchBomUsagePanel .bu-tabs{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0}
    #branchBomUsagePanel .bu-tab{border:1px solid var(--b3,#c6d1e3);border-radius:999px;padding:6px 14px;font-size:13px;background:var(--surface,#fff);color:var(--t700,#2a3c58);cursor:pointer}
    #branchBomUsagePanel .bu-tab[aria-pressed="true"]{background:var(--acc-solid,#2563eb);border-color:var(--acc-solid,#2563eb);color:#fff;font-weight:700}
    #branchBomUsagePanel .bu-wrap{overflow:auto;max-height:440px;border:1px solid var(--b2,#dde4ef);border-radius:12px;background:var(--surface,#fff)}
    #branchBomUsagePanel table{width:100%;border-collapse:collapse;font-size:13px}
    #branchBomUsagePanel th{position:sticky;top:0;background:var(--surface-3,#edf1f8);color:var(--t800,#182841);text-align:left;padding:10px;white-space:nowrap}
    #branchBomUsagePanel td{padding:10px;border-top:1px solid var(--b2,#dde4ef);vertical-align:top;color:var(--t800,#182841)}
    #branchBomUsagePanel td.r,#branchBomUsagePanel th.r{text-align:right;white-space:nowrap}
    #branchBomUsagePanel tr.bu-row{cursor:pointer}
    #branchBomUsagePanel tr.bu-row:hover td{background:var(--row-hover,#f5f8fe)}
    #branchBomUsagePanel tr.bu-detail td{background:var(--surface-2,#f6f8fc)}
    #branchBomUsagePanel .bu-sub{color:var(--t500,#5f7190);font-size:12px}
    #branchBomUsagePanel .bu-badge{display:inline-block;padding:2px 9px;border-radius:999px;font-size:12px;font-weight:700;white-space:nowrap}
    #branchBomUsagePanel .bu-red{background:var(--badge-short-bg,#ffe4e6);color:var(--badge-short-t,#be123c)}
    #branchBomUsagePanel .bu-amb{background:var(--badge-pend-bg,#fef3c7);color:var(--badge-pend-t,#92400e)}
    #branchBomUsagePanel .bu-blue{background:var(--badge-sur-bg,#dbeafe);color:var(--badge-sur-t,#1d4ed8)}
    #branchBomUsagePanel .bu-send{border:1px solid var(--acc-b300,#93b8fb);background:var(--acc-bg-50,#eff5ff);color:var(--acc-t700,#1d4ed8);border-radius:10px;padding:5px 12px;font-size:13px;font-weight:700;cursor:pointer;white-space:nowrap}
    #branchBomUsagePanel .bu-ghost{background:transparent;font-weight:600}
    #branchBomUsagePanel .bu-link{border:0;background:none;padding:0 2px;font-size:12px;color:var(--acc-t700,#1d4ed8);text-decoration:underline;cursor:pointer}
    #branchBomUsagePanel .bu-link:disabled{opacity:.5;cursor:default}
    #branchBomUsagePanel .bu-send:disabled{opacity:.5;cursor:default}
    #branchBomUsagePanel .bu-sum{font-size:14px;color:var(--t800,#182841)}
    #branchBomUsagePanel .bu-bar{height:8px;border-radius:999px;background:var(--surface-4,#e1e8f3);overflow:hidden;margin:8px 0 2px}
    #branchBomUsagePanel .bu-bar span{display:block;height:100%;background:var(--acc-solid,#2563eb)}
  `;
  document.head.appendChild(style);

  function analyse() {
    const file = branchCheckState && branchCheckState.files && branchCheckState.files.SALES;
    if (!file) return null;
    const tracked = new Map(orderedCountEntries());
    const menus = new Map();
    for (const r of file.rows) {
      if (!r.code || !r.qty) continue;
      let m = menus.get(r.code);
      if (!m) { m = { code: r.code, name: r.name || r.code, qty: 0, xl: [] }; menus.set(r.code, m); }
      m.qty += r.qty; m.xl.push(r.excelRow);
    }
    const used = [], notUsed = [];
    for (const m of menus.values()) {
      const recipe = appState.recipes.find(r => r.menuCode === m.code);
      if (recipe && recipe.menuName) m.name = recipe.menuName;
      if (!recipe) { notUsed.push({ ...m, kind: 'none', reason: 'ยังไม่มีเมนูในสูตร BOM' }); continue; }
      if (!recipe.items || !recipe.items.length) { notUsed.push({ ...m, kind: 'none', reason: 'มีเมนูแล้ว แต่ยังไม่มีวัตถุดิบในสูตร' }); continue; }
      const deductions = new Map(), problems = new Set();
      let untrackedItems = 0;
      for (const item of recipe.items) {
        if (item.unitReview) { problems.add('สูตรเดิมต่างหน่วยจากต้นแบบ รอตรวจปริมาณต่อหน่วยขาย'); continue; }
        const mat = appState.materials.find(x => x.id === item.materialId);
        if (!mat || !Number.isFinite(item.netQty) || item.netQty <= 0 || !Number.isFinite(item.yieldRate) || item.yieldRate <= 0 || item.yieldRate > 100) {
          problems.add('สูตร BOM หรือ Yield ไม่ถูกต้อง' + (mat ? ' (' + mat.name + ')' : '')); continue;
        }
        const s = tracked.get(mat.code);
        if (!s) { untrackedItems++; continue; }
        const amount = m.qty * item.netQty / (item.yieldRate / 100);
        const d = deductions.get(mat.code) || { code: mat.code, name: s.name || mat.name, unit: s.unit, amount: 0 };
        d.amount += amount; deductions.set(mat.code, d);
      }
      const list = [...deductions.values()].sort((a, b) => b.amount - a.amount);
      if (list.length) used.push({ ...m, items: list, problems: [...problems], total: recipe.items.length });
      else if (problems.size) notUsed.push({ ...m, kind: 'problem', reason: [...problems].join(' / ') });
      else notUsed.push({ ...m, kind: 'untracked', reason: 'มีสูตรครบ แต่ไม่มีวัตถุดิบที่อยู่ในรายการติดตามของหน้านี้ (' + untrackedItems + ' รายการ)' });
    }
    used.sort((a, b) => b.qty - a.qty || a.code.localeCompare(b.code, undefined, { numeric: true }));
    notUsed.sort((a, b) => b.qty - a.qty || a.code.localeCompare(b.code, undefined, { numeric: true }));
    const qUsed = used.reduce((t, x) => t + x.qty, 0), qNot = notUsed.reduce((t, x) => t + x.qty, 0);
    return { file, used, notUsed, qUsed, qNot };
  }

  function xlRows(list) {
    const s = list.slice(0, 5).join(', ');
    return s + (list.length > 5 ? ' …+' + (list.length - 5) : '');
  }

  let query = '';
  let notice = '';
  const isBranch = () => typeof bomCloud !== 'undefined' && bomCloud.member && bomCloud.member.role === 'branch';
  const reportable = m => m.kind === 'none' || m.kind === 'problem';
  const reported = code => !!(window.menuRequests && window.menuRequests.isReported(code));
  const nameOf = code => { const m = lastModel && lastModel.notUsed.find(x => x.code === code); return m ? m.name : ''; };
  const rowOf = code => (lastModel && lastModel.notUsed.find(x => x.code === code)) || null;

  // ---- "แจ้ง Admin" popup: attach recipe photos (optional), then confirm ----
  let modal = null;
  function closeModal() {
    if (!modal) return;
    const el = modal.el; modal = null;
    try { if (el.open) el.close(); } catch (e) { /* already closed */ }
    el.remove();
  }
  function paintModal() {
    if (!modal) return;
    const m = modal, rq = window.menuRequests, max = rq.maxImages, dis = m.busy ? ' disabled' : '';
    const edit = m.mode === 'edit';
    const thumbs = m.images.map((src, i) => `<div style="position:relative"><img src="${E(src)}" alt="รูปที่ ${i + 1}" style="width:112px;height:112px;object-fit:cover;border-radius:10px;border:1px solid var(--b3,#c6d1e3);background:#fff;display:block"><button type="button" data-mrm="${i}"${dis} aria-label="ลบรูปที่ ${i + 1}" style="position:absolute;top:-8px;right:-8px;width:24px;height:24px;border-radius:50%;border:1px solid var(--b3,#c6d1e3);background:var(--surface,#fff);color:var(--t800,#182841);font-size:12px;cursor:pointer">✕</button></div>`).join('');
    const add = m.images.length < max ? `<button type="button" data-madd${dis} style="width:112px;height:112px;border:2px dashed var(--b3,#c6d1e3);border-radius:10px;background:var(--surface-2,#f6f8fc);color:var(--t700,#2a3c58);font-size:13px;cursor:pointer;line-height:1.5">📎<br>เลือกรูป /<br>ถ่ายรูป</button>` : '';
    $('bmBody').innerHTML = `
      <h2 style="margin:0 0 12px;font-size:18px;font-weight:700">${edit ? '🖼️ แก้ไขรูปสูตรที่แจ้งไว้' : '🔔 แจ้ง Admin: ยังไม่มีสูตร BOM'}</h2>
      <div style="border:1px solid var(--b2,#dde4ef);border-radius:12px;padding:10px 12px;background:var(--surface-2,#f6f8fc);margin-bottom:14px">
        <div style="font-size:12px;opacity:.7">${E(m.code)}</div>
        <div style="font-weight:700">${E(m.name)}</div>
        <div style="font-size:13px;margin-top:2px">จำนวนขายในไฟล์ <b>${num(m.qty)}</b>${m.kind === 'none' ? '' : ' · <span style="opacity:.8">' + E(m.reason) + '</span>'}</div>
      </div>
      <div style="font-weight:700">แนบรูปสูตร <span style="font-weight:400;opacity:.7">(ไม่บังคับ · สูงสุด ${max} รูป)</span></div>
      <div style="font-size:12px;opacity:.7;margin:2px 0 8px">ถ่ายหรือเลือกรูปใบสูตรของเมนูนี้ให้เห็นวัตถุดิบและปริมาณชัดเจน เหมือนใบสูตรมาตรฐานที่ใช้อยู่</div>
      <div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:8px;min-height:30px">${m.loading ? '<span style="font-size:13px;opacity:.7">กำลังโหลดรูปที่แนบไว้…</span>' : thumbs + add}</div>
      <div role="status" style="font-size:13px;min-height:20px;${m.error ? 'color:var(--rose-t600,#e11d48)' : 'opacity:.8'}">${E(m.error || m.status || '')}</div>
      <div style="display:flex;justify-content:flex-end;gap:10px;margin-top:14px">
        <button type="button" data-mcancel${dis} style="color:inherit;background:transparent;border:1px solid var(--b3,#c6d1e3);border-radius:10px;padding:9px 18px">ยกเลิก</button>
        <button type="button" data-mok class="ponn-dialog-ok" style="background:var(--acc-solid,#2563eb);color:#fff;border:1px solid var(--acc-solid,#2563eb);border-radius:10px;padding:9px 18px;font-weight:700"${dis || (m.loading ? ' disabled' : '')}>${m.busy ? 'กำลังส่ง…' : edit ? 'บันทึกรูป' : 'ยืนยันการแจ้ง Admin'}</button>
      </div>`;
  }
  async function openNotify(code, mode) {
    const rq = window.menuRequests, row = rowOf(code);
    if (!rq || !row || modal) return;
    const el = document.createElement('dialog');
    el.className = 'ponn-dialog';
    el.style.cssText = 'width:min(94vw,560px);max-width:94vw;max-height:92vh;overflow:auto;padding:20px';
    el.innerHTML = '<div id="bmBody"></div><input id="bmFile" type="file" accept="image/*" multiple hidden>';
    document.body.appendChild(el);
    modal = { el, code, name: row.name, qty: row.qty, reason: row.reason, kind: row.kind, mode: mode || 'new', images: [], busy: false, loading: mode === 'edit', error: '', status: '' };
    el.addEventListener('cancel', ev => { if (modal && modal.busy) ev.preventDefault(); });
    el.addEventListener('close', () => { if (modal && modal.el === el) closeModal(); else el.remove(); });
    el.addEventListener('click', ev => {
      if (!modal) return;
      const t = ev.target;
      if (t === el && !modal.busy) { closeModal(); return; }
      if (t.closest('[data-mcancel]')) { closeModal(); return; }
      if (t.closest('[data-madd]')) { $('bmFile').value = ''; $('bmFile').click(); return; }
      const rm = t.closest('[data-mrm]');
      if (rm) { modal.images.splice(Number(rm.dataset.mrm), 1); modal.error = ''; paintModal(); return; }
      if (t.closest('[data-mok]')) confirmModal();
    });
    $('bmFile').addEventListener('change', async ev => {
      if (!modal) return;
      const room = rq.maxImages - modal.images.length, picked = [...ev.target.files];
      if (!picked.length) return;
      const files = picked.slice(0, room);
      modal.busy = true; modal.error = ''; modal.status = 'กำลังย่อรูป…'; paintModal();
      try {
        for (const f of files) { const u = await rq.compressImage(f); if (!modal) return; modal.images.push(u); }
        modal.status = picked.length > room ? 'แนบได้สูงสุด ' + rq.maxImages + ' รูป ระบบเลือกเฉพาะรูปแรก ๆ ที่ยังว่างอยู่' : '';
      } catch (e) { if (modal) modal.error = e.message; }
      if (modal) { modal.busy = false; paintModal(); }
    });
    el.showModal(); paintModal();
    if (modal && modal.mode === 'edit') {
      try {
        const groups = await rq.fetchImages(code);
        if (modal) modal.images = ((groups[0] && groups[0].images) || []).slice(0, rq.maxImages);
      } catch (e) { if (modal) modal.error = 'โหลดรูปที่แนบไว้ไม่สำเร็จ: ' + e.message; }
      if (modal) { modal.loading = false; paintModal(); }
    }
  }
  async function confirmModal() {
    const m = modal, rq = window.menuRequests;
    if (!m || m.busy) return;
    m.busy = true; m.error = ''; m.status = m.mode === 'edit' ? 'กำลังบันทึกรูป…' : 'กำลังแจ้ง Admin…'; paintModal();
    try {
      if (m.mode === 'edit') {
        await rq.setImages(m.code, m.images);
        notice = m.images.length ? '✓ อัปเดตรูปสูตรของ ' + m.code + ' แล้ว (' + m.images.length + ' รูป)' : 'ลบรูปสูตรของ ' + m.code + ' แล้ว';
      } else {
        const ctx = { file: (lastModel && lastModel.file.file) || '', date: '' };
        try { ctx.date = branchCheckScope().date; } catch (e) { /* date is optional */ }
        const item = { code: m.code, name: m.name, qty: m.qty, reason: m.reason };
        if (m.images.length) item.images = m.images;
        await rq.report([item], ctx);
        notice = '✓ แจ้ง Admin เรื่อง ' + m.code + ' แล้ว' + (m.images.length ? ' พร้อมรูปสูตร ' + m.images.length + ' รูป' : '') + ' ระบบจะนำไปแสดงในหน้า 2 ของ Admin';
      }
      closeModal(); render(true);
    } catch (e) {
      if (modal === m) { m.busy = false; m.status = ''; m.error = (m.mode === 'edit' ? 'บันทึกรูปไม่สำเร็จ: ' : 'แจ้งไม่สำเร็จ: ') + e.message; paintModal(); }
    }
  }

  function notifyCell(m) {
    if (!reportable(m)) return '<span class="bu-sub">—</span>';
    if (!reported(m.code)) return `<button type="button" class="bu-send" data-bunotify="${E(m.code)}">🔔 แจ้ง Admin</button>`;
    const n = window.menuRequests ? window.menuRequests.imageCount(m.code) : 0;
    return `<div style="display:flex;flex-direction:column;align-items:flex-end;gap:4px"><span class="bu-badge bu-blue">✓ แจ้งแล้ว</span><button type="button" class="bu-link" data-buedit="${E(m.code)}">${n ? '📎 ' + n + ' รูป · ดู/แก้ไข' : '📎 แนบรูปสูตรเพิ่ม'}</button></div>`;
  }
  function bodyHtml(model) {
    const { used, notUsed } = model;
    const canNotify = isBranch();
    const q = query.trim().toLocaleLowerCase();
    const hit = m => !q || (m.code + ' ' + m.name).toLocaleLowerCase().includes(q);
    let body;
    if (!used.length && !notUsed.length) {
      body = '<p class="bu-sub">ไฟล์ยอดขายไม่มีรายการที่มีจำนวนขาย</p>';
    } else if (tab === 'used') {
      body = used.length ? `<div class="bu-wrap"><table><thead><tr><th class="r">#</th><th>รหัสเมนู</th><th>ชื่อเมนู</th><th class="r">จำนวนขาย</th><th>วัตถุดิบที่ถูกตัด (ที่ติดตามในหน้านี้)</th></tr></thead><tbody>${
        used.map((m, i) => [m, i]).filter(([m]) => hit(m)).map(([m, i]) => `<tr class="bu-row" data-bu="${i}" tabindex="0" aria-expanded="false"><td class="r">${i + 1}</td><td>${E(m.code)}</td><td><strong>${E(m.name)}</strong>${m.problems.length ? ' <span class="bu-badge bu-amb">บางวัตถุดิบยังไม่ถูกคิด</span>' : ''}<div class="bu-sub">แถว Excel ${E(xlRows(m.xl))}</div></td><td class="r"><strong>${num(m.qty)}</strong></td><td>${m.items.slice(0, 3).map(d => E(d.name) + ' <b>' + amt(d.amount) + '</b> ' + E(d.unit)).join('<br>')}${m.items.length > 3 ? '<div class="bu-sub">…และอีก ' + (m.items.length - 3) + ' รายการ (กดแถวเพื่อดูทั้งหมด)</div>' : '<div class="bu-sub">กดแถวเพื่อดูรายละเอียด</div>'}</td></tr><tr class="bu-detail" data-bud="${i}" hidden><td></td><td colspan="4"><table><thead><tr><th>รหัสวัตถุดิบ</th><th>ชื่อวัตถุดิบ</th><th class="r">ตัดรวม</th><th>หน่วย</th></tr></thead><tbody>${m.items.map(d => `<tr><td>${E(d.code)}</td><td>${E(d.name)}</td><td class="r"><strong>${amt(d.amount)}</strong></td><td>${E(d.unit)}</td></tr>`).join('')}</tbody></table>${m.problems.length ? '<div class="bu-sub" style="margin-top:6px">⚠️ ' + E(m.problems.join(' / ')) + '</div>' : ''}<div class="bu-sub" style="margin-top:6px">สูตรนี้มีวัตถุดิบ ${m.total} รายการ แสดงเฉพาะที่อยู่ในรายการติดตามของหน้านี้</div></td></tr>`).join('')
      }</tbody></table></div>` : '<p class="bu-sub">ยังไม่มีเมนูในไฟล์ที่ถูกนำไปตัดวัตถุดิบที่ติดตาม</p>';
    } else {
      const badge = { none: ['bu-red', 'ไม่มีสูตร BOM'], problem: ['bu-amb', 'สูตรไม่พร้อมคิด'], untracked: ['bu-blue', 'ไม่กระทบวัตถุดิบที่ติดตาม'] };
      body = notUsed.length ? `<div class="bu-wrap"><table><thead><tr><th class="r">#</th><th>รหัสเมนู</th><th>ชื่อเมนู</th><th class="r">จำนวนขาย</th><th>สถานะ</th>${canNotify ? '<th>แจ้ง Admin</th>' : ''}</tr></thead><tbody>${
        notUsed.map((m, i) => [m, i]).filter(([m]) => hit(m)).map(([m, i]) => `<tr><td class="r">${i + 1}</td><td>${E(m.code)}</td><td><strong>${E(m.name)}</strong><div class="bu-sub">แถว Excel ${E(xlRows(m.xl))}</div></td><td class="r"><strong>${num(m.qty)}</strong></td><td><span class="bu-badge ${badge[m.kind][0]}">${badge[m.kind][1]}</span>${m.kind === 'none' ? '' : '<div class="bu-sub" style="margin-top:4px">' + E(m.reason) + '</div>'}</td>${canNotify ? '<td class="r">' + notifyCell(m) + '</td>' : ''}</tr>`).join('')
      }</tbody></table></div>` : '<p class="bu-sub">✅ ทุกเมนูที่ขายถูกนำไปคิดตามสูตรแล้ว</p>';
    }
    const shown = (tab === 'used' ? used : notUsed).filter(hit).length;
    if (q && !shown && (used.length || notUsed.length)) body = '<p class="bu-sub">ไม่พบรายการที่ตรงกับ “' + E(query.trim()) + '” ในแท็บนี้</p>';
    return body;
  }
  function infoText(model) {
    const { used, notUsed } = model;
    const q = query.trim().toLocaleLowerCase();
    if (!q) return '';
    const hit = m => (m.code + ' ' + m.name).toLocaleLowerCase().includes(q);
    const list = tab === 'used' ? used : notUsed;
    return 'พบ ' + list.filter(hit).length + ' จาก ' + list.length + ' รายการ';
  }
  let lastKey = null;
  function render(force = true) {
    const panel = $('branchBomUsagePanel');
    if (!panel) return;
    const f = typeof branchCheckState !== 'undefined' && branchCheckState.files && branchCheckState.files.SALES;
    const key = f ? [f.file, f.rows.length, f.rows.reduce((t, r) => t + r.qty, 0), appState.recipes.length, appState.materials.length, orderedCountEntries().map(e => e[0]).join('|')].join('~') : '';
    if (!force && key === lastKey) return;
    if (key !== lastKey) notice = '';
    lastKey = key;
    let model = null;
    try { model = analyse(); } catch (e) { model = null; }
    lastModel = model;
    panel.hidden = !model;
    if (!model) return;
    const { file, used, notUsed, qUsed, qNot } = model;
    const total = qUsed + qNot, pct = total ? Math.round(qUsed / total * 1000) / 10 : 0;
    panel.innerHTML = `<h3 class="font-semibold">📋 เมนูในไฟล์ยอดขายกับสูตร BOM</h3>
      <p class="bu-sum" style="margin-top:6px">ยอดขาย ${used.length + notUsed.length} เมนู · <b>ถูกนำมาคิด ${used.length}</b> · <b>ไม่ถูกนำมาคิด ${notUsed.length}</b> · ครอบคลุม ${pct}% ของจำนวนที่ขาย</p>
      <div class="bu-bar" aria-hidden="true"><span style="width:${pct}%"></span></div>
      <div class="bu-sub">ไฟล์ ${E(file.file || '')} · เรียงตามจำนวนขายมากไปน้อย</div>
      <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:10px"><input id="buSearch" type="search" value="${E(query)}" placeholder="🔍 ค้นหารหัสเมนูหรือชื่อเมนู" aria-label="ค้นหาเมนูในไฟล์ยอดขาย" style="flex:1;min-width:220px;max-width:420px;padding:8px 12px;border:1px solid var(--b3,#c6d1e3);border-radius:10px;background:var(--input-bg,#fff);color:var(--t900,#0f1b2d);font-size:14px"><div class="bu-tabs" style="margin:0" role="group" aria-label="เลือกรายการ">
        <button type="button" class="bu-tab" data-butab="used" aria-pressed="${tab === 'used'}">✅ ถูกนำมาคิด (${used.length})</button>
        <button type="button" class="bu-tab" data-butab="not" aria-pressed="${tab === 'not'}">⚠️ ไม่ถูกนำมาคิด (${notUsed.length})</button>
      </div><span id="buInfo" class="bu-sub">${E(infoText(model))}</span><span id="buMsg" class="bu-sub">${E(notice || (isBranch() && window.menuRequests && window.menuRequests.error) || '')}</span></div>
      <div id="buBody">${bodyHtml(model)}</div>`;
  }

  document.addEventListener('input', e => {
    if (!e.target || e.target.id !== 'buSearch' || !lastModel) return;
    query = e.target.value;
    $('buBody').innerHTML = bodyHtml(lastModel);
    $('buInfo').textContent = infoText(lastModel);
  });
  document.addEventListener('click', e => {
    const ed = e.target.closest && e.target.closest('#branchBomUsagePanel [data-buedit]');
    if (ed) { openNotify(ed.dataset.buedit, 'edit'); return; }
    const nb = e.target.closest && e.target.closest('#branchBomUsagePanel [data-bunotify]');
    if (nb) { openNotify(nb.dataset.bunotify, 'new'); return; }
    const t = e.target.closest && e.target.closest('#branchBomUsagePanel [data-butab]');
    if (t) { tab = t.dataset.butab; render(true); return; }
    const row = e.target.closest && e.target.closest('#branchBomUsagePanel tr.bu-row');
    if (row) toggle(row);
  });
  document.addEventListener('keydown', e => {
    if (e.key !== 'Enter' && e.key !== ' ') return;
    const row = e.target.closest && e.target.closest('#branchBomUsagePanel tr.bu-row');
    if (row) { e.preventDefault(); toggle(row); }
  });
  function toggle(row) {
    const d = document.querySelector('#branchBomUsagePanel tr[data-bud="' + row.dataset.bu + '"]');
    if (!d) return;
    d.hidden = !d.hidden;
    row.setAttribute('aria-expanded', String(!d.hidden));
  }

  // Refresh whenever the page-5 calculation runs (upload, clear, change branch/date, edit counts).
  const original = window.calculateBranchCheck;
  if (typeof original === 'function') {
    window.calculateBranchCheck = function () {
      const result = original.apply(this, arguments);
      try { render(false); } catch (e) { /* panel is informational only */ }
      return result;
    };
  }
  window.renderBranchBomUsage = render;
})();
