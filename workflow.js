/* Monthly central imports and branch physical counts. No network data submission. */
'use strict';
let workflow = { version: 1, batches: {}, branches: {}, counts: {}, settings: {} };
let importPreview = null;
let workflowReady = false;
let workflowDb;
let workflowSaveQueue = Promise.resolve();
const countDrafts = {};
let activeCountScope = null;
let activeCountCodes = [];
let importReadVersion = 0;
const $w = id => document.getElementById(id);
const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const labels = { SALES: 'ยอดขาย', STOCK_IN: 'รับเข้า', STOCK_OUT: 'โอน/จ่ายออกอื่น', EXTRA_IN: 'รับเข้าเพิ่มเติม / เครื่องดื่ม' };
const sourceQuantityField = type => type === 'STOCK_OUT' ? 'Sum Number' : 'Number';
const fmt = value => value == null ? '—' : Number(value).toLocaleString('th-TH', { maximumFractionDigits: 3 });

function normalizedDate(value) {
  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) throw Error('วันที่ Excel ไม่ถูกต้อง');
    value = `${parsed.y}-${String(parsed.m).padStart(2,'0')}-${String(parsed.d).padStart(2,'0')}`;
  }
  const match = String(value ?? '').match(/^(\d{4})-(\d{2})-(\d{2})(?:\s|T|$)/);
  if (!match) throw Error('วันที่ต้องเป็น YYYY-MM-DD');
  const year = Number(match[1]) > 2400 ? Number(match[1]) - 543 : Number(match[1]);
  const date = `${year}-${match[2]}-${match[3]}`;
  const check = new Date(date + 'T00:00:00Z');
  if (!Number.isFinite(check.getTime()) || check.toISOString().slice(0,10) !== date) throw Error('วันที่ไม่มีอยู่จริง');
  return date;
}
function priorDay(date) { const d = new Date(date + 'T00:00:00Z'); d.setUTCDate(d.getUTCDate()-1); return d.toISOString().slice(0,10); }
function dateRange(start, end) {
  if (!start || !end || start > end) return [];
  const dates = [];
  for (let date = start; date <= end && dates.length <= 366; ) {
    dates.push(date); const d = new Date(date + 'T00:00:00Z'); d.setUTCDate(d.getUTCDate()+1); date = d.toISOString().slice(0,10);
  }
  return dates;
}

async function openWorkflowStore() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open('restaurant-bom-workflow', 1);
    req.onupgradeneeded = () => req.result.createObjectStore('state');
    req.onerror = () => reject(req.error);
    req.onsuccess = () => { workflowDb = req.result; resolve(); };
  });
}
async function storeWorkflow() {
  if(typeof BOM_SUPABASE!=='undefined') {
    if(!workflowReady)throw Error('กรุณารอให้เชื่อมต่อฐานข้อมูลก่อนบันทึก');
    return cloudSave(cloudSnapshot());
  }
  if (!workflowReady || !workflowDb) throw Error('ยังไม่พร้อมบันทึกข้อมูลในเบราว์เซอร์');
  const snapshot = structuredClone({ ...workflow, materials: appState.materials, recipes: appState.recipes });
  const save = () => new Promise((resolve, reject) => {
    const tx = workflowDb.transaction('state', 'readwrite');
    tx.objectStore('state').put(snapshot, 'current');
    tx.oncomplete = resolve; tx.onerror = () => reject(tx.error); tx.onabort = () => reject(tx.error || Error('บันทึกไม่สำเร็จ'));
  });
  workflowSaveQueue = workflowSaveQueue.catch(() => {}).then(save);
  return workflowSaveQueue;
}
function storageError(error) { alert('บันทึกไม่สำเร็จ: ' + error.message); }
function parseSourceWorkbook(book, type, quantityField) {
  quantityField = sourceQuantityField(type);
  const result = [], errors = [];
  let matchedSheets = 0;
  for (const name of book.SheetNames) {
    const grid = XLSX.utils.sheet_to_json(book.Sheets[name], { header: 1, defval: null, raw: true });
    if (!grid.some(row => row.some(v => v !== null))) continue;
    const headers = grid[0].map(v => String(v ?? '').trim()).map(v => /^sum\s*number$/i.test(v) ? 'Sum Number' : v);
    // The template's beverage day helper uses Create Time, not Doc Date.
    const extraTransfer = type === 'EXTRA_IN' && headers.includes('BranchFrom (ID)');
    const dateColumn = type === 'STOCK_IN' || extraTransfer ? 'DocDate' : type === 'EXTRA_IN' ? 'Create Time' : 'Doc Date';
    const typeColumns = type === 'SALES' ? ['Sell ID', 'หน่วย'] : type === 'STOCK_IN' || extraTransfer ? ['BranchFrom (ID)'] : type === 'EXTRA_IN' ? ['Supplier (Name)', 'Doc Date'] : ['Branch To (ID)'];
    const required = ['Branch (ID)', 'Branch (Name)', 'Product (Code)', 'Product (Name)', dateColumn, quantityField, ...typeColumns];
    if (required.some(h => !headers.includes(h))) { errors.push(`ชีต ${name}: ขาดคอลัมน์ ${required.filter(h => !headers.includes(h)).join(', ')}`); continue; }
    matchedSheets++;
    grid.slice(1).forEach((row, i) => {
      if (!row.some(v => v !== null && v !== '')) return;
      const item = Object.fromEntries(headers.map((h, j) => [h, row[j]]));
      try {
        const branch = String(item['Branch (ID)'] ?? '').trim(), code = String(item['Product (Code)'] ?? '').trim();
        const raw = item[quantityField];
        const qty = typeof raw === 'number' ? raw : Number(String(raw ?? '').replace(/,/g,''));
        if (!branch || !code || raw == null || raw === '' || !Number.isFinite(qty) || qty < 0) throw Error('รหัสสาขา/สินค้า หรือจำนวนไม่ถูกต้อง');
        if (item.Status === true || String(item.Status).toLowerCase() === 'true') throw Error('พบ Status=true ต้องตรวจสถานะเอกสารก่อนนำเข้า');
        result.push({ branch, branchName: String(item['Branch (Name)'] || branch), code, name: String(item['Product (Name)'] || code), date: normalizedDate(item[dateColumn]), qty, unit: String(item['หน่วย'] || ''), sheet: name, row: i+2 });
      } catch (error) { errors.push(`ชีต ${name} แถว ${i+2}: ${error.message}`); }
    });
  }
  if (!matchedSheets || !result.length) errors.push('ไม่พบรายการข้อมูลที่นำเข้าได้');
  if (result.length && result.every(r=>r.qty===0)) errors.push('คอลัมน์จำนวนเป็นศูนย์ทุกแถว กรุณาตรวจว่าเลือกประเภทไฟล์และคอลัมน์จำนวนถูกต้อง');
  const months = [...new Set(result.map(r => r.date.slice(0,7)))];
  if (months.length !== 1) errors.push('กรุณาใช้ไฟล์ของเดือนเดียวต่อการนำเข้า');
  return { rows: result, errors, month: months[0] };
}
function handleImportTypeChange() {
  importReadVersion++;
  importPreview = null;
  $w('excelFileInput').value = '';
  $w('commitImportBtn').disabled = true;
  $w('stockOutMapping').classList.toggle('hidden', $w('importTypeSelect').value !== 'STOCK_OUT');
  $w('extraInHelp').classList.toggle('hidden', $w('importTypeSelect').value !== 'EXTRA_IN');
  $w('importSummaryText').textContent = 'เลือกไฟล์รวมทุกสาขา ระบบอ่านรหัสสาขาจากแต่ละแถว';
  $w('previewTableBody').innerHTML = '';
  $w('importStatusBadge').textContent = 'รอไฟล์';
  $w('validationErrorsBox').classList.add('hidden');
}
async function handleFileSelected(event) {
  const readVersion=++importReadVersion;
  importPreview = null; $w('commitImportBtn').disabled = true;
  const file = event.target.files[0]; if (!file) return;
  const type = $w('importTypeSelect').value;
  const quantityField = sourceQuantityField(type);
  try {
    if (typeof XLSX === 'undefined') throw Error('ไม่พบตัวอ่าน Excel กรุณาตรวจไฟล์ vendor/xlsx.full.min.js');
    $w('importStatusBadge').textContent = 'กำลังอ่าน';
    const parsed = parseSourceWorkbook(XLSX.read(await file.arrayBuffer(), { type: 'array' }), type, quantityField);
    if ($w('importTypeSelect').value !== type || readVersion!==importReadVersion) return;
    const dates = parsed.rows.map(r => r.date).sort();
    importPreview = { ...parsed, type, file: file.name, quantityField, start: dates[0], end: dates[dates.length-1] };
    $w('importStatusBadge').textContent = parsed.errors.length ? 'ตรวจพบข้อผิดพลาด' : 'พร้อมตรวจทาน';
    $w('importSummaryText').textContent = `${file.name}: ${parsed.rows.length.toLocaleString()} แถว / ${new Set(parsed.rows.map(r=>r.branch)).size} สาขา / ${importPreview.start || '—'} ถึง ${importPreview.end || '—'} · ใช้ ${quantityField} · แสดงตัวอย่าง 30 แถว`;
    $w('validationErrorsBox').classList.toggle('hidden', !parsed.errors.length);
    $w('validationErrorsList').innerHTML = parsed.errors.slice(0,30).map(e=>`<li>${esc(e)}</li>`).join('');
    $w('previewTableHead').innerHTML = '<tr><th>สาขา</th><th>วันที่</th><th>รหัสสินค้า</th><th>ชื่อสินค้า</th><th>จำนวนต้นทาง</th></tr>';
    $w('previewTableBody').innerHTML = parsed.rows.slice(0,30).map(r=>`<tr><td class="p-2">${esc(r.branchName)} (${esc(r.branch)})</td><td>${r.date}</td><td>${esc(r.code)}</td><td>${esc(r.name)}</td><td>${fmt(r.qty)}</td></tr>`).join('');
    $w('commitImportBtn').disabled = parsed.errors.length > 0 || !workflowReady;
  } catch (error) { $w('importStatusBadge').textContent = 'อ่านไม่สำเร็จ'; alert(error.message); }
}
async function commitImportData() {
  if (!importPreview || importPreview.errors.length || !workflowReady) return;
  const p = importPreview, key = `${p.type}:${p.month}`;
  if (workflow.batches[key] && !confirm(`แทนที่ไฟล์${labels[p.type]}ของเดือน ${p.month} ทั้งชุดหรือไม่? ใช้ไฟล์รวมฉบับล่าสุด ยอดตรวจนับจะคงอยู่`)) return;
  const before = structuredClone(workflow);
  $w('commitImportBtn').disabled = true;
  const aggregate = new Map();
  p.rows.forEach(r => {
    workflow.branches[r.branch] = r.branchName;
    const id = JSON.stringify([r.branch,r.date,r.code]);
    if (!aggregate.has(id)) aggregate.set(id, { branch:r.branch,date:r.date,code:r.code,name:r.name,qty:0,unit:r.unit });
    aggregate.get(id).qty += r.qty;
  });
  workflow.batches[key] = { type:p.type,month:p.month,file:p.file,quantityField:p.quantityField,start:p.start,end:p.end,branches:[...new Set(p.rows.map(r=>r.branch))],rows:[...aggregate.values()],sourceRows:p.rows.length, importedAt:new Date().toISOString() };
  try {
    await storeWorkflow(); importPreview = null;
    refreshWorkflowSelectors(); renderImportHistory(); renderCriticalSettings();
    $w('importStatusBadge').textContent = 'บันทึกแล้ว';
    $w('importSummaryText').textContent = `บันทึก ${p.rows.length.toLocaleString()} รายการ แยกตามสาขาแล้ว ช่วงวันที่ ${p.start} ถึง ${p.end}`;
  } catch (error) { workflow = before; storageError(error); $w('commitImportBtn').disabled = false; }
}
function renderImportHistory() {
  $w('importHistory').innerHTML = Object.values(workflow.batches).map(b=>`<li>${esc(labels[b.type])} · ${esc(b.file)} · ${b.start} ถึง ${b.end} · ${b.branches.length} สาขา · ${fmt(b.sourceRows)} แถว</li>`).join('') || '<li>ยังไม่มีไฟล์ส่วนกลางที่บันทึก</li>';
  if(typeof refreshCleanupPeriods==='function')refreshCleanupPeriods();
}
function refreshWorkflowSelectors() {
  for (const id of ['stockCountBranchInput','varianceBranchInput']) {
    const select = $w(id), previous = select.value;
    select.innerHTML = (id === 'varianceBranchInput' ? '<option value="ALL">ทุกสาขา (แสดงแยกสาขา)</option>' : '<option value="">เลือกสาขาที่ตรวจนับ</option>') + Object.entries(workflow.branches).sort((a,b)=>a[1].localeCompare(b[1])).map(([code,name])=>`<option value="${esc(code)}">${esc(name)} (${esc(code)})</option>`).join('');
    if ([...select.options].some(o=>o.value===previous)) select.value=previous;
  }
  const months = [...new Set(Object.values(workflow.batches).map(b=>b.month))].sort().reverse();
  const selected = $w('varianceMonthInput').value;
  $w('varianceMonthInput').innerHTML = months.map(m=>`<option>${m}</option>`).join('');
  if (months.includes(selected)) $w('varianceMonthInput').value=selected;
}
function sourceCatalog() {
  const catalog = new Map();
  Object.values(workflow.batches).filter(b=>b.type!=='SALES').forEach(b=>b.rows.forEach(r=>catalog.set(r.code,r.name)));
  return catalog;
}
function renderCriticalSettings() {
  const select = $w('criticalProduct'), previous=select.value;
  const catalog = sourceCatalog();
  appState.materials.forEach(m=>catalog.set(m.code,m.name));
  select.innerHTML = '<option value="">เลือกวัตถุดิบสำคัญ</option>'+[...catalog].sort((a,b)=>a[0].localeCompare(b[0])).map(([code,name])=>`<option value="${esc(code)}">${esc(code)} — ${esc(name)}</option>`).join('');
  if (catalog.has(previous)) select.value=previous;
  $w('criticalList').innerHTML = Object.entries(workflow.settings).filter(([,s])=>s.critical).map(([code,s])=>`<li class="flex flex-wrap items-center justify-between gap-2 border-b border-slate-100 py-2"><span>${esc(code)} · ${esc(s.name)} · หน่วย ${esc(s.unit)} · รับเข้า × ${s.inFactor} / จ่ายออก × ${s.outFactor}</span><button type="button" data-code="${esc(code)}" onclick="removeCriticalProduct(this.dataset.code, this)" class="shrink-0 rounded-lg border border-rose-200 px-3 py-1.5 text-rose-600 hover:bg-rose-50 disabled:opacity-50" aria-label="ลบ ${esc(s.name)} ออกจากวัตถุดิบสำคัญ">ลบออกจากรายการ</button></li>`).join('') || '<li>ยังไม่ได้เลือกวัตถุดิบสำคัญ กรุณากำหนดก่อนตรวจนับและดูผลต่าง</li>';
}
async function removeCriticalProduct(code, button) {
  const setting = workflow.settings[code];
  if (!setting?.critical) return;
  if (!workflowReady) { alert('กรุณารอให้ระบบเปิดข้อมูลเสร็จก่อน'); return; }
  if (button) button.disabled = true;
  rememberCountDraft();
  setting.critical = false;
  try {
    await storeWorkflow();
  } catch (error) {
    setting.critical = true;
    if (button) button.disabled = false;
    storageError(error);
    return;
  }
  renderCriticalSettings();
  if ($w('criticalProduct').value === code) selectCriticalProduct();
  renderStockCountTable();
  renderVarianceDashboard();
}
function selectCriticalProduct() {
  const s = workflow.settings[$w('criticalProduct').value];
  $w('criticalEnabled').checked = s?.critical ?? true;
  for (const [id,key] of [['criticalUnit','unit'],['criticalInFactor','inFactor'],['criticalOutFactor','outFactor']]) $w(id).value = s?.[key] ?? '';
}
async function saveCriticalProduct() {
  const code=$w('criticalProduct').value, unit=$w('criticalUnit').value.trim();
  const inFactor=Number($w('criticalInFactor').value), outFactor=Number($w('criticalOutFactor').value);
  const cost=workflow.settings[code]?.cost??null;
  if (!code || !unit || !Number.isFinite(inFactor) || !Number.isFinite(outFactor) || inFactor<=0 || outFactor<=0) { alert('ระบุหน่วยฐานและอัตราแปลงรับเข้า/จ่ายออกเป็นค่ามากกว่า 0'); return; }
  const existing=appState.materials.find(m=>m.code===code), name=sourceCatalog().get(code)||existing?.name||code;
  const old=workflow.settings[code];
  if (old && old.unit!==unit && Object.values(workflow.counts).some(c=>c[code]!=null)) { alert('วัตถุดิบนี้มียอดนับแล้ว ไม่สามารถเปลี่ยนหน่วยฐานได้'); return; }
  const before=structuredClone(workflow), materialsBefore=structuredClone(appState.materials);
  workflow.settings[code]={...old,name,unit,inFactor,outFactor,cost,critical:$w('criticalEnabled').checked};
  if (existing) {
    if (existing.baseUnit!==unit && appState.recipes.some(r=>r.items.some(i=>i.materialId===existing.id))) { workflow=before; alert('หน่วยฐานไม่ตรงกับสูตร BOM เดิม กรุณาใช้หน่วยเดียวกัน'); return; }
    Object.assign(existing,{baseUnit:unit, mainUnit:unit, mainFactor:1, subUnit:unit, subFactor:1, costPerBase:cost??0});
  } else appState.materials.push({id:'source-'+code,code,name,category:'วัตถุดิบสำคัญ',baseUnit:unit,mainUnit:unit,mainFactor:1,subUnit:unit,subFactor:1,costPerBase:cost??0});
  try { await storeWorkflow(); renderCriticalSettings(); renderBomModule(); } catch(error) {workflow=before;appState.materials=materialsBefore;storageError(error);}
}
function criticalEntries() {return Object.entries(workflow.settings).filter(([,s])=>s.critical);}
let editingMaterialId = null;
let editingRecipeId = null;
let materialEditSaving = false;
function materialEditLocks(mat) {
  const counted = Object.values(workflow.counts).some(count => count[mat.code] != null);
  const configured = Boolean(workflow.settings[mat.code]);
  const imported = sourceCatalog().has(mat.code);
  const usedInRecipe = appState.recipes.some(recipe => recipe.items.some(item => item.materialId === mat.id));
  return { code: counted || configured || imported, unit: counted || configured || usedInRecipe };
}
function openSelectedMaterialEditor() {
  const mat = appState.materials.find(m => m.id === $w('ingredientSelectInput').value);
  if (!mat) { alert('กรุณาเลือกวัตถุดิบที่ต้องการแก้ไข'); return; }
  editingMaterialId = mat.id;
  const recipe = appState.recipes.find(r => r.menuId === appState.selectedMenuId);
  if (!recipe) { alert('กรุณาเลือกเมนูอาหารก่อน'); return; }
  editingRecipeId = recipe.menuId;
  const item = recipe.items.find(i => i.materialId === mat.id);
  $w('editMaterialRecipe').textContent = `ปริมาณสำหรับเมนู: ${recipe.menuName} (ต่อ 1 จาน)`;
  $w('editMaterialQty').value = item?.netQty ?? '';
  $w('editMaterialYield').value = item?.yieldRate ?? 100;
  const locks = materialEditLocks(mat);
  for (const [id, value] of [['editMaterialCode',mat.code],['editMaterialName',mat.name],['editMaterialUnit',mat.baseUnit]]) $w(id).value = value;
  $w('editMaterialCode').readOnly = locks.code;
  $w('editMaterialUnit').readOnly = locks.unit;
  $w('editMaterialCodeHint').textContent = locks.code ? 'รหัสนี้เชื่อมกับข้อมูลนำเข้าหรือยอดตรวจนับแล้ว จึงคงรหัสเดิมไว้' : 'รหัสต้องไม่ซ้ำกับวัตถุดิบอื่น';
  $w('editMaterialUnitHint').textContent = locks.unit ? 'มีสูตรหรือข้อมูลตรวจนับใช้หน่วยนี้อยู่ จึงคงหน่วยเดิมเพื่อไม่ให้ปริมาณคลาดเคลื่อน' : 'ระบุหน่วยที่ใช้ในสูตร เช่น g, ml หรือชิ้น';
  $w('editMaterialError').textContent = '';
  $w('modalEditMaterial').classList.remove('hidden');
}
function closeSelectedMaterialEditor() {
  if (materialEditSaving) return;
  editingMaterialId = null;
  editingRecipeId = null;
  $w('modalEditMaterial').classList.add('hidden');
}
async function saveSelectedMaterial() {
  if (materialEditSaving) return;
  const mat = appState.materials.find(m => m.id === editingMaterialId);
  if (!mat) return;
  const code = $w('editMaterialCode').value.trim(), name = $w('editMaterialName').value.trim(), unit = $w('editMaterialUnit').value.trim();
  const recipe = appState.recipes.find(r => r.menuId === editingRecipeId);
  const qty = Number($w('editMaterialQty').value), yieldRate = Number($w('editMaterialYield').value);
  const fail = text => { $w('editMaterialError').textContent = text; };
  if (!recipe || !code || !name || !unit || !Number.isFinite(qty) || qty <= 0 || !Number.isFinite(yieldRate) || yieldRate <= 0 || yieldRate > 100) { fail('กรอกปริมาณต่อจานมากกว่า 0 และ Yield มากกว่า 0 ถึง 100% ให้ครบ'); return; }
  if (appState.materials.some(m => m.id !== mat.id && m.code === code) || (code !== mat.code && (sourceCatalog().has(code) || workflow.settings[code]))) { fail('รหัสนี้มีอยู่แล้ว กรุณาใช้รหัสอื่น'); return; }
  const locks = materialEditLocks(mat);
  if ((locks.code && code !== mat.code) || (locks.unit && unit !== mat.baseUnit)) { fail('ไม่สามารถเปลี่ยนรหัสหรือหน่วยที่มีข้อมูลอ้างอิงอยู่แล้ว'); return; }
  const before = structuredClone(mat), settingBefore = structuredClone(workflow.settings[mat.code] || null);
  const itemsBefore = structuredClone(recipe.items);
  const item = recipe.items.find(i => i.materialId === mat.id);
  if (item) {Object.assign(item, { netQty: qty, yieldRate });delete item.unitReview;}
  else recipe.items.push({ materialId: mat.id, netQty: qty, yieldRate });
  Object.assign(mat, { code, name, baseUnit: unit });
  if (before.baseUnit !== unit) Object.assign(mat, { mainUnit: unit, subUnit: unit, mainFactor: 1, subFactor: 1 });
  if (workflow.settings[code]) Object.assign(workflow.settings[code], { name });
  materialEditSaving = true;
  $w('editMaterialSave').disabled = $w('editMaterialCancel').disabled = true;
  try {
    await storeWorkflow();
  } catch (error) {
    Object.assign(mat, before);
    recipe.items = itemsBefore;
    if (settingBefore) workflow.settings[before.code] = settingBefore;
    fail('บันทึกไม่สำเร็จ: ' + error.message);
    return;
  } finally {
    materialEditSaving = false;
    $w('editMaterialSave').disabled = $w('editMaterialCancel').disabled = false;
  }
  closeSelectedMaterialEditor();
  renderBomModule();
  $w('ingredientSelectInput').value = mat.id;
  handleIngredientSelectChange(mat.id);
  renderCriticalSettings();
  if ($w('criticalProduct').value === code) selectCriticalProduct();
  rememberCountDraft();
  renderStockCountTable();
  renderVarianceDashboard();
}
const batchRowIndexes = new WeakMap();
function rowsFor(batch,branch,date) {
  if (!batch) return [];
  if (!batchRowIndexes.has(batch)) {
    const index = new Map();
    batch.rows.forEach(row=>{
      const key=countKey(row.branch,row.date);
      if (!index.has(key)) index.set(key,[]);
      index.get(key).push(row);
    });
    batchRowIndexes.set(batch,index);
  }
  return batchRowIndexes.get(batch).get(countKey(branch,date))||[];
}
function countKey(branch,date) {return `${branch}:${date}`;}
async function applyPonnTemplate() {
  if (typeof PONN_TEMPLATE === 'undefined' || workflow.templateId === PONN_TEMPLATE.id) return;
  const previous=structuredClone({workflow,materials:appState.materials,recipes:appState.recipes});
  workflow.templateArchive=structuredClone({settings:workflow.settings,materials:appState.materials,recipes:appState.recipes});
  for (const source of PONN_TEMPLATE.materials) {
    let mat=appState.materials.find(m=>m.code===source.code);
    if(mat && mat.baseUnit!==source.unit) {
      Object.values(workflow.counts).forEach(count=>{if(count[source.code]) count[source.code].unitReview=mat.baseUnit;});
      appState.recipes.forEach(recipe=>recipe.items.forEach(item=>{if(item.materialId===mat.id)item.unitReview=mat.baseUnit;}));
    }
    if(!mat){mat={id:'ponn-'+source.code,code:source.code};appState.materials.push(mat);}
    Object.assign(mat,{name:source.name,baseUnit:source.unit,mainUnit:source.unit,mainFactor:1,subUnit:source.subUnit,subFactor:1/source.countDivisor,category:'วัตถุดิบสำคัญ PONN',costPerBase:0});
    workflow.settings[source.code]={...source};
  }
  const templateCodes=new Set(PONN_TEMPLATE.materials.map(m=>m.code));
  for(const source of PONN_TEMPLATE.recipes){
    let recipe=appState.recipes.find(r=>r.menuCode===source.menuCode);
    if(!recipe){recipe={menuId:'ponn-'+source.menuCode,menuCode:source.menuCode,sellingPrice:0,items:[]};appState.recipes.push(recipe);}
    recipe.menuName=source.menuName;
    recipe.items=recipe.items.filter(i=>!templateCodes.has(appState.materials.find(m=>m.id===i.materialId)?.code));
    recipe.items.push(...source.items.map(i=>({materialId:appState.materials.find(m=>m.code===i.code).id,netQty:i.netQty,yieldRate:100,templateRows:i.sourceRows})));
  }
  workflow.branches[PONN_TEMPLATE.branch]=PONN_TEMPLATE.branchName;
  workflow.templateId=PONN_TEMPLATE.id;
  try {await storeWorkflow();} catch(error){workflow=previous.workflow;appState.materials=previous.materials;appState.recipes=previous.recipes;throw error;}
}
function rememberCountDraft() {
  if (!activeCountScope) return;
  const draft={};
  activeCountCodes.forEach((code,i)=>{
    const field=$w('count-value-'+i);
    if (field) draft[code]={main:field.value,sub:$w('count-sub-'+i).value,loss:$w('count-loss-'+i).value,remarks:$w('count-note-'+i).value};
  });
  countDrafts[activeCountScope]=draft;
}
function handleCountScopeChange() {rememberCountDraft();renderStockCountTable();}
function renderStockCountTable() {
  const branch=$w('stockCountBranchInput').value,date=$w('stockCountDateInput').value;
  activeCountScope=countKey(branch,date);activeCountCodes=criticalEntries().map(([code])=>code);
  const count=countDrafts[activeCountScope]||workflow.counts[activeCountScope]||{};
  $w('stockCountTableBody').innerHTML=criticalEntries().map(([code,s],i)=>{
    const c=count[code]?.unitReview?{}:count[code]||{};
    const input=(id,value,placeholder)=>`<input id="${id}-${i}" type="number" min="0" step="any" value="${esc(value)}" oninput="rememberCountDraft()" placeholder="${placeholder}" class="border rounded-lg p-2 w-28" ${!branch?'disabled':''}>`;
    return `<tr data-count-code="${esc(code)}"><td class="p-3">${esc(s.name)}<br><small>${esc(code)}</small>${count[code]?.unitReview?'<p class="text-amber-700">ยอดเดิมต่างหน่วย กรุณาตรวจนับใหม่</p>':''}</td><td class="p-3">${esc(s.unit)}<br><small>1 หน่วยหลัก = ${s.countDivisor||1} ${esc(s.subUnit||s.unit)}</small></td><td class="p-3">${input('count-value',c.main??c.qty??'','หน่วยหลัก')}</td><td class="p-3">${input('count-sub',c.sub??'','เศษ')}</td><td class="p-3">${input('count-loss',c.loss??'','สูญเสีย')}</td><td class="p-3"><input id="count-note-${i}" value="${esc(c.remarks||'')}" oninput="rememberCountDraft()" placeholder="หมายเหตุ" class="border rounded-lg p-2" ${!branch?'disabled':''}></td></tr>`;
  }).join('')||'<tr><td colspan="6" class="p-6">กำหนดวัตถุดิบสำคัญในเมนู 1 ก่อน</td></tr>';
  $w('countSaveStatus').textContent=branch?`กรอกยอดคงเหลือจริงปลายวันของ ${workflow.branches[branch]} · ช่องว่าง = ยังไม่ได้นับ / 0 = นับแล้วไม่เหลือ`:'เลือกสาขาก่อนกรอกยอดตรวจนับ';
  filterStockCountRows();
}
function filterStockCountRows() {
  const search=$w('stockCountSearch').value.toLowerCase();
  $w('stockCountTableBody').querySelectorAll('tr[data-count-code]').forEach(tr=>tr.hidden=!tr.textContent.toLowerCase().includes(search));
}
async function savePhysicalStockCount() {
  const branch=$w('stockCountBranchInput').value,date=$w('stockCountDateInput').value;
  if (!branch || !date || !criticalEntries().length) {alert('เลือกสาขา วันที่ และกำหนดวัตถุดิบสำคัญก่อนบันทึก');return;}
  const key=countKey(branch,date), previous=structuredClone(workflow.counts[key]||{}), next={...previous};
  for (const [[code],i] of criticalEntries().map((entry,i)=>[entry,i])) {
    const raw=$w('count-value-'+i).value, subRaw=$w('count-sub-'+i).value, lossRaw=$w('count-loss-'+i).value;
    if (raw==='' && subRaw==='' && lossRaw==='') {if(!previous[code]?.unitReview)delete next[code];continue;}
    if(raw==='' && subRaw===''){alert('กรอกยอดคงเหลือด้วยเมื่อระบุของสูญเสีย');return;}
    const main=Number(raw),sub=Number(subRaw),loss=Number(lossRaw),divisor=workflow.settings[code].countDivisor||1;
    if([main,sub,loss].some(n=>!Number.isFinite(n)||n<0)){alert('ยอดนับและสูญเสียต้องเป็นตัวเลขตั้งแต่ 0 ขึ้นไป');return;}
    next[code]={qty:main+sub/divisor,main,sub,loss,waste:loss/divisor,remarks:$w('count-note-'+i).value,updatedAt:new Date().toISOString()};
  }
  workflow.counts[key]=next;
  try {
    await storeWorkflow();delete countDrafts[key];$w('countSaveStatus').textContent=`บันทึกแล้ว ${workflow.branches[branch]} · ${date} · ${Object.keys(next).length} รายการ`;
    $w('varianceBranchInput').value=branch;$w('varianceDailyInput').value=date;
    appState.varianceFilter.branch=branch;appState.varianceFilter.dailyDate=date;
  } catch(error) {workflow.counts[key]=previous;storageError(error);}
}

function buildReport(branch,start,end) {
  const dates=dateRange(start,end), entries=typeof orderedCountEntries==='function'?orderedCountEntries():criticalEntries();
  if (!dates.length || dates.length>366) return [];
  const issues=new Set(), totals=new Map(entries.map(([code])=>[code,{in:0,out:0,bom:0}]));
  for (const date of dates) {
    const batches={};
    for (const type of Object.keys(labels)) {
      let b=workflow.batches[`${type}:${date.slice(0,7)}`];
      // RLS hides other branches' rows. Coverage confirms an upload even when
      // this branch has no movement; keep metadata out of persisted documents.
      if(!b && type!=='SALES' && typeof bomCloud!=='undefined' && bomCloud.member?.role==='branch' && bomCloud.member.branch===branch) {
        const coverage=bomCloud.coverage.find(c=>c.kind===type && c.period===date.slice(0,7) && c.start_date<=date && c.end_date>=date);
        if(coverage)b={start:coverage.start_date,end:coverage.end_date,quantityField:coverage.quantity_field,branches:[],rows:[]};
      }
      // Central movement files cover all branches. No matching transaction for
      // this branch is zero movement, not an absent upload. Sales remain scoped.
      if (!b || date<b.start || date>b.end || (type==='SALES' && !b.branches.includes(branch))) issues.add(`ขาดข้อมูล${labels[type]}ของสาขาในช่วงที่เลือก`);
      else if (b.quantityField !== sourceQuantityField(type)) issues.add(`กรุณานำเข้า${labels[type]}ใหม่ โดยใช้คอลัมน์ ${sourceQuantityField(type)}`);
      else batches[type]=b;
    }
    for (const type of ['STOCK_IN','EXTRA_IN','STOCK_OUT']) {
      rowsFor(batches[type],branch,date).forEach(r=>{
        if (totals.has(r.code)) totals.get(r.code)[type==='STOCK_OUT'?'out':'in']+=r.qty*workflow.settings[r.code][type==='STOCK_OUT'?'outFactor':'inFactor'];
      });
    }
    rowsFor(batches.SALES,branch,date).filter(r=>r.qty!==0).forEach(sale=>{
      const recipe=appState.recipes.find(r=>r.menuCode===sale.code);
      // Only sales codes with a populated BOM participate in theoretical usage.
      if (!recipe?.items.length) return;
      recipe.items.forEach(item=>{
        if(item.unitReview){issues.add('สูตรเดิมต่างหน่วยจากต้นแบบ กรุณาตรวจปริมาณต่อจานใหม่');return;}
        const mat=appState.materials.find(m=>m.id===item.materialId);
        if (!mat || !Number.isFinite(item.netQty) || item.netQty<=0 || !Number.isFinite(item.yieldRate) || item.yieldRate<=0 || item.yieldRate>100) {issues.add('สูตร BOM หรือ Yield ไม่ถูกต้อง');return;}
        if (totals.has(mat.code)) totals.get(mat.code).bom+=sale.qty*item.netQty/(item.yieldRate/100);
      });
    });
  }
  return entries.map(([code,s])=>{
    const rowIssues=new Set(issues), first=workflow.counts[countKey(branch,priorDay(start))]?.[code], last=workflow.counts[countKey(branch,end)]?.[code];
    if(!Number.isFinite(s.inFactor)||s.inFactor<=0||!Number.isFinite(s.outFactor)||s.outFactor<=0)rowIssues.add('รหัสใหม่ตามรายการ: รอตรวจอัตราแปลงรับเข้า/จ่ายออก');
    if(typeof COUNT_UNIT_LIST!=='undefined' && ['RM-101-0256','RM-101-0290','RM-102-0069'].includes(code) && !appState.recipes.some(r=>r.items.some(i=>appState.materials.find(m=>m.id===i.materialId)?.code===code)))rowIssues.add('รอผูกสูตรกับรหัสวัตถุดิบนี้');
    const opening=first?.unitReview?undefined:first?.qty,actual=last?.unitReview?undefined:last?.qty;
    let waste=0;
    dates.forEach(date=>{const count=workflow.counts[countKey(branch,date)]?.[code];if(!count||count.qty==null||count.unitReview)rowIssues.add('รอข้อมูลตรวจนับ/สูญเสียรายวันให้ครบ');else waste+=count.waste||0;});
    if (opening==null) rowIssues.add(`รอยอดนับยกมา ${priorDay(start)}`);
    if (actual==null) rowIssues.add(`รอยอดนับปลายวัน ${end}`);
    const t=totals.get(code), ready=rowIssues.size===0, theory=ready?opening+t.in-t.out-waste-t.bom:null;
    const actualUsed=ready?opening+t.in-t.out-waste-actual:null, overuse=ready?actualUsed-t.bom:null;
    return {branch,code,name:s.name,unit:s.unit,start,end,opening,actual,...t,waste,ready,theory,actualUsed,overuse,issues:[...rowIssues]};
  });
}
let latestReport=[];
let visibleVarianceRows=[];
function renderVarianceDashboard() {
  const branch=$w('varianceBranchInput').value||'ALL';
  const mode=appState.varianceFilter.periodMode;
  let start,end;
  if(mode==='daily') start=end=$w('varianceDailyInput').value;
  else if(mode==='monthly') {
    const month=$w('varianceMonthInput').value;
    start=month?month+'-01':'';
    end=month?new Date(Date.UTC(Number(month.slice(0,4)),Number(month.slice(5)),0)).toISOString().slice(0,10):'';
  } else {start=$w('varianceStartDateInput').value;end=$w('varianceEndDateInput').value;}
  const branches=branch==='ALL'?Object.keys(workflow.branches):[branch];
  latestReport=branches.flatMap(b=>buildReport(b,start,end));
  visibleVarianceRows=latestReport;
  $w('activeBranchLabel').textContent=branch==='ALL'?'ทุกสาขา (แยกผลต่าง ไม่หักกลบกัน)':workflow.branches[branch]||'เลือกสาขา';
  $w('activePeriodLabel').textContent=start&&end?`${start} ถึง ${end}`:'เลือกช่วงวันที่';
  const ready=latestReport.filter(r=>r.ready), incomplete=latestReport.length-ready.length;
  $w('reportReadiness').textContent=!latestReport.length?'ยังไม่มีรายการ: นำเข้าข้อมูล เลือกวัตถุดิบสำคัญ และเลือกช่วงวันที่ไม่เกิน 366 วัน':`คำนวณได้ ${ready.length} รายการ · รอข้อมูล ${incomplete} รายการ · ไม่ใช้ศูนย์แทนข้อมูลที่ขาด`;
  $w('kpiTheoryCost').textContent=String(ready.length);
  $w('kpiNetVarianceCost').textContent=String(ready.filter(r=>r.overuse < -0.000001).length);
  $w('kpiVarianceDirectionText').textContent='นับแยกตามสาขาและวัตถุดิบ ไม่รวมปริมาณต่างหน่วยเข้าด้วยกัน';
  $w('kpiVariancePct').textContent=String(incomplete);
  $w('kpiWarningItemsCount').textContent=String(ready.filter(r=>r.overuse>0.000001).length);
  $w('varianceTableBody').innerHTML=reportRowsHtml(latestReport,false);
  if(appState.varianceFilter.viewType==='trend') {
    const dates=dateRange(start,end);
    const rows=dates.length<=366?dates.flatMap(d=>branches.flatMap(b=>buildReport(b,d,d))):[];
    visibleVarianceRows=rows;
    $w('varianceDailyBreakdownTableBody').innerHTML=reportRowsHtml(rows,true);
  }
}
function stockDifference(r) {
  if(!r.ready)return {qty:null,pct:null,label:'รอข้อมูล',tone:'pending'};
  const qty=Math.abs(r.overuse)<0.000001?0:-r.overuse;
  return {qty,pct:r.bom>0?qty/r.bom*100:qty===0?0:null,label:qty<0?'ใช้เกินสูตร':qty>0?'ใช้น้อยกว่าสูตร':'ตรงตามสูตร',tone:qty<0?'short':qty>0?'surplus':'equal'};
}
function reportRowsHtml(rows,withDate) {
  const signed=(value,sign)=>value==null?'—':value===0?'0':sign+fmt(value);
  return rows.map(r=>{
    const d=stockDifference(r);
    const pct=d.pct==null?'':(d.pct>0?'+':'')+d.pct.toLocaleString('th-TH',{maximumFractionDigits:1})+'%';
    return `<tr class="stock-result-row">${withDate?'<td class="p-3 whitespace-nowrap">'+r.end+'</td>':''}
      <td class="p-3"><strong>${esc(r.name)}</strong><br><small class="text-slate-400">${esc(r.code)}</small><br><small class="text-slate-500">${esc(workflow.branches[r.branch]||r.branch)} (${esc(r.branch)})</small></td>
      <td class="p-3 text-right whitespace-nowrap">${fmt(r.opening)} ${esc(r.unit)}</td>
      <td class="p-3 text-right stock-in">${signed(r.in,'+')}</td>
      <td class="p-3 text-right stock-out">${signed(r.out,'−')}</td>
      <td class="p-3 text-right stock-out">${signed(r.waste,'−')}</td>
      <td class="p-3 text-right stock-bom">${signed(r.bom,'−')}</td>
      <td class="p-3 text-right stock-theory">${fmt(r.theory)}</td>
      <td class="p-3 text-right stock-actual">${fmt(r.actual)}</td>
      <td class="p-3 text-right whitespace-nowrap stock-diff-${d.tone}">${d.qty==null?'—':(d.qty>0?'+':'')+fmt(d.qty)+' '+esc(r.unit)}</td>
      <td class="p-3"><span class="stock-badge stock-badge-${d.tone}">${d.label}${pct?' ('+pct+')':''}</span>${!r.ready?'<div class="stock-issues">'+esc(r.issues.join(' / '))+'</div>':d.pct==null?'<div class="stock-issues">BOM เป็น 0 จึงคำนวณ % ไม่ได้</div>':''}</td></tr>`;
  }).join('')||`<tr><td colspan="${withDate?11:10}" class="p-6">ยังไม่มีข้อมูลสำหรับรายงาน</td></tr>`;
}
const originalSetVarianceViewType=setVarianceViewType;
setVarianceViewType=function(type){originalSetVarianceViewType(type);renderVarianceDashboard();};
function exportVarianceReport() {
  // Refresh from the current controls, then export exactly the displayed view.
  renderVarianceDashboard();
  if(!visibleVarianceRows.length){alert('ยังไม่มีข้อมูลรายงาน');return;}
  const daily=appState.varianceFilter.viewType==='trend';
  const rows=visibleVarianceRows.map(r=>{
    const d=stockDifference(r),pct=d.pct==null?'':(d.pct>0?'+':'')+d.pct.toLocaleString('th-TH',{maximumFractionDigits:1})+'%';
    return {...(daily?{'วันที่':r.end}:{'ตั้งแต่วันที่':r.start,'ถึงวันที่':r.end}),
      'สาขา':workflow.branches[r.branch]||r.branch,'รหัสสาขา':r.branch,'วัตถุดิบ':r.name,'รหัสวัตถุดิบ':r.code,'หน่วย':r.unit,
      'ยกมา (Opening)':r.opening??'','รับเข้า (+)':r.in,'จ่ายออก (−)':r.out===0?0:-r.out,'สูญเสีย (−)':r.waste===0?0:-r.waste,
      'ตัดตามสูตร (−)':r.bom===0?0:-r.bom,'สต็อกควรมี':r.theory??'','ยอดนับจริง':r.actual??'',
      'ผลต่าง (Qty)':d.qty??'','% ผลต่างเทียบ BOM':d.pct==null?'':d.pct/100,
      'สถานะ':d.label+(pct?' ('+pct+')':''),'รายละเอียด':!r.ready?r.issues.join(' / '):d.pct==null?'BOM เป็น 0 จึงคำนวณ % ไม่ได้':''};
  });
  const sheet=XLSX.utils.json_to_sheet(rows),headers=Object.keys(rows[0]);
  const numeric=['ยกมา (Opening)','รับเข้า (+)','จ่ายออก (−)','สูญเสีย (−)','ตัดตามสูตร (−)','สต็อกควรมี','ยอดนับจริง','ผลต่าง (Qty)','% ผลต่างเทียบ BOM'];
  for(let row=1;row<=rows.length;row++)for(const header of numeric){
    const cell=sheet[XLSX.utils.encode_cell({r:row,c:headers.indexOf(header)})];
    if(cell?.t==='n')cell.z=header==='% ผลต่างเทียบ BOM'?'+0.0%;-0.0%;0.0%':header==='รับเข้า (+)'||header==='ผลต่าง (Qty)'?'+#,##0.###;-#,##0.###;0':'#,##0.###;-#,##0.###;0';
  }
  sheet['!cols']=headers.map(h=>({wch:h==='วัตถุดิบ'?42:h==='รายละเอียด'?65:h==='สถานะ'?30:19}));
  sheet['!autofilter']={ref:sheet['!ref']};
  const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,sheet,daily?'ผลต่างรายวัน':'ผลต่างสรุปช่วง');
  XLSX.writeFile(book,daily?'ผลต่างวัตถุดิบ_รายวัน.xlsx':'ผลต่างวัตถุดิบ_สรุปช่วง.xlsx');
}
window.addEventListener('DOMContentLoaded',async()=>{
  if(typeof BOM_SUPABASE!=='undefined'){initializeCloudUI();return;}
  try {
    await openWorkflowStore();
    const saved=await new Promise((resolve,reject)=>{const req=workflowDb.transaction('state').objectStore('state').get('current');req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error);});
    if(saved){workflow=saved;if(saved.materials)appState.materials=saved.materials;if(saved.recipes)appState.recipes=saved.recipes;}
    workflowReady=true;
    await applyPonnTemplate();
    if(typeof applyCountUnitList==='function')await applyCountUnitList();
    $w('workflowStorage').textContent='ข้อมูลบันทึกในเบราว์เซอร์เครื่องนี้ · ยังไม่เชื่อมข้อมูลข้ามเครื่องของสาขา';
  }catch(error){$w('workflowStorage').textContent='เปิดที่เก็บข้อมูลไม่สำเร็จ: '+error.message;}
  refreshWorkflowSelectors();renderImportHistory();renderCriticalSettings();renderBomModule();handleImportTypeChange();
});
// Persist changes made through the existing BOM editor as well as branch data.
for(const name of ['saveNewMenu','saveNewMaterial','addIngredientToCurrentRecipe','deleteIngredientFromRecipe']){
  const original=window[name];
  window[name]=function(...args){
    if(typeof BOM_SUPABASE!=='undefined' && (!workflowReady || !['admin','branch'].includes(bomCloud.member?.role))){alert('ยังไม่พร้อมแก้ไขสูตร กรุณาตรวจการเชื่อมต่อฐานข้อมูล');return;}
    const before=structuredClone({materials:appState.materials,recipes:appState.recipes});
    const result=original(...args);
    if(workflowReady)storeWorkflow().catch(error=>{appState.materials=before.materials;appState.recipes=before.recipes;renderBomModule();storageError(error);});
    return result;
  };
}
