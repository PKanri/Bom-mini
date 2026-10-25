'use strict';
let restorePlan=null,restoreBusy=false,restoreBackupReady=false,restoredPeriod='';
function resetRestore() {
  restorePlan=null;restoreBackupReady=false;restoredPeriod='';
  $w('restoreFile').value='';$w('restoreOverwrite').checked=false;
  $w('restoreBackupSaved').checked=false;$w('restoreBackupSaved').disabled=true;
  $w('restoreBackup').disabled=true;$w('restoreCommit').disabled=true;$w('restoreReport').hidden=true;
  $w('restorePreview').textContent='เลือกไฟล์ JSON เพื่อเริ่มตรวจสอบ';
}
function readRestoreDocuments(backup) {
  if(backup?.format!=='bom-period-backup-v1'||backup.project!==BOM_SUPABASE.url||!Array.isArray(backup.documents))throw Error('ต้องเป็นไฟล์สำรอง JSON ของระบบและโปรเจกต์นี้');
  const docs=new Map(),object=v=>v&&typeof v==='object'&&!Array.isArray(v);
  const validDate=d=>typeof d==='string'&&/^\d{4}-\d{2}-\d{2}$/.test(d)&&!isNaN(Date.parse(d))&&new Date(d).toISOString().slice(0,10)===d;
  for(const r of [...backup.documents,...(backup.preservedClosingCounts||[])]) {
    if(!object(r)||!['SALES','STOCK_IN','STOCK_OUT','EXTRA_IN','count'].includes(r.kind)||typeof r.branch!=='string'||!r.branch||r.id!==`${r.kind}:${r.branch}:${r.period}`||!object(r.payload))throw Error('รูปแบบแถวจัดเก็บไม่ถูกต้อง');
    if(r.kind==='count') {
      if(!validDate(r.period)||Object.values(r.payload).some(c=>!object(c)||c.qty!=null&&(!Number.isFinite(c.qty)||c.qty<0)))throw Error('ยอดตรวจนับในไฟล์ไม่ถูกต้อง');
    } else {
      const p=r.payload;
      if(!/^\d{4}-(0[1-9]|1[0-2])$/.test(r.period)||p.type!==r.kind||p.month!==r.period||!validDate(p.start)||!validDate(p.end)||p.start>p.end||p.start.slice(0,7)!==r.period||p.end.slice(0,7)!==r.period||!Array.isArray(p.rows)||!Array.isArray(p.branches)||p.branches.length!==1||p.branches[0]!==r.branch)throw Error('ข้อมูลไฟล์รายเดือนไม่ถูกต้อง');
      if(p.rows.some(row=>!object(row)||row.branch!==r.branch||!validDate(row.date)||row.date.slice(0,7)!==r.period||typeof row.code!=='string'||!Number.isFinite(row.qty)))throw Error('รายการเคลื่อนไหวในไฟล์ไม่ถูกต้อง');
      if(p.excludedDates!=null&&(!Array.isArray(p.excludedDates)||p.excludedDates.some(d=>!validDate(d))))throw Error('วันที่ในไฟล์ไม่ถูกต้อง');
    }
    if(docs.has(r.id))throw Error('พบรหัสแถวซ้ำในไฟล์สำรอง');
    docs.set(r.id,{id:r.id,kind:r.kind,branch:r.branch,period:r.period,payload:r.payload});
  }
  if(!docs.size||docs.size>10000)throw Error('ไฟล์ต้องมีข้อมูล 1–10,000 แถวจัดเก็บ');
  return [...docs.values()];
}
function restoreCanonical(v) {
  if(Array.isArray(v))return '['+v.map(restoreCanonical).join(',')+']';
  if(v&&typeof v==='object')return '{'+Object.keys(v).sort().map(k=>JSON.stringify(k)+':'+restoreCanonical(v[k])).join(',')+'}';
  return JSON.stringify(v);
}
function buildRestorePlan(documents,current,period) {
  const changes=[],conflicts=[],unchanged=[];
  for(const doc of documents) {
    const old=current.get(doc.id);
    if(old&&restoreCanonical(old.payload)===restoreCanonical(doc.payload)){unchanged.push(doc);continue;}
    if(old)conflicts.push(old);
    changes.push({...doc,expected:old?.revision??null});
  }
  return {changes,conflicts,unchanged,period,documents};
}
async function previewRestoreFile(file) {
  if(restoreBusy)return;
  resetRestore();
  if(!file)return;
  if(!workflowReady||bomCloud.member?.role!=='admin'){appAlert('เฉพาะส่วนกลางที่เข้าสู่ระบบแล้ว');return;}
  restoreBusy=true;$w('restoreFile').disabled=true;
  try {
    if(file.size>100*1024*1024)throw Error('ไฟล์ใหญ่เกิน 100 MB');
    const backup=JSON.parse(await file.text()),documents=readRestoreDocuments(backup);
    await bomCloud.queue;
    const current=new Map();
    for(let offset=0;;offset+=500){const rows=await cloudRequest(`/rest/v1/bom_documents?select=*&kind=neq.master&order=id&limit=500&offset=${offset}`);rows.forEach(r=>current.set(r.id,r));if(rows.length<500)break;}
    restorePlan=buildRestorePlan(documents,current,String(backup.period||''));
    const p=restorePlan;
    $w('restorePreview').innerHTML=`<strong>${esc(file.name)}</strong><p>ช่วง ${esc(p.period)} · ${new Set(documents.map(d=>d.branch)).size} สาขา · ${documents.length} แถวจัดเก็บ</p><p>เพิ่มใหม่ ${p.changes.length-p.conflicts.length} · ต่างจากปัจจุบัน ${p.conflicts.length} · เหมือนเดิม ${p.unchanged.length}</p><div class="max-h-64 overflow-auto mt-3">${p.changes.map(d=>`<div>${esc(d.kind)} · สาขา ${esc(d.branch)} · ${esc(d.period)} · ${d.expected==null?'เพิ่มใหม่':'แทนที่ข้อมูลปัจจุบัน'}</div>`).join('')}</div>`;
    $w('restoreBackup').disabled=!p.changes.length;
    if(!p.changes.length){restoredPeriod=p.period;$w('restoreReport').hidden=false;}
  } catch(error){restorePlan=null;$w('restorePreview').textContent='ตรวจไฟล์ไม่สำเร็จ: '+error.message;}
  finally {restoreBusy=false;$w('restoreFile').disabled=false;updateRestoreButton();}
}
function updateRestoreButton() {
  $w('restoreCommit').disabled=restoreBusy||bomCloud.member?.role!=='admin'||!restorePlan?.changes.length||!restoreBackupReady||!$w('restoreBackupSaved').checked||(restorePlan.conflicts.length>0&&!$w('restoreOverwrite').checked);
}
function downloadRestoreSafetyBackup() {
  if(!restorePlan||restoreBusy||bomCloud.member?.role!=='admin')return;
  const data={format:'bom-period-backup-v1',project:BOM_SUPABASE.url,exportedAt:new Date().toISOString(),period:restorePlan.period,documents:restorePlan.conflicts,preservedClosingCounts:[]};
  const url=URL.createObjectURL(new Blob([JSON.stringify(data,null,2)],{type:'application/json'})),a=document.createElement('a');
  a.href=url;a.download='bom-before-restore-'+new Date().toISOString().slice(0,10)+'.json';document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),30000);
  restoreBackupReady=true;$w('restoreBackupSaved').disabled=false;updateRestoreButton();
}
async function commitRestore() {
  updateRestoreButton();if($w('restoreCommit').disabled||bomCloud.busy||bomCloud.maintenance)return;
  const plan=restorePlan;
  if(!await appConfirm(`กู้คืน ${plan.changes.length} แถวจัดเก็บเข้า Supabase (แทนที่ ${plan.conflicts.length} แถว)? สูตร BOM ไม่เปลี่ยน ข้อมูลที่ยังไม่บันทึกในหน้าจอจะถูกยกเลิก`))return;
  restoreBusy=true;updateRestoreButton();$w('restoreFile').disabled=true;
  let committed=false;
  try {
    await bomCloud.queue;bomCloud.maintenance=true;
    document.querySelectorAll('section.tab-content').forEach(el=>el.inert=true);
    await cloudRequest('/rest/v1/rpc/bom_save',{method:'POST',body:JSON.stringify({changes:plan.changes})});committed=true;
    bomCloud.maintenance=false;await cloudLoad();
    restoredPeriod=plan.period;restorePlan=null;latestReport=[];visibleVarianceRows=[];
    $w('restorePreview').textContent='กู้คืนสำเร็จแล้ว เปิดหน้า 4 เพื่อเลือกสาขาและวันที่ดูรายงาน';$w('restoreReport').hidden=false;$w('restoreBackup').disabled=true;
  } catch(error) {
    restorePlan=null;$w('restoreBackup').disabled=true;
    $w('restorePreview').textContent=(committed?'กู้คืนแล้ว แต่โหลดข้อมูลไม่สำเร็จ: ':'ยังยืนยันการกู้คืนไม่ได้: ')+error.message+' · โหลดข้อมูลล่าสุดและตรวจไฟล์ใหม่ก่อนทำต่อ';
    workflowReady=false;
  } finally {bomCloud.maintenance=false;restoreBusy=false;$w('restoreFile').disabled=false;document.querySelectorAll('section.tab-content').forEach(el=>el.inert=!workflowReady);updateRestoreButton();}
}
function openRestoredReport() {
  if(!workflowReady||bomCloud.member?.role!=='admin')return;
  switchTab('tab-variance');
  const match=restoredPeriod.match(/^(\d{4}-\d{2}-\d{2}) ถึง (\d{4}-\d{2}-\d{2})$/);
  if(match){setVariancePeriodMode('custom');$w('varianceStartDateInput').value=match[1];$w('varianceEndDateInput').value=match[2];}
  else if(/^\d{4}-\d{2}$/.test(restoredPeriod)){setVariancePeriodMode('monthly');$w('varianceMonthInput').value=restoredPeriod;}
  else if(/^\d{4}$/.test(restoredPeriod)){setVariancePeriodMode('custom');$w('varianceStartDateInput').value=restoredPeriod+'-01-01';$w('varianceEndDateInput').value=restoredPeriod+'-12-31';}
  renderVarianceDashboard();
}
