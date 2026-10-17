'use strict';
let cleanupPlan=null;
let cleanupBackupRequested=false;
let cleanupBusy=false;
function cleanupPeriods() {
  if(typeof bomCloud==='undefined')return [];
  return [...new Set([...bomCloud.baseline.values()].filter(r=>r.kind!=='master').map(r=>r.period.slice(0,7)).filter(p=>/^\d{4}-\d{2}$/.test(p)))].sort().reverse();
}
function refreshCleanupPeriods() {
  const select=$w('cleanupPeriod');if(!select)return;
  const previous=select.value,months=cleanupPeriods(),years=[...new Set(months.map(m=>m.slice(0,4)))];
  select.innerHTML='<option value="">เลือกเดือนหรือปี</option><optgroup label="ทั้งปี">'+years.map(y=>`<option value="${y}">ทั้งปี ${y}</option>`).join('')+'</optgroup><optgroup label="รายเดือน">'+months.map(m=>`<option value="${m}">${m}</option>`).join('')+'</optgroup>';
  if([...select.options].some(o=>o.value===previous))select.value=previous;
  resetCleanupPreview();
}
function resetCleanupPreview() {
  cleanupPlan=null;cleanupBackupRequested=false;
  $w('cleanupPreview').textContent='เลือกช่วง แล้วกดตรวจข้อมูลก่อนลบ';
  $w('cleanupBackup').disabled=true;$w('cleanupDelete').disabled=true;
  $w('cleanupBackupConfirmed').checked=false;$w('cleanupBackupConfirmed').disabled=true;
  $w('cleanupConfirmText').value='';
}
function makeCleanupPlan(documents,period) {
  if(!/^\d{4}(-(?:0[1-9]|1[0-2]))?$/.test(period)||Number(period.slice(0,4))<1900)throw Error('เลือกเดือนหรือปีให้ถูกต้อง');
  const year=Number(period.slice(0,4));
  const end=period.length===4?period+'-12-31':new Date(Date.UTC(year,Number(period.slice(5,7)),0)).toISOString().slice(0,10);
  const kinds=new Set(['SALES','STOCK_IN','STOCK_OUT','EXTRA_IN','count']);
  const matching=documents.filter(r=>kinds.has(r.kind)&&(r.period===period||r.period.startsWith(period+'-')));
  const retained=matching.filter(r=>r.kind==='count'&&r.period===end);
  const remove=matching.filter(r=>!(r.kind==='count'&&r.period===end));
  return {period,end,remove,retained};
}
async function prepareCleanup() {
  if(cleanupBusy)return;
  if(!workflowReady||bomCloud.member?.role!=='admin'){alert('เฉพาะผู้ดูแลระบบเท่านั้น');return;}
  const period=$w('cleanupPeriod').value;
  if(!period){alert('เลือกเดือนหรือปีก่อน');return;}
  resetCleanupPreview();cleanupBusy=true;
  $w('cleanupPreview').textContent='กำลังตรวจข้อมูลล่าสุดใน Supabase…';
  try {
    await bomCloud.queue;
    const docs=[];
    for(let offset=0;;offset+=500){
      const rows=await cloudRequest(`/rest/v1/bom_documents?select=*&kind=neq.master&order=id&limit=500&offset=${offset}`);
      docs.push(...rows);if(rows.length<500)break;
    }
    if($w('cleanupPeriod').value!==period)return;
    cleanupPlan=makeCleanupPlan(docs,period);
    const groups={};cleanupPlan.remove.forEach(r=>groups[r.kind]=(groups[r.kind]||0)+1);
    const branches=new Set(cleanupPlan.remove.map(r=>r.branch));
    $w('cleanupPreview').innerHTML=`<strong>ช่วง ${esc(period)} · ทุกสาขา · จะลบ ${cleanupPlan.remove.length} แถวจัดเก็บ จาก ${branches.size} สาขา</strong><ul class="list-disc pl-5 mt-2">${Object.entries(groups).map(([kind,n])=>`<li>${esc(kind==='count'?'ยอดนับรายวัน':labels[kind])}: ${n} แถวจัดเก็บ</li>`).join('')}</ul><p class="mt-2">คงยอดนับวันที่ ${cleanupPlan.end}: ${cleanupPlan.retained.length} แถว เพื่อใช้เป็นยอดยกมาของช่วงถัดไป และคงข้อมูลสูตร BOM/วัตถุดิบทั้งหมด</p><p>เมื่อไม่มีข้อมูลช่วงนี้แล้ว จะเปิดรายงานย้อนหลังของช่วงนี้ไม่ได้ครบเหมือนเดิม</p>`;
    if(cleanupPlan.remove.length>10000)throw Error('ข้อมูลเกิน 10,000 แถวจัดเก็บ กรุณาเลือกช่วงเป็นรายเดือน');
    $w('cleanupBackup').disabled=!cleanupPlan.remove.length;
  }catch(error){cleanupPlan=null;$w('cleanupPreview').textContent='ตรวจข้อมูลไม่สำเร็จ: '+error.message;}
  finally{cleanupBusy=false;}
}
function downloadCleanupBackup() {
  if(!cleanupPlan?.remove.length||cleanupBusy)return;
  const backup={format:'bom-period-backup-v1',project:BOM_SUPABASE.url,exportedAt:new Date().toISOString(),period:cleanupPlan.period,preservedClosingDate:cleanupPlan.end,documents:cleanupPlan.remove,preservedClosingCounts:cleanupPlan.retained};
  const url=URL.createObjectURL(new Blob([JSON.stringify(backup,null,2)],{type:'application/json'}));
  const link=document.createElement('a');link.href=url;link.download=`bom-backup-${cleanupPlan.period}.json`;
  document.body.appendChild(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),30000);
  cleanupBackupRequested=true;$w('cleanupBackupConfirmed').disabled=false;updateCleanupConfirmation();
}
function updateCleanupConfirmation() {
  $w('cleanupDelete').disabled=bomCloud.member?.role!=='admin'||cleanupBusy||!cleanupPlan?.remove.length||!cleanupBackupRequested||!$w('cleanupBackupConfirmed').checked||$w('cleanupConfirmText').value.trim()!==cleanupPlan.period;
}
async function deleteCleanupPeriod() {
  updateCleanupConfirmation();if($w('cleanupDelete').disabled)return;
  const plan=cleanupPlan;
  if($w('cleanupPeriod').value!==plan.period){resetCleanupPreview();return;}
  if(!confirm(`ลบข้อมูล ${plan.period} ทุกสาขา จำนวน ${plan.remove.length} แถวจัดเก็บจาก Supabase?\nเก็บยอดนับ ${plan.end} และสูตร BOM ไว้\nข้อมูลที่กรอกแต่ยังไม่บันทึกในหน้านี้จะถูกยกเลิก`))return;
  cleanupBusy=true;updateCleanupConfirmation();
  let committed=false;
  try {
    await bomCloud.queue;
    bomCloud.maintenance=true;
    document.querySelectorAll('section.tab-content').forEach(el=>el.inert=true);
    // Only the reviewed records are deleted, in a single revision-checked transaction.
    const changes=plan.remove.map(r=>({id:r.id,kind:r.kind,branch:r.branch,period:r.period,expected:r.revision,remove:true}));
    await cloudRequest('/rest/v1/rpc/bom_save',{method:'POST',body:JSON.stringify({changes})});
    committed=true;bomCloud.maintenance=false;
    await cloudLoad();
    latestReport=[];visibleVarianceRows=[];renderVarianceDashboard();
    resetCleanupPreview();
    $w('cleanupPreview').textContent=`ลบข้อมูล ${plan.period} แล้ว ${changes.length} แถวจัดเก็บ · คงยอดนับ ${plan.end} และสูตร BOM`;
  }catch(error){
    workflowReady=false;cleanupPlan=null;cleanupBackupRequested=false;
    $w('cleanupPreview').textContent=(committed?'ลบสำเร็จแล้ว แต่โหลดข้อมูลล่าสุดไม่สำเร็จ: ':'ยังยืนยันผลการลบไม่ได้: ')+error.message+' · กดโหลดข้อมูลล่าสุดและตรวจสอบก่อนทำรายการใหม่';
    cloudStatus($w('cleanupPreview').textContent);
  }finally{
    bomCloud.maintenance=false;cleanupBusy=false;
    document.querySelectorAll('section.tab-content').forEach(el=>el.inert=!workflowReady);
    updateCleanupConfirmation();
  }
}
