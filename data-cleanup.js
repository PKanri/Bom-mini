'use strict';
let cleanupPlan=null;
let cleanupBackupRequested=false;
let cleanupBusy=false;
function cleanupSelection() {
  return $w('cleanupMode').value==='range'?`${$w('cleanupStart').value} ถึง ${$w('cleanupEnd').value}`:$w('cleanupPeriod').value;
}
function makeRangeCleanupPlan(documents,start,end) {
  const valid=d=>/^\d{4}-\d{2}-\d{2}$/.test(d)&&Number(d.slice(0,4))>=1900&&!isNaN(Date.parse(d))&&new Date(d).toISOString().slice(0,10)===d;
  if(!valid(start)||!valid(end)||start>end)throw Error('เลือกวันที่เริ่มต้นและสิ้นสุดให้ถูกต้อง');
  const remove=[],updates=[];
  for(const doc of documents) {
    if(doc.kind==='count') {if(doc.period>=start&&doc.period<=end)remove.push(doc);continue;}
    if(!['SALES','STOCK_IN','STOCK_OUT','EXTRA_IN'].includes(doc.kind))continue;
    const p=doc.payload;
    if(!valid(p.start)||!valid(p.end)||!Array.isArray(p.rows))throw Error('ข้อมูลไฟล์นำเข้าไม่สมบูรณ์ กรุณาตรวจสอบก่อนลบ');
    const days=dateRange(p.start,p.end),excluded=new Set(p.excludedDates||[]);
    const deleting=days.filter(d=>d>=start&&d<=end&&!excluded.has(d));
    if(!deleting.length)continue;
    deleting.forEach(d=>excluded.add(d));
    const remaining=days.filter(d=>!excluded.has(d));
    if(!remaining.length){remove.push(doc);continue;}
    updates.push({before:doc,after:{...doc,payload:{...p,start:remaining[0],end:remaining.at(-1),excludedDates:[...excluded].sort(),rows:p.rows.filter(r=>r.date<start||r.date>end)}}});
  }
  return {period:`${start} ถึง ${end}`,end,range:true,remove,updates,retained:[]};
}
function cleanupAffected(plan) {return [...plan.remove,...(plan.updates||[]).map(u=>u.before)];}
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
  const range=$w('cleanupMode').value==='range';
  $w('cleanupStartField').hidden=!range;$w('cleanupEndField').hidden=!range;
  $w('cleanupStart').disabled=!range;$w('cleanupEnd').disabled=!range;
  $w('cleanupPeriodField').hidden=range;$w('cleanupPeriod').disabled=range;
  $w('cleanupConfirmHint').textContent=range?'2. พิมพ์ช่วงที่เลือก เช่น 2026-10-01 ถึง 2026-10-02':'2. พิมพ์เดือนหรือปีที่เลือก เช่น 2026-10 หรือ 2026';
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
  const period=cleanupSelection(),range=$w('cleanupMode').value==='range',start=$w('cleanupStart').value,end=$w('cleanupEnd').value;
  if(range&&(!start||!end)){alert('เลือกวันที่เริ่มต้นและวันที่สิ้นสุดก่อน');return;}
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
    if(cleanupSelection()!==period)return;
    cleanupPlan=range?makeRangeCleanupPlan(docs,start,end):makeCleanupPlan(docs,period);
    const affected=cleanupAffected(cleanupPlan);
    const groups={};affected.forEach(r=>groups[r.kind]=(groups[r.kind]||0)+1);
    const branches=new Set(cleanupPlan.remove.map(r=>r.branch));
    $w('cleanupPreview').innerHTML=`<strong>ช่วง ${esc(period)} · ทุกสาขา · จะลบ ${cleanupPlan.remove.length} แถวจัดเก็บ จาก ${branches.size} สาขา</strong><ul class="list-disc pl-5 mt-2">${Object.entries(groups).map(([kind,n])=>`<li>${esc(kind==='count'?'ยอดนับรายวัน':labels[kind])}: ${n} แถวจัดเก็บ</li>`).join('')}</ul><p class="mt-2">คงยอดนับวันที่ ${cleanupPlan.end}: ${cleanupPlan.retained.length} แถว เพื่อใช้เป็นยอดยกมาของช่วงถัดไป และคงข้อมูลสูตร BOM/วัตถุดิบทั้งหมด</p><p>เมื่อไม่มีข้อมูลช่วงนี้แล้ว จะเปิดรายงานย้อนหลังของช่วงนี้ไม่ได้ครบเหมือนเดิม</p>`;
    if(range)$w('cleanupPreview').innerHTML=`<strong>ลบช่วง ${esc(period)} · ทุกสาขา</strong><p>ลบทั้งแถวจัดเก็บ ${cleanupPlan.remove.length} แถว และลบเฉพาะวันที่ในไฟล์รายเดือน ${cleanupPlan.updates.length} แถวจัดเก็บ</p><p>รวมยอดนับวันสิ้นสุดด้วย เก็บข้อมูลนอกช่วงและสูตร BOM ไว้ รายงานที่ต้องใช้ยอดนับที่ลบเป็นยอดยกมาจะรอข้อมูล</p><p>พิมพ์ยืนยัน: <strong>${esc(period)}</strong></p>`;
    if(affected.length>10000)throw Error('ข้อมูลเกิน 10,000 แถวจัดเก็บ กรุณาเลือกช่วงสั้นลง');
    $w('cleanupBackup').disabled=!affected.length;
  }catch(error){cleanupPlan=null;$w('cleanupPreview').textContent='ตรวจข้อมูลไม่สำเร็จ: '+error.message;}
  finally{cleanupBusy=false;}
}
function downloadCleanupBackup() {
  if(!cleanupPlan||!cleanupAffected(cleanupPlan).length||cleanupBusy)return;
  const backup={format:'bom-period-backup-v1',project:BOM_SUPABASE.url,exportedAt:new Date().toISOString(),period:cleanupPlan.period,preservedClosingDate:cleanupPlan.end,documents:cleanupPlan.remove,preservedClosingCounts:cleanupPlan.retained};
  backup.documents=cleanupAffected(cleanupPlan);backup.updates=cleanupPlan.updates||[];
  const url=URL.createObjectURL(new Blob([JSON.stringify(backup,null,2)],{type:'application/json'}));
  const link=document.createElement('a');link.href=url;link.download=`bom-backup-${cleanupPlan.period}.json`;
  document.body.appendChild(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),30000);
  cleanupBackupRequested=true;$w('cleanupBackupConfirmed').disabled=false;updateCleanupConfirmation();
}
function updateCleanupConfirmation() {
  $w('cleanupDelete').disabled=bomCloud.member?.role!=='admin'||cleanupBusy||!cleanupPlan||!cleanupAffected(cleanupPlan).length||!cleanupBackupRequested||!$w('cleanupBackupConfirmed').checked||$w('cleanupConfirmText').value.trim()!==cleanupPlan.period;
}
async function deleteCleanupPeriod() {
  updateCleanupConfirmation();if($w('cleanupDelete').disabled)return;
  const plan=cleanupPlan;
  if(cleanupSelection()!==plan.period){resetCleanupPreview();return;}
  if(!confirm(`ลบข้อมูล ${plan.period} ทุกสาขา กระทบ ${cleanupAffected(plan).length} แถวจัดเก็บจาก Supabase?\n${plan.range?'ลบยอดนับทุกวันในช่วง รวมวันสิ้นสุด':`เก็บยอดนับ ${plan.end} ไว้`} · เก็บสูตร BOM\nข้อมูลที่กรอกแต่ยังไม่บันทึกในหน้านี้จะถูกยกเลิก`))return;
  cleanupBusy=true;updateCleanupConfirmation();
  let committed=false;
  try {
    await bomCloud.queue;
    bomCloud.maintenance=true;
    document.querySelectorAll('section.tab-content').forEach(el=>el.inert=true);
    // Only the reviewed records are deleted, in a single revision-checked transaction.
    const changes=plan.remove.map(r=>({id:r.id,kind:r.kind,branch:r.branch,period:r.period,expected:r.revision,remove:true}));
    for(const u of plan.updates||[])changes.push({id:u.before.id,kind:u.before.kind,branch:u.before.branch,period:u.before.period,expected:u.before.revision,payload:u.after.payload});
    await cloudRequest('/rest/v1/rpc/bom_save',{method:'POST',body:JSON.stringify({changes})});
    committed=true;bomCloud.maintenance=false;
    await cloudLoad();
    latestReport=[];visibleVarianceRows=[];renderVarianceDashboard();
    resetCleanupPreview();
    $w('cleanupPreview').textContent=`ลบข้อมูล ${plan.period} แล้ว ${changes.length} แถวจัดเก็บ · ${plan.range?'คงข้อมูลนอกช่วงและสูตร BOM':`คงยอดนับ ${plan.end} และสูตร BOM`}`;
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
