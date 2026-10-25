'use strict';
function resetStockOverview() {
  for(const id of ['stockOverviewHead','stockOverviewBody'])$w(id).innerHTML='';
  $w('stockOverviewSummary').textContent='';
  const controls=$w('stockOverviewScrollControls');if(controls)controls.hidden=true;
}
function compareStockOverviewBranches(a,b){
  const numeric=id=>/^\d+$/.test(String(id).trim());
  const an=numeric(a[0]),bn=numeric(b[0]);
  if(an!==bn)return an?-1:1;
  return an?Number(a[0])-Number(b[0]):String(a[0]).localeCompare(String(b[0]),'th',{numeric:true});
}
function stockOverviewAllowed(){return workflowReady&&typeof bomCloud!=='undefined'&&bomCloud.member?.role==='admin';}
function stockOverviewSnapshot(branch,date) {
  const hasCount=counts=>Object.values(counts||{}).some(c=>c.qty!=null&&Number.isFinite(c.qty));
  const exact=workflow.counts[countKey(branch,date)];
  if(hasCount(exact))return {date,counts:exact,current:true};
  const previous=Object.keys(workflow.counts).filter(k=>k.startsWith(branch+':')&&k.slice(branch.length+1)<date&&hasCount(workflow.counts[k])).sort().reverse()[0];
  return {date:previous?previous.slice(branch.length+1):null,counts:previous?workflow.counts[previous]:{},current:false};
}
function openStockOverview() {
  if(!stockOverviewAllowed()){resetStockOverview();return;}
  loadStockOverviewLayout();
  renderStockOverviewPicker();
  if(!$w('stockOverviewDate').value)$w('stockOverviewDate').value=localDateString();
  renderStockOverview();
}
function stockOverviewUnits(count,setting,code) {
  const units=countUnits(setting,code),saved=count.unitInputs;
  // Preserve the main/sub quantities actually entered when their units and sum still match.
  if(saved?.length===units.length && saved.every((u,i)=>u.label===units[i].label&&u.factor===units[i].factor&&Number.isFinite(Number(u.value))&&Number(u.value)>=0) && Math.abs(saved.reduce((sum,u)=>sum+Number(u.value)*u.factor,0)-count.qty)<1e-8)
    return units.map((u,i)=>({...u,value:Number(saved[i].value)}));
  const main=units[0],sub=units[1];
  if(!sub||sub.factor>=main.factor)return [{...main,value:count.qty/main.factor}];
  const mainValue=Math.floor(count.qty/main.factor+1e-10);
  const remainder=Math.max(0,count.qty-mainValue*main.factor);
  return [{...main,value:mainValue},{...sub,value:Math.round(remainder/sub.factor*1e9)/1e9}];
}
function renderStockOverview() {
  resetStockOverview();
  if(!stockOverviewAllowed())return;
  try {
    const date=normalizedDate($w('stockOverviewDate').value),selected=$w('stockOverviewBranch').value||'ALL';
    const branches=stockOverviewSelectedBranches(selected);
    const snapshots=branches.map(([id])=>stockOverviewSnapshot(id,date));
    const search=$w('stockOverviewSearch').value.trim().toLocaleLowerCase();
    const entries=orderedCountEntries().filter(([code,s])=>(code+' '+s.name).toLocaleLowerCase().includes(search));
    $w('stockOverviewHead').innerHTML='<tr><th>วัตถุดิบ / หน่วย</th>'+branches.map(([id,name],i)=>{
      const snap=snapshots[i];
      const known=entries.filter(([code])=>snap.counts[code]?.qty!=null&&!snap.counts[code]?.unitReview).length;
      return `<th>${esc(name)} (${esc(id)})<small>${snap.current?'บันทึกวันที่ '+esc(date):'ยังไม่บันทึกวันที่ '+esc(date)}</small><small>${snap.date?(snap.current?'':'แสดงยอดล่าสุด '+esc(snap.date)+' · ')+known+'/'+entries.length+' รายการ':'ไม่มีประวัติก่อนวันที่เลือก'}</small></th>`;
    }).join('')+'</tr>';
    $w('stockOverviewBody').innerHTML=entries.map(([code,s])=>{
      const unit=countUnits(s,code)[0];
      return `<tr><td>${esc(s.name)}<small>${esc(code)} · ${esc(unit.label)}</small></td>`+snapshots.map(snap=>{
        const c=snap.counts[code];
        if(c?.unitReview)return '<td class="pending">รอตรวจหน่วย</td>';
        if(c?.qty==null||!Number.isFinite(c.qty))return '<td class="pending">ยังไม่บันทึก</td>';
        return '<td>'+stockOverviewUnits(c,s,code).map(u=>`<div>${esc(fmt(u.value))} ${esc(u.label)}</div>`).join('')+'</td>';
      }).join('')+'</tr>';
    }).join('')||`<tr><td colspan="${branches.length+1}">ไม่พบวัตถุดิบ</td></tr>`;
    const current=snapshots.filter(s=>s.current).length;
    $w('stockOverviewSummary').textContent=`วันที่ ${date} · ${branches.length} สาขา · มีบันทึกในวันนี้ ${current} สาขา · ยังไม่บันทึก ${branches.length-current} สาขา · ยอดย้อนหลังอ้างอิงวันที่ที่ระบุบนหัวคอลัมน์`;
  }catch(e){$w('stockOverviewSummary').textContent=e.message;}
  if(typeof requestAnimationFrame==='function')requestAnimationFrame(syncStockOverviewScroll);
}

function exportStockOverview() {
  if(!stockOverviewAllowed())return;
  try {
    if(typeof XLSX==='undefined')throw Error('ไม่พบตัวส่งออก Excel กรุณาโหลดหน้าเว็บใหม่');
    const date=normalizedDate($w('stockOverviewDate').value),selected=$w('stockOverviewBranch').value||'ALL';
    const search=$w('stockOverviewSearch').value.trim().toLocaleLowerCase();
    const entries=orderedCountEntries().filter(([code,s])=>(code+' '+s.name).toLocaleLowerCase().includes(search));
    const branches=stockOverviewSelectedBranches(selected);
    if(!entries.length||!branches.length)throw Error('ไม่มีข้อมูลตามตัวกรองที่เลือกให้ส่งออก');
    const rows=[['สต็อกคงเหลือแต่ละสาขา'],['วันที่ '+date+' · เฉพาะยอดบันทึกวันที่เลือก · ช่องว่าง = ยังไม่บันทึกหรือรอตรวจหน่วย'],
      ['วัตถุดิบ / หน่วย',...branches.map(([id,name])=>name+' ('+id+')')]];
    for(const [code,s]of entries){
      const main=countUnits(s,code)[0];
      rows.push([s.name+'\n'+code+' · '+main.label,...branches.map(([id])=>{
        const count=workflow.counts[countKey(id,date)]?.[code];
        if(count?.qty==null||!Number.isFinite(count.qty)||count.unitReview)return null;
        return stockOverviewUnits(count,s,code).map(u=>fmt(u.value)+' '+u.label).join('\n');
      })]);
    }
    const sheet=XLSX.utils.aoa_to_sheet(rows);
    sheet['!cols']=[{wch:52},...branches.map(()=>({wch:30}))];
    sheet['!rows']=[{hpt:24},{hpt:30},{hpt:30},...entries.map(()=>({hpt:48}))];
    sheet['!merges']=[{s:{r:0,c:0},e:{r:0,c:branches.length}},{s:{r:1,c:0},e:{r:1,c:branches.length}}];
    sheet['!autofilter']={ref:XLSX.utils.encode_range({s:{r:2,c:0},e:{r:rows.length-1,c:branches.length}})};
    const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,sheet,'สต็อกคงเหลือ');
    XLSX.writeFile(book,'สต็อกคงเหลือ_ตาราง_'+date+'_'+(selected==='ALL'?'ทุกสาขา':selected==='CUSTOM'?'สาขาที่เลือก':selected)+'.xlsx');
  }catch(e){appAlert(e.message);}
}

// Keep a compact horizontal control within reach throughout the tall table.
function syncStockOverviewScroll(){
  const viewport=$w('stockOverviewScroll'),controls=$w('stockOverviewScrollControls');
  if(!viewport||!controls)return;
  const max=Math.max(0,viewport.scrollWidth-viewport.clientWidth);
  controls.hidden=!stockOverviewAllowed()||max<=1;
  const range=$w('stockOverviewScrollRange');
  range.max=String(max);range.value=String(viewport.scrollLeft);
  $w('stockOverviewScrollLeft').disabled=viewport.scrollLeft<=1;
  $w('stockOverviewScrollRight').disabled=viewport.scrollLeft>=max-1;
}
function setStockOverviewScroll(value){
  const viewport=$w('stockOverviewScroll');if(!viewport)return;
  viewport.scrollLeft=Number(value);syncStockOverviewScroll();
}
function moveStockOverviewScroll(direction){
  const viewport=$w('stockOverviewScroll');if(!viewport)return;
  viewport.scrollBy({left:direction*Math.max(200,viewport.clientWidth-300),behavior:'smooth'});
}
if(typeof window!=='undefined')window.addEventListener('resize',syncStockOverviewScroll);

let stockOverviewLayoutKey=null;
let stockOverviewLayoutSource=null;
let stockOverviewLayoutSaving=false;
let stockOverviewBranchOrder=[];
function stockOverviewStorageKey(){
  return 'ponn.stockOverview.layout.v1.'+(bomCloud.session?.user?.id||'local');
}
function loadStockOverviewLayout(){
  const key=stockOverviewStorageKey();
  if(stockOverviewLayoutKey===key&&stockOverviewLayoutSource===workflow)return;
  stockOverviewLayoutSource=workflow;
  stockOverviewLayoutKey=key;stockOverviewBranchOrder=[];
  $w('stockOverviewBranch').value='ALL';
  $w('stockOverviewLayoutStatus').textContent='';
  try{
    const cloudSaved=workflow.stockOverviewLayouts?.[bomCloud.session?.user?.id];
    let saved=cloudSaved;
    if(!saved){try{saved=JSON.parse(localStorage.getItem(key)||'null');}catch{}}
    $w('stockOverviewLayoutEditor').open=!cloudSaved;
    if(saved&&Array.isArray(saved.branches)){
      stockOverviewBranchOrder=[...new Set(saved.branches.filter(id=>typeof id==='string'))];
      $w('stockOverviewBranch').value=saved.mode==='CUSTOM'?'CUSTOM':'ALL';
      $w('stockOverviewLayoutStatus').textContent=cloudSaved?'ใช้รูปแบบที่บันทึกกับบัญชีแล้ว':'พบรูปแบบเดิมในเครื่องนี้ กดบันทึกเพื่อใช้ข้ามเบราว์เซอร์';
    }
  }catch(e){$w('stockOverviewLayoutStatus').textContent='อ่านรูปแบบที่บันทึกไว้ไม่ได้ กรุณาเลือกสาขาใหม่';}
}
function stockOverviewSelectedBranches(mode){
  if(mode==='CUSTOM')return stockOverviewBranchOrder.filter(id=>Object.hasOwn(workflow.branches,id)).map(id=>[id,workflow.branches[id]]);
  return Object.entries(workflow.branches).filter(([id])=>mode==='ALL'||mode===id).sort(compareStockOverviewBranches);
}
function renderStockOverviewPicker(){
  const custom=$w('stockOverviewBranch').value==='CUSTOM';
  $w('stockOverviewCustomBranches').hidden=!custom;
  const available=Object.entries(workflow.branches).filter(([id])=>!stockOverviewBranchOrder.includes(id)).sort(compareStockOverviewBranches);
  $w('stockOverviewAddBranch').innerHTML='<option value="">เลือกสาขาที่จะเพิ่ม</option>'+available.map(([id,name])=>`<option value="${esc(id)}">${esc(name)} (${esc(id)})</option>`).join('');
  $w('stockOverviewBranchList').innerHTML=stockOverviewBranchOrder.map((id,i)=>`<div class="stock-layout-row"><span>${i+1}. ${esc(workflow.branches[id]||'ไม่พบสาขา')} (${esc(id)})</span><button type="button" aria-label="เลื่อนสาขาลำดับ ${i+1} ขึ้น" onclick="editStockOverviewBranch(${i},-1)" ${i===0?'disabled':''}>↑</button><button type="button" aria-label="เลื่อนสาขาลำดับ ${i+1} ลง" onclick="editStockOverviewBranch(${i},1)" ${i===stockOverviewBranchOrder.length-1?'disabled':''}>↓</button><button type="button" onclick="editStockOverviewBranch(${i},0)" aria-label="ลบสาขาลำดับ ${i+1}">ลบ</button></div>`).join('')||'<p class="text-sm text-slate-500">เลือกสาขาด้านบนเพื่อเริ่มจัดลำดับ</p>';
}
function stockOverviewLayoutChanged(){
  if(!stockOverviewAllowed())return;
  $w('stockOverviewLayoutEditor').open=true;
  $w('stockOverviewLayoutStatus').textContent='ยังไม่ได้บันทึกรูปแบบ';
  renderStockOverviewPicker();renderStockOverview();
}
function addStockOverviewBranch(){
  if(!stockOverviewAllowed())return;
  const id=$w('stockOverviewAddBranch').value;
  if(!Object.hasOwn(workflow.branches,id)||stockOverviewBranchOrder.includes(id))return;
  stockOverviewBranchOrder.push(id);stockOverviewLayoutChanged();
}
function editStockOverviewBranch(index,direction){
  if(!stockOverviewAllowed()||!Number.isInteger(index)||index<0||index>=stockOverviewBranchOrder.length)return;
  if(direction===0)stockOverviewBranchOrder.splice(index,1);
  else{
    const next=index+direction;
    if(next<0||next>=stockOverviewBranchOrder.length)return;
    [stockOverviewBranchOrder[index],stockOverviewBranchOrder[next]]=[stockOverviewBranchOrder[next],stockOverviewBranchOrder[index]];
  }
  stockOverviewLayoutChanged();
}
async function saveStockOverviewLayout(){
  if(!stockOverviewAllowed()||stockOverviewLayoutSaving)return;
  const userId=bomCloud.session?.user?.id;
  if(!userId){$w('stockOverviewLayoutStatus').textContent='กรุณาเข้าสู่ระบบก่อนบันทึก';return;}
  const source=workflow,previous=source.stockOverviewLayouts;
  const saved={mode:$w('stockOverviewBranch').value,branches:[...stockOverviewBranchOrder]};
  stockOverviewLayoutSaving=true;
  const button=$w('stockOverviewSaveLayout');button.disabled=true;
  $w('stockOverviewLayoutStatus').textContent='กำลังบันทึกรูปแบบกับบัญชี…';
  source.stockOverviewLayouts={...previous,[userId]:saved};
  try{
    await storeWorkflow();
    if(workflow!==source||bomCloud.session?.user?.id!==userId)return;
    $w('stockOverviewLayoutStatus').textContent='บันทึกแล้ว เปิดเบราว์เซอร์อื่นด้วยบัญชีเดียวกันเพื่อใช้รูปแบบนี้';
    $w('stockOverviewLayoutEditor').open=false;
    try{localStorage.setItem(stockOverviewStorageKey(),JSON.stringify(saved));}catch{}
  }catch(e){
    if(previous===undefined)delete source.stockOverviewLayouts;
    else source.stockOverviewLayouts=previous;
    if(workflow===source&&bomCloud.session?.user?.id===userId){
      $w('stockOverviewLayoutEditor').open=true;
      $w('stockOverviewLayoutStatus').textContent='บันทึกไม่สำเร็จ: '+e.message;
    }
  }finally{stockOverviewLayoutSaving=false;button.disabled=false;}
}
