'use strict';
function resetStockOverview() {
  for(const id of ['stockOverviewHead','stockOverviewBody'])$w(id).innerHTML='';
  $w('stockOverviewSummary').textContent='';
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
  const selected=$w('stockOverviewBranch').value;
  $w('stockOverviewBranch').innerHTML='<option value="ALL">ทุกสาขา</option>'+Object.entries(workflow.branches).sort((a,b)=>a[1].localeCompare(b[1])).map(([id,name])=>`<option value="${esc(id)}">${esc(name)} (${esc(id)})</option>`).join('');
  $w('stockOverviewBranch').value=workflow.branches[selected]?selected:'ALL';
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
    const branches=Object.entries(workflow.branches).filter(([id])=>selected==='ALL'||selected===id).sort((a,b)=>a[1].localeCompare(b[1]));
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
}
