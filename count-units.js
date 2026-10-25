// Display units follow the user's list. Factors preserve the existing BOM base.
const COUNT_UNIT_LIST = {
  'RM-101-0040':['แพ็ค (1000 กรัม)','ถุง',1,0.1],
  'RM-101-0267':['แพ็ค (2000 กรัม)','ชุด',1,0.1],
  'RM-101-0348':['ชั่งน้ำหนัก (กรัม)','กรัม',1/500,1/500],
  'RM-101-0039':['แพ็ค (500 กรัม/แพ็ค)','กรัม',1,1/500],
  'RM-101-0263':['แพ็ค (1000 กรัม)','กรัม',1,1/1000],
  'RM-101-0002':['แพ็ค (35 ไม้)','ไม้',1,1/35],
  'RM-101-0019':['แพ็ค (5 ชิ้น)','กรัม',1,1/1000],
  'RM-100-0411':['ชุด (150 กรัม)','-',1/3,null],
  'RM-101-0023':['แพ็ค (800 กรัม)','-',1,null],
  'RM-101-0289':['แพ็ค (5 ชุด)','ชุด',1,1/5],
  'RM-101-0324':['แพ็ค (5 ชิ้น)','กรัม',600/400,1/400],
  'RM-101-0239':['แพ็ค (2 ตัว)','ตัว',1,1/2],
  'RM-101-0217':['แพ็ค (20 ชิ้น)','ชิ้น',1,1/20],
  'RM-101-0322':['แพ็ค (5 ชุด/แพ็ค)','ชุด',1,1/5],
  'RM-101-0276':['แพ็ค (225 กรัม)','ถุง',225/450,45/450],
  'RM-101-0258':['แพ็ค (450 กรัม)','ชิ้น',1,1/10],
  'RM-100-0010':['แพ็ค (35 ตัว)','ตัว',1,1/35],
  'RM-101-0310':['แพ็ค (4 ตัว)','ตัว',1,1/4],
  'RM-101-0309':['แพ็ค (250 กรัม)','-',1,null],
  'RM-101-0339':['แพ็ค (250 กรัม)','-',1,null],
  'RM-101-0330':['แพ็ค (35 ไม้)','ไม้',1,1/35],
  'RM-101-0219':['แพ็ค (5 ชุด)','ชุด',1,1/5],
  'RM-101-0508':['แพ็ค (3 ชุด)','ชุด',1,1/3],
  'RM-101-0512':['แพ็ค (3 ชุด)','ชุด',1,1/3],
  'RM-102-0104':['แพ็ค (5 ชุด)','ชุด',1,1/5],
  'RM-102-0131':['แพ็ค (5 ถุง)','ถุง',1,1/5],
  'RM-300-0001':['ฟอง','-',1,null],
  'FG-100-0265':['ขวด','-',1,null],
  'FG-100-0240':['ขวด','-',1,null],
  'FG-100-0266':['กระป๋อง','-',1,null],
  'FG-100-0267':['กระป๋อง','-',1,null],
  'FG-100-0242':['ขวด','-',1,null],
  'FG-100-0241':['ขวด','-',1,null],
  'FG-100-0247':['ขวด','-',1,null]
};
async function applyCountUnitList() {
  if(workflow.countUnitListVersion===2)return;
  const previous=structuredClone(workflow),materials=structuredClone(appState.materials);
  workflow.countUnitListV2Archive=structuredClone(workflow.settings);
  for(const code of ['RM-101-0508','RM-101-0512','RM-102-0104']){
    const source=PONN_TEMPLATE.materials.find(m=>m.code===code),spec=COUNT_UNIT_LIST[code];
    workflow.settings[code]={...workflow.settings[code],...source,unit:spec[0],subUnit:'ชุด',critical:true};
    let mat=appState.materials.find(m=>m.code===code);
    if(!mat){mat={id:'ponn-'+code,code};appState.materials.push(mat);}
    Object.assign(mat,{name:source.name,baseUnit:spec[0],mainUnit:spec[0],mainFactor:1,subUnit:'ชุด',subFactor:spec[3]});
  }
  for(const code of ['RM-101-0256','RM-101-0290','RM-102-0069'])if(workflow.settings[code])workflow.settings[code].critical=false;
  if(!previous.countUnitListVersion)for(const code of Object.keys(COUNT_UNIT_LIST))if(workflow.settings[code])workflow.settings[code].critical=true;
  workflow.countUnitListVersion=2;
  try{await storeWorkflow();}catch(error){workflow=previous;appState.materials=materials;throw error;}
}
function countUnits(setting, code) {
  const spec=COUNT_UNIT_LIST[code];
  if(spec) return [{label:spec[0],factor:spec[2]},...(spec[1]==='-'?[]:[{label:spec[1],factor:spec[3]}])];
  const units = [];
  const add = (label, factor) => { if (!units.some(u => u.label === label)) units.push({label, factor}); };
  const main = setting.unit.startsWith('แพ็ค') ? 'แพ็ค' : setting.unit;
  // The template's weighed pork row is stored in 500 g equivalents.
  if (code !== 'RM-101-0348') add(main, 1);
  if (setting.subUnit && setting.subUnit !== '-') add(setting.subUnit, 1/(setting.countDivisor || 1));
  return units;
}
function displayCountInputs(count, units, setting) {
  if (!count || count.unitReview || (!count.unitInputs && count.qty==null && count.main==null && count.sub==null)) return units.map(()=>'');
  if (count.unitInputs?.length===units.length && count.unitInputs.every((u,i)=>units[i].label===u.label && units[i].factor===u.factor))
    return count.unitInputs.map(u=>u.value);
  const total = count.qty ?? ((Number(count.main)||0)+(Number(count.sub)||0)/(setting.countDivisor||1));
  // Existing records remain numerically identical; do not split or round stored quantities.
  return units.map((u,i)=>i===0 ? total/u.factor : '');
}
// Explicit count-sheet order, independent of JSON/JSONB property ordering.
function orderedCountEntries() {
  const codes=[...new Set([...(workflow.criticalOrder||[]),...Object.keys(COUNT_UNIT_LIST)])];
  const order=new Map(codes.map((code,index)=>[code,index]));
  return criticalEntries().sort(([a],[b])=>(order.get(a)??Number.MAX_SAFE_INTEGER)-(order.get(b)??Number.MAX_SAFE_INTEGER)||a.localeCompare(b));
}
function rememberCountDraft() {
  if (!activeCountScope) return;
  const draft={};
  activeCountCodes.forEach((code,i)=>{
    const units=countUnits(workflow.settings[code],code);
    draft[code]={unitInputs:units.map((u,j)=>({...u,value:$w(`count-unit-${i}-${j}`).value})),loss:$w('count-loss-'+i).value,lossUnit:$w('count-loss-unit-'+i).value,remarks:$w('count-note-'+i).value};
  });
  countDrafts[activeCountScope]=draft;
}
function updateCountTotal(index) {
  rememberCountDraft();
  const code=activeCountCodes[index], c=countDrafts[activeCountScope][code];
  const entered=c.unitInputs.some(u=>u.value!=='');
  const total=c.unitInputs.reduce((sum,u)=>sum+Number(u.value)*u.factor,0);
  const main=countUnits(workflow.settings[code],code)[0];
  $w('count-total-'+index).textContent=entered&&Number.isFinite(total)?`${fmt(total/main.factor)} ${main.label}`:'ยังไม่ได้นับ';
}
function openSavedCountDate(date) {
  if(!date)return;
  rememberCountDraft();
  $w('stockCountDateInput').value=date;
  renderStockCountTable();
}
function renderStockCountTable() {
  const branch=$w('stockCountBranchInput').value,date=$w('stockCountDateInput').value;
  const savedDates=Object.keys(workflow.counts).filter(key=>key.startsWith(branch+':')&&Object.values(workflow.counts[key]).some(c=>c.qty!=null)).map(key=>key.slice(branch.length+1)).sort().reverse();
  $w('savedCountDateSelect').innerHTML='<option value="">'+(!branch?'เลือกสาขาก่อน':savedDates.length?'เลือกวันที่เพื่อดูข้อมูลเดิม':'ยังไม่มีวันที่บันทึก')+'</option>'+savedDates.map(d=>`<option value="${esc(d)}">${esc(d)}</option>`).join('');
  $w('savedCountDateSelect').disabled=!branch||!savedDates.length;
  $w('savedCountDateSelect').value=savedDates.includes(date)?date:'';
  const entries=orderedCountEntries();
  activeCountScope=countKey(branch,date);activeCountCodes=entries.map(([code])=>code);
  const counts=countDrafts[activeCountScope]||workflow.counts[activeCountScope]||{};
  $w('stockCountTableBody').innerHTML=entries.map(([code,s],i)=>{
    const count=counts[code], c=count?.unitReview?{}:count||{}, units=countUnits(s,code), values=displayCountInputs(c,units,s);
    const savedLossUnit=c.lossUnit && units.some(u=>u.label===c.lossUnit)?c.lossUnit:units[0].label;
    const lossFactor=units.find(u=>u.label===savedLossUnit).factor;
    const loss=c.unitInputs?c.loss??'':c.waste!=null?c.waste/lossFactor:c.loss!=null?Number(c.loss)/(s.countDivisor||1)/lossFactor:'';
    const total=values.some(v=>v!=='')?values.reduce((sum,v,j)=>sum+Number(v)*units[j].factor,0):null;
    return `<tr data-count-code="${esc(code)}"><td class="p-3"><strong style="display:block;font-size:16px;font-weight:700;line-height:1.5">${esc(s.name)}</strong><small>${esc(code)}</small>${count?.unitReview?'<p class="text-amber-700">ยอดเดิมต่างหน่วย กรุณาตรวจนับใหม่</p>':''}</td>
      ${[0,1].map(j=>{const u=units[j];return u?`<td class="p-3"><label class="text-xs font-semibold">${esc(u.label)}<input id="count-unit-${i}-${j}" type="number" min="0" step="any" value="${esc(values[j])}" oninput="updateCountTotal(${i})" placeholder="จำนวน${esc(u.label)}" aria-label="คงเหลือ ${esc(s.name)} หน่วย ${esc(u.label)}" class="block border rounded-lg p-2 w-28 mt-1" ${!branch?'disabled':''}></label></td>`:'<td class="p-3 text-slate-400"><input disabled placeholder="– ไม่มีเศษ" class="border rounded-lg p-2 w-28 bg-slate-100"></td>';}).join('')}
      <td id="count-total-${i}" class="p-3 text-emerald-700 font-semibold">${total==null?'ยังไม่ได้นับ':fmt(total/units[0].factor)+' '+esc(units[0].label)}</td>
      <td class="p-3"><input id="count-loss-${i}" type="number" min="0" step="any" value="${esc(loss)}" oninput="rememberCountDraft()" placeholder="จำนวนสูญเสีย" class="border rounded-lg p-2 w-24" ${!branch?'disabled':''}><select id="count-loss-unit-${i}" onchange="rememberCountDraft()" class="border rounded-lg p-2" ${!branch?'disabled':''}>${units.map(u=>`<option value="${esc(u.label)}" ${u.label===savedLossUnit?'selected':''}>${esc(u.label)}</option>`).join('')}</select></td>
      <td class="p-3"><input id="count-note-${i}" value="${esc(c.remarks||'')}" oninput="rememberCountDraft()" placeholder="หมายเหตุ" class="border rounded-lg p-2" ${!branch?'disabled':''}></td></tr>`;
  }).join('')||'<tr><td colspan="6" class="p-6">กำหนดวัตถุดิบสำคัญในเมนู 1 ก่อน</td></tr>';
  const savedItems=Object.values(workflow.counts[activeCountScope]||{}).filter(c=>c.qty!=null).length;
  $w('countSaveStatus').textContent=branch?`สาขา ${workflow.branches[branch]} · ${date} · ${savedItems?'มีข้อมูลบันทึกแล้ว '+savedItems+' รายการ':'ยังไม่มีข้อมูลบันทึกของวันนี้'}${countDrafts[activeCountScope]?' · กำลังแสดงฉบับร่างในหน้านี้':''} · ว่าง = ยังไม่ได้นับ / 0 = นับแล้วไม่เหลือ`:'เลือกสาขาก่อนกรอกยอดตรวจนับ';
  filterStockCountRows();
}
function countAnomalyWarnings(branch,date,next,previous,entries) {
  const warnings=[];
  const dates=Object.keys(workflow.counts).filter(k=>k.startsWith(branch+':')&&k.slice(branch.length+1)<date).sort().reverse();
  for(const [code,s] of entries) {
    const c=next[code],old=previous[code],units=countUnits(s,code),main=units[0];
    const name=s.name+' ('+code+')';
    if(old?.qty!=null&&!old.unitReview&&c?.qty==null){warnings.push(name+': ล้างยอดที่เคยบันทึกเป็น “ยังไม่ได้นับ”');continue;}
    if(c?.qty==null)continue;
    if(old?.qty===c.qty&&old?.waste===c.waste)continue;
    let ref=old?.qty!=null&&!old.unitReview?{value:old,date:date+' (ก่อนแก้ไข)'}:null;
    if(!ref)for(const k of dates){const v=workflow.counts[k][code];if(v?.qty!=null&&!v.unitReview){ref={value:v,date:k.slice(branch.length+1)};break;}}
    const reasons=[];
    if(ref&&Number.isFinite(ref.value.qty)) {
      const before=ref.value.qty,diff=c.qty-before;
      if(before>0&&Math.abs(diff)>=main.factor&&(diff/before>=1||diff/before<=-0.75))reasons.push(`ยอด ${fmt(before/main.factor)} → ${fmt(c.qty/main.factor)} ${main.label} (${diff>0?'+':''}${fmt(diff/before*100)}%) เทียบ ${ref.date}`);
      else if(before===0&&c.qty>=10*main.factor)reasons.push(`จาก 0 เป็น ${fmt(c.qty/main.factor)} ${main.label} เทียบ ${ref.date}`);
    }
    if(c.waste>0&&c.waste>c.qty)reasons.push('จำนวนสูญเสียมากกว่ายอดคงเหลือ กรุณาตรวจหน่วย');
    for(const u of (c.unitInputs||[]).slice(1))if(u.factor<main.factor&&Number(u.value)*u.factor>=main.factor)reasons.push(`ช่องเศษ ${fmt(Number(u.value))} ${u.label} มีปริมาณตั้งแต่ 1 ${main.label} ขึ้นไป ตรวจว่าไม่ได้นับซ้ำกับหน่วยหลัก`);
    if(reasons.length)warnings.push(name+': '+reasons.join(' / '));
  }
  return warnings;
}
async function savePhysicalStockCount() {
  const branch=$w('stockCountBranchInput').value,date=$w('stockCountDateInput').value;
  if(!branch||!date||!criticalEntries().length){appAlert('เลือกสาขา วันที่ และวัตถุดิบก่อนบันทึก');return;}
  const key=countKey(branch,date),previous=structuredClone(workflow.counts[key]||{}),next={...previous};
  for(const [[code,s],i] of orderedCountEntries().map((entry,i)=>[entry,i])){
    const units=countUnits(s,code),unitInputs=units.map((u,j)=>({...u,value:$w(`count-unit-${i}-${j}`).value}));
    const loss=$w('count-loss-'+i).value,lossUnit=$w('count-loss-unit-'+i).value;
    const hasCount=unitInputs.some(u=>u.value!=='');
    if(!hasCount){
      if(loss!==''&&Number(loss)!==0){appAlert('กรอกยอดคงเหลือด้วยเมื่อระบุของสูญเสีย');return;}
      if(!previous[code]?.unitReview){
        if(previous[code]?.receiptQty!=null){next[code]={receipt:previous[code].receipt,receiptUnit:previous[code].receiptUnit,receiptQty:previous[code].receiptQty};}
        else delete next[code];
      }
      continue;
    }
    const lossFactor=units.find(u=>u.label===lossUnit)?.factor;
    if(unitInputs.some(u=>!Number.isFinite(Number(u.value))||Number(u.value)<0)||!Number.isFinite(Number(loss))||Number(loss)<0||!lossFactor){appAlert('จำนวนต้องเป็นตัวเลขตั้งแต่ 0 ขึ้นไป และเลือกหน่วยสูญเสีย');return;}
    next[code]={...(previous[code]?.receiptQty!=null?{receipt:previous[code].receipt,receiptUnit:previous[code].receiptUnit,receiptQty:previous[code].receiptQty}:{}),qty:unitInputs.reduce((sum,u)=>sum+Number(u.value)*u.factor,0),unitInputs,loss,lossUnit,waste:Number(loss)*lossFactor,remarks:$w('count-note-'+i).value,updatedAt:new Date().toISOString()};
  }
  if(Object.values(next).some(c=>c.qty!=null&&(!Number.isFinite(c.qty)||!Number.isFinite(c.waste)))){appAlert('ยอดรวมสูงเกินไปหรือหน่วยแปลงไม่ถูกต้อง กรุณาตรวจตัวเลข');return;}
  const warnings=countAnomalyWarnings(branch,date,next,previous,orderedCountEntries());
  if(warnings.length&&!await appConfirm(`พบ ${warnings.length} รายการที่ควรตรวจซ้ำ · สาขา ${workflow.branches[branch]||branch} · ${date}\n\n${warnings.join('\n\n')}\n\nการรับสินค้า/การใช้งานจริงอาจทำให้ยอดเปลี่ยนได้ หากตรวจแล้วถูกต้อง กดตกลงเพื่อบันทึก หรือยกเลิกเพื่อแก้ไข`))return;
  workflow.counts[key]=next;
  try{await storeWorkflow();delete countDrafts[key];$w('countSaveStatus').textContent=`บันทึกแล้ว ${workflow.branches[branch]} · ${date} · ระบบแปลงหน่วยให้แล้ว`;$w('varianceBranchInput').value=branch;$w('varianceDailyInput').value=date;appState.varianceFilter.branch=branch;appState.varianceFilter.dailyDate=date;}
  catch(error){workflow.counts[key]=previous;storageError(error);}
}

