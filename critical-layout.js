let criticalLayoutDraft=[];
let criticalLayoutBusy=false;
function canEditCriticalLayout(){return workflowReady&&typeof bomCloud!=='undefined'&&bomCloud.member?.role==='admin';}
function openCriticalLayout(){
  if(!$w('criticalLayoutEditor').open)return;
  if(!canEditCriticalLayout()){$w('criticalLayoutEditor').open=false;return;}
  criticalLayoutDraft=orderedCountEntries().map(([code])=>code);
  $w('criticalLayoutStatus').textContent='';renderCriticalLayout();
}
function renderCriticalLayout(){
  const query=$w('criticalLayoutSearch').value.trim().toLocaleLowerCase();
  $w('criticalLayoutAdd').innerHTML='<option value="">เลือกวัตถุดิบเพื่อเพิ่มท้ายรายการ</option>'+Object.entries(workflow.settings).filter(([code,s])=>!criticalLayoutDraft.includes(code)&&(code+' '+s.name).toLocaleLowerCase().includes(query)).map(([code,s])=>`<option value="${esc(code)}">${esc(s.name)} (${esc(code)})</option>`).join('');
  $w('criticalLayoutRows').innerHTML=criticalLayoutDraft.map((code,i)=>`<div class="critical-layout-row"><span><b>${i+1}. ${esc(workflow.settings[code]?.name||code)}</b><small>${esc(code)}</small></span><button type="button" onclick="moveCriticalLayout(${i},-1)" aria-label="เลื่อนขึ้น" ${i===0?'disabled':''}>↑</button><button type="button" onclick="moveCriticalLayout(${i},1)" aria-label="เลื่อนลง" ${i===criticalLayoutDraft.length-1?'disabled':''}>↓</button><button type="button" onclick="moveCriticalLayout(${i},0)">เอาออก</button></div>`).join('')||'ยังไม่ได้เลือกวัตถุดิบ';
}
function addCriticalLayout(){
  if(!canEditCriticalLayout()||criticalLayoutBusy)return;
  const code=$w('criticalLayoutAdd').value;
  if(!workflow.settings[code]||criticalLayoutDraft.includes(code))return;
  criticalLayoutDraft.push(code);renderCriticalLayout();
}
function moveCriticalLayout(index,direction){
  if(!canEditCriticalLayout()||criticalLayoutBusy)return;
  if(!Number.isInteger(index)||index<0||index>=criticalLayoutDraft.length)return;
  if(direction===0)criticalLayoutDraft.splice(index,1);
  else{const next=index+direction;if(next<0||next>=criticalLayoutDraft.length)return;[criticalLayoutDraft[index],criticalLayoutDraft[next]]=[criticalLayoutDraft[next],criticalLayoutDraft[index]];}
  renderCriticalLayout();
}
async function saveCriticalLayout(){
  if(!canEditCriticalLayout()||criticalLayoutBusy)return;
  criticalLayoutBusy=true;$w('criticalLayoutSave').disabled=true;
  const source=workflow,oldSettings=structuredClone(workflow.settings),oldOrder=workflow.criticalOrder;
  try{
    rememberCountDraft();
    const selected=new Set(criticalLayoutDraft);
    for(const [code,s]of Object.entries(workflow.settings))s.critical=selected.has(code);
    workflow.criticalOrder=[...criticalLayoutDraft];
    await storeWorkflow();
  }catch(error){
    source.settings=oldSettings;
    if(oldOrder===undefined)delete source.criticalOrder;else source.criticalOrder=oldOrder;
    $w('criticalLayoutStatus').textContent='บันทึกไม่สำเร็จ: '+error.message;return;
  }finally{criticalLayoutBusy=false;$w('criticalLayoutSave').disabled=false;}
  if(workflow!==source)return;
  renderCriticalSettings();renderStockCountTable();renderVarianceDashboard();
  if(typeof renderBranchCheck==='function')renderBranchCheck();
  if(typeof renderStockOverview==='function')renderStockOverview();
  $w('criticalLayoutStatus').textContent='บันทึกแล้ว ทั้ง 4 หน้าและ Excel ใช้รายการและลำดับเดียวกัน';
  $w('criticalLayoutEditor').open=false;
}
