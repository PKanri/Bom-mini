let masterEditing=null,masterSaving=false,masterCreating=false;
function masterAllowed(){return workflowReady&&['admin','branch'].includes(bomCloud.member?.role);}
function showMaterialMaster(){
  if(!masterAllowed())return;
  masterCreating=false;
  $w('masterTitle').textContent='แก้ไขข้อมูลวัตถุดิบหลัก';
  $w('masterChoose').hidden=false;
  $w('masterUnitHelp').textContent='หน่วยนี้ใช้สำหรับตรวจนับและแสดงรายงาน การเปลี่ยนชื่อหน่วยหลักต้องยังหมายถึงปริมาณเดิม สูตร BOM และยอดรวมที่บันทึกแล้วคงเดิม';
  const selected=$w('masterMaterialSelect').value;
  $w('masterMaterialSelect').innerHTML=appState.materials.map(m=>`<option value="${esc(m.id)}">${esc(m.code)} · ${esc(m.name)}</option>`).join('');
  if(appState.materials.some(m=>m.id===selected))$w('masterMaterialSelect').value=selected;
  loadMaterialMaster();$w('materialMasterDialog').showModal();
}
function loadMaterialMaster(){
  const mat=appState.materials.find(m=>m.id===$w('masterMaterialSelect').value);masterEditing=mat?.id;
  if(!mat)return;
  const setting=workflow.settings[mat.code]||{unit:mat.baseUnit};
  const units=countUnits(setting,mat.code);
  $w('masterCode').value=mat.code;$w('masterName').value=mat.name;
  $w('masterMain').value=units[0].label;$w('masterSub').value=units[1]?.label||'';
  $w('masterRatio').value=units[1]?units[0].factor/units[1].factor:'';
  $w('masterCode').readOnly=materialEditLocks(mat).code;
  $w('masterHint').textContent=$w('masterCode').readOnly?'รหัสนี้เชื่อมกับข้อมูลเดิมแล้ว จึงคงรหัสไว้เพื่อให้ข้อมูลตรงกัน':'แก้รหัสได้ โดยต้องไม่ซ้ำกับวัตถุดิบอื่น';
  $w('masterStatus').textContent='';
}
async function saveMaterialMaster(){
  if(!masterAllowed()||masterSaving)return;
  if(masterCreating){await createMaterialMaster();return;}
  const mat=appState.materials.find(m=>m.id===masterEditing);if(!mat)return;
  const code=$w('masterCode').value.trim(),name=$w('masterName').value.trim(),main=$w('masterMain').value.trim(),sub=$w('masterSub').value.trim(),ratio=Number($w('masterRatio').value);
  const status=$w('masterStatus');
  if(!code||!name||!main||(sub&&(!Number.isFinite(ratio)||ratio<=0))||(sub&&sub===main)){status.textContent='กรอกรหัส ชื่อ หน่วยหลัก และจำนวนหน่วยย่อยต่อ 1 หน่วยหลักให้ถูกต้อง';return;}
  if(code!==mat.code&&(materialEditLocks(mat).code||appState.materials.some(m=>m.code===code)||workflow.settings[code]||sourceCatalog().has(code))){status.textContent='เปลี่ยนรหัสไม่ได้ รหัสเดิมมีข้อมูลอ้างอิง หรือรหัสใหม่ซ้ำ';return;}
  const previous=structuredClone(mat),settings=structuredClone(workflow.settings);
  const setting=workflow.settings[mat.code]||{unit:mat.baseUnit,critical:false,inFactor:1,outFactor:1};
  const factor=countUnits(setting,mat.code)[0].factor;
  const units=[{label:main,factor},...(sub?[{label:sub,factor:factor/ratio}]:[])];
  masterSaving=true;$w('masterFields').disabled=true;status.textContent='กำลังบันทึก…';
  try{
    rememberCountDraft();
    Object.assign(mat,{code,name,mainUnit:main,mainFactor:factor,subUnit:sub,subFactor:sub?factor/ratio:1});
    workflow.settings[code]={...setting,name,subUnit:sub,countUnits:units};
    await storeWorkflow();
  }catch(e){Object.assign(mat,previous);workflow.settings=settings;status.textContent='บันทึกไม่สำเร็จ: '+e.message;return;}
  finally{masterSaving=false;$w('masterFields').disabled=false;}
  $w('materialMasterDialog').close();renderCriticalSettings();renderBomModule();renderStockCountTable();renderVarianceDashboard();
  if(typeof renderBranchCheck==='function')renderBranchCheck();
  if(typeof renderStockOverview==='function')renderStockOverview();
}

function showNewMaterialMaster(){
  if(!masterAllowed()||masterSaving)return;
  masterCreating=true;masterEditing=null;
  $w('masterTitle').textContent='เพิ่มวัตถุดิบหลัก';$w('masterChoose').hidden=true;
  for(const id of ['masterCode','masterName','masterMain','masterSub','masterRatio'])$w(id).value='';
  $w('masterCode').readOnly=false;
  $w('masterHint').textContent='รหัสต้องไม่ซ้ำกับวัตถุดิบที่มีอยู่';
  $w('masterUnitHelp').textContent='ตัวอย่าง: หน่วยหลัก แพ็ค (500 กรัม) หน่วยย่อย กรัม และ 1 หน่วยหลักมี 500 หน่วยย่อย โดยสูตร BOM ของวัตถุดิบใหม่นี้จะใช้หน่วยหลัก';
  $w('masterStatus').textContent='';$w('materialMasterDialog').showModal();$w('masterCode').focus();
}
async function createMaterialMaster(){
  const code=$w('masterCode').value.trim(),name=$w('masterName').value.trim(),main=$w('masterMain').value.trim(),sub=$w('masterSub').value.trim(),ratio=Number($w('masterRatio').value),status=$w('masterStatus');
  if(!code||!name||!main||(sub&&(!Number.isFinite(ratio)||ratio<=0||sub===main))){status.textContent='กรอกรหัส ชื่อ หน่วยหลัก และจำนวนหน่วยย่อยต่อ 1 หน่วยหลักให้ถูกต้อง';return;}
  if(appState.materials.some(m=>m.code===code)||workflow.settings[code]){status.textContent='รหัสนี้มีอยู่แล้ว กรุณาใช้หน้าแก้ไขข้อมูลวัตถุดิบ';return;}
  const id='mat-'+crypto.randomUUID();
  const mat={id,code,name,baseUnit:main,mainUnit:main,mainFactor:1,subUnit:sub,subFactor:sub?1/ratio:1,costPerBase:0,category:'วัตถุดิบหลัก'};
  const source=workflow,materials=appState.materials;
  masterSaving=true;$w('masterFields').disabled=true;status.textContent='กำลังบันทึก…';
  try{
    materials.push(mat);
    source.settings[code]={name,unit:main,subUnit:sub,countDivisor:sub?ratio:1,inFactor:1,outFactor:1,critical:false,countUnits:[{label:main,factor:1},...(sub?[{label:sub,factor:1/ratio}]:[])]};
    await storeWorkflow();
  }catch(e){const index=materials.indexOf(mat);if(index>=0)materials.splice(index,1);delete source.settings[code];status.textContent='บันทึกไม่สำเร็จ: '+e.message;return;}
  finally{masterSaving=false;$w('masterFields').disabled=false;}
  $w('materialMasterDialog').close();renderCriticalSettings();renderBomModule();
  $w('ingredientSelectInput').value=id;filterIngredientOptions();
}
