'use strict';
// Remember tokens in this tab only; never store the password or membership.
const cloudSessionKey='bom-session:'+BOM_SUPABASE.url;
function cloudRememberSession() {
  try {sessionStorage.setItem(cloudSessionKey,JSON.stringify(bomCloud.session));} catch { /* Storage may be disabled. */ }
}
const bomCloud = { session:null, member:null, coverage:[], baseline:new Map(), initializing:false, busy:false, queue:Promise.resolve() };
function cloudStatus(message) { $w('workflowStorage').textContent=message; $w('cloudMessage').textContent=message; }
async function cloudRequest(path, options={}, authenticated=true) {
  if(authenticated && !bomCloud.session) throw Error('กรุณาเข้าสู่ระบบก่อน');
  if(authenticated && Date.now()>bomCloud.session.expiresAt-60000) {
    const fresh=await cloudRequest('/auth/v1/token?grant_type=refresh_token',{method:'POST',body:JSON.stringify({refresh_token:bomCloud.session.refresh_token})},false);
    bomCloud.session={...fresh,expiresAt:Date.now()+fresh.expires_in*1000};
    cloudRememberSession();
  }
  const response=await fetch(BOM_SUPABASE.url+path,{...options,headers:{apikey:BOM_SUPABASE.publishableKey,'Content-Type':'application/json',
    ...(authenticated?{Authorization:'Bearer '+bomCloud.session.access_token}:{}),...options.headers}});
  const body=await response.text();
  let data;try{data=body?JSON.parse(body):null;}catch{data=null;}
  if(!response.ok) {
    if(data?.code==='40001') throw Error('มีผู้อื่นบันทึกข้อมูลชุดนี้แล้ว กดโหลดข้อมูลล่าสุดแล้วกรอกใหม่ เพื่อไม่เขียนทับกัน');
    throw Error(data?.message||data?.msg||data?.error_description||`Supabase HTTP ${response.status}`);
  }
  return data;
}
function cloudSnapshot() { return structuredClone({...workflow,materials:appState.materials,recipes:appState.recipes}); }
function cloudDocuments(snapshot) {
  const {batches={},counts={},...master}=snapshot;
  const docs=new Map([['master',{id:'master',kind:'master',branch:'',period:'',payload:master}]]);
  for(const batch of Object.values(batches)) {
    for(const branch of batch.branches) {
      const id=`${batch.type}:${branch}:${batch.month}`;
      docs.set(id,{id,kind:batch.type,branch,period:batch.month,payload:{...batch,branches:[branch],rows:batch.rows.filter(r=>r.branch===branch)}});
    }
  }
  for(const [key,payload] of Object.entries(counts)) {
    const at=key.lastIndexOf(':'),branch=key.slice(0,at),period=key.slice(at+1),id=`count:${key}`;
    docs.set(id,{id,kind:'count',branch,period,payload});
  }
  return docs;
}
function cloudRestore(docs) {
  const master=docs.get('master')?.payload;
  workflow=structuredClone(master||{version:1,settings:{},branches:{}});
  workflow.batches={};workflow.counts={};
  appState.materials=workflow.materials||[];appState.recipes=workflow.recipes||[];
  for(const row of docs.values()) {
    if(row.kind==='master') continue;
    if(row.kind==='count') workflow.counts[countKey(row.branch,row.period)]=structuredClone(row.payload);
    else {
      const key=`${row.kind}:${row.period}`,p=structuredClone(row.payload);
      if(!workflow.batches[key]) workflow.batches[key]={...p,branches:[],rows:[]};
      workflow.batches[key].branches.push(row.branch);workflow.batches[key].rows.push(...p.rows);
    }
  }
  appState.selectedMenuId=appState.recipes[0]?.menuId;
}
function cloudRender() {
  for(const key of Object.keys(countDrafts))delete countDrafts[key];
  activeCountScope=null;activeCountCodes=[];
  refreshWorkflowSelectors();renderImportHistory();renderCriticalSettings();renderBomModule();handleImportTypeChange();
  const branch=bomCloud.member?.role==='branch'?bomCloud.member.branch:null;
  for(const id of ['stockCountBranchInput','varianceBranchInput']) {
    if(branch){$w(id).innerHTML=`<option value="${esc(branch)}">${esc(workflow.branches[branch]||branch)} (${esc(branch)})</option>`;$w(id).value=branch;}
  }
  $w('nav-tab-bom').hidden=false;
  $w('nav-tab-import').hidden=Boolean(branch);
  $w('nav-tab-restore').hidden=bomCloud.member?.role!=='admin';
  if(branch)switchTab('tab-count');
  $w('cloudMigrate').hidden=bomCloud.member?.role!=='admin';
  $w('cloudActions').hidden=!bomCloud.member;
  renderStockCountTable();
  if(typeof refreshCleanupPeriods==='function')refreshCleanupPeriods();
}
async function cloudLoad() {
  bomCloud.coverage=[];
  let coverageWarning='';
  const docs=new Map();
  for(let offset=0;;offset+=500){
    const rows=await cloudRequest(`/rest/v1/bom_documents?select=*&order=id&limit=500&offset=${offset}`);
    rows.forEach(r=>docs.set(r.id,r));if(rows.length<500)break;
  }
  if(bomCloud.member.role==='branch') {
    try {
      for(let offset=0;;offset+=500) {
        const rows=await cloudRequest(`/rest/v1/rpc/bom_import_coverage?order=kind,period,start_date,end_date,quantity_field&limit=500&offset=${offset}`,{method:'POST',body:'{}'});
        bomCloud.coverage.push(...rows);if(rows.length<500)break;
      }
    }
    catch {coverageWarning=' · ยังตรวจสถานะไฟล์ส่วนกลางไม่ได้ ให้ผู้ดูแลรัน supabase-branch-coverage.sql แล้วโหลดใหม่';}
  }
  bomCloud.baseline=docs;cloudRestore(docs);workflowReady=true;
  if(!docs.has('master')) {
    if(bomCloud.member.role!=='admin') {workflowReady=false;throw Error('รอส่วนกลางตั้งค่าฐานข้อมูลครั้งแรก');}
    bomCloud.initializing=true;
    try {await applyPonnTemplate();await applyCountUnitList();} finally {bomCloud.initializing=false;}
    await cloudSave(cloudSnapshot());
  }
  if(bomCloud.member.role==='admin' && typeof applyCountUnitList==='function')await applyCountUnitList();
  cloudRender();cloudStatus('เชื่อมต่อ Supabase · โหลดข้อมูลล่าสุดแล้ว'+coverageWarning);
}
function cloudSave(snapshot) {
  if(bomCloud.maintenance)return Promise.reject(Error('กำลังล้างข้อมูลย้อนหลัง กรุณารอสักครู่'));
  if(bomCloud.initializing)return Promise.resolve();
  const save=async()=>{
    if(!bomCloud.member)throw Error('ยังไม่เชื่อมต่อฐานข้อมูล กรุณากดโหลดข้อมูลล่าสุด');
    const docs=cloudDocuments(snapshot),changes=[];
    for(const [id,doc] of docs) {
      const old=bomCloud.baseline.get(id);
      if(!old || JSON.stringify(doc.payload)!==JSON.stringify(old.payload))changes.push({...doc,expected:old?.revision??null});
    }
    for(const [id,old] of bomCloud.baseline)if(!docs.has(id))changes.push({...old,expected:old.revision,remove:true});
    if(bomCloud.member.role!=='admin' && changes.some(c=>!(c.kind==='master' && c.id==='master' && !c.remove) && !(c.kind==='count' && c.branch===bomCloud.member.branch)))throw Error('บัญชีสาขาแก้ไขสูตรกลางและยอดนับของสาขาตนเองได้เท่านั้น');
    if(!changes.length)return;
    const result=await cloudRequest('/rest/v1/rpc/bom_save',{method:'POST',body:JSON.stringify({changes})});
    for(const change of changes){if(change.remove)bomCloud.baseline.delete(change.id);else bomCloud.baseline.set(change.id,{...docs.get(change.id),revision:result.find(r=>r.id===change.id).revision});}
    cloudStatus('บันทึกลง Supabase แล้ว · '+new Date().toLocaleTimeString('th-TH'));
  };
  const pending=bomCloud.queue.catch(()=>{}).then(save);
  bomCloud.queue=pending;
  return pending;
}
async function cloudReload() {
  if(!bomCloud.session||!bomCloud.member){await initializeCloudUI();return;}
  if(bomCloud.busy || !confirm('โหลดข้อมูลล่าสุดจากฐานข้อมูล? ข้อมูลที่กรอกแต่ยังไม่ได้บันทึกจะถูกยกเลิก'))return;
  bomCloud.busy=true;
  try{await bomCloud.queue;await cloudLoad();}catch(error){cloudStatus(error.message);}finally{bomCloud.busy=false;}
}
async function cloudMigrateLocal() {
  if(bomCloud.busy||bomCloud.member?.role!=='admin')return;
  if(!confirm('นำข้อมูลเดิมจากเบราว์เซอร์นี้ขึ้น Supabase? จะทำได้เมื่อฐานข้อมูลยังไม่มีไฟล์นำเข้าหรือยอดตรวจนับ และจะแทนที่สูตร/การตั้งค่าด้วยข้อมูลเดิม'))return;
  bomCloud.busy=true;
  try {
    await bomCloud.queue;await cloudLoad();
    if([...bomCloud.baseline.values()].some(r=>r.kind!=='master'))throw Error('ฐานข้อมูลมีข้อมูลใช้งานแล้ว จึงไม่ย้ายทับ');
    await openWorkflowStore();
    const saved=await new Promise((resolve,reject)=>{const req=workflowDb.transaction('state').objectStore('state').get('current');req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error);});
    if(!saved)throw Error('ไม่พบข้อมูลเดิมในเบราว์เซอร์นี้');
    await cloudSave(saved);cloudRestore(bomCloud.baseline);await applyCountUnitList();cloudRender();cloudStatus('ย้ายข้อมูลเดิมขึ้น Supabase แล้ว');
  }catch(error){cloudStatus(error.message);}finally{bomCloud.busy=false;}
}
async function initializeCloudUI() {
  let saved=null;
  try {saved=JSON.parse(sessionStorage.getItem(cloudSessionKey));} catch { /* No usable saved session. */ }
  cloudClearSession();
  cloudStatus('กรุณาเข้าสู่ระบบเพื่อเปิดข้อมูล');
  if(!saved?.access_token||!saved?.refresh_token||!Number.isFinite(saved.expiresAt))return;
  bomCloud.busy=true;$w('cloudSignInButton').disabled=true;
  cloudStatus('กำลังตรวจสอบการเข้าสู่ระบบ…');
  try {
    bomCloud.session=saved;
    const user=await cloudRequest('/auth/v1/user');
    bomCloud.session.user=user;
    await cloudLoadMembership(user.id);
    await cloudLoad();cloudRememberSession();cloudAuthView(true);
  } catch(error) {
    cloudClearSession();cloudStatus('เข้าสู่ระบบต่อไม่สำเร็จ กรุณาเข้าสู่ระบบใหม่: '+error.message);
  } finally {bomCloud.busy=false;$w('cloudSignInButton').disabled=false;}
}
async function cloudLoadMembership(userId) {
  const members=await cloudRequest('/rest/v1/bom_members?select=role,branch&user_id=eq.'+encodeURIComponent(userId));
  if(members.length!==1||!['admin','branch'].includes(members[0].role))throw Error('บัญชีนี้ยังไม่ได้รับสิทธิ์ใช้งาน กรุณาติดต่อส่วนกลาง');
  bomCloud.member=members[0];
}
function cloudAuthView(unlocked) {
  document.body.classList.toggle('auth-locked',!unlocked);
  document.querySelectorAll('section.tab-content, header').forEach(el=>el.inert=!unlocked);
  $w('cloudLogin').hidden=unlocked;
  $w('cloudGate').hidden=unlocked;
  $w('cloudActions').hidden=!unlocked;
  $w('cloudActions').open=false;
  $w('cloudAccount').textContent=unlocked
    ? `${bomCloud.session.user.email} · ${bomCloud.member.role==='admin'?'ส่วนกลาง / ผู้ดูแลระบบ':'สาขา '+bomCloud.member.branch}` : '';
  $w('cloudAccount').title=$w('cloudAccount').textContent;
}
function cloudClearSession() {
  if(typeof resetRestore==='function')resetRestore();
  try {sessionStorage.removeItem(cloudSessionKey);} catch { /* Storage may be disabled. */ }
  bomCloud.session=null;bomCloud.member=null;bomCloud.coverage=[];bomCloud.baseline=new Map();workflowReady=false;
  workflow={version:1,settings:{},branches:{},batches:{},counts:{}};
  appState.materials=[];appState.recipes=[];appState.selectedMenuId=null;
  for(const key of Object.keys(countDrafts))delete countDrafts[key];
  activeCountScope=null;activeCountCodes=[];
  latestReport=[];visibleVarianceRows=[];
  cloudAuthView(false);
}
async function cloudSignIn() {
  if(bomCloud.busy)return;
  bomCloud.busy=true;$w('cloudSignInButton').disabled=true;
  cloudStatus('กำลังเข้าสู่ระบบ…');
  try {
    const session=await cloudRequest('/auth/v1/token?grant_type=password',{
      method:'POST',body:JSON.stringify({email:$w('cloudEmail').value.trim(),password:$w('cloudPassword').value})
    },false);
    bomCloud.session={...session,expiresAt:Date.now()+session.expires_in*1000};
    await cloudLoadMembership(session.user.id);
    await cloudLoad();cloudRememberSession();cloudAuthView(true);
  } catch(error) {
    cloudClearSession();
    cloudStatus(/Invalid login credentials/i.test(error.message)?'อีเมลหรือรหัสผ่านไม่ถูกต้อง':'เข้าสู่ระบบไม่สำเร็จ: '+error.message);
  } finally {
    $w('cloudPassword').value='';bomCloud.busy=false;$w('cloudSignInButton').disabled=false;
  }
}
async function cloudSignOut() {
  if(bomCloud.busy||bomCloud.maintenance)return;
  if(!confirm('ออกจากระบบ? ข้อมูลที่ยังไม่ได้กดบันทึกจะถูกยกเลิก'))return;
  bomCloud.busy=true;
  try {
    await bomCloud.queue.catch(()=>{});
    try {await cloudRequest('/auth/v1/logout?scope=local',{method:'POST'});}catch { /* Clear this device even when offline. */ }
    cloudClearSession();cloudStatus('ออกจากระบบแล้ว');
  } finally {bomCloud.busy=false;}
}
