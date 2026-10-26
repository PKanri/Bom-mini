/* A branch-only working preview. No cloud writes or changes to official imports. */
'use strict';
const branchCheckLabels={SALES:'ยอดขาย',STOCK_IN:'itec in',STOCK_OUT:'itec out',EXTRA_IN:'เครื่องดื่ม in'};
let branchCheckState={scope:'',files:{},counts:{},rows:[],generation:0};
function resetBranchCheck() {
  branchCheckState={scope:'',files:{},counts:{},rows:[],generation:branchCheckState.generation+1};
  if($w('branchSaveStatus'))$w('branchSaveStatus').textContent='';
  const body=$w('branchCheckBody');if(body)body.innerHTML='';
  if($w('branchCheckSummary'))$w('branchCheckSummary').textContent='';
  if($w('branchCheckSources'))$w('branchCheckSources').innerHTML='';
  if($w('branchFileDetailHead'))$w('branchFileDetailHead').innerHTML='';
  if($w('branchFileDetailBody'))$w('branchFileDetailBody').innerHTML='';
  if($w('branchFileDetailInfo'))$w('branchFileDetailInfo').textContent='ยังไม่ได้อัปโหลดไฟล์';
  if($w('branchFileDetailSearch'))$w('branchFileDetailSearch').value='';
  if($w('branchFileDetailPage'))$w('branchFileDetailPage').textContent='';
  if($w('branchFileDetailPrev'))$w('branchFileDetailPrev').disabled=true;
  if($w('branchFileDetailNext'))$w('branchFileDetailNext').disabled=true;
}
function branchCheckScope() {
  if(!workflowReady)throw Error('กรุณารอโหลดข้อมูลและเข้าสู่ระบบก่อน');
  const member=typeof bomCloud!=='undefined'?bomCloud.member:null;
  if(typeof BOM_SUPABASE!=='undefined' && !member)throw Error('กรุณาเข้าสู่ระบบ');
  const branch=member?.role==='branch'?member.branch:$w('branchCheckBranch').value;
  const date=normalizedDate($w('branchCheckDate').value);
  if(!branch || branch==='ALL')throw Error('เลือกสาขาก่อน');
  return {branch,date,key:branch+':'+date};
}
function openBranchCheck() {
  const member=typeof bomCloud!=='undefined'?bomCloud.member:null;
  const selected=$w('branchCheckBranch').value;
  const branches=member?.role==='branch'?[[member.branch,workflow.branches[member.branch]||member.branch]]:Object.entries(workflow.branches).sort(compareBranchIds);
  $w('branchCheckBranch').innerHTML='<option value="">เลือกสาขา</option>'+branches.map(([id,name])=>`<option value="${esc(id)}">${esc(name)} (${esc(id)})</option>`).join('');
  $w('branchCheckBranch').value=member?.role==='branch'?member.branch:selected;
  $w('branchCheckBranch').disabled=member?.role==='branch';
  if(!$w('branchCheckDate').value){const now=new Date();$w('branchCheckDate').value=`${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')}`;}
  renderBranchCheck();
}
function branchCheckDate(value) {
  if(typeof value==='number' || /^\d{4}-/.test(String(value)))return normalizedDate(value);
  // iTech exports Thai weekday prefixes and Buddhist years in day/month/year order.
  const match=String(value).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if(!match)throw Error('อ่านวันที่ไม่ได้: '+value);
  return normalizedDate(`${match[3]}-${match[2].padStart(2,'0')}-${match[1].padStart(2,'0')}`);
}
function parseBranchCheckFile(book,type,scope,allowUnscoped=false,allowEmpty=false) {
  const preferred={SALES:'ยอดขาย (วาง)',STOCK_IN:'iTec In (วาง)',STOCK_OUT:'iTec Out (วาง)',EXTRA_IN:'เครื่องดื่ม In (วาง)'}[type];
  const candidates=[];
  for(const name of book.SheetNames){
    const sheet=book.Sheets[name],range=XLSX.utils.decode_range(sheet['!ref']||'A1');
    // Templates may have formatting extending to XFD. Only source columns matter.
    range.e.c=Math.min(range.e.c,79);
    const grid=XLSX.utils.sheet_to_json(sheet,{header:1,defval:null,raw:true,range});
    const index=grid.slice(0,10).findIndex(row=>{const r=row.map(v=>String(v??'').trim());return r.includes('Product (Code)')&&r.includes(type==='STOCK_IN'&&r.includes('Branch To (ID)')&&!r.includes('BranchFrom (ID)')?'Sum Number':sourceQuantityField(type));});
    if(index<0)continue;
    const headers=grid[index].map(v=>String(v??'').trim());
    const marker=type==='STOCK_OUT'?'Branch To (ID)':type==='STOCK_IN'?'BranchFrom (ID)':type==='SALES'?'ราคาขาย/หน่วย':'Supplier (Name)';
    if(name!==preferred && !headers.includes(marker) && !(type==='STOCK_IN'&&headers.includes('Branch To (ID)')))continue;
    candidates.push({name,headers,grid:grid.slice(index+1),offset:range.s.r+index+2,columnOffset:range.s.c});
  }
  const matches=candidates.filter(c=>c.name===preferred);
  const chosen=matches.length?matches:candidates;
  if(chosen.length!==1)throw Error(chosen.length?'พบหลายตาราง กรุณาใช้ไฟล์ประเภทเดียว':'ไม่พบตาราง '+branchCheckLabels[type]+' ที่มีคอลัมน์ '+sourceQuantityField(type));
  const {name,headers,grid,offset,columnOffset}=chosen[0],rows=[],detailRows=[];
  if(type==='STOCK_IN' && !headers.includes('BranchFrom (ID)') && !headers.includes('Branch To (ID)'))
    throw Error('ไฟล์ itec in ต้องมี BranchFrom (ID) หรือ Branch To (ID)');
  const quantityField=type==='STOCK_IN'&&headers.includes('Branch To (ID)')&&!headers.includes('BranchFrom (ID)')?'Sum Number':sourceQuantityField(type);
  if(type==='STOCK_OUT' && !headers.includes('Branch To (ID)'))throw Error('ไฟล์จ่ายออกต้องมี Branch To (ID)');
  // Branch sales summaries omit branch/date on every row. The selected scope
  // owns such a file; mixed scoped/unscoped exports still require confirmation.
  const codeColumn=headers.indexOf('Product (Code)'),branchColumn=headers.indexOf('Branch (ID)');
  const aggregateSales=type==='SALES' && grid.filter(row=>String(row[codeColumn]??'').trim())
    .every(row=>!String(row[branchColumn]??'').trim());
  let skipped=0,unscoped=0;
  for(let i=0;i<grid.length;i++){
    const r=Object.fromEntries(headers.map((h,j)=>[h,grid[i][j]]));
    const code=String(r['Product (Code)']??'').trim();
    if(!code)continue;
    const branch=String(r['Branch (ID)']??'').trim();
    if(branch && branch!==scope.branch){skipped++;continue;}
    // Branch previews intentionally assign every matching row to the selected day.
    if(!branch){
      if(!allowUnscoped && !aggregateSales)throw Error('ตารางไม่มีรหัสสาขา กรุณายืนยันว่าไฟล์เป็นของสาขาที่เลือก แล้วเลือกไฟล์อีกครั้ง');
      unscoped++;
    }
    const raw=r[quantityField],qty=typeof raw==='number'?raw:Number(String(raw??'').replace(/,/g,''));
    if(raw==null||raw===''||!Number.isFinite(qty)||qty<0)throw Error(`แถว ${i+offset}: จำนวนไม่ถูกต้อง`);
    if(r.Status===true||String(r.Status).toLowerCase()==='true')throw Error(`แถว ${i+offset}: Status=true กรุณาตรวจเอกสาร`);
    rows.push({branch:scope.branch,date:scope.date,code,qty,name:String(r['Product (Name)']||code)});
    detailRows.push({row:i+offset,qty,values:headers.map((_,j)=>{
      const cell=book.Sheets[name][XLSX.utils.encode_cell({r:i+offset-1,c:j+columnOffset})];
      return cell?XLSX.utils.format_cell(cell):'';
    })});
  }
  if(!rows.length&&!allowEmpty)throw Error('ไม่พบรายการของสาขาที่เลือก หากไม่มีรายการจริงให้กด “วันนี้ไม่มีรายการ”');
  return {type,rows,skipped,unscoped,aggregateSales,quantityField,sheet:name,columns:headers,detailRows};
}
function parseBranchCheckWorkbook(book,scope,allowUnscoped=false) {
  const normalize=name=>name.toLowerCase().replace(/itech/g,'itec').replace(/\(วาง\)/g,'').replace(/\s/g,'');
  const files={},warnings=[];
  for(const [type,label]of Object.entries(branchCheckLabels)) {
    const names=book.SheetNames.filter(n=>normalize(n)===normalize(label));
    if(names.length!==1)throw Error('ต้องมีชีต '+label+' เพียง 1 ชีต (ไม่มีรายการให้เว้นชีตว่างไว้)');
    const name=names[0],sheet=book.Sheets[name],range=XLSX.utils.decode_range(sheet['!ref']||'A1');range.e.c=Math.min(range.e.c,79);
    const grid=XLSX.utils.sheet_to_json(sheet,{header:1,defval:null,range});
    let parsed;
    if(!grid.some(row=>row.some(v=>v!=null&&String(v).trim()!=='')))parsed={type,rows:[],detailRows:[],columns:[],sheet:name,skipped:0};
    else parsed=parseBranchCheckFile({SheetNames:[name],Sheets:{[name]:sheet}},type,scope,allowUnscoped,true);
    if(!parsed.rows.length&&parsed.skipped){warnings.push(label+': พบแต่ข้อมูลต่างสาขา กรุณาตรวจไฟล์');continue;}
    files[type]=parsed;
  }
  return {files,warnings};
}
async function loadBranchCheckWorkbook(input) {
  const file=input.files?.[0];if(!file)return;
  try {
    const scope=branchCheckScope();ensureBranchCheckScope(scope);
    const generation=branchCheckState.generation,token=Symbol('workbook');branchCheckState.workbookRead=token;
    if(typeof XLSX==='undefined')throw Error('ไม่พบตัวอ่าน Excel');
    if(file.size>100*1024*1024)throw Error('ไฟล์ใหญ่เกิน 100 MB');
    $w('branchCheckMessage').textContent='กำลังอ่าน '+file.name;
    const book=XLSX.read(await file.arrayBuffer(),{type:'array'});
    if(generation!==branchCheckState.generation||branchCheckState.workbookRead!==token||scope.key!==branchCheckScope().key)return;
    const parsed=parseBranchCheckWorkbook(book,scope,$w('branchCheckUnscoped').checked);
    branchCheckState.files=Object.fromEntries(Object.entries(parsed.files).map(([type,value])=>[type,{...value,file:file.name}]));
    branchCheckState.reads={};branchCheckState.detailPage=0;
    $w('branchFileDetailSearch').value='';$w('branchFileDetailType').value='SALES';
    $w('branchCheckMessage').textContent='อ่าน '+file.name+' แล้ว · ชีตว่างถือว่าไม่มีรายการ · ใช้วันที่ '+scope.date+(parsed.warnings.length?' · '+parsed.warnings.join(' / '):'');
    renderBranchCheck();
  }catch(e){$w('branchCheckMessage').textContent=e.message;}finally{input.value='';}
}
function clearBranchCheckWorkbook() {
  branchCheckState.workbookRead=null;branchCheckState.reads={};branchCheckState.files={};
  $w('branchCheckMessage').textContent='ล้างไฟล์แล้ว กรุณาเลือกไฟล์ใหม่';renderBranchCheck();
}
async function loadBranchCheckFile(type,input) {
  const file=input.files?.[0];if(!file)return;
  try{
    const scope=branchCheckScope();ensureBranchCheckScope(scope);
    const generation=branchCheckState.generation;
    const token=Symbol(type);
    branchCheckState.reads||={};branchCheckState.reads[type]=token;
    if(typeof XLSX==='undefined')throw Error('ไม่พบตัวอ่าน Excel');
    if(file.size>100*1024*1024)throw Error('ไฟล์ใหญ่เกิน 100 MB');
    $w('branchCheckMessage').textContent='กำลังอ่าน '+file.name;
    const book=XLSX.read(await file.arrayBuffer(),{type:'array'});
    if(generation!==branchCheckState.generation || branchCheckState.reads?.[type]!==token || scope.key!==branchCheckScope().key)return;
    const parsed=parseBranchCheckFile(book,type,scope,$w('branchCheckUnscoped').checked);
    branchCheckState.files[type]={...parsed,file:file.name};
    $w('branchFileDetailType').value=type;
    $w('branchFileDetailSearch').value='';
    branchCheckState.detailPage=0;
    $w('branchCheckMessage').textContent=`อ่าน ${file.name}: ${parsed.rows.length} แถว · ข้ามต่างสาขา ${parsed.skipped} แถว · ใช้วันที่ ${scope.date} ตามหน้าจอ`+(parsed.aggregateSales?` · ยอดขายสรุป ใช้สาขา ${scope.branch} ที่เลือก`:parsed.unscoped?' · ใช้สาขาตามที่ยืนยัน':'');
    renderBranchCheck();
  }catch(error){$w('branchCheckMessage').textContent=error.message;}finally{input.value='';}
}
function ensureBranchCheckScope(scope) {
  if(branchCheckState.scope!==scope.key){
    resetBranchCheck();branchCheckState.scope=scope.key;
    $w('branchCheckUnscoped').checked=false;
    $w('branchCheckMessage').textContent='เลือกไฟล์ใหม่เมื่อเปลี่ยนสาขาหรือวันที่ ข้อมูลทดลองของวันเดิมถูกล้างแล้ว';
  }
}
async function zeroBranchCheck(type) {
  try{
    const scope=branchCheckScope();ensureBranchCheckScope(scope);
    if(!await appConfirm(`ยืนยันว่าสาขา ${scope.branch} วันที่ ${scope.date} ไม่มีรายการ${branchCheckLabels[type]}?`))return;
    if(branchCheckState.reads)delete branchCheckState.reads[type];
    branchCheckState.files[type]={type,rows:[],file:'ยืนยันไม่มีรายการ'};
    renderBranchCheck();
  }catch(e){$w('branchCheckMessage').textContent=e.message;}
}
function clearBranchCheckFile(type) {if(branchCheckState.reads)delete branchCheckState.reads[type];delete branchCheckState.files[type];renderBranchCheck();}
function branchCheckInput(index,field,part,value) {
  $w('branchSaveStatus').textContent='มีการแก้ไขยอด กรุณาบันทึกอีกครั้ง';
  const [code,s]=orderedCountEntries()[index]||[];
  if(!code)return;
  const entry=branchCheckState.counts[code]||={};
  (entry[field]||=branchCheckValues(branchCheckScope(),code,s,field).slice())[part]=value;
  calculateBranchCheck();
}
function branchCheckValues(scope,code,s,field) {
  const units=countUnits(s,code);
  const saved=workflow.counts[countKey(scope.branch,field==='opening'?priorDay(scope.date):scope.date)]?.[code];
  if(field==='waste')return [saved?.unitReview?'':saved?.waste??''];
  return displayCountInputs(saved,units,s);
}
function branchCheckQuantity(values,units) {
  if(!values.some(v=>v!==''&&v!=null))return null;
  let total=0;
  for(let i=0;i<values.length;i++){
    if(values[i]===''||values[i]==null)continue;
    const n=Number(values[i]);if(!Number.isFinite(n)||n<0)throw Error('ยอดตรวจนับต้องเป็นตัวเลขตั้งแต่ 0 ขึ้นไป');
    total+=n*units[i].factor;
  }
  if(!Number.isFinite(total))throw Error('จำนวนมากเกินไป');
  return total;
}
function branchCheckModel(scope) {
  const source={...workflow,batches:{},counts:{...workflow.counts}};
  const opening={},closing={};
  for(const [code,s]of orderedCountEntries()){
    const units=countUnits(s,code),draft=branchCheckState.counts[code]||{};
    const qty=field=>branchCheckQuantity(draft[field]??branchCheckValues(scope,code,s,field),field==='waste'?[{factor:1}]:units);
    const start=qty('opening'),end=qty('actual'),waste=qty('waste');
    if(start!=null)opening[code]={qty:start};
    if(end!=null)closing[code]={qty:end,waste:waste??0};
  }
  source.counts[countKey(scope.branch,priorDay(scope.date))]=opening;
  source.counts[countKey(scope.branch,scope.date)]=closing;
  for(const [type,file]of Object.entries(branchCheckState.files)){
    source.batches[type+':'+scope.date.slice(0,7)]={type,start:scope.date,end:scope.date,branches:[scope.branch],rows:file.rows,quantityField:sourceQuantityField(type)};
  }
  return buildReport(scope.branch,scope.date,scope.date,source,false);
}
function calculateBranchCheck() {
  try{
    const scope=branchCheckScope(),rows=branchCheckModel(scope);
    branchCheckState.rows=rows;
    rows.forEach((r,index)=>{
      const d=stockDifference(r);
      for(const [field,value]of [['in',r.in],['out',r.out],['bom',r.bom],['theory',r.theory],['diff',d.qty]])$w(`bc-${field}-${index}`).textContent=fmt(value);
    });
    const ready=rows.filter(r=>r.ready);
    $w('branchCheckSummary').textContent=`คำนวณได้ ${ready.length} / ${rows.length} รายการ · มีผลต่าง ${ready.filter(r=>Math.abs(r.overuse)>1e-6).length} รายการ · รอข้อมูล ${rows.length-ready.length} รายการ`;
  }catch(e){branchCheckState.rows=[];$w('branchCheckSummary').textContent=e.message;
    for(const cell of document.querySelectorAll('#branchCheckBody [data-result]'))cell.textContent='—';
  }
}
function renderBranchCheck() {
  try{
    const scope=branchCheckScope();ensureBranchCheckScope(scope);
    $w('branchOpeningDate').textContent='ยอดบันทึก '+priorDay(scope.date);
    $w('branchCheckSources').innerHTML=Object.keys(branchCheckLabels).map(type=>{const file=branchCheckState.files[type];return `<div class="bc-source"><strong>${esc(branchCheckLabels[type])}</strong><p>${file?esc(file.rows.length?file.rows.length+' รายการ · '+file.sheet:'ไม่มีรายการ (ชีตว่าง)'):'รอไฟล์ของสาขา'}</p></div>`;}).join('');
    $w('branchCheckBody').innerHTML=orderedCountEntries().map(([code,s],index)=>{
      const units=countUnits(s,code);
      const inputs=field=>{
        const specs=field==='waste'?[{label:s.unit,factor:1}]:units;
        const values=branchCheckState.counts[code]?.[field]??branchCheckValues(scope,code,s,field);
        return specs.map((u,part)=>`<label class="bc-count">${esc(u.label)}<input aria-label="${esc(s.name+' '+field+' '+u.label)}" type="number" min="0" step="any" placeholder="${field==='waste'?'0':'ยังไม่นับ'}" value="${esc(values[part]??'')}" oninput="branchCheckInput(${index},'${field}',${part},this.value)"></label>`).join('');
      };
      return `<tr><td><strong>${esc(s.name)}</strong><small>${esc(code)} · รายงานเป็น ${esc(s.unit)}</small></td><td>${inputs('opening')}</td><td id="bc-in-${index}" data-result></td><td id="bc-out-${index}" data-result></td><td id="bc-bom-${index}" data-result></td><td>${inputs('waste')}</td><td id="bc-theory-${index}" data-result></td><td>${inputs('actual')}</td><td id="bc-diff-${index}" data-result></td></tr>`;
    }).join('');
    calculateBranchCheck();
    renderBranchFileDetails();
  }catch(e){$w('branchCheckSummary').textContent=e.message;$w('branchCheckBody').innerHTML='';renderBranchFileDetails();}
}
function renderBranchFileDetails(resetPage=false) {
  const head=$w('branchFileDetailHead'),body=$w('branchFileDetailBody'),info=$w('branchFileDetailInfo');
  head.innerHTML='';body.innerHTML='';
  $w('branchFileDetailPrev').disabled=true;$w('branchFileDetailNext').disabled=true;
  $w('branchFileDetailPage').textContent='';
  try{
    const scope=branchCheckScope();
    if(branchCheckState.scope!==scope.key)throw Error('เลือกสาขาและวันที่ แล้วอัปโหลดไฟล์เพื่อดูรายละเอียด');
    const type=$w('branchFileDetailType').value||'SALES',file=branchCheckState.files[type];
    if(!file?.detailRows){info.textContent=file?'ยืนยันว่าไม่มีรายการ จึงไม่มีไฟล์ให้แสดง':'ยังไม่ได้อัปโหลดไฟล์ประเภทนี้ กรุณาเลือกไฟล์ของสาขา';return;}
    const query=$w('branchFileDetailSearch').value.trim().toLocaleLowerCase();
    const matched=file.detailRows.filter(r=>!query||r.values.some(v=>String(v).toLocaleLowerCase().includes(query))||String(r.row)===query);
    const pages=Math.max(1,Math.ceil(matched.length/50));
    const page=resetPage?0:Math.min(branchCheckState.detailPage||0,pages-1);
    branchCheckState.detailPage=page;
    info.textContent=`${file.file} · ชีต ${file.sheet} · รายการของสาขา ${scope.branch} ${file.detailRows.length} แถว · พบตามคำค้น ${matched.length} แถว · ข้ามต่างสาขา ${file.skipped||0} แถว · จำนวนที่ใช้คำนวณมาจาก ${file.quantityField||sourceQuantityField(type)}`;
    head.innerHTML='<tr><th>แถวใน Excel</th><th>จำนวนที่ใช้คำนวณ</th>'+file.columns.map((name,i)=>`<th>${esc(name||'คอลัมน์ '+(i+1))}</th>`).join('')+'</tr>';
    body.innerHTML=matched.slice(page*50,(page+1)*50).map(r=>'<tr><td>'+r.row+'</td><td>'+esc(fmt(r.qty))+'</td>'+r.values.map(v=>'<td>'+esc(v)+'</td>').join('')+'</tr>').join('')||`<tr><td colspan="${file.columns.length+2}">ไม่พบรายการตามคำค้น</td></tr>`;
    $w('branchFileDetailPage').textContent=`หน้า ${page+1} / ${pages} · หน้าละ 50 รายการ`;
    $w('branchFileDetailPrev').disabled=page===0;
    $w('branchFileDetailNext').disabled=page>=pages-1;
  }catch(e){info.textContent=e.message;}
}
function changeBranchFileDetailPage(step) {
  branchCheckState.detailPage=Math.max(0,(branchCheckState.detailPage||0)+step);
  renderBranchFileDetails();
}
function exportBranchCheck() {
  calculateBranchCheck();
  if(!branchCheckState.rows.length)return;
  const scope=branchCheckScope(),book=XLSX.utils.book_new();
  const rows=branchCheckState.rows.map(r=>({วันที่:scope.date,สาขา:scope.branch,รหัส:r.code,วัตถุดิบ:r.name,หน่วย:r.unit,ยกมา:r.opening??'',รับเข้า:r.in,จ่ายออก:r.out,ใช้ตามสูตร:r.bom,สูญเสีย:r.waste,ควรเหลือ:r.theory??'',นับจริง:r.actual??'',ผลต่าง:stockDifference(r).qty??'',สถานะ:r.ready?'ตรวจแล้ว':r.issues.join(' / '),ประเภท:'ผลตรวจจากไฟล์สาขา (ไฟล์เคลื่อนไหวใช้ตรวจในหน้านี้)'}));
  XLSX.utils.book_append_sheet(book,XLSX.utils.json_to_sheet(rows),'ตรวจสิ้นวัน');
  XLSX.writeFile(book,`ตรวจสิ้นวัน_${scope.branch}_${scope.date}.xlsx`);
}

let branchCountSaving=false;
async function saveBranchCheckCounts() {
  if(branchCountSaving)return;
  let key,previous,next,scope,generation;
  try {
    scope=branchCheckScope();generation=branchCheckState.generation;
    if(branchCheckState.scope!==scope.key)throw Error('เลือกสาขาและวันที่ก่อนบันทึก');
    key=countKey(scope.branch,scope.date);previous=workflow.counts[key];next=structuredClone(previous||{});
    const entries=orderedCountEntries();let entered=0;
    for(const [code,s]of entries){
      const units=countUnits(s,code),draft=branchCheckState.counts[code]||{};
      const values=draft.actual??branchCheckValues(scope,code,s,'actual');
      const qty=branchCheckQuantity(values,units);
      const waste=branchCheckQuantity(draft.waste??branchCheckValues(scope,code,s,'waste'),[{factor:1}])??0;
      if(qty==null){if(waste)throw Error('กรอกยอดนับจริงของ '+s.name+' ก่อนบันทึกของสูญเสีย');continue;}
      entered++;
      next[code]={...(previous?.[code]||{}),qty,unitInputs:units.map((u,i)=>({...u,value:values[i]??''})),waste,loss:String(waste/units[0].factor),lossUnit:units[0].label,updatedAt:new Date().toISOString()};
      delete next[code].unitReview;
    }
    if(!entered)throw Error('กรอกยอดนับจริงอย่างน้อย 1 รายการก่อนบันทึก');
    const warnings=countAnomalyWarnings(scope.branch,scope.date,next,previous||{},entries);
    if(!await appConfirm(`บันทึกยอดนับจริง ${entered} รายการ · สาขา ${scope.branch} · วันที่ ${scope.date} หรือไม่?\nบันทึกยอดนับและสูญเสีย ช่องที่ยังไม่กรอกจะคงข้อมูลเดิมไว้ ยอดยกมาและไฟล์อัปโหลดไม่ถูกบันทึก${warnings.length?'\n\nรายการที่ควรตรวจซ้ำ:\n'+warnings.join('\n'):''}`))return;
    branchCountSaving=true;$w('branchSaveCounts').disabled=true;
    workflow.counts[key]=next;
    try{await storeWorkflow();}catch(e){if(workflow.counts[key]===next){if(previous===undefined)delete workflow.counts[key];else workflow.counts[key]=previous;}throw e;}
    delete countDrafts[key];
    if(activeCountScope===key)renderStockCountTable();
    if(branchCheckState.generation===generation&&branchCheckState.scope===scope.key)
      $w('branchSaveStatus').textContent=`บันทึกยอดนับจริงแล้ว ${entered} รายการ · สาขา ${scope.branch} · ${scope.date} ดูยอดนี้ได้ในหน้า 3`;
  }catch(e){$w('branchSaveStatus').textContent='บันทึกไม่สำเร็จ: '+e.message;}
  finally{branchCountSaving=false;$w('branchSaveCounts').disabled=false;}
}
