// The report records these contributions while calculating each total.
function showReportDetails(index,field) {
  openCalculationDetails(latestReport[index],field);
}
function showBranchReportDetails(index,field) {
  openCalculationDetails(branchCheckState.rows[index],field,branchCheckState.files);
}
function openCalculationDetails(source,field,files) {
  const titles=files?{in:'itec in + เครื่องดื่ม in',out:'itec out',bom:'ใช้ตาม BOM'}:{in:'รับเข้า',out:'จ่ายออก',waste:'สูญเสีย',bom:'ตัดตามสูตร BOM'};
  if(!source||!titles[field])return;
  const row=files?source:reportDisplayRow(source);
  const factor=files?1:countUnits(workflow.settings[source.code],source.code)[0]?.factor||1;
  const excelRows=new Map();
  if(files){
    for(const [type,file] of Object.entries(files)){
      excelRows.set(type,new Map((file.detailRows||[]).map(record=>[record.row,record])));
    }
  }
  function sourceCell(entry) {
    if(!files)return esc(field==='waste'?'ยอดสูญเสียที่สาขาบันทึก':labels[entry.type]||entry.type);
    const file=files[entry.type],record=excelRows.get(entry.type)?.get(entry.excelRow);
    if(!file||!record)return esc(branchCheckLabels[entry.type]);
    const cells=file.columns.map((label,i)=>label?`<div><b>${esc(label)}</b> · ${esc(record.values[i]??'')}</div>`:'').join('');
    return `<details><summary>${esc(file.sheet)} · แถว ${record.row}</summary><small>${esc(file.file||'')}</small>${cells}</details>`;
  }
  const entries=source.details?.[field]||[];
  const number=n=>Number(n).toLocaleString('th-TH',{maximumFractionDigits:8});
  const date=d=>typeof displayDate==='function'?displayDate(d):d;
  const dialog=document.createElement('dialog');
  dialog.className='ponn-dialog report-details-dialog';
  dialog.setAttribute('aria-labelledby','report-details-title');
  const isBom=field==='bom';
  const material=isBom?appState.materials.find(mat=>mat.code===source.code):null;
  const header=isBom
    ? ['วันที่ / แหล่งข้อมูล','เมนูที่ขาย','จำนวนขาย','ปริมาณสุทธิต่อจาน','ใช้จริงรวม','ยอดที่ระบบตัด']
    : field==='waste' ? ['วัตถุดิบ','หมายเหตุ','วันที่ / แหล่งข้อมูล','ยอดที่ระบบคิด']
    : ['วันที่ / แหล่งข้อมูล','วัตถุดิบ','จำนวนจากไฟล์','อัตราแปลง','ยอดที่ระบบคิด'];
  const body=entries.map(e=>{
    // Movement files carry their own product name; show the master-data name (same as the report row) and keep the file's name as a note.
    const useMaster=!isBom&&field!=='waste'&&row.name;
    const fileName=useMaster&&e.name&&e.name!==row.name&&e.name!==e.code?`<br><small>ชื่อในไฟล์: ${esc(e.name)}</small>`:'';
    const product=`<td><strong>${esc(useMaster?row.name:(e.name||e.code))}</strong><br><small>${esc(e.code)}</small>${fileName}</td>`;
    const origin=`<td>${esc(date(e.date))}<br>${sourceCell(e)}</td>`;
    const system=`<td>${number(e.amount/factor)} ${esc(row.unit)}</td>`;
    if(isBom){
      const recipeQty=(amount)=>material
        ? recipeQuantityText({...e,netQty:amount},material)
        : number(amount/factor)+' '+esc(row.unit);
      const yieldInfo=e.yieldRate!==100?`<br><small>Yield ${number(e.yieldRate)}% · เบิกจริง ${recipeQty(e.factor)} ต่อจาน</small>`:'';
      return `<tr>${origin}${product}<td>${number(e.qty)} หน่วย</td><td><strong>${recipeQty(e.netQty)}</strong>${yieldInfo}</td><td>${recipeQty(e.amount)}</td>${system}</tr>`;
    }
    const amounts=field==='waste'?`<td>${esc(e.remarks||'—')}</td>`:`<td>${number(e.qty)} ${esc(e.sourceUnit||'')}</td><td>${number(e.factor/factor)}</td>`;
    if(field==='waste')return `<tr>${product}${amounts}${origin}${system}</tr>`;
    return `<tr>${origin}${product}${amounts}${system}</tr>`;
  }).join('');
  dialog.innerHTML=`<h2 id="report-details-title">รายละเอียด${titles[field]}</h2>
    <p><strong>${esc(row.name)}</strong> · ${esc(row.code)}</p>
    <p>สาขา ${esc(workflow.branches[row.branch]||row.branch)} (${esc(row.branch)}) · ${esc(date(row.start))} ถึง ${esc(date(row.end))}</p>
    <p>หน่วยที่ระบบใช้คำนวณ: <strong>${esc(row.unit)}</strong></p>
    <div class="report-details-table"><table><thead><tr>${header.map(label=>`<th>${label}</th>`).join('')}</tr></thead><tbody>${body||`<tr><td colspan="${header.length}">ไม่มีรายการที่นำมาคำนวณในช่วงนี้</td></tr>`}</tbody></table></div>
    <p class="report-details-total">รวม ${number(row[field])} ${esc(row.unit)}</p>
    <p class="report-details-note">${field==='bom'?'ยอดที่นำมาคิด = จำนวนขาย × ปริมาณสุทธิต่อจาน ÷ (Yield ÷ 100)':field==='waste'?'แสดงยอดสูญเสียที่บันทึก โดยแปลงเป็นหน่วยเดียวกับตารางรายงาน':files?'ยอดที่นำมาคิด = จำนวนจากไฟล์ × อัตราแปลง':'ยอดที่นำมาคิด = จำนวนจากไฟล์ × อัตราแปลง · ข้อมูลที่นำเข้าถูกรวมตามวันที่ สาขา และรหัสสินค้าแล้ว'}</p>
    <div class="ponn-dialog-actions"><button type="button" autofocus>ปิด</button></div>`;
  dialog.querySelector('button').onclick=()=>dialog.close();
  dialog.addEventListener('close',()=>dialog.remove(),{once:true});
  dialog.addEventListener('click',event=>{if(event.target===dialog){const r=dialog.getBoundingClientRect();if(event.clientX<r.left||event.clientX>r.right||event.clientY<r.top||event.clientY>r.bottom)dialog.close();}});
  document.body.appendChild(dialog);dialog.showModal();
}
(function(){
  const style=document.createElement('style');
  style.textContent='.report-detail-number{font:inherit;color:inherit;background:transparent;border:0;cursor:pointer;padding:6px 8px;border-radius:6px;text-decoration:underline;text-decoration-style:dotted;text-underline-offset:4px}.report-detail-number:hover{background:#f1f5f9}.report-detail-number:focus-visible{outline:2px solid #64748b}.report-details-dialog{width:min(1100px,calc(100vw - 32px))}.report-details-dialog p{margin:8px 0}.report-details-table{overflow:auto;max-height:48vh;margin-top:18px}.report-details-table table{width:100%;border-collapse:collapse;font-size:14px}.report-details-table th,.report-details-table td{border:1px solid #cbd5e1;padding:10px;text-align:left}.report-details-table th{background:#f1f5f9;position:sticky;top:0}.report-details-total{font-weight:700;text-align:right}.report-details-note{font-size:12px;color:#64748b}';
  document.head.appendChild(style);
})();
