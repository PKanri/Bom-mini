// The report records these contributions while calculating each total.
function showReportDetails(index,field) {
  const source=latestReport[index],titles={in:'รับเข้า',out:'จ่ายออก',waste:'สูญเสีย',bom:'ตัดตามสูตร BOM'};
  if(!source||!titles[field])return;
  const row=reportDisplayRow(source),main=countUnits(workflow.settings[source.code],source.code)[0];
  const factor=main?.factor>0?main.factor:1;
  const entries=source.details?.[field]||[];
  const number=n=>Number(n).toLocaleString('th-TH',{maximumFractionDigits:8});
  const date=d=>typeof displayDate==='function'?displayDate(d):d;
  const dialog=document.createElement('dialog');
  dialog.className='ponn-dialog report-details-dialog';
  dialog.setAttribute('aria-labelledby','report-details-title');
  const header=field==='bom'?'<th>จำนวนขาย</th><th>ใช้สุทธิต่อจาน</th><th>Yield</th><th>เบิกจริงต่อจาน</th>':field==='waste'?'<th>หมายเหตุ</th>':'<th>จำนวนจากไฟล์</th><th>แปลงต่อ 1 จำนวนในไฟล์</th>';
  const body=entries.map(e=>{
    const extra=field==='bom'?`<td>${number(e.qty)}</td><td>${number(e.netQty/factor)}</td><td>${number(e.yieldRate)}%</td><td>${number(e.factor/factor)}</td>`:field==='waste'?`<td>${esc(e.remarks||'—')}</td>`:`<td>${number(e.qty)}</td><td>${number(e.factor/factor)}</td>`;
    return `<tr><td>${esc(date(e.date))}</td><td>${esc(field==='waste'?'ยอดสูญเสียที่สาขาบันทึก':labels[e.type]||e.type)}</td><td>${esc(e.code)}<br>${esc(e.name||'')}</td>${extra}<td><strong>${number(e.amount/factor)}</strong></td></tr>`;
  }).join('');
  dialog.innerHTML=`<h2 id="report-details-title">รายละเอียด${titles[field]}</h2>
    <p><strong>${esc(row.name)}</strong> · ${esc(row.code)}</p>
    <p>สาขา ${esc(workflow.branches[row.branch]||row.branch)} (${esc(row.branch)}) · ${esc(date(row.start))} ถึง ${esc(date(row.end))}</p>
    <p>หน่วยที่ใช้แสดงปริมาณวัตถุดิบ: <strong>${esc(row.unit)}</strong></p>
    <div class="report-details-table"><table><thead><tr><th>วันที่</th><th>แหล่งข้อมูล</th><th>${field==='bom'?'เมนูที่ขาย':'วัตถุดิบ'}</th>${header}<th>ยอดที่นำมาคิด</th></tr></thead><tbody>${body||`<tr><td colspan="${field==='bom'?8:field==='waste'?5:6}">ไม่มีรายการที่นำมาคำนวณในช่วงนี้</td></tr>`}</tbody></table></div>
    <p class="report-details-total">รวม ${number(row[field])} ${esc(row.unit)}</p>
    <p class="report-details-note">${field==='bom'?'ยอดที่นำมาคิด = จำนวนขาย × ปริมาณสุทธิต่อจาน ÷ (Yield ÷ 100)':field==='waste'?'แสดงยอดสูญเสียที่บันทึก โดยแปลงเป็นหน่วยเดียวกับตารางรายงาน':'ยอดที่นำมาคิด = จำนวนจากไฟล์ × อัตราแปลง · ข้อมูลที่นำเข้าถูกรวมตามวันที่ สาขา และรหัสสินค้าแล้ว'}</p>
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
