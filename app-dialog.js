// Shared top-centred dialogs. Confirmations resolve only after the user's choice.
(function(){
  const queue=[];let active=false;
  function showNext(){
    if(active||!queue.length)return;
    active=true;
    const {message,confirm,resolve}=queue.shift(),previous=document.activeElement;
    const dialog=document.createElement('dialog');
    dialog.className='ponn-dialog';
    dialog.setAttribute('aria-labelledby','ponn-dialog-title');
    dialog.setAttribute('aria-describedby','ponn-dialog-message');
    const title=document.createElement('h2');title.id='ponn-dialog-title';title.textContent=confirm?'ยืนยันการทำรายการ':'แจ้งเตือน';
    const body=document.createElement('div');body.id='ponn-dialog-message';body.className='ponn-dialog-message';body.textContent=String(message);
    const actions=document.createElement('div');actions.className='ponn-dialog-actions';
    let done=false;
    function finish(value){if(done)return;done=true;dialog.close();dialog.remove();active=false;if(previous?.isConnected)previous.focus();resolve(value);showNext();}
    const ok=document.createElement('button');ok.type='button';ok.textContent='ตกลง';ok.className='ponn-dialog-ok';ok.onclick=()=>finish(true);
    if(confirm){const cancel=document.createElement('button');cancel.type='button';cancel.textContent='ยกเลิก';cancel.onclick=()=>finish(false);actions.append(cancel);}
    actions.append(ok);dialog.append(title,body,actions);document.body.append(dialog);
    dialog.addEventListener('cancel',e=>{e.preventDefault();finish(false);});
    dialog.showModal();(confirm?actions.firstChild:ok).focus();
  }
  function enqueue(message,confirm){return new Promise(resolve=>{queue.push({message,confirm,resolve});showNext();});}
  window.appAlert=message=>enqueue(message,false);
  window.appConfirm=message=>enqueue(message,true);
  const style=document.createElement('style');style.textContent=`
    .ponn-dialog{position:fixed;inset:clamp(20px,8vh,80px) auto auto 50%;transform:translateX(-50%);margin:0;width:min(520px,calc(100vw - 32px));max-height:80vh;overflow:auto;box-sizing:border-box;border:1px solid #e2e8f0;border-radius:18px;padding:24px;background:#fff;color:#0f172a;box-shadow:0 20px 60px #0f172a33;font-family:inherit}
    .ponn-dialog::backdrop{background:#0f172a55}.ponn-dialog h2{margin:0 0 14px;font-size:18px;font-weight:700}.ponn-dialog-message{white-space:pre-wrap;overflow-wrap:anywhere;line-height:1.7;font-size:14px}.ponn-dialog-actions{display:flex;justify-content:flex-end;gap:10px;margin-top:24px}.ponn-dialog-actions button{border:1px solid #cbd5e1;border-radius:10px;padding:9px 22px;background:#fff;color:#475569;cursor:pointer;font:inherit}.ponn-dialog-actions .ponn-dialog-ok{background:#1e293b;color:white;border-color:#1e293b}.ponn-dialog-actions button:focus-visible{outline:2px solid #64748b;outline-offset:3px}
  `;document.head.append(style);
})();
