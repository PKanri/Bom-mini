/* Theme switch: "light" (normal) or "dark". Remembered per browser. */
(function(){
  var KEY='bom.theme',root=document.documentElement;
  function read(){try{var v=localStorage.getItem(KEY);return v==='dark'||v==='light'?v:null;}catch(e){return null;}}
  function save(v){try{localStorage.setItem(KEY,v);}catch(e){}}
  function apply(theme){
    root.setAttribute('data-theme',theme);
    var meta=document.querySelector('meta[name="theme-color"]');
    if(!meta){meta=document.createElement('meta');meta.name='theme-color';document.head.appendChild(meta);}
    meta.content=theme==='dark'?'#0a111f':'#eef2f8';
    var buttons=document.querySelectorAll('#themeSwitch button[data-theme-value]');
    for(var i=0;i<buttons.length;i++)buttons[i].setAttribute('aria-pressed',String(buttons[i].getAttribute('data-theme-value')===theme));
  }
  apply(read()||'light'); // set before first paint (this script runs in <head>)
  function init(){
    apply(root.getAttribute('data-theme')||'light');
    var box=document.getElementById('themeSwitch');
    if(!box)return;
    box.addEventListener('click',function(ev){
      var b=ev.target.closest&&ev.target.closest('button[data-theme-value]');
      if(!b)return;
      var t=b.getAttribute('data-theme-value');save(t);apply(t);
    });
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
