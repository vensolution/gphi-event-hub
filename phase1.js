/* Event Hub - Phase 1: dark/light button, Philippine time, Caps Lock warning, button press effects, phone polish */
(function(){
  // 1) Philippine Time everywhere (any date shown with toLocale*String uses Asia/Manila unless a zone is given)
  ['toLocaleString','toLocaleDateString','toLocaleTimeString'].forEach(function(fn){
    var orig=Date.prototype[fn];
    Date.prototype[fn]=function(loc,opt){opt=Object.assign({},opt||{});if(!opt.timeZone)opt.timeZone='Asia/Manila';return orig.call(this,loc,opt)};
  });

  // 2) Dark / Light mode button (remembered on this device)
  var root=document.documentElement,KEY='hub-theme';
  function saved(){try{return localStorage.getItem(KEY)}catch(e){return null}}
  function apply(m){root.setAttribute('data-theme',m);try{localStorage.setItem(KEY,m)}catch(e){};if(btn)btn.textContent=m==='dark'?'☀️ Light':'🌙 Dark'}
  var start=saved()||(window.matchMedia&&matchMedia('(prefers-color-scheme:dark)').matches?'dark':'light');
  var btn;
  function mount(){
    btn=document.createElement('button');btn.id='themeBtn';btn.type='button';btn.setAttribute('aria-label','Switch dark or light mode');
    btn.onclick=function(){apply(root.getAttribute('data-theme')==='dark'?'light':'dark')};
    document.body.appendChild(btn);apply(start);
  }

  // 3) Caps Lock warning on every password box
  var warn=document.createElement('div');warn.id='capsWarn';warn.textContent='⚠️ Caps Lock is ON';
  function caps(e){
    var t=e.target;if(!t||t.type!=='password'){warn.style.display='none';return}
    var on=e.getModifierState&&e.getModifierState('CapsLock');
    if(on){t.insertAdjacentElement('afterend',warn);warn.style.display='block'}else warn.style.display='none';
  }
  ['keydown','keyup','focusin'].forEach(function(ev){document.addEventListener(ev,caps,true)});
  document.addEventListener('focusout',function(){warn.style.display='none'},true);

  // 4) Look and feel: button press effect, smooth transitions, bigger tap targets on phones
  var css=document.createElement('style');css.textContent=
   'button,.tile,[role=button]{transition:transform .12s ease,box-shadow .15s ease,filter .15s ease}'+
   'button:hover:not(:disabled){filter:brightness(1.06);box-shadow:0 3px 10px rgba(0,0,0,.18)}'+
   'button:active:not(:disabled){transform:scale(.95);box-shadow:none}'+
   'body{transition:background .2s,color .2s}'+
   '#themeBtn{position:fixed;right:14px;bottom:calc(14px + env(safe-area-inset-bottom,0px));z-index:9999;border:1px solid var(--bd);background:var(--card);color:var(--tx);border-radius:999px;padding:10px 16px;font-size:14px;cursor:pointer;box-shadow:0 4px 14px rgba(0,0,0,.25)}'+
   '#capsWarn{display:none;margin:4px 0;padding:6px 10px;border-radius:8px;background:#fef3c7;color:#92400e;border:1px solid #f59e0b;font-size:13px;font-weight:700}'+
   '@media (max-width:700px){button,input,select{min-height:44px}input,select,textarea{font-size:16px}}';
  document.head.appendChild(css);
  if(document.body)mount();else document.addEventListener('DOMContentLoaded',mount);
})();
