/* Event Hub UI extras: Philippine Time everywhere, dark/light theme, friendly buttons, uniform search boxes, Caps Lock warning, phone-friendly sizes.
   Loaded by index.html and bybet/admin.html. */
(function () {
  /* 1. every date and time is shown in Philippine Time (Asia/Manila), whatever the device clock or time zone is */
  var TZ = 'Asia/Manila';
  ['toLocaleString', 'toLocaleDateString', 'toLocaleTimeString'].forEach(function (k) {
    var o = Date.prototype[k];
    Date.prototype[k] = function (l, opt) { return o.call(this, l, Object.assign({ timeZone: TZ }, opt || {})); };
  });

  /* 2. theme: saved choice (light / dark) or Auto; the parent page tells the Event Games window when it changes */
  var root = document.documentElement;
  function applyTheme(t) { if (t === 'light' || t === 'dark') root.setAttribute('data-theme', t); else root.removeAttribute('data-theme'); }
  try { applyTheme(localStorage.getItem('gphi-theme')); } catch (e) {}
  window.addEventListener('storage', function (e) { if (e.key === 'gphi-theme') applyTheme(e.newValue); });
  window.addEventListener('message', function (e) { if (e.origin === location.origin && e.data && e.data.type === 'gphi-theme') applyTheme(e.data.theme); });

  /* 3. look and feel */
  var css = [
    'button,.btn,.g,.bb,.tab{transition:transform .12s ease,box-shadow .15s ease,filter .15s ease,background-color .15s ease;-webkit-tap-highlight-color:transparent;touch-action:manipulation}',
    'button:not(:disabled):hover,.btn:hover,.bb:hover{filter:brightness(1.07);box-shadow:0 2px 10px rgba(0,0,0,.18)}',
    'button:not(:disabled):active,.btn:active,.bb:active{transform:scale(.96);filter:brightness(.95)}',
    'button:disabled{opacity:.55;cursor:not-allowed}',
    '.rpl-host{position:relative;overflow:hidden}',
    '.rpl{position:absolute;border-radius:50%;width:24px;height:24px;margin:-12px;background:rgba(255,255,255,.45);transform:scale(0);animation:rpl .55s ease-out;pointer-events:none}',
    '@keyframes rpl{to{transform:scale(8);opacity:0}}',
    '@media (prefers-reduced-motion:reduce){.rpl{display:none}button,.btn{transition:none}}',
    'button:focus-visible,input:focus-visible,select:focus-visible,textarea:focus-visible,a:focus-visible{outline:2px solid var(--ac,var(--a,#2563eb));outline-offset:2px}',
    /* every search box looks the same */
    'input[type=search],input[data-search]{width:100%;max-width:100%;min-height:44px;border-radius:99px!important;padding:10px 14px 10px 40px!important;background-repeat:no-repeat!important;background-position:13px center!important;background-size:18px!important;',
    'background-image:url("data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'18\' height=\'18\' fill=\'none\' stroke=\'%23888\' stroke-width=\'2\' stroke-linecap=\'round\'%3E%3Ccircle cx=\'8\' cy=\'8\' r=\'6\'/%3E%3Cpath d=\'M13 13l4 4\'/%3E%3C/svg%3E")!important}',
    '.caps-warn{color:#92400e;background:#fef3c7;border:1px solid #f59e0b;padding:6px 10px;border-radius:8px;font-size:12px;margin-top:4px}',
    /* phones: bigger tap targets, no zoom when typing */
    '@media (max-width:820px){input,select,textarea{font-size:16px!important}button,select,input:not([type=checkbox]):not([type=radio]){min-height:40px}}'
  ].join('\n');
  var st = document.createElement('style'); st.id = 'uix'; st.textContent = css; (document.head || root).appendChild(st);

  document.addEventListener('click', function (e) {
    var b = e.target.closest && e.target.closest('button,.btn,.bb,.tab');
    if (!b || b.disabled || b.querySelector('.nbadge') || b.closest('.mods')) return;
    var r = b.getBoundingClientRect(), s = document.createElement('span');
    s.className = 'rpl'; s.style.left = (e.clientX - r.left) + 'px'; s.style.top = (e.clientY - r.top) + 'px';
    b.classList.add('rpl-host'); b.appendChild(s);
    setTimeout(function () { s.remove(); b.classList.remove('rpl-host'); }, 560);
  }, true);

  /* 4. one search style: mark every search box and remove the old magnifier emoji from the hint text */
  function tidy(n) {
    if (!n.querySelectorAll) return;
    var list = n.matches && n.matches('input') ? [n] : n.querySelectorAll('input[placeholder]');
    [].forEach.call(list, function (i) {
      var p = i.getAttribute('placeholder') || '';
      if (/^\uD83D\uDD0D\s*/.test(p)) { p = p.replace(/^\uD83D\uDD0D\s*/, ''); i.setAttribute('placeholder', p); }
      if (/search/i.test(p) && i.type !== 'password') i.setAttribute('data-search', '1');
    });
  }
  function startObs() { tidy(document.body); new MutationObserver(function (ms) { ms.forEach(function (m) { [].forEach.call(m.addedNodes, function (n) { if (n.nodeType === 1) tidy(n); }); }); }).observe(document.body, { childList: true, subtree: true }); }
  if (document.body) startObs(); else document.addEventListener('DOMContentLoaded', startObs);

  /* 5. Caps Lock warning under every password box */
  function caps(e) {
    var t = e.target; if (!t || t.type !== 'password' || !t.parentElement) return;
    var on = e.getModifierState && e.getModifierState('CapsLock'), w = t.parentElement.querySelector('.caps-warn');
    if (on && !w) { w = document.createElement('div'); w.className = 'caps-warn'; w.setAttribute('role', 'alert'); w.textContent = '\u26A0 Caps Lock is ON. Passwords are case-sensitive.'; t.insertAdjacentElement('afterend', w); }
    else if (!on && w) w.remove();
  }
  document.addEventListener('keydown', caps, true); document.addEventListener('keyup', caps, true);
  document.addEventListener('focusout', function (e) { if (e.target && e.target.type === 'password') { var w = e.target.parentElement && e.target.parentElement.querySelector('.caps-warn'); if (w) w.remove(); } }, true);
})();
