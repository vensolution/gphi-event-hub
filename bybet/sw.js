// Lets the page open without internet after the first visit. Spins still need internet.
const C='spin-v12',SHELL=['./','index.html','logo.webp'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(C).then(c=>c.addAll(SHELL)).then(()=>self.skipWaiting()))});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!=C).map(x=>caches.delete(x)))).then(()=>self.clients.claim()))});
self.addEventListener('fetch',e=>{const r=e.request,u=new URL(r.url);
 if(r.method!='GET'||u.origin!=location.origin||u.pathname.endsWith('/admin.html')||u.pathname.endsWith('/tracker.html'))return;
 if(r.mode=='navigate'){e.respondWith(fetch(r).then(x=>{const k=x.clone();caches.open(C).then(c=>c.put('index.html',k));return x}).catch(()=>caches.match('index.html')));return}
 e.respondWith(caches.match(r).then(h=>h||fetch(r).then(x=>{const k=x.clone();caches.open(C).then(c=>c.put(r,k));return x})))});
