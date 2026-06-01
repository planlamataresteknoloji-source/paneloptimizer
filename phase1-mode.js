/* PanelOptimizer v30 — Basit/Pro mode controller. Classic script, iOS Safari compatible. */
/* === PHASE 1 v29 — ERKEN VE BAĞIMSIZ MOD MOTORU === */
(function(){
  window.phase1SetMode=function(mode){
    mode=(mode==='pro')?'pro':'basic';
    var b=document.body;if(!b)return;
    b.classList.remove('phase1-basic','phase1-pro','basic-mode');
    b.classList.add(mode==='pro'?'phase1-pro':'phase1-basic');
    var btns=document.querySelectorAll('[data-phase1-mode]');
    for(var i=0;i<btns.length;i++){btns[i].classList.toggle('active',btns[i].getAttribute('data-phase1-mode')===mode)}
    try{localStorage.setItem('phase1Mode',mode)}catch(e){}
    // Basit modda kullanıcının direkt çalışabilmesi için temel bölümleri açık tut.
    ['sec-sheet','sec-parts','sec-calc'].forEach(function(id){
      var el=document.getElementById(id); if(!el)return;
      el.classList.add('open');
      var body=el.querySelector('.section-body'); if(body)body.style.maxHeight='5000px';
    });
  };
  window.phase1Status=function(msg,type){
    var el=document.getElementById('phase1QuickStatus'); if(!el)return;
    el.className='phase1-status '+(type||'ok'); el.textContent=msg;
    clearTimeout(window.__phase1StatusTimer);
    window.__phase1StatusTimer=setTimeout(function(){el.className='phase1-status'},4500);
  };
  window.phase1ParseLine=function(line){
    var m=String(line||'').trim().match(/^(\d+(?:[\.,]\d+)?)\s*[x×*]\s*(\d+(?:[\.,]\d+)?)(?:\s*[x×*]\s*(\d+))?\s*(.*)$/i);
    if(!m)return null;
    return {w:m[1].replace(',','.'),h:m[2].replace(',','.'),q:m[3]||'1',name:(m[4]||'Parça').trim()||'Parça'};
  };
  window.phase1QuickAdd=function(){
    var ta=document.getElementById('phase1QuickText'); if(!ta)return;
    var lines=ta.value.split(/\n+/).map(function(x){return x.trim()}).filter(Boolean);
    if(!lines.length){phase1Status('Önce hızlı giriş alanına parça yaz.', 'warn');return;}
    var added=0,bad=[];
    lines.forEach(function(line,idx){
      var p=phase1ParseLine(line); if(!p){bad.push(idx+1);return;}
      var name=document.getElementById('partName'), pw=document.getElementById('pw'), ph=document.getElementById('ph'), pq=document.getElementById('pq');
      if(name)name.value=p.name; if(pw)pw.value=p.w; if(ph)ph.value=p.h; if(pq)pq.value=Math.max(1,parseInt(p.q,10)||1);
      try{ if(typeof window.addPart==='function'){window.addPart();added++;} else {bad.push(idx+1);} }catch(e){console.error(e);bad.push(idx+1);}
    });
    if(added) ta.value='';
    phase1Status(bad.length ? (added+' parça eklendi. Okunamayan satır: '+bad.join(', ')) : (added+' parça başarıyla eklendi.'), bad.length?'warn':'ok');
  };
  window.phase1Sample=function(){
    var ta=document.getElementById('phase1QuickText');
    if(ta)ta.value='450x720x2 Kapak Beyaz\n320x560x4 Raf\n600x800x1 Yan Panel';
  };
  window.phase1SmartCalculate=function(){
    try{ if(typeof window.setAlgo==='function') window.setAlgo('hybrid'); }catch(e){}
    try{ if(typeof window.calculate==='function') window.calculate(); else phase1Status('Hesaplama motoru henüz yüklenmedi.', 'warn'); }catch(e){console.error(e);phase1Status('Hesaplama sırasında hata oluştu: '+e.message,'err');}
  };
  window.addEventListener('error',function(e){
    var m=(e&&e.message)?e.message:'Bilinmeyen JS hatası';
    console.error('PanelOptimizer hata:',m,e.error||'');
    var el=document.getElementById('phase1QuickStatus');
    if(el){el.className='phase1-status err';el.textContent='Teknik hata: '+m;}
  });
  document.addEventListener('DOMContentLoaded',function(){
    var saved='basic';try{saved=localStorage.getItem('phase1Mode')||'basic'}catch(e){}
    phase1SetMode(saved==='pro'?'pro':'basic');
  });
})();
