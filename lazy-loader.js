/* PanelOptimizer v30 — optional library loader. Classic script, iOS Safari compatible. */
/* ── LAZY SCRIPT LOADER ── */
function loadScript(url,cb){
  var s=document.createElement('script');
  s.src=url;
  s.onload=function(){if(cb)cb(null);};
  s.onerror=function(){if(cb)cb(new Error('Yüklenemedi: '+url));};
  document.head.appendChild(s);
}
var _xlsxLoaded=false,_pdfLoaded=false;
function ensureXLSX(cb){
  if(typeof XLSX!=='undefined'){cb(null);return;}
  if(_xlsxLoaded){setTimeout(function(){cb(null);},300);return;}
  _xlsxLoaded=true;
  loadScript('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',cb);
}
function ensurePDF(cb){
  if(typeof window.jspdf!=='undefined'){cb(null);return;}
  if(_pdfLoaded){setTimeout(function(){cb(null);},300);return;}
  _pdfLoaded=true;
  loadScript('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js',cb);
}
