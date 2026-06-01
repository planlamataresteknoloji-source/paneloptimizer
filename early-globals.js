/* PanelOptimizer v30 — early globals. Classic script, iOS Safari compatible. */
/* ── ERKEN GLOBAL DEĞİŞKENLER ──
   HTML parse edilmeden önce tanımlanmalı — TDZ önlemi.
   inline onclick'ler (satır ~900+) bunlara erişir. ── */
var LANG_EARLY={tr:{},en:{}};  // LANG asıl bloğu aşağıda doldurulur
var currentLang=(function(){try{return localStorage.getItem('pop9_lang')||'tr';}catch(e){return'tr';}})();
function t(key){
  if(typeof LANG==='undefined')return key;
  return(LANG[currentLang]&&LANG[currentLang][key])||(LANG.tr&&LANG.tr[key])||key;
}
var SHEET_W=1830,SHEET_H=3660;
var parts=[],colorIdx=0;
var currentAlgo='hybrid';
var currentGrain='any';
var allPanelsResult=[],currentPanelIdx=0;
var currentMaterial='';
var COLORS=['#818cf8','#34d399','#fbbf24','#f87171','#38bdf8','#a78bfa','#4ade80','#fb923c','#e879f9','#67e8f9','#a3e635','#f472b6','#60a5fa','#facc15','#86efac'];
var zoomLevel=1,fitScale=1;
var ZOOM_MIN=.2,ZOOM_MAX=6,ZOOM_STEP=.2;
var DR=window.devicePixelRatio||2;
var RULER_W=52, RULER_H=22; // px — cetrul alanı (DPR'siz)
var _hatchCache={};
var _kbToastShown=false;
var _obStep=0;
var OB_TOTAL=3;
