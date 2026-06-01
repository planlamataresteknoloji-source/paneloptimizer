/* PanelOptimizer v30 — original app core. Classic script, iOS Safari compatible. */
/* ── GLOBAL DEĞİŞKENLER — en başta tanımlanmalı (TDZ önlemi) ── */

// (moved to top)

// (moved to top)

/* ═══════════════════════════════════════════
   MOBİL UYUMLU İNDİRME YARDIMCILARI
   iOS Safari: DOM'a eklemeden a.click() çalışmaz,
   data: URL'leri de çoğu zaman inmiyor.
   Blob + createObjectURL en güvenilir yöntem.
═══════════════════════════════════════════ */
function mobileDownload(href, filename){
  try {
    // data: URL'yi Blob'a çevir (iOS Safari için en güvenilir yol)
    if(href.startsWith('data:')){
      const [meta, b64orText] = href.split(',');
      const mime = meta.split(':')[1].split(';')[0];
      const isBase64 = meta.includes('base64');
      let blob;
      if(isBase64){
        const bin = atob(b64orText);
        const arr = new Uint8Array(bin.length);
        for(let i=0;i<bin.length;i++) arr[i]=bin.charCodeAt(i);
        blob = new Blob([arr], {type: mime});
      } else {
        blob = new Blob([decodeURIComponent(b64orText)], {type: mime});
      }
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = filename; a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      setTimeout(()=>{ URL.revokeObjectURL(url); document.body.removeChild(a); }, 500);
    } else {
      const a = document.createElement('a');
      a.href = href; a.download = filename; a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      setTimeout(()=> document.body.removeChild(a), 500);
    }
  } catch(e) {
    // Fallback: yeni sekmede aç
    window.open(href, '_blank');
  }
}

function mobileDownloadBlob(blob, filename){
  try {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename; a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    setTimeout(()=>{ URL.revokeObjectURL(url); document.body.removeChild(a); }, 500);
  } catch(e) {
    showMsg(t('msgDownloadFail'),'error');
  }
}

function mobileDownloadXLSX(wb, filename){
  try {
    const wbout = XLSX.write(wb, {bookType:'xlsx', type:'array'});
    mobileDownloadBlob(new Blob([wbout], {type:'application/octet-stream'}), filename);
  } catch(e) {
    XLSX.writeFile(wb, filename); // fallback
  }
}

function mobileDownloadPDF(pdf, filename){
  try {
    const blob = pdf.output('blob');
    mobileDownloadBlob(blob, filename);
  } catch(e) {
    pdf.save(filename); // fallback
  }
}

/* ═══════════════════════════════════════════
   MALZEME & KALINLIK
═══════════════════════════════════════════ */
function setMaterial(v){
  currentMaterial=v;
  try{localStorage.setItem('pop9_material',v)}catch(e){}
  updateSubtitle();
  loadMaterialPrice(v); // kayıtlı fiyatı otomatik yükle
}
function setThickness(v){
  currentThickness=parseInt(v)||18;
  try{localStorage.setItem('pop9_thickness',currentThickness)}catch(e){}
  updateSubtitle();
}
/* ═══════════════════════════════════════════
   PARA BİRİMİ
═══════════════════════════════════════════ */
const CURRENCY_SYMBOLS={TRY:'₺',USD:'$',EUR:'€'};
const CURRENCY_NAMES={TRY:'Türk Lirası (₺)',USD:'Dolar ($)',EUR:'Euro (€)'};
function getCurrencySymbol(){return CURRENCY_SYMBOLS[currentCurrency]||'₺';}
function setCurrency(cur){
  currentCurrency=cur;
  ['TRY','USD','EUR'].forEach(k=>{
    const el=document.getElementById('cur-'+k);
    if(el)el.className='grain-opt'+(k===cur?' selected':'');
  });
  // Tüm fiyat label'larını güncelle
  document.querySelectorAll('.currency-symbol').forEach(el=>el.textContent=getCurrencySymbol());
  _cachedPrices.valid=false;
  updateCost();
  renderParts();
  if(allPanelsResult.length)buildCutList(allPanelsResult);
  try{localStorage.setItem('pop9_currency',cur)}catch(e){}
}

/* ═══════════════════════════════════════════
   ONBOARDING
═══════════════════════════════════════════ */
/* ═══════════════════════════════════════════
   ONBOARDING WIZARD
═══════════════════════════════════════════ */// (moved to top)

function obStep(dir){
  const next=_obStep+dir;
  if(next>=OB_TOTAL){closeOnboarding(false);return;}
  if(next<0)return;
  // Sayfaları değiştir
  document.getElementById(`ob-page-${_obStep}`).classList.remove('active');
  _obStep=next;
  document.getElementById(`ob-page-${_obStep}`).classList.add('active');
  // Dotları güncelle
  for(let i=0;i<OB_TOTAL;i++){
    const d=document.getElementById(`ob-dot-${i}`);
    d.classList.toggle('active',i===_obStep);
    d.classList.toggle('done',i<_obStep);
  }
  // Butonları güncelle
  document.getElementById('ob-back').style.display=_obStep>0?'':'none';
  const nxtBtn=document.getElementById('ob-next');
  nxtBtn.textContent=_obStep===OB_TOTAL-1?t('obFinish'):t('obNext');
}

function showOnboarding(){
  _obStep=0;
  // Sayfaları sıfırla
  for(let i=0;i<OB_TOTAL;i++){
    var _op=document.getElementById('ob-page-'+i);if(_op)_op.classList.toggle('active',i===0);
    const d=document.getElementById(`ob-dot-${i}`);
    if(d)d.classList.toggle('active',i===0);
    if(d)d.classList.remove('done');
  }
  document.getElementById('ob-back').style.display='none';
  document.getElementById('ob-next').textContent=t('obNext');
  document.getElementById('onboardOverlay').classList.add('visible');
}

function closeOnboarding(neverShow=false){
  document.getElementById('onboardOverlay').classList.remove('visible');
  if(neverShow)try{localStorage.setItem('pop9_onboarded','1')}catch(e){}
}

// Skip linki
var _obSkipEl=document.getElementById('ob-skip');if(_obSkipEl)_obSkipEl.addEventListener('click',function(){closeOnboarding(true);});

function checkOnboarding(){
  const done=localStorage.getItem('pop9_onboarded');
  if(!done&&!parts.length){
    setTimeout(showOnboarding,700);
  }
}
/* ═══════════════════════════════════════════
   DİL SİSTEMİ (TR / EN)
═══════════════════════════════════════════ */
var LANG={
  tr:{
    calculate:'▶\u00a0Hesapla & Yerleştir',calculating:'⏳ Hesaplanıyor…',
    save:'✓\u00a0Kaydet',cancel:'✕\u00a0İptal',edit:'Düzenle',del:'✕ Sil',
    duplicate:'⧉ Çoğalt',undo:'↩ Geri Al',redo:'↪ İleri Al',
    addPart:'＋\u00a0Parça Ekle',optimize:'Optimize Et',recalculate:'⚠️ Yeniden Hesapla',
    panelSetup:'Ana Panel',algorithm:'Algoritma',partsSection:'Parça Ekle',
    quickCalc:'Hızlı Hesap',resultsSection:'Yerleştir & İstatistik',
    projectsSection:'Projeler',history:'Hesap Geçmişi',
    backups:'Otomatik Yedekler',share:'Paylaş',
    cutList:'✂️ Kesim Listesi',costTitle:'Maliyet & Malzeme',
    width:'Genişlik (mm)',height:'Yükseklik (mm)',qty:'Adet',name:'Ad',
    widthShort:'Genişlik',heightShort:'Yükseklik',
    partName:'Parça Adı',projectName:'Proje Adı',
    note:'Not / Açıklama',noteOpt:'Not / Açıklama (opsiyonel)',
    customer:'Müşteri / Firma Adı',orderNo:'Sipariş / Teklif No (opsiyonel)',
    material:'Malzeme',thickness:'Kalınlık (mm)',
    edgeMargin:'Kenar Payı (mm)',cutMargin:'Kesim Payı (mm)',
    stockQty:'📦 Eldeki Panel Adedi',
    algoMaxRects:'MaxRects (BSSF): Parçaları en küçük artık boşluğa yerleştirir, yoğun paketleme sağlar. Genel kullanım için önerilir.',
    algoFFD:'FFD — First Fit Decreasing: Parçaları büyük→küçük sıralar, ilk sığan boşluğa yerleştirir. Hızlı & güvenilir.',
    algoGuillotine:'Guillotine (BSSF + SLA): Akıllı bölme yöntemi. Tek geçişli kesim planlamasına uygun.',
    grainAny:'Serbest (90° döndürmeye izin ver)',grainH:'→ Yatay tahıl: döndürme yok',grainV:'↓ Dikey tahıl: döndürme yok',
    symmetric:'🪞 Simetrik Yerleşim',
    cutOrderPlacement:'Yerleşim',cutOrderLR:'→↓ Sol→Sağ',cutOrderSize:'↕ Boyut',
    totalParts:'Toplam Parça',placed:'Yerleştirilen',areaUsage:'Alan Kullanımı',
    waste:'Fire',panels:'Panel',unplaced:'Yerleştirilemeyen',
    sqmPrice:'Parça Malzemesi (m²)',cutPrice:'İşçilik / Kesim',sheetPrice:'Panel Fiyatı',
    fireCost:'🔥 Fire Maliyeti',totalCost:'Toplam Maliyet',currency:'Para Birimi',vat:'KDV',
    costBreakdown:'Birim maliyet kırılımı',costNote1:'— parça alanı üzerinden',
    costNote2:'— toplam kesim adedi × ücret',costNote3:'— toplam panel maliyeti',
    costNote4:'(panel payı fire maliyetini içerir)',
    exportSVG:'SVG',exportPDF:'PDF',exportPNG:'PNG',exportExcel:'📊 Excel',
    exportFullXLSX:'📊 Tam XLSX',exportJSON:'📤 JSON',loadJSON:'📥 JSON Yükle',
    exportDXF:'📐 DXF',exportCSV:'📄 CSV',
    quotePDF:'📋\u00a0Teklif PDF Oluştur',shareLink:'🔗\u00a0Paylaşım Linki Oluştur',
    msgCalcFirst:'Önce hesap yapın.',msgAddParts:'Önce parça ekleyin.',
    msgNoUndo:'Geri alınacak işlem yok.',msgNoRedo:'İleri alınacak işlem yok.',
    msgSaved:'Proje kaydedildi.',msgPartAdded:'Parça eklendi.',
    msgPartDeleted:' parça silindi.',msgPartDuplicated:' parça kopyalandı.',
    msgPartsMerged:' parça birleştirildi → ',msgNoMerge:'Birleştirilecek aynı boyutlu parça bulunamadı.',
    msgImported:' parça içe aktarıldı.',msgExcelDone:'Excel dosyası indirildi.',
    msgExcelLoading:'Excel kütüphanesi yükleniyor, tekrar deneyin.',
    msgInvalidJSON:'Geçersiz JSON dosyası.',msgInvalidLink:'Geçersiz paylaşım linki.',
    msgInvalidSize:'Geçerli boyut girin.',msgStorageFull:'Depolama alanı doldu. Eski projeleri silin.',
    msgLinkCreated:'Paylaşım linki oluşturuldu.',msgNoFit:t('msgNoFit'),
    msgUndo:'" geri alındı',msgRedo:'" ileri alındı',msgLoaded:'" yüklendi.',
    msgUpdated:' güncellendi.',msgPriceSaved:t('msgPriceSaved'),msgPriceLoaded:t('msgPriceLoaded'),
    placeholderMain:'Parça ekleyin ve',placeholderBtn:'Hesapla',placeholderSub:'butonuna tıklayın',
    placeholderHint:'Scroll: zoom · Sürükle: kaydır · Ctrl+0: sıfırla',
    emptyParts:'Henüz parça eklenmedi',emptyPartsHint:'Boyutları girin ve ekleyin,\nya da CSV dosyası yükleyin.',
    emptyPartsAdd:'＋ İlk parçayı ekle',emptyProjects:'Kayıtlı proje yok.',
    emptyHistory:'Henüz hesap yapılmadı.',
    emptyBackups:'Henüz yedek yok. Hesap yapıldıkça otomatik oluşur.',
    noResult:'Sonuç yok',legendHint:'Hesap sonrası lejant burada görünür',
    kbTitle:'⌨️ Klavye Kısayolları',kbGeneral:'Genel',kbCanvas:'Canvas',
    kbCalculate:'Hesapla & Yerleştir',kbUndo:'Geri al',kbRedo:'İleri al (Redo)',
    kbSave:'Projeyi kaydet',kbDuplicate:'Parçayı kopyala',kbFullscreen:'Tam ekran aç/kapat',
    kbHelp:'Bu yardımı aç/kapat',kbEscape:'İptal / Kapat',
    kbZoomIn:'Yaklaştır',kbZoomOut:'Uzaklaştır',kbZoomReset:'Sıfırla / Fit',
    kbScroll:'Zoom (fare tekerleği)',kbPan:'Kaydır',
    obTitle:'Panel Optimizasyonu',obSub:'Birkaç adımda ilk yerleşim planınızı oluşturun.',
    obStep1Title:'Ana Paneli Ayarlayın',
    obStep1Body:'Keseceğiniz levhanın genişlik ve yüksekliğini girin. Standart boyutlar için hazır presetleri kullanabilirsiniz.',
    obStep1Tip:'💡 MDF 2440×1220, Kontraplak 2500×1250 gibi presetler sol üst köşede hazır.',
    obStep2Title:'Parçaları Ekleyin',
    obStep2Body:'Her parçanın adını, genişlik/yükseklik (mm) ve adetini girin. Farklı parça türleri için farklı renkler atayabilirsiniz.',
    obStep2Tip:'💡 CSV veya JSON dosyası sürükleyerek toplu parça aktarımı yapabilirsiniz.',
    obStep3Title:'Hesapla ve İndir',
    obStep3Body:'«Hesapla & Yerleştir» butonuna basın. Sonucu PNG, PDF, SVG veya Excel olarak dışa aktarın.',
    obStep3Tip:'💡 Ctrl+Enter ile hızlı hesaplama, Ctrl+Z ile geri alma yapabilirsiniz.',
    obSkip:'Bir daha gösterme',obBack:'← Geri',obNext:'İleri →',obFinish:'Başlayalım ✓',
    canvasTitle:'Yerleşim Planı',canvasEmpty:'Henüz hesap yapılmadı',
    rotate:'↻\u00a0Döndür (W ↔ H)',mergeBtn:'⊕ Aynıları Birleştir',
    maxQty:'📐 Maks. Çıkacak Adet (tek tür)',allParts:'Tüm Parçalar',selected:'seçili',
    clearSel:'✕ Seçimi Kaldır',sortSize:'↕ Boyuta Göre Sırala',sortUp:'▲ Yukarı',sortDown:'▼ Aşağı',
    showAll:'▼ Tümünü Göster',csvImport:'CSV / TXT İçe Aktar',dragDrop:'ya da CSV dosyası yükleyin.',
    rotAllowed:'Döndürmeye izin ver (90°)',grainLocked:'(tahıl yönü kilitli)',
    normalDim:'Normal (W×H)',rotDim:'Döndürülmüş (H×W)',partDetail:'▶ Parça detayı göster',
    editSave:'Düzenle: Kaydet',editCancel:'Düzenle: İptal',customColor:'Özel renk',
    totalPerPiece:'Toplam / adet',totalCapacity:'Toplam kapasite',singleCut:'Tek kesim',
    areaPerType:'Alan kullanımı (tek tür)',theoreticalNote:'Tüm parçalar için teorik maks. adet hesaplanır.',
    panelSum:'Panel toplamı:',fireLabel:'FİRE',
    kbToast:'⌨ Klavye kısayolları için <strong>?</strong> tuşuna basın',
    designedBy:'DESIGNED BY',themeLight:'☀️ Light',themeDark:'🌙 Dark',
    msgDownloadFail:'İndirme başlatılamadı. Tarayıcınız bu formatı desteklemiyor olabilir.',
    msgFullscreenUnsupported:'Tam ekran bu tarayıcıda desteklenmiyor.',
    msgFullscreenExit:'Tam ekrandan çık (F / Esc)',
    msgEnterDims:'Önce genişlik ve yükseklik girin.',
    msgPresetExists:'zaten kayıtlı.',
    msgSortDone:'Parçalar büyükten küçüğe sıralandı.',
    msgMergeNone:'Birleştirilecek aynı boyutlu parça bulunamadı.',
    msgMaxSize:'Boyut 10.000mm\'den büyük olamaz.',
    msgPartUpdated:' güncellendi.',
    msgLinkCopied:'Link kopyalandı!',
    msgClipboardFail:'Linki kopyalayamadık, lütfen manuel kopyalayın.',
    msgSharedLoaded:'Paylaşılan proje yüklendi!',msgPanelNotFound:'Panel verisi bulunamadı.',
    msgPreparing:'⏳ Hazırlanıyor...',msgJsonOnly:'Sadece .json dosyası kabul edilir.',
    msgDropLoaded:'Proje sürüklenerek yüklendi: ',
    msgPdfLoading:'PDF kütüphanesi yükleniyor, lütfen bekleyin...',
    rotatedLabel:'döndürülmüş',
    msgUnplaced:'parça yerleştirilemedi! Boyutları kontrol edin.',
    stockOk:'✅ Stok yeterli:',stockUsed:'panel kullanılıyor,',stockAvail:'panel mevcut.',
    showAllLabel:'▼ Tümünü Göster',partUnit:'parça',
    stockPlaceholder:'ör. 10 — boş bırakırsanız sınırsız',
    searchPlaceholder:'Parça ara...',
    projectPlaceholder:'ör. Mutfak dolabı',
    cutPriceLabel:'Kesim Ücreti',cutPriceNote:'— toplam kesim adedi × ücret',
    sheetPriceLabel:'Panel Fiyatı',sheetPriceNote:'— toplam panel maliyeti',
  },
  en:{
    calculate:'▶\u00a0Calculate & Place',calculating:'⏳ Calculating…',
    save:'✓\u00a0Save',cancel:'✕\u00a0Cancel',edit:'Edit',del:'✕ Delete',
    duplicate:'⧉ Duplicate',undo:'↩ Undo',redo:'↪ Redo',
    addPart:'＋\u00a0Add Part',optimize:'Optimize',recalculate:'⚠️ Recalculate',
    panelSetup:'Main Panel',algorithm:'Algorithm',partsSection:'Add Part',
    quickCalc:'Quick Calc',resultsSection:'Place & Stats',
    projectsSection:'Projects',history:'Calc History',
    backups:'Auto Backups',share:'Share',
    cutList:'✂️ Cut List',costTitle:'Cost & Material',
    width:'Width (mm)',height:'Height (mm)',qty:'Qty',name:'Name',
    widthShort:'Width',heightShort:'Height',
    partName:'Part Name',projectName:'Project Name',
    note:'Note / Description',noteOpt:'Note / Description (optional)',
    customer:'Customer / Company',orderNo:'Order / Quote No (optional)',
    material:'Material',thickness:'Thickness (mm)',
    edgeMargin:'Edge Margin (mm)',cutMargin:'Kerf (mm)',
    stockQty:'📦 Panels in Stock',
    algoMaxRects:'MaxRects (BSSF): Places parts in the smallest remaining space. Recommended for general use.',
    algoFFD:'FFD — First Fit Decreasing: Sorts large→small, fits into first available space. Fast & reliable.',
    algoGuillotine:'Guillotine (BSSF + SLA): Smart split method. Suitable for single-pass cutting plans.',
    grainAny:'Free (allow 90° rotation)',grainH:'→ Horizontal grain: no rotation',grainV:'↓ Vertical grain: no rotation',
    symmetric:'🪞 Symmetric Placement',
    cutOrderPlacement:'Placement',cutOrderLR:'→↓ Left→Right',cutOrderSize:'↕ Size',
    totalParts:'Total Parts',placed:'Placed',areaUsage:'Area Usage',
    waste:'Waste',panels:'Panels',unplaced:'Unplaced',
    sqmPrice:'Part Material (m²)',cutPrice:'Labour / Cutting',sheetPrice:'Sheet Price',
    fireCost:'🔥 Waste Cost',totalCost:'Total Cost',currency:'Currency',vat:'VAT',
    costBreakdown:'Unit cost breakdown',costNote1:'— based on part area',
    costNote2:'— total cuts × rate',costNote3:'— total sheet cost',
    costNote4:'(sheet share includes waste cost)',
    exportSVG:'SVG',exportPDF:'PDF',exportPNG:'PNG',exportExcel:'📊 Excel',
    exportFullXLSX:'📊 Full XLSX',exportJSON:'📤 JSON',loadJSON:'📥 Load JSON',
    exportDXF:'📐 DXF',exportCSV:'📄 CSV',
    quotePDF:'📋\u00a0Generate Quote PDF',shareLink:'🔗\u00a0Create Share Link',
    msgCalcFirst:'Please calculate first.',msgAddParts:'Please add parts first.',
    msgNoUndo:'Nothing to undo.',msgNoRedo:'Nothing to redo.',
    msgSaved:'Project saved.',msgPartAdded:'Part added.',
    msgPartDeleted:' part(s) deleted.',msgPartDuplicated:' part(s) duplicated.',
    msgPartsMerged:' parts merged → ',msgNoMerge:'No matching parts to merge.',
    msgImported:' part(s) imported.',msgExcelDone:'Excel file downloaded.',
    msgExcelLoading:'Loading Excel library, please try again.',
    msgInvalidJSON:'Invalid JSON file.',msgInvalidLink:'Invalid share link.',
    msgInvalidSize:'Please enter a valid size.',msgStorageFull:'Storage full. Delete old projects.',
    msgLinkCreated:'Share link created.',msgNoFit:'does not fit on sheet!',
    msgUndo:"' undone",msgRedo:"' redone",msgLoaded:"' loaded.",
    msgUpdated:' updated.',msgPriceSaved:' m² price saved: ',msgPriceLoaded:' price loaded: ',
    placeholderMain:'Add parts and click',placeholderBtn:'Calculate',placeholderSub:'',
    placeholderHint:'Scroll: zoom · Drag: pan · Ctrl+0: reset',
    emptyParts:'No parts added yet',emptyPartsHint:'Enter dimensions and add,\nor drop a CSV file.',
    emptyPartsAdd:'＋ Add first part',emptyProjects:'No saved projects.',
    emptyHistory:'No calculations yet.',
    emptyBackups:'No backups yet. Created automatically after calculations.',
    noResult:'No results',legendHint:'Legend appears here after calculation',
    kbTitle:'⌨️ Keyboard Shortcuts',kbGeneral:'General',kbCanvas:'Canvas',
    kbCalculate:'Calculate & Place',kbUndo:'Undo',kbRedo:'Redo',
    kbSave:'Save project',kbDuplicate:'Duplicate part',kbFullscreen:'Toggle fullscreen',
    kbHelp:'Toggle this help',kbEscape:'Cancel / Close',
    kbZoomIn:'Zoom in',kbZoomOut:'Zoom out',kbZoomReset:'Reset / Fit',
    kbScroll:'Zoom (scroll wheel)',kbPan:'Pan',
    obTitle:'Panel Optimizer',obSub:'Create your first layout plan in a few steps.',
    obStep1Title:'Set Up Main Panel',
    obStep1Body:'Enter the width and height of the sheet you want to cut. Use presets for standard sizes.',
    obStep1Tip:'💡 Presets like MDF 2440×1220, Plywood 2500×1250 are available in the top left.',
    obStep2Title:'Add Parts',
    obStep2Body:'Enter the name, width/height (mm) and quantity for each part. Assign different colours to different part types.',
    obStep2Tip:'💡 Drag and drop a CSV or JSON file for bulk import.',
    obStep3Title:'Calculate & Export',
    obStep3Body:'Press «Calculate & Place». Export the result as PNG, PDF, SVG or Excel.',
    obStep3Tip:'💡 Ctrl+Enter for quick calculation, Ctrl+Z to undo.',
    obSkip:"Don't show again",obBack:'← Back',obNext:'Next →',obFinish:"Let's Start ✓",
    canvasTitle:'Layout Plan',canvasEmpty:'No calculation yet',
    rotate:'↻\u00a0Rotate (W ↔ H)',mergeBtn:'⊕ Merge Identical',
    maxQty:'📐 Max Possible Qty (single type)',allParts:'All Parts',selected:'selected',
    clearSel:'✕ Clear Selection',sortSize:'↕ Sort by Size',sortUp:'▲ Up',sortDown:'▼ Down',
    showAll:'▼ Show All',csvImport:'CSV / TXT Import',dragDrop:'or drop a CSV file.',
    rotAllowed:'Allow rotation (90°)',grainLocked:'(grain direction locked)',
    normalDim:'Normal (W×H)',rotDim:'Rotated (H×W)',partDetail:'▶ Show part detail',
    editSave:'Edit: Save',editCancel:'Edit: Cancel',customColor:'Custom colour',
    totalPerPiece:'Total / piece',totalCapacity:'Total capacity',singleCut:'Single cut',
    areaPerType:'Area usage (single type)',theoreticalNote:'Theoretical max qty calculated for all parts.',
    panelSum:'Panel total:',fireLabel:'WASTE',
    kbToast:'⌨ Press <strong>?</strong> for keyboard shortcuts',
    designedBy:'DESIGNED BY',themeLight:'☀️ Light',themeDark:'🌙 Dark',
    msgDownloadFail:'Download could not start. Your browser may not support this format.',
    msgFullscreenUnsupported:'Fullscreen is not supported in this browser.',
    msgFullscreenExit:'Exit fullscreen (F / Esc)',
    msgEnterDims:'Please enter width and height first.',
    msgPresetExists:'already saved.',
    msgSortDone:'Parts sorted large to small.',
    msgMergeNone:'No matching parts found to merge.',
    msgMaxSize:'Size cannot exceed 10,000mm.',
    msgPartUpdated:' updated.',
    msgLinkCopied:'Link copied!',
    msgClipboardFail:'Could not copy link, please copy manually.',
    msgSharedLoaded:'Shared project loaded!',msgPanelNotFound:'Panel data not found.',
    msgPreparing:'⏳ Preparing...',msgJsonOnly:'Only .json files are accepted.',
    msgDropLoaded:'Project loaded via drag & drop: ',
    msgPdfLoading:'Loading PDF library, please wait...',
    rotatedLabel:'rotated',
    msgUnplaced:'part(s) could not be placed! Check dimensions.',
    stockOk:'✅ Stock sufficient:',stockUsed:'panel(s) used,',stockAvail:'available.',
    showAllLabel:'▼ Show All',partUnit:'part(s)',
    stockPlaceholder:'e.g. 10 — leave blank for unlimited',
    searchPlaceholder:'Search parts...',
    projectPlaceholder:'e.g. Kitchen cabinet',
    cutPriceLabel:'Cutting Rate',cutPriceNote:'— total cuts × rate',
    sheetPriceLabel:'Sheet Price',sheetPriceNote:'— total sheet cost',
  }
};
// (moved to top)

// t() fonksiyonu script başında tanımlı

function setLang(lang){
  currentLang=lang;
  try{localStorage.setItem('pop9_lang',lang)}catch(e){}
  document.getElementById('langBtnTR').classList.toggle('active',lang==='tr');
  document.getElementById('langBtnEN').classList.toggle('active',lang==='en');
  const L=LANG[lang];

  // data-i18n elementleri
  document.querySelectorAll('[data-i18n]').forEach(el=>{
    const key=el.dataset.i18n;
    if(L[key]!==undefined)el.innerHTML=L[key];
  });
  // data-i18n-title
  document.querySelectorAll('[data-i18n-title]').forEach(el=>{
    const key=el.dataset.i18nTitle;
    if(L[key]!==undefined)el.title=L[key];
  });
  // CTA butonu
  const ctaBtn=document.getElementById('calcBtn');
  if(ctaBtn&&!ctaBtn.disabled)ctaBtn.innerHTML=L.calculate;
  // Canvas başlık
  const ct=document.querySelector('.canvas-title');
  if(ct)ct.textContent=L.canvasTitle;
  const ci=document.getElementById('canvasInfo');
  if(ci&&(ci.textContent===LANG[lang==='tr'?'en':'tr'].canvasEmpty||ci.textContent==='Henüz hesap yapılmadı'||ci.textContent==='No calculation yet'))ci.textContent=L.canvasEmpty;
  // Placeholder
  const phText=document.querySelector('.ph-text');
  if(phText)phText.innerHTML=`${L.placeholderMain}<br><strong style="color:var(--accent)">${L.placeholderBtn}</strong>${L.placeholderSub?' '+L.placeholderSub:''}`;
  const phHint=document.querySelector('.ph-hint');
  if(phHint)phHint.textContent=L.placeholderHint;
  // zoom-hint
  const zh=document.getElementById('zoomHint');
  if(zh)zh.textContent=L.placeholderHint;
  // KB toast
  const kt=document.getElementById('kbToast');
  if(kt)kt.innerHTML=L.kbToast;
  // data-kb elementleri (klavye modal)
  document.querySelectorAll('[data-kb]').forEach(el=>{
    const key=el.dataset.kb;
    if(L[key]!==undefined)el.textContent=L[key];
  });
  // Onboarding
  const obT=document.querySelector('.onboard-title');
  if(obT)obT.innerHTML=`<span>${L.obTitle}</span>`;
  const obS=document.querySelector('.onboard-sub');
  if(obS)obS.textContent=L.obSub;
  document.querySelectorAll('[data-ob]').forEach(el=>{
    const key=el.dataset.ob;
    if(L[key]!==undefined)el.innerHTML=L[key];
  });
  const obSkipEl=document.getElementById('ob-skip');
  if(obSkipEl)obSkipEl.textContent=L.obSkip;
  const obBackEl=document.getElementById('ob-back');
  if(obBackEl)obBackEl.textContent=L.obBack;
  const obNextEl=document.getElementById('ob-next');
  if(obNextEl)obNextEl.textContent=(_obStep===OB_TOTAL-1)?L.obFinish:L.obNext;
  // Parça listesi
  renderParts();
  // Tema butonu
  const themeBtn=document.getElementById('themeToggleBtn');
  if(themeBtn){const isLight=document.body.classList.contains('light-mode');themeBtn.textContent=isLight?L.themeLight:L.themeDark;}
  // input placeholder'ları
  const stockEl=document.getElementById('sheetStock');
  if(stockEl)stockEl.placeholder=L.stockPlaceholder;
  const searchEl=document.getElementById('partsSearch');
  if(searchEl)searchEl.placeholder=L.searchPlaceholder;
  const projEl=document.getElementById('projectName');
  if(projEl)projEl.placeholder=L.projectPlaceholder;
}

/* ═══════════════════════════════════════════
   GLOBAL HATA YAKALAYICI
═══════════════════════════════════════════ */
window.onerror=function(msg,src,line,col,err){
  var detail=String(msg)+(line?' L'+line:'')+(col?':'+col:'')+(err&&err.stack?'\n'+String(err.stack).substring(0,200):'');
  console.error('Panel Optimizasyonu hata:',detail);
  var box=document.getElementById('msgBox');
  if(box){box.textContent='Hata: '+String(msg).substring(0,120)+(line?' (L'+line+')':'');box.className='msg error';}
  return false;
};
window.addEventListener('unhandledrejection',e=>{
  console.error('Unhandled promise:',e.reason);
});

/* ═══════════════════════════════════════════
   STATE
═══════════════════════════════════════════ */

// (moved to top)
let currentCutOrder='placement'; // 'placement' | 'lr-tb' | 'size'
// (moved to top)
let undoStack=[];
let redoStack=[];
let calcHistory=[];
let _idCounter=1;           // güvenli ID sayacı
let currentCurrency='TRY';  // TRY | USD | EUR
// (moved to top)
let currentThickness=18;   // mm

function getColor(){return COLORS[colorIdx++%COLORS.length]}
function genId(){return ++_idCounter}

/* XSS güvenli HTML kaçış */
function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;')}

/* ZOOM */

/* Hatch pattern cache */
// (moved to top)

/* ═══════════════════════════════════════════
   MARGIN HELPERS
═══════════════════════════════════════════ */
function getMargins(){
  const edge=parseInt(document.getElementById('marginEdge').value)||0;
  const cut=parseInt(document.getElementById('marginCut').value)||0;
  return{edge,cut};
}
function getEffectiveSheet(){
  const{edge}=getMargins();
  return{w:SHEET_W-edge*2, h:SHEET_H-edge*2};
}

/* ═══════════════════════════════════════════
   STORAGE
═══════════════════════════════════════════ */
function save(){
  try{
    const _data=JSON.stringify(parts);
    // Boyut tahmini: karakter sayısı × 2 byte (UTF-16)
    if(_data.length*2>300000)console.warn('localStorage: parça verisi büyük (~',Math.round(_data.length*2/1024),'KB)');
    localStorage.setItem('pop9_parts',_data);
    localStorage.setItem('pop9_ci',colorIdx);
    localStorage.setItem('pop9_sw',SHEET_W);
    localStorage.setItem('pop9_sh',SHEET_H);
    localStorage.setItem('pop9_algo',currentAlgo);
    localStorage.setItem('pop9_grain',currentGrain);
    localStorage.setItem('pop9_cutorder',currentCutOrder);
    localStorage.setItem('pop9_marginEdge',document.getElementById('marginEdge').value||0);
    localStorage.setItem('pop9_marginCut',document.getElementById('marginCut').value||0);
    localStorage.setItem('pop9_price',document.getElementById('sheetPrice').value||0);
    localStorage.setItem('pop9_sqmprice',document.getElementById('sqmPrice').value||0);
    localStorage.setItem('pop9_cutprice',document.getElementById('cutPrice').value||0);
    localStorage.setItem('pop9_currency',currentCurrency);
    localStorage.setItem('pop9_stock',document.getElementById('sheetStock').value||'');
    localStorage.setItem('pop9_symmetric',document.getElementById('symmetricMode').checked?'1':'0');
    localStorage.setItem('pop9_material',currentMaterial);
    localStorage.setItem('pop9_thickness',currentThickness);
    localStorage.setItem('pop9_customer',document.getElementById('customerName').value||'');
    localStorage.setItem('pop9_orderno',document.getElementById('orderNo').value||'');
    ['projectDelivery','projectNote','operatorName','workshopNote'].forEach(id=>{const el=document.getElementById(id);if(el)localStorage.setItem('pop9_'+id,el.value||'');});
    localStorage.setItem('pop9_kdv',document.getElementById('kdvEnabled').checked?'1':'0');
    localStorage.setItem('pop9_kdvrate',document.getElementById('kdvRate').value||'20');
    localStorage.setItem('pop9_idcounter',_idCounter);
    const n=new Date();
    document.getElementById('lastSaveLbl').textContent='Kaydedildi '+n.toLocaleTimeString('tr-TR',{hour:'2-digit',minute:'2-digit'});
    updateStorageInfo();
    showStaleBanner(); // parça/ayar değişti, hesap eskidi
  }catch(e){
    if(e.name==='QuotaExceededError'||e.code===22){
      showMsg(t('msgStorageFull'),'error');
    }
  }
}
function load(){
  try{
    const p=localStorage.getItem('pop9_parts'),ci=localStorage.getItem('pop9_ci');
    const sw=localStorage.getItem('pop9_sw'),sh=localStorage.getItem('pop9_sh');
    const algo=localStorage.getItem('pop9_algo');
    const grain=localStorage.getItem('pop9_grain');
    const cutorder=localStorage.getItem('pop9_cutorder');
    const me=localStorage.getItem('pop9_marginEdge'),mc=localStorage.getItem('pop9_marginCut');
    const pr=localStorage.getItem('pop9_price');
    const sqmp=localStorage.getItem('pop9_sqmprice');
    const cutp=localStorage.getItem('pop9_cutprice');
    const curr=localStorage.getItem('pop9_currency');
    const idc=localStorage.getItem('pop9_idcounter');
    if(p)parts=JSON.parse(p);
    if(ci)colorIdx=parseInt(ci)||0;
    if(idc)_idCounter=parseInt(idc)||parts.length+1;
    if(sw){SHEET_W=parseInt(sw)||1830;document.getElementById('sheetW').value=SHEET_W}
    if(sh){SHEET_H=parseInt(sh)||3660;document.getElementById('sheetH').value=SHEET_H}
    if(algo)setAlgo(algo);
    if(grain)setGrain(grain);
    if(cutorder)setCutOrder(cutorder);
    if(me)document.getElementById('marginEdge').value=me;
    if(mc)document.getElementById('marginCut').value=mc;
    if(pr)document.getElementById('sheetPrice').value=pr;
    if(sqmp)document.getElementById('sqmPrice').value=sqmp;
    if(cutp)document.getElementById('cutPrice').value=cutp;
    if(curr)setCurrency(curr);
    const stk=localStorage.getItem('pop9_stock');if(stk!==null&&document.getElementById('sheetStock'))document.getElementById('sheetStock').value=stk;
    const sym=localStorage.getItem('pop9_symmetric');if(sym==='1'&&document.getElementById('symmetricMode'))document.getElementById('symmetricMode').checked=true;
    const cust=localStorage.getItem('pop9_customer');if(cust&&document.getElementById('customerName'))document.getElementById('customerName').value=cust;
    const ordn=localStorage.getItem('pop9_orderno');if(ordn&&document.getElementById('orderNo'))document.getElementById('orderNo').value=ordn;
    ['projectDelivery','projectNote','operatorName','workshopNote'].forEach(id=>{const v=localStorage.getItem('pop9_'+id);const el=document.getElementById(id);if(v&&el)el.value=v;});
    const mat=localStorage.getItem('pop9_material');if(mat){currentMaterial=mat;const ms=document.getElementById('materialType');if(ms)ms.value=mat;}
    const thi=localStorage.getItem('pop9_thickness');if(thi){currentThickness=parseInt(thi)||18;const ts=document.getElementById('sheetThickness');if(ts)ts.value=currentThickness;}
    const kdvS=localStorage.getItem('pop9_kdv');if(kdvS==='1')document.getElementById('kdvEnabled').checked=true;
    const kdvR=localStorage.getItem('pop9_kdvrate');if(kdvR)document.getElementById('kdvRate').value=kdvR;
    if(parts.length)renderParts();
    updateSubtitle();
    // Geçmişi localStorage'dan yükle (sadece burada bir kez okunur)
    try{const h=localStorage.getItem('pop9_history');if(h)calcHistory=JSON.parse(h);}catch(e){}
    renderProjectList();
    renderHistory();
    renderBackupList();
    loadShareFromURL();
    loadPartPresets();
    checkOnboarding();
    // Theme
    const theme=localStorage.getItem('pop9_theme')||'dark';
    if(theme==='light'){document.body.classList.add('light-mode');document.getElementById('themeToggleBtn').textContent='☀️ Light'}
    // Dil
    setLang(currentLang);
  }catch(e){}
}

/* ═══════════════════════════════════════════
   ALGORITHM
═══════════════════════════════════════════ */
const ALGO_DESC={
  hybrid:'HYBRID GRID + MAXRECTS: Kare/seri parçaları önce düzenli grid-shelf mantığıyla dizer, kalan parçaları MaxRects BSSF+BAF ile boşluklara yerleştirir. Kare yoğun işler için önerilen varsayılan mod.',
  maxrects:'MAXRECTS (BSSF+BAF): En yüksek alan kullanımı. Kısa kenar ve alan skoru kombinasyonu ile en iyi yerleşimi bulur. Karışık ölçülü parçalarda güçlüdür.',
  shelf:'SHELF / ROW PACKING: Aynı veya benzer ölçülü seri parçaları satır/sütun mantığında hızlı ve düzenli dizer. Kare/dikdörtgen seri üretimde okunabilir kesim planı verir.',
  ffd:'FFD — First Fit Decreasing (SLA split): Parçaları büyük→küçük sıralar, ilk sığan boşluğa yerleştirir. Hızlı & güvenilir.',
  guillotine:'Guillotine (BSSF + SLA split): Shorter Leftover Axis akıllı bölme yöntemi ile uzun ince boşluk birikmesini engeller. Tek geçişli kesim planlamasına uygun.'
};
function setAlgo(a){
  if(!ALGO_DESC[a])a='hybrid';
  currentAlgo=a;
  ['hybrid','maxrects','shelf','ffd','guillotine'].forEach(k=>{
    const el=document.getElementById('algo-'+k);
    if(el)el.className='algo-opt'+(k===a?' selected':'');
  });
  document.getElementById('algoLabel').textContent=(a==='hybrid'?'HYBRID':a.toUpperCase());
  document.getElementById('algoDesc').textContent=ALGO_DESC[a];
  // header-sub subtitleTxt ile güncelleniyor
  save();
}

const GRAIN_DESC={
  any:'Serbest: Döndürmeye izin verilir. Algoritma en iyi yerleşimi bulur.',
  horizontal:'Yatay tahıl: Tüm parçalar yatay yönde tutulur, döndürülmez. Ahşap/laminat deseni için.',
  vertical:'Dikey tahıl: Tüm parçalar dikey yönde tutulur, döndürülmez. Ahşap/laminat deseni için.'
};
function setGrain(g){
  currentGrain=g;
  ['any','horizontal','vertical'].forEach(k=>{
    document.getElementById('grain-'+k).className='grain-opt'+(k===g?' selected':'');
  });
  document.getElementById('grainDesc').textContent=GRAIN_DESC[g];
  save();
}

const CUTORDER_OPTIONS=['placement','lr-tb','size'];
function setCutOrder(o){
  currentCutOrder=o;
  CUTORDER_OPTIONS.forEach(k=>{
    document.getElementById('cutorder-'+k).className='grain-opt'+(k===o?' selected':'');
  });
  // Eğer sonuç varsa kesim listesini yeniden sırala
  if(allPanelsResult.length)buildCutList(allPanelsResult);
  save();
}

function showKbModal(){document.getElementById('kbModal').classList.add('visible')}
function hideKbModal(){document.getElementById('kbModal').classList.remove('visible')}
function toggleFullscreen(){
  const wrap=document.getElementById('canvasWrap');
  const btn=document.getElementById('fsBtn');
  if(!document.fullscreenElement){
    const reqFS=wrap.requestFullscreen||wrap.webkitRequestFullscreen||wrap.mozRequestFullScreen||wrap.msRequestFullscreen;
    if(!reqFS){showMsg(t('msgFullscreenUnsupported'),'info');return;}
    reqFS.call(wrap).then(()=>{
      btn.innerHTML='<svg width="12" height="12" viewBox="0 0 24 24" fill="none"><path d="M8 3H5a2 2 0 00-2 2v3M21 8V5a2 2 0 00-2-2h-3M3 16v3a2 2 0 002 2h3M16 21h3a2 2 0 002-2v-3" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>';
      btn.title=t('msgFullscreenExit');
    }).catch(()=>{});
  } else {
    (document.exitFullscreen||document.webkitExitFullscreen||document.mozCancelFullScreen||document.msExitFullscreen).call(document);
  }
}
document.addEventListener('fullscreenchange',()=>{
  if(!document.fullscreenElement){
    const btn=document.getElementById('fsBtn');
    btn.innerHTML='<svg width="12" height="12" viewBox="0 0 24 24" fill="none"><path d="M3 3h6M3 3v6M21 3h-6M21 3v6M3 21h6M3 21v-6M21 21h-6M21 21v-6" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>';
    btn.title='Tam ekran (F)';
    if(allPanelsResult.length)setTimeout(()=>renderPanel(currentPanelIdx),100);
  }
});
function toggleTheme(){
  const isLight=document.body.classList.toggle('light-mode');
  document.getElementById('themeToggleBtn').textContent=isLight?'☀️ Light':'🌙 Dark';
  try{localStorage.setItem('pop9_theme',isLight?'light':'dark')}catch(e){}
}

/* ═══════════════════════════════════════════
   SECTIONS
═══════════════════════════════════════════ */
function toggleSection(id){document.getElementById(id).classList.toggle('open')}

/* ═══════════════════════════════════════════
   SHEET
═══════════════════════════════════════════ */
function getSheet(){
  const w=parseInt(document.getElementById('sheetW').value)||1830;
  const h=parseInt(document.getElementById('sheetH').value)||3660;
  SHEET_W=w;SHEET_H=h;return{w,h};
}
function rotateSheet(){
  const a=document.getElementById('sheetW'),b=document.getElementById('sheetH'),t=a.value;
  a.value=b.value;b.value=t;updateSubtitle();
}
function setSheet(w,h){
  document.getElementById('sheetW').value=w;
  document.getElementById('sheetH').value=h;
  updateSubtitle();
}
function updateSubtitle(){
  const w=parseInt(document.getElementById('sheetW').value)||'?';
  const h=parseInt(document.getElementById('sheetH').value)||'?';
  const me=parseInt(document.getElementById('marginEdge').value)||0;
  const mc=parseInt(document.getElementById('marginCut').value)||0;
  let sub=`Ana panel: ${w} × ${h} mm — ${currentAlgo.toUpperCase()}`;
  if(currentMaterial)sub+=` | ${currentMaterial}`;
  if(currentThickness&&currentThickness!==18)sub+=` ${currentThickness}mm`;
  if(me||mc)sub+=` | Kenar:${me}mm Kesim:${mc}mm`;
  document.getElementById('subtitleTxt').textContent=sub;
  if(parts.length)renderParts();
  calcQuickSingle();
}

/* ═══════════════════════════════════════════
   MSG
═══════════════════════════════════════════ */
function showMsg(m,type){
  const e=document.getElementById('msgBox');
  e.textContent=m;e.className='msg '+type;
  clearTimeout(e._t);e._t=setTimeout(()=>e.className='msg',3200);
}

/* ═══════════════════════════════════════════
   UNDO SYSTEM
═══════════════════════════════════════════ */
function pushUndo(snapshot,label){
  undoStack.push({parts:JSON.parse(JSON.stringify(snapshot)),label});
  if(undoStack.length>30)undoStack.shift();
  redoStack=[];  // yeni işlem — redo geçmişini temizle
  updateUndoRedoBtns();
}
function undoLast(){
  if(!undoStack.length){showMsg(t('msgNoUndo'),'error');return}
  redoStack.push({parts:JSON.parse(JSON.stringify(parts)),label:undoStack[undoStack.length-1].label});
  if(redoStack.length>30)redoStack.shift();
  const snap=undoStack.pop();
  parts=snap.parts;colorIdx=parts.length;
  _rpCache={};
  renderParts();save();
  updateUndoRedoBtns();
  showUndo((currentLang==='en'?'↩ Undone: ':'↩ '+'"'+snap.label+'" geri alındı'));
}
function redoLast(){
  if(!redoStack.length){showMsg(t('msgNoRedo'),'error');return}
  const snap=redoStack.pop();
  undoStack.push({parts:JSON.parse(JSON.stringify(parts)),label:snap.label});
  parts=snap.parts;colorIdx=parts.length;
  _rpCache={};
  renderParts();save();
  updateUndoRedoBtns();
  showUndo((currentLang==='en'?'↪ Redone: ':'↪ '+'"'+snap.label+'" ileri alındı'));
}
function updateUndoRedoBtns(){
  const ub=document.getElementById('undoBtn2');
  const rb=document.getElementById('redoBtn2');
  if(ub){
    ub.disabled=!undoStack.length;
    ub.style.opacity=undoStack.length?'1':'.35';
    const lastUndo=undoStack.length?undoStack[undoStack.length-1].label:'';
    ub.title=lastUndo?`${t('undo')} — "${lastUndo}" (Ctrl+Z)`:t('undo')+' (Ctrl+Z)';
    ub.setAttribute('data-label',lastUndo);
  }
  if(rb){
    rb.disabled=!redoStack.length;
    rb.style.opacity=redoStack.length?'1':'.35';
    const lastRedo=redoStack.length?redoStack[redoStack.length-1].label:'';
    rb.title=lastRedo?`${t('redo')} — "${lastRedo}" (Ctrl+Y)`:t('redo')+' (Ctrl+Y)';
    rb.setAttribute('data-label',lastRedo);
  }
  // Undo/redo butonlarının yanında işlem adını göster
  const ul=document.getElementById('undoActionLbl');
  const rl=document.getElementById('redoActionLbl');
  if(ul)ul.textContent=undoStack.length?`"${undoStack[undoStack.length-1].label}"`:'' ;
  if(rl)rl.textContent=redoStack.length?`"${redoStack[redoStack.length-1].label}"`:'' ;
}
let undoHideT=null;
function showUndo(msg){
  document.getElementById('undoMsg').textContent=msg;
  const bar=document.getElementById('undoBar');
  bar.classList.add('visible');
  clearTimeout(undoHideT);undoHideT=setTimeout(hideUndo,4000);
}
function hideUndo(){document.getElementById('undoBar').classList.remove('visible')}

/* ═══════════════════════════════════════════
   ADD / REMOVE PARTS
═══════════════════════════════════════════ */
function setPartSize(w,h){
  document.getElementById('pw').value=w;
  document.getElementById('ph').value=h;
  document.getElementById('pw').focus();
}

/* ═══ KULLANICI PRESET KAYDETME ═══ */
_partPresets=[]; // re-init
function loadPartPresets(){
  try{_partPresets=JSON.parse(localStorage.getItem('pop9_partpresets')||'[]');}catch(e){_partPresets=[];}
  renderPartPresets();
}
function savePartPreset(){
  const w=parseInt(document.getElementById('pw').value);
  const h=parseInt(document.getElementById('ph').value);
  if(!w||!h||w<=0||h<=0){showMsg(t('msgEnterDims'),'error');return}
  // Zaten var mı?
  if(_partPresets.some(p=>p.w===w&&p.h===h)){showMsg(`${w}×${h} ${t('msgPresetExists')}`,'info');return}
  if(_partPresets.length>=12){_partPresets.shift();} // max 12 preset
  _partPresets.push({w,h});
  try{localStorage.setItem('pop9_partpresets',JSON.stringify(_partPresets));}catch(e){}
  renderPartPresets();
  showMsg(`${w}×${h} preset olarak kaydedildi.`,'success');
}
function deletePartPreset(w,h){
  _partPresets=_partPresets.filter(p=>!(p.w===w&&p.h===h));
  try{localStorage.setItem('pop9_partpresets',JSON.stringify(_partPresets));}catch(e){}
  renderPartPresets();
}
function renderPartPresets(){
  const wrap=document.getElementById('partPresets');
  if(!wrap)return;
  // Sabit presetler + kullanıcı presetleri
  const fixed=[{w:600,h:400},{w:800,h:600},{w:500,h:300},{w:1200,h:600},{w:400,h:400}];
  const allPresets=[...fixed,..._partPresets.filter(u=>!fixed.some(f=>f.w===u.w&&f.h===u.h))];
  wrap.innerHTML=allPresets.map(p=>{
    const isUser=_partPresets.some(u=>u.w===p.w&&u.h===p.h)&&!fixed.some(f=>f.w===p.w&&f.h===p.h);
    return `<button class="btn btn-chip" onclick="setPartSize(${p.w},${p.h})" title="${isUser?'Sağ tık: sil':''}">
      ${p.w}×${p.h}${isUser?` <span onclick="event.stopPropagation();deletePartPreset(${p.w},${p.h})" style="margin-left:2px;opacity:.5;font-size:.6rem" title="Sil">✕</span>`:''}
    </button>`;
  }).join('');
}
/* ═══ STALE BANNER ═══ */
function showStaleBanner(){
  if(!allPanelsResult.length)return; // henüz hesap yapılmadıysa gösterme
  const b=document.getElementById('staleBanner');
  if(b)b.classList.add('visible');
}
function hideStaleBanner(){
  const b=document.getElementById('staleBanner');
  if(b)b.classList.remove('visible');
}
function addPart(){
  getSheet();
  const name=document.getElementById('partName').value.trim();
  const note=document.getElementById('partNote').value.trim();
  const labelCode=(document.getElementById('partLabelCode')||{value:''}).value.trim();
  const materialCode=(document.getElementById('partMaterialCode')||{value:''}).value.trim();
  const edgeBand=(document.getElementById('partEdgeBand')||{value:''}).value.trim();
  const w=parseInt(document.getElementById('pw').value);
  const h=parseInt(document.getElementById('ph').value);
  const q=Math.max(1,parseInt(document.getElementById('pq').value)||1);
  if(!w||!h||isNaN(w)||isNaN(h)||w<=0||h<=0){showMsg(t('msgInvalidSize'),'error');return}
  if(w>10000||h>10000){showMsg('Boyut 10.000mm den buyuk olamaz.','error');return}
  // Tahıl yönü kısıtı: serbest değilse döndürme kapalı
  const rot=currentGrain==='any'?document.getElementById('allowRotate').checked:false;
  if(!(w<=SHEET_W&&h<=SHEET_H)&&!(rot&&h<=SHEET_W&&w<=SHEET_H)){showMsg(`(${w}×${h}) ${t('msgNoFit')}`,'error');return}
  pushUndo(parts,'Parça ekleme');
  parts.push({id:genId(),name:name||`Parça ${parts.length+1}`,w,h,qty:q,note,labelCode,materialCode,edgeBand,color:getColor(),rotate:rot});
  renderParts();
  document.getElementById('partName').value='';
  document.getElementById('pw').value='';
  document.getElementById('ph').value='';
  document.getElementById('pq').value=1;
  document.getElementById('partNote').value='';
  ['partLabelCode','partMaterialCode','partEdgeBand'].forEach(id=>{const el=document.getElementById(id);if(el)el.value='';});
  showMsg(esc(parts[parts.length-1].name)+' eklendi.','success');
  save();
  setTimeout(()=>document.getElementById('partName').focus(),50);
}
function removePart(id){
  pushUndo(parts,'Parça silme');
  parts=parts.filter(p=>p.id!==id);
  renderParts();save();
}

/* ═══════════════════════════════════════════
   CSV IMPORT
═══════════════════════════════════════════ */
function importCSV(input){
  const f=input.files[0];if(!f)return;
  const reader=new FileReader();
  reader.onload=e=>{
    const lines=e.target.result.split('\n').filter(l=>l.trim());
    let added=0;
    pushUndo(parts,'CSV içe aktarım');
    for(const line of lines){
      // Tırnaklı CSV alanlarını doğru parse et
      const cols=[];
      let cur='',inQ=false;
      for(let i=0;i<line.length;i++){
        const ch=line[i];
        if(ch==='"'){if(inQ&&line[i+1]==='"'){cur+='"';i++;}else{inQ=!inQ;}}
        else if(ch===','&&!inQ){cols.push(cur.trim());cur='';}
        else{cur+=ch;}
      }
      cols.push(cur.trim());

      if(cols.length<2)continue;

      // Format 1: Dışa aktarılan format — #,Ad,Genişlik,Yükseklik,X,Y,Panel[,Maliyet][,Not]
      // Format 2: Basit format     — Ad,Genişlik,Yükseklik[,Adet][,Not]
      // Ayırt et: cols[0] sayı mı?
      const isExportFormat = !isNaN(parseInt(cols[0])) && !isNaN(parseInt(cols[2]));

      let name, w, h, q, note;
      if(isExportFormat){
        // #,Ad,Genişlik,Yükseklik,X,Y,Panel[,Maliyet[,Not]]
        name = cols[1]||`Parça ${parts.length+1}`;
        w    = parseInt(cols[2]);
        h    = parseInt(cols[3]);
        q    = 1; // export formatında adet yok, 1 varsayılan
        note = cols[8]||cols[7]||''; // Not en sonda
      } else {
        // Ad,Genişlik,Yükseklik[,Adet][,Not] — başlık satırını atla
        if(isNaN(parseInt(cols[1])))continue; // w sayı değilse başlık satırı
        name = cols[0]||`Parça ${parts.length+1}`;
        w    = parseInt(cols[1]);
        h    = parseInt(cols[2]);
        q    = parseInt(cols[3])||1;
        note = cols[4]||'';
      }

      if(!w||!h||w<=0||h<=0||w>10000||h>10000)continue;
      const rot=currentGrain==='any';
      parts.push({id:genId(),name,w,h,qty:q,note,color:getColor(),rotate:rot});
      added++;
    }
    renderParts();save();
    showMsg(added+t('msgImported'),'success');
  };
  reader.readAsText(f);
  input.value='';
}

/* ═══════════════════════════════════════════
   RENK DEĞİŞTİRME
═══════════════════════════════════════════ */
let _colorPickerPartId=null;

function applyColor(col){
  if(_colorPickerPartId===null)return;
  const p=parts.find(x=>x.id===_colorPickerPartId);
  if(!p)return;
  pushUndo(parts,'Renk değiştir');
  p.color=col;
  renderParts();save();
  closeColorPicker();
}
function applyCustomColor(col){
  if(_colorPickerPartId===null)return;
  const p=parts.find(x=>x.id===_colorPickerPartId);
  if(!p)return;
  p.color=col;
  // Grid güncelle
  document.querySelectorAll('.color-swatch-btn').forEach(el=>el.classList.remove('active'));
  renderParts();save();
}
function closeColorPicker(){
  document.getElementById('colorPickerPopup').classList.remove('visible');
  _colorPickerPartId=null;
}
document.addEventListener('click',e=>{
  const popup=document.getElementById('colorPickerPopup');
  if(popup&&popup.classList.contains('visible')&&!popup.contains(e.target))closeColorPicker();
});
/* ═══════════════════════════════════════════
   PARÇA BAŞI MALİYET HESABI
   Hesap yapılmışsa gerçek yerleşim verisi,
   yapılmamışsa teorik alan oranı kullanılır.
═══════════════════════════════════════════ */
// Fiyat önbelleği — renderParts başında güncellenir, döngüde DOM sorgusu yapılmaz
let _cachedPrices={sqm:0,cut:0,sheet:0,valid:false};
function refreshPriceCache(){
  _cachedPrices={
    sqm:parseFloat(document.getElementById('sqmPrice').value)||0,
    cut:parseFloat(document.getElementById('cutPrice').value)||0,
    sheet:parseFloat(document.getElementById('sheetPrice').value)||0,
    valid:true
  };
}
function calcPartUnitCost(p){
  if(!_cachedPrices.valid)refreshPriceCache();
  const sqmPrice=_cachedPrices.sqm, cutPrice=_cachedPrices.cut, sheetPrice=_cachedPrices.sheet;

  // Hiçbir fiyat girilmemişse null döndür
  if(!sqmPrice && !cutPrice && !sheetPrice) return null;

  const areaM2 = (p.w * p.h) / 1e6; // mm² → m²

  /* ── 1. m² MALZEME PAYI (parça alanı × m² birim fiyatı) ── */
  const sqmPart = areaM2 * sqmPrice;

  /* ── 2. KESİM PAYI (kesim başı ücret × 1 adet) ── */
  const cutPart = cutPrice;

  /* ── 3. PANEL MALZEME PAYI ──────────────────────────────
     Hesap yapılmışsa: (bu parça toplam alanı / tüm yerleşen alan) × toplam panel maliyeti / adet
     Yapılmamışsa: teorik oran = parça alanı / panel alanı × panel fiyatı                    */
  let panelPart = 0;
  if(sheetPrice > 0){
    if(allPanelsResult.length){
      // Gerçek yerleşim: bu parçanın kaplama oranına düşen panel maliyeti
      const allPlaced      = allPanelsResult.flatMap(pan=>pan.placed);
      const totalUsedArea  = allPlaced.reduce((s,r)=>s+r.w*r.h, 0); // mm²
      const panelCount     = allPanelsResult.length;
      const totalPanelCost = sheetPrice * panelCount;
      if(totalUsedArea > 0){
        const shareOfTotal = (p.w * p.h) / totalUsedArea; // bu parça 1 adedin oranı
        panelPart = shareOfTotal * totalPanelCost;
      }
    } else {
      // Henüz hesap yapılmamış: teorik oran (parça/panel alanı)
      const panelArea = SHEET_W * SHEET_H;
      panelPart = (p.w * p.h / panelArea) * sheetPrice;
    }
  }

  const unitTotal = sqmPart + cutPart + panelPart;
  return { sqmPart, cutPart, panelPart, unitTotal };
}

/* ═══════════════════════════════════════════
   RENDER PARTS — diff-bazlı artımlı güncelleme
   Tüm listeyi sıfırdan çizmek yerine:
   1. Silinen parçaları DOM'dan kaldır
   2. Yeni eklenen parçaları ekle
   3. Değişen parçaları yerinde güncelle
   Bu sayede 50+ parçada scroll pozisyonu ve
   açık edit satırları korunur, CPU yükü düşer.
═══════════════════════════════════════════ */
let _rpCache = {}; // id → serialized snapshot

function _partSnapshot(p){
  // Parçanın render edilen tüm alanlarını tek string'e sıkıştır
  // (maliyet dahil — fiyat değişiminde de güncelleme tetiklenir)
  refreshPriceCache();
  const cost = calcPartUnitCost(p);
  return `${p.name}|${p.w}|${p.h}|${p.qty}|${p.note||''}|${p.color}|${p.rotate}|${currentGrain}|${cost?cost.unitTotal.toFixed(4):'0'}|${getCurrencySymbol()}|${currentLang}`;
}

function _buildPartHTML(p){
  refreshPriceCache();
  const fmtTL = v => v.toLocaleString(currentLang==='en'?'en-GB':'tr-TR',{minimumFractionDigits:2,maximumFractionDigits:2});
  const hasResult = allPanelsResult.length > 0;
  const cost = calcPartUnitCost(p);
  const hasCost = cost && cost.unitTotal > 0;

  let costBreakdownHtml = '';
  if(hasCost){
    const sqmPrice  = _cachedPrices.sqm;
    const cutPrice  = _cachedPrices.cut;
    const sheetPrice= _cachedPrices.sheet;
    const rows = [];
    const isEN=currentLang==='en';
    if(sqmPrice  > 0) rows.push(`<span>${isEN?'Material (m²)':'Malzeme (m²)'}</span><span>${getCurrencySymbol()}${fmtTL(cost.sqmPart)}</span>`);
    if(sheetPrice> 0) rows.push(`<span>${isEN?'Sheet share':'Panel payı'}${hasResult?'':isEN?' <i>(theoretical)</i>':' <i>(teorik)</i>'}</span><span>${getCurrencySymbol()}${fmtTL(cost.panelPart)}</span>`);
    if(cutPrice  > 0) rows.push(`<span>${isEN?'Cutting':'Kesim'}</span><span>${getCurrencySymbol()}${fmtTL(cost.cutPart)}</span>`);
    costBreakdownHtml = `
      <div class="part-cost-breakdown" id="pcb-${p.id}">
        <div class="pcb-inner">
          <div class="pcb-title">${currentLang==='en'?'Unit cost breakdown':'Birim maliyet kırılımı'}</div>
          ${rows.map(r=>`<div class="pcb-row">${r}</div>`).join('')}
          <div class="pcb-total"><span>${currentLang==='en'?'Total / unit':'Toplam / adet'}</span><span>${getCurrencySymbol()}${fmtTL(cost.unitTotal)}</span></div>
          ${p.qty>1?`<div class="pcb-total-qty"><span>${p.qty} ${currentLang==='en'?'units total':'adet toplam'}</span><span>${getCurrencySymbol()}${fmtTL(cost.unitTotal*p.qty)}</span></div>`:''}
        </div>
      </div>`;
  }

  const itemHTML = `
  <div class="part-item" draggable="true" ondragstart="dragStart(event,${p.id})" id="pi-${p.id}" style="--part-color:${p.color};background:${p.color}12">
    <input type="checkbox" class="part-checkbox" data-id="${p.id}" onchange="updateBulkToolbar()" onclick="event.stopPropagation()">
    <div class="part-prio">
      <button class="prio-btn edit-btn" onclick="movePart(event,${p.id},-1)" title="Önce yerleştir (yukarı)">▲</button>
      <button class="prio-btn edit-btn" onclick="movePart(event,${p.id},1)" title="Sonra yerleştir (aşağı)">▼</button>
    </div>
    <div class="part-swatch" style="background:${p.color};width:10px;height:10px;border-radius:3px;flex-shrink:0"></div>
    <div style="flex:1;min-width:0">
      <div class="part-name">${esc(p.name)}</div>
      <div class="part-dim">${p.w}×${p.h} mm${p.qty>1?` · <span style="color:var(--cyan)">${p.qty} adet</span>`:''}${p.note?` · <span style="color:var(--muted);font-size:.59rem" title="${esc(p.note)}">📝 ${esc(p.note.length>22?p.note.substring(0,22)+'…':p.note)}</span>`:''}${p.labelCode?` · <span style="color:var(--accent);font-size:.59rem">🏷 ${esc(p.labelCode)}</span>`:''}${p.materialCode?` · <span style="color:var(--warning);font-size:.59rem">🎨 ${esc(p.materialCode)}</span>`:''}${p.edgeBand?` · <span style="color:var(--success);font-size:.59rem">▰ ${esc(p.edgeBand)}</span>`:''}</div>
    </div>
    ${hasCost?`
      <button class="part-cost-btn" onclick="togglePartCost(event,${p.id})" title="Birim maliyet detayı">
        ${getCurrencySymbol()}${fmtTL(cost.unitTotal)}<span class="part-cost-chevron">▾</span>
      </button>`:''}
    <span class="part-max-badge" title="Bu tek türden maksimum kaç adet sığar">max ${calcMax(p.w,p.h,p.rotate)}</span>
    <div class="part-actions">
      <button class="color-btn edit-btn part-aux" title="Renk değiştir" style="background:${p.color}22;border-radius:4px" onclick="event.stopPropagation()">
        🎨<input type="color" value="${p.color}" oninput="changePartColor(event,${p.id})">
      </button>
      <button class="swap-btn edit-btn part-aux" onclick="swapPartDims(event,${p.id})" title="W↔H Döndür">⇄</button>
      <button class="copy-btn edit-btn part-aux" onclick="duplicatePart(event,${p.id})" title="Kopyala">⧉</button>
      <button class="edit-btn" onclick="startEdit(event,${p.id})" title="Düzenle" style="font-size:13px">✎</button>
      <button class="remove-btn" onclick="removePart(${p.id})" title="Sil">✕</button>
    </div>
  </div>
  ${costBreakdownHtml}
  <div class="part-edit-row" id="er-${p.id}">
    <div style="font-family:var(--mono);font-size:.6rem;color:var(--cyan);text-transform:uppercase;letter-spacing:.1em;margin-bottom:10px;display:flex;align-items:center;gap:6px">
      <span style="display:inline-block;width:8px;height:8px;border-radius:2px;background:${p.color}"></span>
      Parça Düzenle — <span style="color:var(--text)">${esc(p.name)}</span>
    </div>
    <div class="row2" style="margin-bottom:6px">
      <div class="field" style="margin-bottom:0"><label data-i18n="partName">Parça Adı</label><input type="text" id="en-${p.id}" value="${esc(p.name)}" placeholder="Parça adı"></div>
      <div class="field" style="margin-bottom:0"><label data-i18n="note">Not / Açıklama</label><input type="text" id="enote-${p.id}" value="${esc(p.note||'')}" placeholder="Opsiyonel"></div>
    </div>
    <div class="row3" style="margin-bottom:6px">
      <div class="field" style="margin-bottom:0"><label>Etiket Kodu</label><input type="text" id="elabel-${p.id}" value="${esc(p.labelCode||'')}" placeholder="A1"></div>
      <div class="field" style="margin-bottom:0"><label>Renk / Kod</label><input type="text" id="emat-${p.id}" value="${esc(p.materialCode||'')}" placeholder="Beyaz 110"></div>
      <div class="field" style="margin-bottom:0"><label>Kenar Bant</label><input type="text" id="eedge-${p.id}" value="${esc(p.edgeBand||'')}" placeholder="Ön 1mm"></div>
    </div>
    <div class="row3" style="margin-bottom:6px">
      <div class="field" style="margin-bottom:0"><label>Genişlik (mm)</label><input type="number" id="ew-${p.id}" value="${p.w}" min="1" max="10000"></div>
      <div class="field" style="margin-bottom:0"><label>Yükseklik (mm)</label><input type="number" id="eh-${p.id}" value="${p.h}" min="1" max="10000"></div>
      <div class="field" style="margin-bottom:0"><label>Adet</label><input type="number" id="eq-${p.id}" value="${p.qty}" min="1" max="9999"></div>
    </div>
    <div class="toggle-row" style="margin-bottom:10px"><input type="checkbox" id="er-rot-${p.id}" ${p.rotate&&currentGrain==='any'?'checked':''} ${currentGrain!=='any'?'disabled':''}><label for="er-rot-${p.id}">Döndürmeye izin ver (90°) ${currentGrain!=='any'?'<span style="color:var(--orange);font-size:.6rem">(tahıl yönü kilitli)</span>':''}</label></div>
    <div class="edit-actions">
      <button class="btn-save-edit" onclick="saveEdit(${p.id})">✓ &nbsp;Kaydet</button>
      <button class="btn-cancel-edit" onclick="cancelEdit(${p.id})">✕ &nbsp;İptal</button>
    </div>
  </div>`;
  return itemHTML;
}

function renderParts(){
  const el=document.getElementById('partsList');
  if(!parts.length){
    el.innerHTML=`<div class="parts-empty"><div class="parts-empty-icon">📐</div><div><strong style="color:var(--text2)">${t('emptyParts')}</strong><br>${t('emptyPartsHint').replace('\n','<br>')}</div><span class="parts-empty-cta" data-i18n="emptyPartsAdd" onclick="document.getElementById('partName').focus()">${t('emptyPartsAdd')}</span></div>`;
    document.getElementById('maxSummary').className='max-summary';
    document.getElementById('pngBtn').disabled=true;
    document.getElementById('pdfBtn').disabled=true;
    _rpCache={};
    return;
  }

  // ── 1. Silinenleri DOM'dan kaldır
  const currentIds=new Set(parts.map(p=>p.id));
  Object.keys(_rpCache).forEach(id=>{
    if(!currentIds.has(parseInt(id))){
      ['pi-','pcb-','er-'].forEach(prefix=>{
        const el2=document.getElementById(prefix+id);
        if(el2)el2.remove();
      });
      delete _rpCache[id];
    }
  });

  // ── 2. Sıra değişikliğini veya yeni parçaları uygula
  // Beklenen DOM sırası: parts dizisinin sırası
  let prevSibling=null; // insertAfter referansı
  for(let i=0;i<parts.length;i++){
    const p=parts[i];
    const snap=_partSnapshot(p);
    const existingItem=document.getElementById('pi-'+p.id);

    if(!existingItem){
      // Yeni parça — oluştur ve doğru konuma ekle
      const tmp=document.createElement('div');
      tmp.innerHTML=_buildPartHTML(p);
      const nodes=[...tmp.childNodes].filter(n=>n.nodeType===1||n.nodeType===3&&n.textContent.trim());
      // Gerçek elementleri al
      const elems=[];
      tmp.querySelectorAll(':scope > *').forEach(n=>elems.push(n));
      if(elems.length){
        const ref=prevSibling?prevSibling.nextSibling:el.firstChild;
        elems.forEach(n=>el.insertBefore(n,ref||null));
        prevSibling=elems[elems.length-1];
      }
      _rpCache[p.id]=snap;
    } else {
      // Var olan — sıra kontrolü
      // Doğru konumda değilse taşı
      const expectedPrev=i===0?null:document.getElementById('er-'+parts[i-1].id);
      const actualPrev=existingItem.previousElementSibling;
      // Sıra yanlışsa — tüm üç elementi (pi, pcb, er) birlikte taşı
      if((expectedPrev===null&&el.firstElementChild!==existingItem&&el.firstElementChild&&!el.firstElementChild.classList.contains('parts-empty'))||
         (expectedPrev&&actualPrev!==expectedPrev)){
        const pcb=document.getElementById('pcb-'+p.id);
        const er=document.getElementById('er-'+p.id);
        const anchor=expectedPrev?expectedPrev.nextSibling:el.firstChild;
        el.insertBefore(existingItem,anchor||null);
        if(pcb)el.insertBefore(pcb,existingItem.nextSibling);
        if(er)el.insertBefore(er,pcb?(pcb.nextSibling):(existingItem.nextSibling));
      }
      // İçerik değişti mi?
      if(_rpCache[p.id]!==snap){
        const tmp=document.createElement('div');
        tmp.innerHTML=_buildPartHTML(p);
        const newItem=tmp.querySelector('#pi-'+p.id);
        const newPcb=tmp.querySelector('#pcb-'+p.id);
        const newEr=tmp.querySelector('#er-'+p.id);
        // Sadece edit satırı açık değilse item'ı güncelle (yoksa inputlar sıfırlanır)
        const erEl=document.getElementById('er-'+p.id);
        const editOpen=erEl&&erEl.classList.contains('visible');
        if(!editOpen&&newItem){existingItem.replaceWith(newItem);}
        // cost breakdown her zaman güncelle
        const pcbEl=document.getElementById('pcb-'+p.id);
        if(newPcb&&pcbEl){pcbEl.replaceWith(newPcb);}
        else if(newPcb&&!pcbEl){
          const itemEl=document.getElementById('pi-'+p.id);
          if(itemEl)itemEl.after(newPcb);
        }
        _rpCache[p.id]=snap;
      }
      prevSibling=document.getElementById('er-'+p.id)||document.getElementById('pcb-'+p.id)||existingItem;
    }
  }

  // ── 3. Ortak güncellemeler
  document.getElementById('partsSearchWrap').style.display=parts.length>=3?'block':'none';
  const _sq=document.getElementById('partsSearch');
  if(_sq&&_sq.value)filterParts(_sq.value);
  updateMaxSummary();
}

function filterParts(q){
  const clear=document.getElementById('partsSearchClear');
  if(clear)clear.className='parts-search-clear'+(q?'  visible':'');
  const items=document.querySelectorAll('#partsList .part-item, #partsList .part-cost-breakdown, #partsList .part-edit-row');
  if(!q){items.forEach(el=>el.style.display='');return;}
  const ql=q.toLowerCase();
  // part-item'ları gizle/göster — ardışık cost-breakdown ve edit-row da eşleşene göre
  const partItems=document.querySelectorAll('#partsList .part-item');
  partItems.forEach(item=>{
    const id=item.id.replace('pi-','');
    const p=parts.find(x=>x.id===parseInt(id));
    const match=p&&(p.name.toLowerCase().includes(ql)||(p.note||'').toLowerCase().includes(ql)||String(p.w).includes(ql)||String(p.h).includes(ql));
    item.style.display=match?'':'none';
    const bd=document.getElementById('pcb-'+id);
    const er=document.getElementById('er-'+id);
    if(bd)bd.style.display=match?'':'none';
    if(er){
      if(!match && er.classList.contains('visible')) cancelEdit(parseInt(id));
      er.style.display='none'; // edit row her zaman gizli filtreli
    }
  });
}
function clearPartsSearch(){
  const inp=document.getElementById('partsSearch');
  if(inp)inp.value='';
  filterParts('');
}

function togglePartCost(e, id){
  e.stopPropagation();
  const box = document.getElementById('pcb-'+id);
  if(!box) return;
  const isOpen = box.classList.contains('open');
  // diğerlerini kapat
  document.querySelectorAll('.part-cost-breakdown.open').forEach(el=>el.classList.remove('open'));
  document.querySelectorAll('.part-cost-btn.active').forEach(el=>el.classList.remove('active'));
  if(!isOpen){
    box.classList.add('open');
    e.currentTarget.classList.add('active');
  }
}

/* ═══════════════════════════════════════════
   INLINE EDIT
═══════════════════════════════════════════ */
let activeEdit=null;
function startEdit(e,id){
  e.stopPropagation();
  if(activeEdit&&activeEdit!==id)cancelEdit(activeEdit);
  activeEdit=id;
  const row=document.getElementById('er-'+id);
  row.className='part-edit-row visible';
  const item=document.getElementById('pi-'+id);
  item.style.outline='1.5px solid rgba(129,140,248,.4)';
  item.style.borderRadius='7px';
  item.setAttribute('draggable','false');
  // Enter → kaydet, Esc → iptal
  row.querySelectorAll('input').forEach(inp=>{
    inp._editHandler=ev=>{
      if(ev.key==='Enter'){ev.preventDefault();saveEdit(id);}
      if(ev.key==='Escape'){ev.preventDefault();cancelEdit(id);}
    };
    inp.addEventListener('keydown',inp._editHandler);
  });
  // İsim alanına odaklan
  setTimeout(()=>{const n=document.getElementById('en-'+id);if(n){n.focus();n.select();}},50);
}
function cancelEdit(id){
  const r=document.getElementById('er-'+id);
  const i=document.getElementById('pi-'+id);
  if(r){
    // Event listener'ları temizle
    r.querySelectorAll('input').forEach(inp=>{
      if(inp._editHandler){inp.removeEventListener('keydown',inp._editHandler);delete inp._editHandler;}
    });
    r.className='part-edit-row';
  }
  if(i){i.style.outline='';i.style.borderRadius='';i.setAttribute('draggable','true');}
  if(activeEdit===id)activeEdit=null;
}
function duplicatePart(e,id){
  e.stopPropagation();
  const p=parts.find(x=>x.id===id);if(!p)return;
  pushUndo(parts,'Parça kopyalama');
  const copy={...p,id:genId(),name:p.name+' (Kopya)',color:getColor()};
  const idx=parts.findIndex(x=>x.id===id);
  parts.splice(idx+1,0,copy);
  renderParts();save();
  showMsg(esc(copy.name)+' eklendi.','success');
}
function changePartColor(e, id){
  const p=parts.find(x=>x.id===id);if(!p)return;
  p.color=e.target.value;
  // Sadece o parçanın swatch ve arka planını güncelle — tam re-render gereksiz
  const item=document.getElementById('pi-'+id);
  if(item){
    item.style.background=p.color+'12';
    item.style.setProperty('--part-color',p.color);
    const swatch=item.querySelector('.part-swatch');
    if(swatch)swatch.style.background=p.color;
    const colorBtn=item.querySelector('.color-btn');
    if(colorBtn)colorBtn.style.background=p.color+'22';
  }
  save();
}

/* ═══════════════════════════════════════════
   TOPLU DÜZENLEME
═══════════════════════════════════════════ */
function getSelectedIds(){
  return [...document.querySelectorAll('#partsList .part-checkbox:checked')].map(el=>parseInt(el.dataset.id)).filter(Boolean);
}
function updateBulkToolbar(){
  const ids=getSelectedIds();
  const tb=document.getElementById('bulkToolbar');
  const cnt=document.getElementById('bulkCount');
  const sa=document.getElementById('selectAllParts');
  if(tb)tb.className='bulk-toolbar'+(ids.length>0?' visible':'');
  if(cnt)cnt.textContent=ids.length+' '+(currentLang==='en'?'selected':t('selected'));
  if(sa)sa.indeterminate=ids.length>0&&ids.length<parts.length;
  if(sa)sa.checked=ids.length===parts.length&&parts.length>0;
}
function toggleSelectAll(checked){
  document.querySelectorAll('#partsList .part-checkbox').forEach(el=>el.checked=checked);
  updateBulkToolbar();
}
function clearBulkSelection(){
  document.querySelectorAll('#partsList .part-checkbox').forEach(el=>el.checked=false);
  updateBulkToolbar();
}
function bulkDelete(){
  const ids=getSelectedIds();
  if(!ids.length)return;
  mobileConfirm(ids.length+' parça silinsin mi?').then(ok=>{
    if(!ok)return;
    pushUndo(parts,'Toplu silme');
    parts=parts.filter(p=>!ids.includes(p.id));
    renderParts();save();
    showMsg(ids.length+t('msgPartDeleted'),'success');
  });
}
function bulkDuplicate(){
  const ids=getSelectedIds();
  if(!ids.length)return;
  pushUndo(parts,'Toplu kopyalama');
  const newParts=[];
  ids.forEach(id=>{
    const p=parts.find(x=>x.id===id);
    if(p)newParts.push({...p,id:genId(),name:p.name+' (Kopya)',color:getColor()});
  });
  parts=[...parts,...newParts];
  renderParts();save();
  showMsg(ids.length+t('msgPartDuplicated'),'success');
}
function bulkMoveUp(){
  const ids=getSelectedIds();
  if(!ids.length)return;
  pushUndo(parts,'Toplu yukarı');
  const arr=[...parts];
  for(let i=1;i<arr.length;i++){
    if(ids.includes(arr[i].id)&&!ids.includes(arr[i-1].id)){
      [arr[i-1],arr[i]]=[arr[i],arr[i-1]];
    }
  }
  parts=arr;renderParts();save();
}
function bulkMoveDown(){
  const ids=getSelectedIds();
  if(!ids.length)return;
  pushUndo(parts,'Toplu aşağı');
  const arr=[...parts];
  for(let i=arr.length-2;i>=0;i--){
    if(ids.includes(arr[i].id)&&!ids.includes(arr[i+1].id)){
      [arr[i],arr[i+1]]=[arr[i+1],arr[i]];
    }
  }
  parts=arr;renderParts();save();
}

/* ═══════════════════════════════════════════
   GRUPLAMA & SIRALAMA
═══════════════════════════════════════════ */

/* ═══════════════════════════════════════════
   MALZEME FİYAT LİSTESİ
   m² birim fiyatlarını malzeme türüne göre sakla
═══════════════════════════════════════════ */
function saveMaterialPrice(){
  const mat=document.getElementById('materialType').value;
  const price=parseFloat(document.getElementById('sqmPrice').value)||0;
  if(!mat||!price)return;
  try{
    const list=JSON.parse(localStorage.getItem('pop9_pricelist')||'{}');
    list[mat]=price;
    localStorage.setItem('pop9_pricelist',JSON.stringify(list));
    showMsg(mat+t('msgPriceSaved')+getCurrencySymbol()+price.toFixed(2),'success');
  }catch(e){}
}
function loadMaterialPrice(mat){
  if(!mat)return;
  try{
    const list=JSON.parse(localStorage.getItem('pop9_pricelist')||'{}');
    if(list[mat]!==undefined){
      document.getElementById('sqmPrice').value=list[mat];
      _cachedPrices.valid=false;
      updateCost();
      if(allPanelsResult.length)buildCutList(allPanelsResult);
      showMsg(mat+t('msgPriceLoaded')+getCurrencySymbol()+list[mat].toFixed(2),'info');
    }
  }catch(e){}
}
function mergeSameParts(){
  if(!parts.length)return;
  pushUndo(parts,'Aynıları birleştir');
  const groups={};
  parts.forEach(p=>{
    // Aynı boyut + döndürme + renk → aynı grup
    const key=`${Math.min(p.w,p.h)}x${Math.max(p.w,p.h)}_r${p.rotate?1:0}`;  // rotate dahil
    if(!groups[key]){groups[key]={...p,qty:p.qty,name:p.name};}
    else{groups[key].qty+=p.qty;}
  });
  const merged=Object.values(groups);
  const saved=parts.length-merged.length;
  parts=merged;
  _rpCache={};
  renderParts();save();
  if(saved>0)showMsg(saved+t('msgPartsMerged')+merged.length+' '+(currentLang==='tr'?'benzersiz tür':'unique types'),'success');
  else showMsg(t('msgMergeNone'),'info');
}
function sortPartsBySize(){
  if(!parts.length)return;
  pushUndo(parts,'Boyuta göre sırala');
  parts=[...parts].sort((a,b)=>(b.w*b.h)-(a.w*a.h));
  renderParts();save();
  showMsg(t('msgSortDone'),'success');
}
function movePart(e,id,dir){
  e.stopPropagation();
  const idx=parts.findIndex(x=>x.id===id);
  if(idx<0)return;
  const newIdx=idx+dir;
  if(newIdx<0||newIdx>=parts.length)return;
  pushUndo(parts,'Öncelik değiştir');
  const arr=[...parts];
  [arr[idx],arr[newIdx]]=[arr[newIdx],arr[idx]];
  parts=arr;
  renderParts();save();
}
function swapPartDims(e, id){
  e.stopPropagation();
  const p=parts.find(x=>x.id===id);if(!p)return;
  pushUndo(parts,'Boyut döndürme');
  const tmp=p.w; p.w=p.h; p.h=tmp;
  renderParts();save();
  showMsg(esc(p.name)+': '+p.w+'×'+p.h+' mm','info');
}
function saveEdit(id){
  const p=parts.find(x=>x.id===id);if(!p)return;
  const nw=parseInt(document.getElementById('ew-'+id).value);
  const nh=parseInt(document.getElementById('eh-'+id).value);
  const nq=parseInt(document.getElementById('eq-'+id).value)||1;
  const rot=document.getElementById('er-rot-'+id).checked;
  // Boyut doğrulama
  if(nw>0&&nh>0){
    const fw=nw, fh=nh;
    if(!(fw<=SHEET_W&&fh<=SHEET_H)&&!(rot&&fh<=SHEET_W&&fw<=SHEET_H)){
      showMsg(`(${fw}×${fh}) ${t('msgNoFit')}`,'error');return;
    }
    if(fw>10000||fh>10000){showMsg(t('msgMaxSize'),'error');return;}
  }
  pushUndo(parts,'Parça düzenleme');
  const newName = document.getElementById('en-'+id).value.trim();
  p.name = newName || p.name; // boşsa eski ismi koru
  p.note=document.getElementById('enote-'+id).value.trim();
  p.labelCode=(document.getElementById('elabel-'+id)||{value:''}).value.trim();
  p.materialCode=(document.getElementById('emat-'+id)||{value:''}).value.trim();
  p.edgeBand=(document.getElementById('eedge-'+id)||{value:''}).value.trim();
  if(nw>0)p.w=nw;if(nh>0)p.h=nh;p.qty=nq;
  p.rotate=rot;
  cancelEdit(id);renderParts();save();
  showMsg(esc(p.name)+t('msgPartUpdated'),'success');
}

/* ═══════════════════════════════════════════
   DRAG REORDER
═══════════════════════════════════════════ */
let draggedId=null;
function dragStart(e,id){draggedId=id}
document.getElementById('partsList').addEventListener('dragover',e=>e.preventDefault());
document.getElementById('partsList').addEventListener('drop',e=>{
  e.preventDefault();if(draggedId===null)return;
  const idx=parts.findIndex(p=>p.id===draggedId);if(idx<0)return;
  const items=[...document.querySelectorAll('#partsList .part-item')];
  let ti=parts.length;
  for(let i=0;i<items.length;i++){const r=items[i].getBoundingClientRect();if(e.clientY<r.top+r.height/2){ti=i;break}}
  const arr=[...parts],d=arr.splice(idx,1)[0];arr.splice(ti>idx?ti-1:ti,0,d);parts=arr;_rpCache={};renderParts();draggedId=null;save();
});

/* ═══════════════════════════════════════════
   CLEAR
═══════════════════════════════════════════ */
function clearAll(){
  mobileConfirm('Tüm parçalar silinsin mi?').then(ok=>{
    if(!ok)return;
    pushUndo(parts,'Tümünü temizle');
    parts=[];colorIdx=0;renderParts();
    document.getElementById('mainCanvas').style.display='none';
    document.getElementById('scaleBarWrap').className='scale-bar-wrap';
    document.getElementById('placeholder').style.display='flex';
    document.getElementById('legend').innerHTML=`<span style="color:var(--muted);font-size:.62rem;font-family:var(--mono)">${t('legendHint')}</span>`;
    document.getElementById('unfitList').className='unfit-list';
    document.getElementById('maxSummary').className='max-summary';
    document.getElementById('pngBtn').disabled=true;document.getElementById('pdfBtn').disabled=true;
    const _svgBtn=document.getElementById('svgBtn');if(_svgBtn)_svgBtn.disabled=true;
    const _pab=document.getElementById('pngAllBtn');if(_pab){_pab.disabled=true;_pab.style.display='none';}
    document.getElementById('quickSingleResult').className='quick-result';
    document.getElementById('quickAllResult').className='quick-result';
    document.getElementById('canvasInfo').textContent=t('canvasEmpty');
    document.getElementById('zoomHint').className='zoom-hint';
    document.getElementById('cutListWrap').style.display='none';
    document.getElementById('panelNav').style.display='none';
    var _cr0=document.getElementById('costResult');if(_cr0)_cr0.className='cost-result';
    const _sa=document.getElementById('stockAlert');if(_sa)_sa.style.display='none';
    ['s-total','s-placed','s-usage','s-waste','s-panels'].forEach(id=>{
      const e=document.getElementById(id);e.textContent='—';e.className='stat-val'
    });
    document.getElementById('s-waste-area').textContent='—';
    zoomLevel=1;document.getElementById('zoomLevelLbl').textContent='100%';
    document.getElementById('canvasInner').style.transform='scale(1)';
    allPanelsResult=[];
    undoStack=[];redoStack=[];
    updateUndoRedoBtns();
    clearBulkSelection();
    save();
  });
}

/* ═══════════════════════════════════════════
   STAT ANIM
═══════════════════════════════════════════ */
function animStat(id,v,c){
  const e=document.getElementById(id);
  e.className='stat-val animating'+(c?' '+c:'');
  e.textContent=v;
  setTimeout(()=>e.classList.remove('animating'),320);
}

/* ═══════════════════════════════════════════
   MAX CALC
═══════════════════════════════════════════ */
function calcMax(pw,ph,rot){
  const{w:SW,h:SH}=getEffectiveSheet();
  const{cut}=getMargins();
  // Kesim payı eklenerek efektif yer kaplar
  const ew=pw+(cut>0?cut:0), eh=ph+(cut>0?cut:0);
  if(ew>SW||eh>SH){
    // Döndürme olmadan sığmıyor, döndürme dene
    if(rot&&eh<=SW&&ew<=SH){
      return calcMaxGrid(eh,ew,SW,SH,rot);
    }
    return 0;
  }
  return calcMaxGrid(ew,eh,SW,SH,rot);
}
function calcMaxGrid(ew,eh,SW,SH,rot){
  let best=Math.floor(SW/ew)*Math.floor(SH/eh);
  if(rot){
    best=Math.max(best,Math.floor(SW/eh)*Math.floor(SH/ew));
    // Karma yerleşim optimizasyonu
    const rN=Math.floor(SH/eh);
    for(let r=0;r<=rN;r++){
      const rem=SH-r*eh;
      best=Math.max(best,r*Math.floor(SW/ew)+(rem>=ew?Math.floor(rem/ew):0)*Math.floor(SW/eh));
    }
    const cN=Math.floor(SW/ew);
    for(let c=0;c<=cN;c++){
      const rem=SW-c*ew;
      best=Math.max(best,c*Math.floor(SH/eh)+(rem>=eh?Math.floor(rem/eh):0)*Math.floor(SH/ew));
    }
  }
  return best;
}
function updateMaxSummary(){
  const box=document.getElementById('maxSummary');
  const rows=document.getElementById('maxSummaryRows');
  const total=document.getElementById('maxSummaryTotal');
  if(!parts.length){box.className='max-summary';return}
  let h='',g=0;const seen=new Set();
  for(const p of parts){
    if(seen.has(p.id))continue;seen.add(p.id);
    const m=calcMax(p.w,p.h,p.rotate);g+=m;
    h+=`<div class="max-summary-row"><span class="max-summary-name"><span style="display:inline-block;width:7px;height:7px;border-radius:2px;background:${p.color};margin-right:5px;vertical-align:middle"></span>${esc(p.name)} <span style="color:var(--muted2)">(${p.w}×${p.h})</span></span><span class="max-summary-val">${m}</span></div>`;
  }
  rows.innerHTML=h;total.textContent=g+' adet';box.className='max-summary visible';
}

/* ═══════════════════════════════════════════
   QUICK CALC
═══════════════════════════════════════════ */
let quickTab='single';
function switchQuickTab(tab){
  quickTab=tab;
  document.getElementById('tab-single').className='tab-btn'+(tab==='single'?' active':'');
  document.getElementById('tab-all').className='tab-btn'+(tab==='all'?' active':'');
  document.getElementById('quickSingleMode').style.display=tab==='single'?'block':'none';
  document.getElementById('quickAllMode').style.display=tab==='all'?'block':'none';
}
function calcQuickSingle(){
  const pw=parseInt(document.getElementById('qpw').value);
  const ph=parseInt(document.getElementById('qph').value);
  const rot=document.getElementById('qAllowRotate').checked;
  const res=document.getElementById('quickSingleResult');
  if(!pw||!ph){res.className='quick-result';return}
  getSheet();
  const{w:SW,h:SH}=getEffectiveSheet();
  const cN=Math.floor(SW/pw),rN=Math.floor(SH/ph),nc=cN*rN;
  const rc=rot?Math.floor(SW/ph)*Math.floor(SH/pw):0;
  const mx=calcMax(pw,ph,rot),u=((mx*pw*ph)/(SW*SH)*100).toFixed(1);
  document.getElementById('qr-normal').textContent=nc+' adet ('+cN+'×'+rN+')';
  document.getElementById('qr-rotated').textContent=rot?rc+' adet':'—';
  document.getElementById('qr-max').textContent=mx+' adet';
  document.getElementById('qr-usage').textContent=u+'%';
  res.className='quick-result visible';
}
function calcQuickAll(){
  getSheet();const res=document.getElementById('quickAllResult');
  if(!parts.length){res.innerHTML=`<div class="quick-result-row"><span class="qr-label">${t('msgAddParts')}</span></div>`;res.className='quick-result visible';return}
  const seen=new Set();let h='',tot=0;
  const{w:SW,h:SH}=getEffectiveSheet();
  for(const p of parts){
    if(seen.has(p.id))continue;seen.add(p.id);
    const m=calcMax(p.w,p.h,p.rotate);tot+=m;
    const u=((m*p.w*p.h)/(SW*SH)*100).toFixed(1);
    h+=`<div class="quick-result-row"><span class="qr-label"><span style="display:inline-block;width:7px;height:7px;border-radius:2px;background:${p.color};margin-right:4px;vertical-align:middle"></span>${esc(p.name)} <span style="opacity:.5">(${p.w}×${p.h})</span></span><span class="qr-val good">${m} adet</span></div><div class="quick-result-row" style="padding:2px 11px 7px"><span class="qr-label" style="font-size:.64rem">Alan kullanımı (tek tür)</span><span style="color:var(--muted2);font-size:.67rem;font-family:var(--mono)">${u}%</span></div>`;
  }
  h+=`<div class="quick-result-row" style="border-top:1px solid rgba(52,211,153,.2)"><span style="font-weight:700;color:var(--text);font-size:.73rem">Toplam kapasite</span><span class="qr-val good" style="font-size:.98rem">${tot} adet</span></div>`;
  res.innerHTML=h;res.className='quick-result visible';
}

/* ═══════════════════════════════════════════
   PACKING ALGORITHMS
═══════════════════════════════════════════ */

/* --- MAXRECTS --- */
/* ═══════════════════════════════════════════
   YARDIMCI FONKSİYONLAR
═══════════════════════════════════════════ */
function intersects(a,b){return!(b.x>=a.x+a.w||b.x+b.w<=a.x||b.y>=a.y+a.h||b.y+b.h<=a.y)}

/* Dominated (içinde kalan) boşlukları temizle */
function pruneFreeRects(rects){
  const n=rects.length;
  const keep=new Uint8Array(n).fill(1);
  for(let i=0;i<n;i++){
    if(!keep[i])continue;
    const a=rects[i];
    for(let j=0;j<n;j++){
      if(i===j||!keep[j])continue;
      const b=rects[j];
      // b, a'yı tamamen kaplıyorsa a'yı at
      if(b.x<=a.x&&b.y<=a.y&&b.x+b.w>=a.x+a.w&&b.y+b.h>=a.y+a.h){keep[i]=0;break}
    }
  }
  return rects.filter((_,i)=>keep[i]);
}

/* ═══════════════════════════════════════════
   MAXRECTS — BSSF + BAF ikili skor
   Düzeltme 1: Alan bazlı sıralama (w*h desc)
   Düzeltme 2: BSSF (Best Short Side Fit) birincil,
               BAF (Best Area Fit) ikincil skor
═══════════════════════════════════════════ */
function packMaxRects(rects, SW, SH){
  let freeRects=[{x:0,y:0,w:SW,h:SH}];
  const placed=[],unplaced=[];

  // Düzeltme: Alan bazlı sıralama — max(w,h) yerine w*h daha dengeli
  const sorted=[...rects].sort((a,b)=>b.w*b.h-a.w*a.h);

  for(const rect of sorted){
    // İkili skor: [shortSideFit, areaFit] — önce kısa kenar, sonra alan
    let bestSS=Infinity,bestArea=Infinity,bestFr=null,bestRot=false;

    for(const fr of freeRects){
      // Normal
      if(rect.w<=fr.w&&rect.h<=fr.h){
        const ss=Math.min(fr.w-rect.w, fr.h-rect.h);
        const area=fr.w*fr.h-rect.w*rect.h;
        if(ss<bestSS||(ss===bestSS&&area<bestArea)){
          bestSS=ss;bestArea=area;bestFr=fr;bestRot=false;
        }
      }
      // Döndürülmüş
      if(rect.rotate&&rect.h<=fr.w&&rect.w<=fr.h){
        const ss=Math.min(fr.w-rect.h, fr.h-rect.w);
        const area=fr.w*fr.h-rect.h*rect.w;
        if(ss<bestSS||(ss===bestSS&&area<bestArea)){
          bestSS=ss;bestArea=area;bestFr=fr;bestRot=true;
        }
      }
    }

    if(!bestFr){unplaced.push(rect);continue}
    const pw=bestRot?rect.h:rect.w, ph=bestRot?rect.w:rect.h;
    placed.push({x:bestFr.x,y:bestFr.y,w:pw,h:ph,
      id:rect.id,label:rect.label,name:rect.name,
      color:rect.color,rotated:bestRot,
      origW:rect.origW,origH:rect.origH});

    // Kesişen boşlukları böl
    const newFree=[];
    for(const fr of freeRects){
      if(!intersects(fr,{x:bestFr.x,y:bestFr.y,w:pw,h:ph})){newFree.push(fr);continue}
      if(bestFr.x>fr.x)         newFree.push({x:fr.x,       y:fr.y,       w:bestFr.x-fr.x,              h:fr.h});
      if(bestFr.x+pw<fr.x+fr.w) newFree.push({x:bestFr.x+pw,y:fr.y,       w:fr.x+fr.w-(bestFr.x+pw),    h:fr.h});
      if(bestFr.y>fr.y)         newFree.push({x:fr.x,       y:fr.y,       w:fr.w,h:bestFr.y-fr.y});
      if(bestFr.y+ph<fr.y+fr.h) newFree.push({x:fr.x,       y:bestFr.y+ph,w:fr.w,h:fr.y+fr.h-(bestFr.y+ph)});
    }
    freeRects=pruneFreeRects(newFree);
  }
  return{placed,unplaced,freeRects};
}

/* ═══════════════════════════════════════════
   FFD — First Fit Decreasing (gerçek)
   Düzeltme: Artık gerçekten first-fit yapıyor.
   İlk sığan boşluğa yerleştirip durur — best-fit değil.
   Guillotine split ile boşlukları böler, büyük→küçük
   sıralı tutarak sonraki parçaların şansını artırır.
═══════════════════════════════════════════ */
function packFFD(rects, SW, SH){
  const sorted=[...rects].sort((a,b)=>b.w*b.h-a.w*a.h);
  const placed=[],unplaced=[];
  // Boşlukları büyükten küçüğe sıralı tut
  let spaces=[{x:0,y:0,w:SW,h:SH}];

  for(const rect of sorted){
    let fit=false;
    for(let i=0;i<spaces.length;i++){
      const sp=spaces[i];
      let rot=false, pw=rect.w, ph=rect.h;
      if(pw<=sp.w&&ph<=sp.h){
        // normal sığıyor
      } else if(rect.rotate&&ph<=sp.w&&pw<=sp.h){
        rot=true; pw=rect.h; ph=rect.w;
      } else continue; // bu boşluğa sığmıyor, sonrakine geç

      placed.push({x:sp.x,y:sp.y,w:pw,h:ph,
        id:rect.id,label:rect.label,name:rect.name,
        color:rect.color,rotated:rot,
        origW:rect.origW,origH:rect.origH});

      // Shorter Leftover Axis (SLA) split — daha büyük artığı önce koy
      const remW=sp.w-pw, remH=sp.h-ph;
      const newSpaces=[];
      if(remW>0&&ph>0)   newSpaces.push({x:sp.x+pw,y:sp.y,   w:remW,h:ph});
      if(remH>0&&sp.w>0) newSpaces.push({x:sp.x,   y:sp.y+ph,w:sp.w,h:remH});
      // Boşlukları büyükten küçüğe sıralı ekle
      spaces.splice(i,1,...newSpaces.sort((a,b)=>b.w*b.h-a.w*a.h));
      fit=true;
      break; // first fit — ilk sığana yerleştir
    }
    if(!fit)unplaced.push(rect);
  }
  return{placed,unplaced,freeRects:spaces};
}

/* ═══════════════════════════════════════════
   GUILLOTINE — MAXRECTS'e yakın kalite
   Düzeltme: Shorter Leftover Axis split
   Her bölmede kısa artık yatay mı dikey mi
   olduğunu kontrol ederek split yönünü seçer.
   Bu sayede uzun ince boşluk birikmesi azalır.
═══════════════════════════════════════════ */
function packGuillotine(rects, SW, SH){
  // Alan bazlı sıralama
  const sorted=[...rects].sort((a,b)=>b.w*b.h-a.w*a.h);
  let spaces=[{x:0,y:0,w:SW,h:SH}];
  const placed=[],unplaced=[];

  for(const rect of sorted){
    let bs=null,bi=-1,rot=false,bestSS=Infinity,bestArea=Infinity;

    for(let i=0;i<spaces.length;i++){
      const sp=spaces[i];
      // Normal
      if(rect.w<=sp.w&&rect.h<=sp.h){
        const ss=Math.min(sp.w-rect.w,sp.h-rect.h);
        const area=sp.w*sp.h-rect.w*rect.h;
        if(ss<bestSS||(ss===bestSS&&area<bestArea)){
          bestSS=ss;bestArea=area;bs=sp;bi=i;rot=false;
        }
      }
      // Döndürülmüş
      if(rect.rotate&&rect.h<=sp.w&&rect.w<=sp.h){
        const ss=Math.min(sp.w-rect.h,sp.h-rect.w);
        const area=sp.w*sp.h-rect.h*rect.w;
        if(ss<bestSS||(ss===bestSS&&area<bestArea)){
          bestSS=ss;bestArea=area;bs=sp;bi=i;rot=true;
        }
      }
    }

    if(!bs){unplaced.push(rect);continue}
    const pw=rot?rect.h:rect.w, ph=rot?rect.w:rect.h;
    placed.push({x:bs.x,y:bs.y,w:pw,h:ph,
      id:rect.id,label:rect.label,name:rect.name,
      color:rect.color,rotated:rot,
      origW:rect.origW,origH:rect.origH});

    spaces.splice(bi,1);

    // Shorter Leftover Axis split:
    // Sağ artık (bs.w-pw) vs alt artık (bs.h-ph) hangisi kısa?
    const remRight=bs.w-pw, remBottom=bs.h-ph;
    if(remRight<remBottom){
      // Kısa artık sağda → yatay split önce
      if(remRight>0&&ph>0)   spaces.push({x:bs.x+pw,y:bs.y,   w:remRight,h:ph});
      if(remBottom>0&&bs.w>0)spaces.push({x:bs.x,   y:bs.y+ph,w:bs.w,    h:remBottom});
    } else {
      // Kısa artık altta → dikey split önce
      if(remBottom>0&&pw>0)  spaces.push({x:bs.x,   y:bs.y+ph,w:pw,      h:remBottom});
      if(remRight>0&&bs.h>0) spaces.push({x:bs.x+pw,y:bs.y,   w:remRight,h:bs.h});
    }
  }
  return{placed,unplaced,freeRects:spaces};
}


/* ═══════════════════════════════════════════
   SHELF / HYBRID GRID EK ALGORİTMALARI
   Tasarım değişmeden yerleşim kalitesini artırır.
═══════════════════════════════════════════ */
function clonePlaced(rect,x,y,w,h,rotated){
  return {x,y,w,h,id:rect.id,label:rect.label,name:rect.name,color:rect.color,rotated:!!rotated,origW:rect.origW,origH:rect.origH};
}
function _partArea(r){return (r.w||0)*(r.h||0)}
function _isSquareLike(r){
  const mn=Math.min(r.w,r.h),mx=Math.max(r.w,r.h);
  return mn>0 && mx/mn<=1.08;
}
function _splitFreeRectsByPlaced(freeRects,used){
  const out=[];
  for(const fr of freeRects){
    if(!intersects(fr,used)){out.push(fr);continue}
    if(used.x>fr.x) out.push({x:fr.x,y:fr.y,w:used.x-fr.x,h:fr.h});
    if(used.x+used.w<fr.x+fr.w) out.push({x:used.x+used.w,y:fr.y,w:fr.x+fr.w-(used.x+used.w),h:fr.h});
    if(used.y>fr.y) out.push({x:fr.x,y:fr.y,w:fr.w,h:used.y-fr.y});
    if(used.y+used.h<fr.y+fr.h) out.push({x:fr.x,y:used.y+used.h,w:fr.w,h:fr.y+fr.h-(used.y+used.h)});
  }
  return pruneFreeRects(out.filter(r=>r.w>0&&r.h>0));
}
function _placeWithExistingFreeRects(rects,freeRects){
  const placed=[],unplaced=[];
  const sorted=[...rects].sort((a,b)=>_partArea(b)-_partArea(a));
  for(const rect of sorted){
    let best=null,bestRot=false,bestSS=Infinity,bestArea=Infinity;
    for(const fr of freeRects){
      if(rect.w<=fr.w&&rect.h<=fr.h){
        const ss=Math.min(fr.w-rect.w,fr.h-rect.h),area=fr.w*fr.h-_partArea(rect);
        if(ss<bestSS||(ss===bestSS&&area<bestArea)){best=fr;bestRot=false;bestSS=ss;bestArea=area}
      }
      if(rect.rotate&&rect.h<=fr.w&&rect.w<=fr.h){
        const ss=Math.min(fr.w-rect.h,fr.h-rect.w),area=fr.w*fr.h-_partArea(rect);
        if(ss<bestSS||(ss===bestSS&&area<bestArea)){best=fr;bestRot=true;bestSS=ss;bestArea=area}
      }
    }
    if(!best){unplaced.push(rect);continue}
    const pw=bestRot?rect.h:rect.w,ph=bestRot?rect.w:rect.h;
    placed.push(clonePlaced(rect,best.x,best.y,pw,ph,bestRot));
    freeRects=_splitFreeRectsByPlaced(freeRects,{x:best.x,y:best.y,w:pw,h:ph});
  }
  return {placed,unplaced,freeRects};
}
function packShelf(rects,SW,SH){
  const sorted=[...rects].sort((a,b)=>{
    const sq=(_isSquareLike(b)?1:0)-(_isSquareLike(a)?1:0);
    return sq||Math.max(b.w,b.h)-Math.max(a.w,a.h)||_partArea(b)-_partArea(a);
  });
  const placed=[],unplaced=[];
  let y=0,rowH=0,x=0;
  for(const rect of sorted){
    let pw=rect.w,ph=rect.h,rot=false;
    if(!(pw<=SW-x&&ph<=SH-y) && rect.rotate && rect.h<=SW-x&&rect.w<=SH-y){pw=rect.h;ph=rect.w;rot=true}
    if(!(pw<=SW-x&&ph<=SH-y)){
      y+=rowH;x=0;rowH=0;pw=rect.w;ph=rect.h;rot=false;
      if(!(pw<=SW&&ph<=SH-y) && rect.rotate && rect.h<=SW&&rect.w<=SH-y){pw=rect.h;ph=rect.w;rot=true}
    }
    if(pw<=SW-x&&ph<=SH-y){
      placed.push(clonePlaced(rect,x,y,pw,ph,rot));
      x+=pw;rowH=Math.max(rowH,ph);
    }else unplaced.push(rect);
  }
  return {placed,unplaced,freeRects:[]};
}
function packHybrid(rects,SW,SH){
  // 1) Kare/seri parçalar için kontrollü grid-shelf başlangıcı
  let freeRects=[{x:0,y:0,w:SW,h:SH}],placed=[],unplaced=[];
  const groups=new Map();
  const rest=[];
  for(const r of rects){
    if(_isSquareLike(r)){
      const key=r.w+'x'+r.h+'|'+(r.rotate?'r':'n');
      if(!groups.has(key))groups.set(key,[]);
      groups.get(key).push(r);
    }else rest.push(r);
  }
  const groupList=[...groups.values()].sort((a,b)=>_partArea(b[0])-_partArea(a[0]));
  for(const group of groupList){
    let queue=[...group];
    while(queue.length){
      const sample=queue[0];
      let best=null;
      for(let i=0;i<freeRects.length;i++){
        const fr=freeRects[i];
        const orientations=[{w:sample.w,h:sample.h,rot:false}];
        if(sample.rotate&&sample.w!==sample.h)orientations.push({w:sample.h,h:sample.w,rot:true});
        for(const o of orientations){
          const cols=Math.floor(fr.w/o.w),rows=Math.floor(fr.h/o.h),cap=cols*rows;
          if(cap<=0)continue;
          const take=Math.min(cap,queue.length),usedW=Math.min(cols,take)*o.w,usedRows=Math.ceil(take/cols),usedH=usedRows*o.h;
          const score=take*1000000-(fr.w*fr.h-usedW*usedH);
          if(!best||score>best.score)best={i,fr,o,cols,take,usedW,usedH,score};
        }
      }
      if(!best){unplaced.push(...queue);break}
      const batch=queue.splice(0,best.take);
      for(let n=0;n<batch.length;n++){
        const cx=n%best.cols,cy=Math.floor(n/best.cols);
        placed.push(clonePlaced(batch[n],best.fr.x+cx*best.o.w,best.fr.y+cy*best.o.h,best.o.w,best.o.h,best.o.rot));
      }
      freeRects=_splitFreeRectsByPlaced(freeRects,{x:best.fr.x,y:best.fr.y,w:best.usedW,h:best.usedH});
    }
  }
  // 2) Kalan dikdörtgenleri MaxRects boşluk yönetimiyle doldur
  const tail=_placeWithExistingFreeRects([...rest,...unplaced],freeRects);
  placed.push(...tail.placed);
  return {placed,unplaced:tail.unplaced,freeRects:tail.freeRects};
}
function _scorePackResult(res){
  const area=res.placed.reduce((s,r)=>s+r.w*r.h,0);
  return res.placed.length*100000000+area-(res.freeRects?res.freeRects.length*10:0);
}
function packBestHybrid(rects,SW,SH){
  const candidates=[packHybrid(rects,SW,SH),packMaxRects(rects,SW,SH),packShelf(rects,SW,SH),packGuillotine(rects,SW,SH)];
  candidates.sort((a,b)=>_scorePackResult(b)-_scorePackResult(a));
  return candidates[0];
}

/* ═══════════════════════════════════════════
   PACK DISPATCHER
   Düzeltme 1: Symmetric mode artık gerçekten çalışıyor
   Düzeltme 2: Part priority sıralaması iletiliyor
   Düzeltme 3: Grain kısıtı adjRects'te korunuyor
═══════════════════════════════════════════ */
function packAll(rects){
  const{edge,cut}=getMargins();
  const SW=SHEET_W-edge*2, SH=SHEET_H-edge*2;
  const allPanels=[];
  let panelNum=1;

  // Kesim payı ekle, orijinal boyutları sakla, grain kısıtını koru
  const adjRects=rects.map(r=>({
    ...r,
    w:r.w+(cut>0?cut:0),
    h:r.h+(cut>0?cut:0),
    origW:r.w, origH:r.h
    // rotate bayrağı r'den geliyor — grain kısıtı addPart'ta uygulandı ✓
  }));

  // Simetrik mod: çift adetteki parçaları iki yarıya böl
  var _symEl=document.getElementById('symmetricMode');const symMode=(_symEl?_symEl.checked:false)&&currentGrain==='any';

  let remaining=[...adjRects];
  const MAX_PANELS=100;

  while(remaining.length>0&&allPanels.length<MAX_PANELS){
    const prevCount=remaining.length;
    let toPlace=remaining;

    // Simetrik yerleşim: her parçanın yarısını al, paneli simetrik doldur
    if(symMode){
      // Çift adetteki parçaların yarısını bu panele ver
      const half=[];
      const counts={};
      for(const r of toPlace){
        const key=r.id;
        counts[key]=(counts[key]||0)+1;
      }
      const used={};
      for(const r of toPlace){
        const key=r.id;
        used[key]=(used[key]||0)+1;
        const total=counts[key];
        // Tek adetse her zaman ekle, çift adetse sadece ilk yarısını
        if(total<=1||used[key]<=Math.ceil(total/2)) half.push(r);
      }
      toPlace=half;
    }

    let result;
    if(currentAlgo==='hybrid')        result=packBestHybrid(toPlace,SW,SH);
    else if(currentAlgo==='maxrects') result=packMaxRects(toPlace,SW,SH);
    else if(currentAlgo==='shelf')    result=packShelf(toPlace,SW,SH);
    else if(currentAlgo==='ffd')      result=packFFD(toPlace,SW,SH);
    else                              result=packGuillotine(toPlace,SW,SH);

    if(result.placed.length===0)break;

    // Yerleştirilen parçaların id setini bul
    const placedIds=new Set(result.placed.map(r=>r.id+'-'+r.x+'-'+r.y));

    // Simetrik modda: yerleştirilen parçaların kopyasını karşı tarafa ekle
    if(symMode&&result.placed.length>0){
      const mirrored=result.placed.map(r=>({
        ...r,
        x:SW-r.x-r.w, // yatay ayna
        id:r.id+'_sym',
        label:(r.label||r.name)+' (S)'
      }));
      // Çakışma kontrolü — simetrik parçalar mevcut yerleşimlere çarpmıyor mu?
      // cut payı kadar güvenlik marjı bırakarak kontrol et
      const cutBuf=cut>0?cut:0;
      const safe=mirrored.filter(m=>{
        const mb={x:m.x-cutBuf,y:m.y-cutBuf,w:m.w+cutBuf*2,h:m.h+cutBuf*2};
        return !result.placed.some(p=>intersects(p,mb));
      });
      result.placed.push(...safe);
    }

    // Koordinatları gerçek panel koordinatlarına çevir
    const adjustedPlaced=result.placed.map(r=>{
      let rw,rh;
      if(r.origW!==undefined){
        // Döndürülmüşse genişlik/yükseklik yer değiştirir
        rw=r.rotated?r.origH:r.origW;
        rh=r.rotated?r.origW:r.origH;
      } else {
        rw=r.w-(cut>0?cut:0);
        rh=r.h-(cut>0?cut:0);
      }
      return{...r,x:r.x+edge,y:r.y+edge,w:rw,h:rh};
    });

    allPanels.push({panel:panelNum,placed:adjustedPlaced,freeRects:result.freeRects});
    panelNum++;

    // Kalan: yerleştirilemeyen + simetrik modda ikinci yarılar
    if(symMode){
      // Orijinal remaining'den yerleştirilen id'leri çıkar (her biri 1 adet)
      const placedOrigIds=new Map();
      result.placed
        .filter(r=>!String(r.id).endsWith('_sym'))
        .forEach(r=>placedOrigIds.set(r.id,(placedOrigIds.get(r.id)||0)+1));
      const stillRemaining=[];
      const removedCount={};
      for(const r of remaining){
        const removed=removedCount[r.id]||0;
        const toRemove=placedOrigIds.get(r.id)||0;
        if(removed<toRemove){removedCount[r.id]=(removed+1);}
        else{stillRemaining.push(r);}
      }
      remaining=stillRemaining;
    } else {
      remaining=result.unplaced;
    }

    if(remaining.length>=prevCount)break; // sonsuz döngü koruması
  }

  return{panels:allPanels,unplaced:remaining};
}


/* ═══════════════════════════════════════════
   PACKING VALIDATION — canlı kullanım güvenlik kontrolü
   Panel dışına taşma ve parça çakışması kontrol edilir.
═══════════════════════════════════════════ */
function validatePackingResult(panels){
  const issues=[];
  const eps=0.01;
  for(const panel of panels){
    const placed=panel.placed||[];
    for(let i=0;i<placed.length;i++){
      const a=placed[i];
      if(a.x<-eps||a.y<-eps||a.x+a.w>SHEET_W+eps||a.y+a.h>SHEET_H+eps){
        issues.push(`Panel ${panel.panel}: "${esc(a.name||a.label||a.id)}" panel sınırını aşıyor.`);
      }
      for(let j=i+1;j<placed.length;j++){
        const b=placed[j];
        if(intersects(a,b)){
          issues.push(`Panel ${panel.panel}: "${esc(a.name||a.label||a.id)}" ile "${esc(b.name||b.label||b.id)}" çakışıyor.`);
        }
      }
    }
  }
  return issues;
}

/* ═══════════════════════════════════════════
   CALCULATE
═══════════════════════════════════════════ */
function calculate(){
  if(!parts.length){showMsg(t('msgAddParts'),'error');return}
  hideStaleBanner();
  const _t0=performance.now();
  getSheet();
  const rects=[];
  for(const p of parts)for(let i=0;i<p.qty;i++)rects.push({w:p.w,h:p.h,id:p.id,label:p.name+(p.qty>1?` (${i+1}/${p.qty})`:' '),name:p.name,color:p.color,rotate:p.rotate});

  const{panels,unplaced}=packAll(rects);
  allPanelsResult=panels;
  const packingIssues=validatePackingResult(panels);
  if(packingIssues.length){
    console.warn('Panel yerleşim doğrulama uyarıları:', packingIssues);
    setTimeout(()=>showMsg('Yerleşim doğrulama uyarısı: '+packingIssues.slice(0,2).join(' | ')+(packingIssues.length>2?' ...':''),'error'),250);
  }
  currentPanelIdx=0;
  // Layout settle olduktan sonra render — wrap.clientHeight 0 gelmemesi için
  requestAnimationFrame(()=>requestAnimationFrame(()=>renderPanel(0)));

  const total=rects.length;
  const totalPlaced=panels.reduce((s,p)=>s+p.placed.length,0);
  const sa=SHEET_W*SHEET_H;

  // ✅ Düzeltildi: tüm panellerin toplamı üzerinden hesaplama
  const usedTotal=panels.reduce((s,p)=>s+p.placed.reduce((ss,r)=>ss+r.w*r.h,0),0);
  const totalArea=sa*panels.length;
  const upct=panels.length>0?(usedTotal/totalArea*100).toFixed(1):'0.0';
  const wpct=(100-parseFloat(upct)).toFixed(1);
  const wArea=totalArea-usedTotal;

  animStat('s-total',total);
  animStat('s-placed',totalPlaced,totalPlaced===total?'good':'warn');
  animStat('s-usage',upct+'%',parseFloat(upct)>=70?'good':'');
  animStat('s-waste',wpct+'%',parseFloat(wpct)<=30?'good':'warn');
  animStat('s-panels',panels.length,panels.length===1?'good':panels.length<=3?'':'warn');

  // ── VERİM ÇUBUĞU
  updateEfficiencyBar(parseFloat(upct));

  // ── ALGORİTMA KARŞILAŞTIRMA — arka planda çalıştır
  setTimeout(()=>runAlgoComparison(rects,parseFloat(upct),panels.length),100);
  if(panels.length>=10){
    setTimeout(()=>showMsg('Dikkat: '+panels.length+' panel olustu. Parca boyutlarini veya adetleri gozden gecirin.','error'),400);
  }

  const wa=document.getElementById('s-waste-area');
  wa.textContent=`${Number((wArea/100).toFixed(0)).toLocaleString('tr-TR')} cm²`;
  wa.className='stat-val'+(parseFloat(wpct)<=30?' good':' warn');
  document.getElementById('s-placed').closest('.stat-box').style.outline=totalPlaced<total?'2px solid var(--red)':'none';

  const uf=document.getElementById('unfitList');
  if(unplaced.length){
    const{edge,cut}=getMargins();
    const SW=SHEET_W-edge*2, SH=SHEET_H-edge*2;
    const details=unplaced.map(r=>{
      const p=parts.find(x=>x.id===r.id);
      if(!p)return esc(r.label||r.name||'?');
      const fitsNorm=p.w<=SW&&p.h<=SH;
      const fitsRot=p.rotate&&p.h<=SW&&p.w<=SH;
      let reason='';
      if(!fitsNorm&&!fitsRot){
        reason=`<span style="color:var(--orange);font-size:.65rem"> — panel sınırı aşılıyor (${p.w}×${p.h} > ${SW}×${SH})</span>`;
      } else {
        reason=`<span style="color:var(--muted2);font-size:.65rem"> — yeterli alan yok</span>`;
      }
      return `<div style="padding:2px 0">✗ <strong>${esc(p.name)}</strong> (${p.w}×${p.h} mm)${reason}</div>`;
    });
    uf.innerHTML=`<div style="margin-bottom:5px;font-weight:700;color:var(--red)">⚠️ ${unplaced.length} ${t('msgUnplaced').replace('!',':')} </div>${details.join('')}`;
    uf.className='unfit-list visible';
  } else uf.className='unfit-list';

  // Hesap süresi
  const _elapsed=((performance.now()-_t0)/1000).toFixed(2);

  // Maliyet
  updateCost(panels.length);

  // Hesap süresi bilgisini canvasInfo'ya ekle
  setTimeout(()=>{
    const _ci=document.getElementById('canvasInfo');
    if(_ci){const base=_ci._base||_ci.textContent||'';if(base&&!base.includes('·  '))_ci.textContent=base+` · ${_elapsed}s`;}
  },50);

  // Parça listesini gerçek yerleşim verisiyle yeniden render et (birim maliyet güncellenir)
  renderParts();

  // Panel navigation + legend + kesim listesi hemen kur
  buildPanelNav(panels.length);
  buildLegend(panels.flatMap(p=>p.placed),wArea,wpct);
  buildCutList(panels);
  // Toast — ilk hesaplamada kısayol ipucu
  showKbToast();

  // History — maliyet bilgisini de kaydet
  const sqmPriceH  = parseFloat(document.getElementById('sqmPrice').value)||0;
  const cutPriceH  = parseFloat(document.getElementById('cutPrice').value)||0;
  const sheetPriceH= parseFloat(document.getElementById('sheetPrice').value)||0;
  const allPlacedH = panels.flatMap(p=>p.placed);
  const totalPanelAreaH = SHEET_W*SHEET_H*panels.length;
  const totalPanelCostH = sheetPriceH*panels.length;
  let histCost=0;
  allPlacedH.forEach(r=>{
    histCost+=(r.w*r.h/1e6)*sqmPriceH + cutPriceH + (totalPanelAreaH>0?(r.w*r.h/totalPanelAreaH)*totalPanelCostH:0);
  });
  addToHistory({panels:panels.length,placed:totalPlaced,total,usage:upct,algo:currentAlgo,cost:histCost>0?histCost:null,ts:Date.now()});
  autoBackup();

  // Stok kontrolü
  const _stock=parseInt(document.getElementById('sheetStock').value)||0;
  const _stockAlert=document.getElementById('stockAlert');
  if(_stock>0&&panels.length>0){
    if(panels.length>_stock){
      _stockAlert.textContent='⚠️ Dikkat: '+panels.length+' panel gerekiyor, eldeki stok '+_stock+' panel. '+(panels.length-_stock)+' panel eksik!';
      _stockAlert.style.display='block';
      _stockAlert.style.background='var(--red-dim)';
      _stockAlert.style.borderColor='rgba(244,63,94,.35)';
      _stockAlert.style.color='var(--red)';
    } else {
      _stockAlert.textContent=t('stockOk')+' '+panels.length+' '+t('stockUsed')+' '+_stock+' '+t('stockAvail')+' ('+(_stock-panels.length)+' '+t('partUnit')+' '+(currentLang==='en'?'remaining':'artacak')+')';
      _stockAlert.style.display='block';
      _stockAlert.style.background='var(--green-dim)';
      _stockAlert.style.borderColor='rgba(52,211,153,.25)';
      _stockAlert.style.color='var(--green)';
    }
  } else {
    if(_stockAlert)_stockAlert.style.display='none';
  }

  document.getElementById('pdfBtn').disabled=false;
  document.getElementById('pngBtn').disabled=false;
  const _svgb=document.getElementById('svgBtn');if(_svgb)_svgb.disabled=false;
  const pngAllBtn=document.getElementById('pngAllBtn');
  if(panels.length>1){pngAllBtn.style.display='inline-flex';pngAllBtn.disabled=false;pngAllBtn.textContent='PNG ×'+panels.length;}
  else{pngAllBtn.style.display='none';pngAllBtn.disabled=true;}
  const zh=document.getElementById('zoomHint');
  zh.className='zoom-hint visible';
  setTimeout(()=>zh.className='zoom-hint',3500);

  // Mobil: hesap sonrası önce istatistik, sonra canvas'a otomatik scroll
  if(window.innerWidth<=860){
    setTimeout(()=>{
      const target=window.innerWidth<=600
        ? document.getElementById('sec-calc')
        : document.getElementById('canvasWrap');
      target.scrollIntoView({behavior:'smooth',block:'start'});
    },250);
  }
} // calculate() sonu

/* ═══════════════════════════════════════════
   VERİM ÇUBUĞU
═══════════════════════════════════════════ */
function updateEfficiencyBar(pct){
  const wrap=document.getElementById('efficiencyBarWrap');
  const fill=document.getElementById('effBarFill');
  const label=document.getElementById('effBarPct');
  if(!wrap)return;
  wrap.classList.add('visible');
  const cls=pct>=75?'good':pct>=55?'mid':'bad';
  fill.style.width=pct+'%';
  fill.className='efficiency-bar-fill'+(cls==='good'?'':' '+cls);
  label.textContent=pct.toFixed(1)+'%';
  label.className=cls;
  document.getElementById('effBarLabel').textContent=
    currentLang==='en'?'Area efficiency (placed parts / total sheet area)':'Alan verimi (parça / toplam panel alanı)';
}

/* ═══════════════════════════════════════════
   ALGORİTMA KARŞILAŞTIRMA — WebWorker ile
   Ana thread'i bloklamadan arka planda çalışır.
   Worker, packing algoritmalarının saf kopyasını
   Blob URL olarak taşır; DOM erişimi yoktur.
═══════════════════════════════════════════ */
let _algoWorker=null;
let _algoWorkerBusy=false;

function _getAlgoWorkerSrc(){
  // Worker içine gömülecek saf fonksiyonlar (DOM bağımsız)
  return `
function intersects(a,b){return a.x<b.x+b.w&&a.x+a.w>b.x&&a.y<b.y+b.h&&a.y+a.h>b.y}
function pruneFreeRects(rects){
  const n=rects.length,keep=new Uint8Array(n).fill(1);
  for(let i=0;i<n;i++){if(!keep[i])continue;const a=rects[i];for(let j=0;j<n;j++){if(i===j||!keep[j])continue;const b=rects[j];if(b.x<=a.x&&b.y<=a.y&&b.x+b.w>=a.x+a.w&&b.y+b.h>=a.y+a.h){keep[i]=0;break}}}
  return rects.filter((_,i)=>keep[i]);
}
function packMaxRects(rects,SW,SH){
  let freeRects=[{x:0,y:0,w:SW,h:SH}];
  const placed=[],unplaced=[];
  const sorted=[...rects].sort((a,b)=>b.w*b.h-a.w*a.h);
  for(const rect of sorted){
    let bestSS=Infinity,bestArea=Infinity,bestFr=null,bestRot=false;
    for(const fr of freeRects){
      if(rect.w<=fr.w&&rect.h<=fr.h){const ss=Math.min(fr.w-rect.w,fr.h-rect.h),area=fr.w*fr.h-rect.w*rect.h;if(ss<bestSS||(ss===bestSS&&area<bestArea)){bestSS=ss;bestArea=area;bestFr=fr;bestRot=false;}}
      if(rect.rotate&&rect.h<=fr.w&&rect.w<=fr.h){const ss=Math.min(fr.w-rect.h,fr.h-rect.w),area=fr.w*fr.h-rect.h*rect.w;if(ss<bestSS||(ss===bestSS&&area<bestArea)){bestSS=ss;bestArea=area;bestFr=fr;bestRot=true;}}
    }
    if(!bestFr){unplaced.push(rect);continue}
    const pw=bestRot?rect.h:rect.w,ph=bestRot?rect.w:rect.h;
    placed.push({x:bestFr.x,y:bestFr.y,w:pw,h:ph,id:rect.id,rotated:bestRot});
    const newFree=[];
    for(const fr of freeRects){
      if(!intersects(fr,{x:bestFr.x,y:bestFr.y,w:pw,h:ph})){newFree.push(fr);continue}
      if(bestFr.x>fr.x)newFree.push({x:fr.x,y:fr.y,w:bestFr.x-fr.x,h:fr.h});
      if(bestFr.x+pw<fr.x+fr.w)newFree.push({x:bestFr.x+pw,y:fr.y,w:fr.x+fr.w-(bestFr.x+pw),h:fr.h});
      if(bestFr.y>fr.y)newFree.push({x:fr.x,y:fr.y,w:fr.w,h:bestFr.y-fr.y});
      if(bestFr.y+ph<fr.y+fr.h)newFree.push({x:fr.x,y:bestFr.y+ph,w:fr.w,h:fr.y+fr.h-(bestFr.y+ph)});
    }
    freeRects=pruneFreeRects(newFree);
  }
  return{placed,unplaced};
}
function packFFD(rects,SW,SH){
  const sorted=[...rects].sort((a,b)=>b.w*b.h-a.w*a.h);
  const placed=[],unplaced=[];
  let spaces=[{x:0,y:0,w:SW,h:SH}];
  for(const rect of sorted){
    let fit=false;
    for(let i=0;i<spaces.length;i++){
      const sp=spaces[i];let rot=false,pw=rect.w,ph=rect.h;
      if(pw<=sp.w&&ph<=sp.h){}
      else if(rect.rotate&&ph<=sp.w&&pw<=sp.h){rot=true;pw=rect.h;ph=rect.w;}
      else continue;
      placed.push({x:sp.x,y:sp.y,w:pw,h:ph,id:rect.id,rotated:rot});
      const remW=sp.w-pw,remH=sp.h-ph,ns=[];
      if(remW>0&&ph>0)ns.push({x:sp.x+pw,y:sp.y,w:remW,h:ph});
      if(remH>0&&sp.w>0)ns.push({x:sp.x,y:sp.y+ph,w:sp.w,h:remH});
      spaces.splice(i,1,...ns.sort((a,b)=>b.w*b.h-a.w*a.h));
      fit=true;break;
    }
    if(!fit)unplaced.push(rect);
  }
  return{placed,unplaced};
}
function packGuillotine(rects,SW,SH){
  const sorted=[...rects].sort((a,b)=>b.w*b.h-a.w*a.h);
  let spaces=[{x:0,y:0,w:SW,h:SH}];
  const placed=[],unplaced=[];
  for(const rect of sorted){
    let bs=null,bi=-1,rot=false,bestSS=Infinity,bestArea=Infinity;
    for(let i=0;i<spaces.length;i++){
      const sp=spaces[i];
      if(rect.w<=sp.w&&rect.h<=sp.h){const ss=Math.min(sp.w-rect.w,sp.h-rect.h),area=sp.w*sp.h-rect.w*rect.h;if(ss<bestSS||(ss===bestSS&&area<bestArea)){bestSS=ss;bestArea=area;bs=sp;bi=i;rot=false;}}
      if(rect.rotate&&rect.h<=sp.w&&rect.w<=sp.h){const ss=Math.min(sp.w-rect.h,sp.h-rect.w),area=sp.w*sp.h-rect.h*rect.w;if(ss<bestSS||(ss===bestSS&&area<bestArea)){bestSS=ss;bestArea=area;bs=sp;bi=i;rot=true;}}
    }
    if(!bs){unplaced.push(rect);continue}
    const pw=rot?rect.h:rect.w,ph=rot?rect.w:rect.h;
    placed.push({x:bs.x,y:bs.y,w:pw,h:ph,id:rect.id,rotated:rot});
    spaces.splice(bi,1);
    const remR=bs.w-pw,remB=bs.h-ph;
    if(remR<remB){if(remR>0&&ph>0)spaces.push({x:bs.x+pw,y:bs.y,w:remR,h:ph});if(remB>0&&bs.w>0)spaces.push({x:bs.x,y:bs.y+ph,w:bs.w,h:remB});}
    else{if(remB>0&&pw>0)spaces.push({x:bs.x,y:bs.y+ph,w:pw,h:remB});if(remR>0&&bs.h>0)spaces.push({x:bs.x+pw,y:bs.y,w:remR,h:bs.h});}
  }
  return{placed,unplaced};
}

function clonePlaced(rect,x,y,w,h,rotated){return{x,y,w,h,id:rect.id,rotated:!!rotated}}
function area(r){return(r.w||0)*(r.h||0)}
function isSquareLike(r){const mn=Math.min(r.w,r.h),mx=Math.max(r.w,r.h);return mn>0&&mx/mn<=1.08}
function packShelf(rects,SW,SH){
  const sorted=[...rects].sort((a,b)=>(isSquareLike(b)?1:0)-(isSquareLike(a)?1:0)||Math.max(b.w,b.h)-Math.max(a.w,a.h)||area(b)-area(a));
  const placed=[],unplaced=[];let y=0,rowH=0,x=0;
  for(const rect of sorted){let pw=rect.w,ph=rect.h,rot=false;if(!(pw<=SW-x&&ph<=SH-y)&&rect.rotate&&rect.h<=SW-x&&rect.w<=SH-y){pw=rect.h;ph=rect.w;rot=true}if(!(pw<=SW-x&&ph<=SH-y)){y+=rowH;x=0;rowH=0;pw=rect.w;ph=rect.h;rot=false;if(!(pw<=SW&&ph<=SH-y)&&rect.rotate&&rect.h<=SW&&rect.w<=SH-y){pw=rect.h;ph=rect.w;rot=true}}if(pw<=SW-x&&ph<=SH-y){placed.push(clonePlaced(rect,x,y,pw,ph,rot));x+=pw;rowH=Math.max(rowH,ph)}else unplaced.push(rect)}
  return{placed,unplaced};
}
function packHybrid(rects,SW,SH){
  const c=[packMaxRects(rects,SW,SH),packShelf(rects,SW,SH),packGuillotine(rects,SW,SH)];
  c.sort((a,b)=>b.placed.length-a.placed.length||b.placed.reduce((s,r)=>s+r.w*r.h,0)-a.placed.reduce((s,r)=>s+r.w*r.h,0));
  return c[0];
}

function runPack(algo,adj,SW,SH){
  let remaining=[...adj],allP=[],MAX=50;
  while(remaining.length&&allP.length<MAX){
    let res;
    if(algo==='hybrid')res=packHybrid(remaining,SW,SH);
    else if(algo==='maxrects')res=packMaxRects(remaining,SW,SH);
    else if(algo==='shelf')res=packShelf(remaining,SW,SH);
    else if(algo==='ffd')res=packFFD(remaining,SW,SH);
    else res=packGuillotine(remaining,SW,SH);
    if(!res.placed.length)break;
    const pc={};res.placed.forEach(r=>{pc[r.id]=(pc[r.id]||0)+1;});
    const rc={};
    remaining=remaining.filter(r=>{const l=pc[r.id]||0,d=rc[r.id]||0;if(d<l){rc[r.id]=d+1;return false;}return true;});
    allP.push(res);
  }
  return allP;
}
self.onmessage=function(e){
  const{rects,currentAlgo,currentPct,currentPanels,SW,SH,cut,sa}=e.data;
  const algos=[
    {key:'hybrid',name:'Hybrid',sub:'Grid + MaxRects — Kare yoğun'},
    {key:'maxrects',name:'MaxRects',sub:'BSSF — Genel kullanım'},
    {key:'shelf',name:'Shelf',sub:'Satır/Grid — Seri parçalar'},
    {key:'ffd',name:'FFD',sub:'Büyükten küçüğe — Hızlı'},
    {key:'guillotine',name:'Guillotine',sub:'Bölme yöntemi — Kesim odaklı'},
  ];
  const results=[];
  const adj=rects.map(r=>({...r,w:r.w+(cut>0?cut:0),h:r.h+(cut>0?cut:0)}));
  for(const algo of algos){
    if(algo.key===currentAlgo){results.push({...algo,pct:currentPct,panels:currentPanels,current:true});continue;}
    try{
      const allP=runPack(algo.key,adj,SW,SH);
      const placed=allP.flatMap(p=>p.placed);
      const used=placed.reduce((s,r)=>s+r.w*r.h,0);
      const total=sa*Math.max(allP.length,1);
      const pct=allP.length>0?(used/total*100):0;
      results.push({...algo,pct,panels:allP.length,current:false});
    }catch(err){results.push({...algo,pct:0,panels:'?',current:false,err:true});}
  }
  self.postMessage({results});
};`;
}

function _ensureAlgoWorker(){
  if(_algoWorker)return _algoWorker;
  try{
    const blob=new Blob([_getAlgoWorkerSrc()],{type:'application/javascript'});
    const url=URL.createObjectURL(blob);
    _algoWorker=new Worker(url);
    // Blob URL artık gerekli değil — worker oluşturulduktan sonra serbest bırak
    URL.revokeObjectURL(url);
  }catch(e){_algoWorker=null;}
  return _algoWorker;
}

function runAlgoComparison(rects,currentPct,currentPanels){
  const wrap=document.getElementById('algoCompareWrap');
  const grid=document.getElementById('algoCompareGrid');
  if(!wrap||!grid)return;
  wrap.classList.add('visible');

  const{edge,cut}=getMargins();
  const SW=SHEET_W-edge*2, SH=SHEET_H-edge*2;
  const sa=SHEET_W*SHEET_H;

  // Loading göstergesi
  grid.innerHTML='<div style="grid-column:1/-1;text-align:center;color:var(--muted2);font-size:.7rem;padding:10px">⏳ Karşılaştırılıyor…</div>';

  const worker=_ensureAlgoWorker();
  if(!worker||_algoWorkerBusy){
    // Worker yok ya da meşgul — fallback: setTimeout ile ana thread'de çalıştır
    setTimeout(()=>_runAlgoComparisonFallback(rects,currentPct,currentPanels,SW,SH,sa,cut,grid),0);
    return;
  }

  _algoWorkerBusy=true;
  worker.onmessage=function(e){
    _algoWorkerBusy=false;
    _renderAlgoResults(e.data.results,grid);
  };
  worker.onerror=function(){
    _algoWorkerBusy=false;
    _algoWorker=null; // sıfırla, bir dahaki sefere yeniden oluştur
    _runAlgoComparisonFallback(rects,currentPct,currentPanels,SW,SH,sa,cut,grid);
  };
  worker.postMessage({rects,currentAlgo,currentPct,currentPanels,SW,SH,cut,sa});
}

function _runAlgoComparisonFallback(rects,currentPct,currentPanels,SW,SH,sa,cut,grid){
  const algos=[
    {key:'hybrid',name:'Hybrid',sub:'Grid + MaxRects — Kare yoğun'},
    {key:'maxrects',name:'MaxRects',sub:'BSSF — Genel kullanım'},
    {key:'shelf',name:'Shelf',sub:'Satır/Grid — Seri parçalar'},
    {key:'ffd',name:'FFD',sub:'Büyükten küçüğe — Hızlı'},
    {key:'guillotine',name:'Guillotine',sub:'Bölme yöntemi — Kesim odaklı'},
  ];
  const results=[];
  const adj=rects.map(r=>({...r,w:r.w+(cut>0?cut:0),h:r.h+(cut>0?cut:0),origW:r.w,origH:r.h}));
  for(const algo of algos){
    if(algo.key===currentAlgo){results.push({...algo,pct:currentPct,panels:currentPanels,current:true});continue;}
    try{
      let remaining=[...adj],allP=[],MAX=50;
      while(remaining.length&&allP.length<MAX){
        let res;
        if(algo.key==='hybrid')res=packBestHybrid(remaining,SW,SH);
        else if(algo.key==='maxrects')res=packMaxRects(remaining,SW,SH);
        else if(algo.key==='shelf')res=packShelf(remaining,SW,SH);
        else if(algo.key==='ffd')res=packFFD(remaining,SW,SH);
        else res=packGuillotine(remaining,SW,SH);
        if(!res.placed.length)break;
        const pc={};res.placed.forEach(r=>{pc[r.id]=(pc[r.id]||0)+1;});
        const rc={};
        remaining=remaining.filter(r=>{const l=pc[r.id]||0,d=rc[r.id]||0;if(d<l){rc[r.id]=d+1;return false;}return true;});
        allP.push(res);
      }
      const placed=allP.flatMap(p=>p.placed);
      const used=placed.reduce((s,r)=>s+r.w*r.h,0);
      const total=sa*Math.max(allP.length,1);
      const pct=allP.length>0?(used/total*100):0;
      results.push({...algo,pct,panels:allP.length,current:false});
    }catch(e){results.push({...algo,pct:0,panels:'?',current:false,err:true});}
  }
  _renderAlgoResults(results,grid);
}

function _renderAlgoResults(results,grid){
  const best=results.reduce((a,b)=>(!b.err&&b.pct>a.pct)?b:a,results[0]);
  const isEN=currentLang==='en';
  grid.innerHTML=results.map(r=>{
    const cls=r.pct>=75?'good':r.pct>=55?'mid':'bad';
    const isBest=!r.err&&r===best;
    const isActive=r.key===currentAlgo;
    const sub=isEN?{hybrid:'Grid + MaxRects — Square-heavy',maxrects:'BSSF — General purpose',shelf:'Rows/Grid — Batch parts',ffd:'Large→small — Fast',guillotine:'Split method — Cut-plan'}[r.key]:r.sub;
    return`<div class="algo-card${isActive?' active':''}${isBest?' winner':''}" onclick="applyAlgoResult('${r.key}')">
      <div><div class="algo-card-name">${r.name}${isBest?' ★':''}</div><div class="algo-card-sub">${sub}</div></div>
      <div class="algo-card-metric ${r.err?'bad':cls}">${r.err?'—':r.pct.toFixed(1)+'%'}</div>
      <div class="algo-card-panels">${r.err?'?':r.panels+' panel'}</div>
      <button class="algo-card-use${isActive?' current':''}" onclick="event.stopPropagation();applyAlgoResult('${r.key}')">
        ${isActive?(isEN?'Active':'Aktif'):(isEN?'Use':'Kullan')}
      </button>
    </div>`;
  }).join('');
  const titleEl=document.querySelector('.algo-compare-title');
  if(titleEl)titleEl.textContent=isEN?'Algorithm Comparison':'Algoritma Karşılaştırması';
}

function applyAlgoResult(algoKey){
  if(algoKey===currentAlgo)return;
  // Algoritmayı değiştir ve yeniden hesapla
  currentAlgo=algoKey;
  // UI'da da güncelle
  document.querySelectorAll('.algo-opt').forEach(el=>{
    el.classList.toggle('selected',el.dataset.algo===algoKey);
  });
  try{localStorage.setItem('pop9_algo',algoKey)}catch(e){}
  calculate();
}

function showKbToast(){
  if(_kbToastShown||localStorage.getItem('pop9_kbhint'))return;
  _kbToastShown=true;
  const toast=document.getElementById('kbToast');
  if(!toast)return;
  toast.innerHTML=t('kbToast');
  toast.classList.add('show');
  setTimeout(()=>{toast.classList.remove('show');try{localStorage.setItem('pop9_kbhint','1')}catch(e){}},4500);
}


/* ═══════════════════════════════════════════
   MULTI-PANEL NAV
═══════════════════════════════════════════ */
function buildPanelNav(count){
  const nav=document.getElementById('panelNav');
  if(count<=1){nav.style.display='none';return}
  nav.style.display='flex';
  nav.innerHTML=Array.from({length:count},(_,i)=>`
    <button class="panel-tab-btn ${i===0?'active':''}" onclick="switchPanel(${i})">Panel ${i+1}</button>
  `).join('');
}
function switchPanel(idx){
  currentPanelIdx=idx;
  document.querySelectorAll('.panel-tab-btn').forEach((b,i)=>{b.className='panel-tab-btn'+(i===idx?' active':'')});
  renderPanel(idx);
}
function renderPanel(idx){
  const panel=allPanelsResult[idx];
  if(!panel)return;
  const waste=getFreeAsWaste(panel.freeRects||[]);
  const canvas=document.getElementById('mainCanvas');
  const wrap=document.getElementById('canvasWrap');
  const isMobile=window.innerWidth<=600;

  // Güvenilir boyut hesabı — clientHeight 0 gelebilir, alternatif kaynaklardan al
  const aw=Math.max(wrap.clientWidth||300, 200) - 4;
  let ah=wrap.clientHeight;
  if(ah<100){
    // Layout henüz settle olmadı — viewport'tan hesapla
    var _hEl=document.querySelector('.header');const headerH=(_hEl?_hEl.offsetHeight:0)||67;
    var _chEl=document.querySelector('.canvas-header');const canvasHeaderH=(_chEl?_chEl.offsetHeight:0)||80;
    var _lgEl=document.getElementById('legend');const legendH=(_lgEl?_lgEl.offsetHeight:0)||36;
    var _scEl=document.getElementById('scaleBarWrap');const scaleH=(_scEl?_scEl.offsetHeight:0)||0;
    const padding=isMobile?20:32;
    ah=Math.max(window.innerHeight-headerH-canvasHeaderH-legendH-scaleH-padding, isMobile?280:320);
    wrap.style.minHeight=ah+'px';
  }
  ah=Math.max(ah,isMobile?280:320)-4;

  // Cetvel boyutları — fitScale hesabından önce çıkarılmalı
  

  // fitScale: cetrul alanı çıkarıldıktan sonra kalan alana sığacak ölçek
  fitScale=Math.min((aw-RULER_W)/SHEET_W, (ah-RULER_H)/SHEET_H);
  fitScale=Math.max(fitScale, 0.05);

  // Canvas boyutu = panel alanı + cetrul alanı
  const cw=Math.round(SHEET_W*fitScale)+RULER_W;
  const ch=Math.round(SHEET_H*fitScale)+RULER_H;
  canvas.width=cw*DR; canvas.height=ch*DR;
  canvas.style.width=cw+'px'; canvas.style.height=ch+'px';
  canvas.style.display='block';
  document.getElementById('placeholder').style.display='none';
  const ctx=canvas.getContext('2d');
  // FIX: setTransform yerine scale — her renderPanel'de temiz başlangıç
  ctx.setTransform(DR,0,0,DR,0,0);
  draw(ctx,cw,ch,fitScale,panel.placed,waste,false,RULER_W,RULER_H);
  zoomLevel=1; applyZoom(false);
  requestAnimationFrame(()=>{
    const scr=document.getElementById('canvasScroll');
    scr.scrollLeft=0; scr.scrollTop=0;
  });
  drawScaleBar(cw,fitScale);
  const{edge}=getMargins();
  const infoBase=`Panel ${idx+1}/${allPanelsResult.length} · ${panel.placed.length} ${currentLang==='en'?'part(s)':'parça'} · ${currentAlgo.toUpperCase()}${edge?` · ${currentLang==='en'?'Edge':'Kenar'}:${edge}mm`:''}`;
  document.getElementById('canvasInfo').textContent=infoBase;
  document.getElementById('canvasInfo')._base=infoBase;
}
function getFreeAsWaste(freeRects){
  const{edge}=getMargins();
  return freeRects
    .filter(r=>r.w>20&&r.h>20) // çok küçük fire alanlarını çizme
    .map(r=>({
      x: r.x + edge, // placed ile aynı koordinat sistemine çek
      y: r.y + edge,
      w: r.w,
      h: r.h
    }));
}

/* ═══════════════════════════════════════════
   DRAW CANVAS
═══════════════════════════════════════════ */
function draw(ctx,cw,ch,sc,placed,waste,isExport,rOX,rOY){
  // rOX/rOY: cetrul alanı offset (px) — parça ve grid bu kadar sağa/aşağı kayar
  // Export modunda cetrul yok, offset 0
  if(isExport){rOX=0;rOY=0;}
  rOX=rOX||0; rOY=rOY||0;
  // Çizim alanı boyutları (cetrul hariç)
  const pw=cw-rOX, ph=ch-rOY;

  // Reset state
  ctx.globalAlpha=1;
  ctx.shadowColor='transparent';ctx.shadowBlur=0;ctx.shadowOffsetX=0;ctx.shadowOffsetY=0;
  const zs=isExport?1:(zoomLevel||1);
  const mmPerPx=1/sc;

  // ── ARKA PLAN — tüm canvas (cetrul dahil)
  const bgBase='#13151a';
  ctx.fillStyle=bgBase;ctx.fillRect(0,0,cw,ch);

  // Subtle radial glow — sadece içerik alanında
  const bgGlow=ctx.createRadialGradient(rOX+pw*.5,rOY+ph*.35,0,rOX+pw*.5,rOY+ph*.35,Math.max(pw,ph)*.65);
  bgGlow.addColorStop(0,'rgba(129,140,248,.05)');
  bgGlow.addColorStop(.6,'rgba(129,140,248,.012)');
  bgGlow.addColorStop(1,'rgba(0,0,0,0)');
  ctx.fillStyle=bgGlow;ctx.fillRect(rOX,rOY,pw,ph);

  // ── İçerik alanına clip + translate — cetrul alanına taşmaz
  ctx.save();
  ctx.beginPath();ctx.rect(rOX,rOY,pw,ph);ctx.clip();
  ctx.translate(rOX,rOY);
  // ── GRID — içerik alanında (translate ile zaten offset'li)
  ctx.save();
  const gs=Math.round(100*sc);
  if(gs>4){
    ctx.strokeStyle='rgba(255,255,255,.022)';ctx.lineWidth=.5;
    for(let x=gs;x<pw;x+=gs){ctx.beginPath();ctx.moveTo(x,0);ctx.lineTo(x,ph);ctx.stroke()}
    for(let y=gs;y<ph;y+=gs){ctx.beginPath();ctx.moveTo(0,y);ctx.lineTo(pw,y);ctx.stroke()}
  }
  const gb=Math.round(500*sc);
  if(gb>4){
    ctx.strokeStyle='rgba(129,140,248,.10)';ctx.lineWidth=.8;
    for(let x=gb;x<pw;x+=gb){ctx.beginPath();ctx.moveTo(x,0);ctx.lineTo(x,ph);ctx.stroke()}
    for(let y=gb;y<ph;y+=gb){ctx.beginPath();ctx.moveTo(0,y);ctx.lineTo(pw,y);ctx.stroke()}
  }
  ctx.restore();

  // ── EDGE MARGIN
  const{edge,cut}=getMargins();
  if(edge>0){
    const ep=Math.round(edge*sc);
    ctx.fillStyle='rgba(129,140,248,.035)';
    ctx.fillRect(0,0,pw,ep);ctx.fillRect(0,ph-ep,pw,ep);
    ctx.fillRect(0,ep,ep,ph-ep*2);ctx.fillRect(pw-ep,ep,ep,ph-ep*2);
    ctx.save();ctx.strokeStyle='rgba(129,140,248,.45)';ctx.lineWidth=.9;
    ctx.setLineDash([5,5]);ctx.strokeRect(ep+.5,ep+.5,pw-ep*2-1,ph-ep*2-1);ctx.restore();
    if(ep>14){
      const efz=Math.round(Math.max(7,Math.min(11,9/zs)));
      ctx.save();ctx.font=`500 ${efz}px 'DM Mono',monospace`;
      ctx.fillStyle='rgba(129,140,248,.5)';
      ctx.textAlign='left';ctx.textBaseline='top';ctx.fillText(`↤ ${edge}mm`,4,4);ctx.restore();
    }
  }

  // ── WASTE (fire) — amber-kırmızı, belirgin ama zarif
  const hatchPat=makeHatch(ctx,'rgba(244,63,94,.28)',9);
  for(const ws of waste){
    if(ws.w<=0||ws.h<=0)continue;
    const wx=Math.round(ws.x*sc),wy=Math.round(ws.y*sc);
    const ww=Math.round(ws.w*sc),wh=Math.round(ws.h*sc);
    // Zemin dolgusu — gradient: amber → kırmızı
    ctx.globalAlpha=1;
    const wg=ctx.createLinearGradient(wx,wy,wx+ww,wy+wh);
    wg.addColorStop(0,'rgba(245,158,11,.08)');
    wg.addColorStop(1,'rgba(244,63,94,.06)');
    ctx.fillStyle=wg;ctx.fillRect(wx,wy,ww,wh);
    ctx.globalAlpha=.35;ctx.fillStyle=hatchPat;ctx.fillRect(wx,wy,ww,wh);
    ctx.globalAlpha=1;
    ctx.save();ctx.strokeStyle='rgba(244,63,94,.3)';ctx.lineWidth=.8;
    ctx.setLineDash([4,4]);ctx.strokeRect(wx+.5,wy+.5,ww-1,wh-1);ctx.restore();
    if(ww>24&&wh>18){
      const rawFs=Math.min(ww*.14,wh*.22,16);
      const fs=Math.max(7,Math.min(rawFs,rawFs/zs*1.4));
      ctx.save();
      ctx.textAlign='center';ctx.textBaseline='middle';
      ctx.font=`700 ${fs}px 'DM Mono',monospace`;
      ctx.fillStyle='rgba(0,0,0,.5)';
      ctx.fillText(t('fireLabel'),wx+ww/2+.7,wy+wh/2+.7+(wh>fs*2.5?-fs*.6:0));
      ctx.fillStyle='rgba(244,100,120,.85)';
      ctx.fillText(t('fireLabel'),wx+ww/2,wy+wh/2+(wh>fs*2.5?-fs*.6:0));
      if(wh>fs*2.5&&ww>32){
        const fd=Math.max(6,fs*.75);
        ctx.font=`400 ${fd}px 'DM Mono',monospace`;
        ctx.fillStyle='rgba(0,0,0,.4)';
        ctx.fillText(`${ws.w}×${ws.h}`,wx+ww/2+.6,wy+wh/2+fs*.72+.6);
        ctx.fillStyle='rgba(244,160,160,.55)';
        ctx.fillText(`${ws.w}×${ws.h}`,wx+ww/2,wy+wh/2+fs*.72);
      }
      ctx.restore();
    }
  }

  // ── PANEL SINIRI (translate sonrası — 0,0 = içerik sol üstü)
  for(let i=4;i>=1;i--){
    ctx.strokeStyle=`rgba(129,140,248,${0.035*i})`;
    ctx.lineWidth=i*4;ctx.strokeRect(0,0,pw,ph);
  }
  ctx.strokeStyle='rgba(129,140,248,.55)';ctx.lineWidth=1.5;ctx.strokeRect(.75,.75,pw-1.5,ph-1.5);
  const tk=Math.min(22,pw*.06,ph*.06);
  ctx.strokeStyle='rgba(129,140,248,.85)';ctx.lineWidth=2;ctx.lineCap='round';
  [[0,0,tk,0,0,tk],[pw,0,-tk,0,0,tk],[0,ph,tk,0,0,-tk],[pw,ph,-tk,0,0,-tk]].forEach(([x,y,dx1,dy1,dx2,dy2])=>{
    ctx.beginPath();ctx.moveTo(x,y);ctx.lineTo(x+dx1,y+dy1);
    ctx.moveTo(x,y);ctx.lineTo(x+dx2,y+dy2);ctx.stroke();
  });
  ctx.fillStyle='rgba(129,140,248,.9)';
  [[2,2],[pw-2,2],[2,ph-2],[pw-2,ph-2]].forEach(([x,y])=>{
    ctx.beginPath();ctx.arc(x,y,2,0,Math.PI*2);ctx.fill();
  });

  // ── YERLEŞTİRİLEN PARÇALAR
  for(const r of placed){
    const rx=Math.floor(r.x*sc),ry=Math.floor(r.y*sc);
    const rw=Math.floor((r.x+r.w)*sc)-rx,rh=Math.floor((r.y+r.h)*sc)-ry;
    if(rw<1||rh<1)continue;
    const pad=Math.max(1,Math.min(2,rw*.015,rh*.015));
    const rad=Math.min(5,rw*.06,rh*.06);

    // Gölge
    ctx.save();
    ctx.shadowColor='rgba(0,0,0,.55)';ctx.shadowBlur=6;ctx.shadowOffsetX=1;ctx.shadowOffsetY=2;

    // Dolgu: renk üst koyuca, alta şeffaf
    const rg=ctx.createLinearGradient(rx,ry,rx,ry+rh);
    rg.addColorStop(0,r.color+'e8');
    rg.addColorStop(.45,r.color+'c0');
    rg.addColorStop(1,r.color+'60');
    ctx.globalAlpha=.92;ctx.fillStyle=rg;
    ctx.beginPath();
    roundRect(ctx,rx+pad,ry+pad,rw-pad*2,rh-pad*2,rad);
    ctx.fill();
    ctx.restore(); // shadow + globalAlpha ikisi birden temizlenir

    // FIX: restore sonrası globalAlpha kesinlikle 1
    ctx.globalAlpha=1;

    // Cam parlaklık — sol+üst köşeden yayılan
    ctx.globalAlpha=1;
    if(rh>10&&rw>10){
      const shineH=Math.min(rh*.28,14);
      const shine=ctx.createLinearGradient(rx,ry,rx,ry+shineH);
      shine.addColorStop(0,'rgba(255,255,255,.22)');
      shine.addColorStop(1,'rgba(255,255,255,0)');
      ctx.fillStyle=shine;
      ctx.save();ctx.beginPath();
      roundRect(ctx,rx+pad,ry+pad,rw-pad*2,shineH,{tl:rad,tr:rad,bl:0,br:0});
      ctx.fill();ctx.restore();
    }

    // Dış çerçeve — parça rengiyle, ince glow
    ctx.strokeStyle=r.color;ctx.lineWidth=1;
    ctx.save();ctx.shadowColor=r.color;ctx.shadowBlur=4;
    ctx.beginPath();
    roundRect(ctx,rx+pad-.3,ry+pad-.3,rw-pad*2+.6,rh-pad*2+.6,rad);ctx.stroke();
    ctx.restore();

    // İç çerçeve çok ince beyaz — derinlik
    const ipad=pad+1.5;
    ctx.strokeStyle='rgba(255,255,255,.10)';ctx.lineWidth=.7;
    ctx.beginPath();
    roundRect(ctx,rx+ipad,ry+ipad,rw-ipad*2,rh-ipad*2,Math.max(0,rad-1));ctx.stroke();

    // Döndürme ikonu
    if(r.rotated&&rw>18&&rh>18){
      const riFs=Math.max(7,Math.min(12,rh*.16/zs));
      ctx.save();
      ctx.font=`600 ${riFs}px sans-serif`;
      ctx.fillStyle='rgba(255,255,255,.55)';
      ctx.textAlign='right';ctx.textBaseline='top';
      ctx.fillText('↻',rx+rw-4,ry+4);
      ctx.restore();
    }

    // ── ETİKET
    if(rw>18&&rh>9){
      ctx.save();
      // FIX: clip için önce beginPath, sonra roundRect path'i oluştur, sonra clip
      ctx.beginPath();
      roundRect(ctx,rx+pad,ry+pad,rw-pad*2,rh-pad*2,rad);
      ctx.clip();

      const wMm=r.w, hMm=r.h;
      const targetFsMm=Math.min(wMm*.07, hMm*.12, 18);
      const baseFs=Math.max(isExport?8:7, targetFsMm*sc);
      const fs=isExport?Math.min(baseFs,rh*.22):Math.max(7/zs, Math.min(baseFs, baseFs/zs*1.2));
      const screenFs=isExport?fs:(fs*zs);
      if(screenFs<6){ctx.restore();continue;}

      ctx.textAlign='center';ctx.textBaseline='middle';
      const cx=rx+rw/2, cy=ry+rh/2;
      const hasDim=rh>fs*2.8&&rw>30;
      const labelY=hasDim?cy-fs*.55:cy;

      const shadowOff=isExport?1.5:(0.8/zs);
      ctx.font=`700 ${fs}px 'DM Mono',monospace`;
      // Koyu metin gölgesi
      ctx.fillStyle='rgba(0,0,0,.75)';
      ctx.fillText(r.label,cx+shadowOff,labelY+shadowOff);
      // Beyaz metin
      ctx.fillStyle='#ffffff';
      ctx.fillText(r.label,cx,labelY);

      if(hasDim){
        const fd=isExport?Math.max(7,fs*.72):Math.max(5.5/zs, fs*.72);
        if(isExport?(fd>=7):(fd*zs>=5)){
          ctx.font=`400 ${fd}px 'DM Mono',monospace`;
          ctx.fillStyle='rgba(0,0,0,.55)';
          ctx.fillText(`${r.w}×${r.h}`,cx+shadowOff*.75,cy+fs*.65+shadowOff*.75);
          ctx.fillStyle='rgba(255,255,255,.60)';
          ctx.fillText(`${r.w}×${r.h}`,cx,cy+fs*.65);
        }
      }
      ctx.restore();
    }
  }

  // Boyut watermark — içerik alanı ortasında (translate aktifken)
  const wmFs=Math.max(9,Math.min(18,pw*.024));
  ctx.font=`700 ${wmFs}px 'DM Mono',monospace`;
  ctx.textAlign='center';ctx.textBaseline='middle';
  ctx.fillStyle='rgba(129,140,248,.032)';
  ctx.fillText(`${SHEET_W} × ${SHEET_H} mm`,pw/2,ph/2);

  // ── translate/clip bitti — mutlak koordinatlara dön
  ctx.restore();

  // ── CETVEL — translate SONRASI, mutlak koordinatlarla çizilir
  if(!isExport){
    ctx.setLineDash([]);
    const rulerBg='rgba(17,18,20,.97)';
    const rulerW=rOX, rulerH=rOY;

    // Arka planlar
    ctx.fillStyle=rulerBg;
    ctx.fillRect(0,0,rulerW,ch);          // sol cetrul
    ctx.fillRect(rulerW,ch-rulerH,cw-rulerW,rulerH); // alt cetrul
    ctx.fillRect(0,ch-rulerH,rulerW,rulerH);          // köşe

    // Ayırıcı çizgiler — kalın ve net
    ctx.strokeStyle='rgba(129,140,248,.35)';ctx.lineWidth=1.5;
    ctx.beginPath();ctx.moveTo(rulerW-.75,0);ctx.lineTo(rulerW-.75,ch);ctx.stroke();
    ctx.beginPath();ctx.moveTo(0,ch-rulerH+.75);ctx.lineTo(cw,ch-rulerH+.75);ctx.stroke();

    const rlFs=Math.max(6,Math.min(9,8/zs));
    const rlFsBold=Math.max(7,Math.min(10,9/zs));
    const tickStep=sc*zs>0.4?100:200;

    // Yatay cetvel — rulerW'den başlar, alt kenarda
    ctx.font=`500 ${rlFs}px 'DM Mono',monospace`;
    for(let mm=0;mm<=SHEET_W;mm+=100){
      const px=Math.round(mm*sc)+rulerW;
      if(px>=cw)break;
      const maj=mm%tickStep===0;
      ctx.strokeStyle=maj?'rgba(129,140,248,.6)':'rgba(255,255,255,.12)';
      ctx.lineWidth=.8;ctx.setLineDash([]);
      ctx.beginPath();ctx.moveTo(px,ch-rulerH);ctx.lineTo(px,ch-(maj?rulerH:rulerH*.38));ctx.stroke();
      if(maj&&mm>0){
        ctx.fillStyle='rgba(129,140,248,.75)';ctx.textAlign='center';ctx.textBaseline='bottom';
        ctx.fillText(mm,px,ch-3);
      }
    }
    ctx.fillStyle='rgba(129,140,248,.9)';ctx.textAlign='right';ctx.textBaseline='bottom';
    ctx.font=`600 ${rlFsBold}px 'DM Mono',monospace`;
    ctx.fillText(SHEET_W+'mm',cw-4,ch-3);

    // Dikey cetvel — rOY'dan başlar, sol kenarda
    ctx.font=`500 ${rlFs}px 'DM Mono',monospace`;
    for(let mm=0;mm<=SHEET_H;mm+=100){
      const py=Math.round(mm*sc)+rOY;
      if(py>=ch-rulerH)break;
      const maj=mm%tickStep===0;
      ctx.strokeStyle=maj?'rgba(129,140,248,.6)':'rgba(255,255,255,.12)';
      ctx.lineWidth=.8;ctx.setLineDash([]);
      ctx.beginPath();ctx.moveTo(0,py);ctx.lineTo(maj?rulerW*.72:rulerW*.32,py);ctx.stroke();
      if(maj&&mm>0){
        ctx.fillStyle='rgba(129,140,248,.75)';ctx.textAlign='right';ctx.textBaseline='middle';
        ctx.fillText(mm,rulerW-3,py);
      }
    }
    ctx.fillStyle='rgba(129,140,248,.9)';ctx.textAlign='left';ctx.textBaseline='top';
    ctx.font=`600 ${rlFsBold}px 'DM Mono',monospace`;
    ctx.fillText(SHEET_H+'mm',3,3);
  }
}

/* Yardımcı: yuvarlak köşe rect path — beginPath ÇAĞIRMAZ, caller sorumlu */
function roundRect(ctx,x,y,w,h,r){
  if(typeof r==='number')r={tl:r,tr:r,bl:r,br:r};
  const{tl=0,tr=0,bl=0,br=0}=r;
  // Negatif/sıfır boyut için düz rect
  if(w<1||h<1){ctx.rect(x,y,Math.max(w,0),Math.max(h,0));return;}
  // Köşe yarıçapı boyutun yarısını geçemez
  const mxr=Math.min(w/2,h/2);
  const rtl=Math.min(tl,mxr),rtr=Math.min(tr,mxr),rbl=Math.min(bl,mxr),rbr=Math.min(br,mxr);
  ctx.moveTo(x+rtl,y);
  ctx.lineTo(x+w-rtr,y);ctx.arcTo(x+w,y,x+w,y+rtr,rtr);
  ctx.lineTo(x+w,y+h-rbr);ctx.arcTo(x+w,y+h,x+w-rbr,y+h,rbr);
  ctx.lineTo(x+rbl,y+h);ctx.arcTo(x,y+h,x,y+h-rbl,rbl);
  ctx.lineTo(x,y+rtl);ctx.arcTo(x,y,x+rtl,y,rtl);
  ctx.closePath();
}

/* Yardımcı: tarama deseni — DPR-aware, cache ile performans */
function makeHatch(ctx,color,size){
  const key=color+'_'+size;
  if(_hatchCache[key])return ctx.createPattern(_hatchCache[key],'repeat');
  // FIX: pattern canvas'ını fiziksel piksel olarak oluştur, CSS px olarak göster
  // Böylece hem ekranda hem export'ta keskin görünür
  const o=document.createElement('canvas');
  o.width=size;o.height=size;
  const c=o.getContext('2d');
  c.strokeStyle=color;c.lineWidth=1;
  c.beginPath();c.moveTo(0,size);c.lineTo(size,0);c.stroke();
  c.beginPath();c.moveTo(-size/2,size/2);c.lineTo(size/2,-size/2);c.stroke();
  _hatchCache[key]=o;
  return ctx.createPattern(o,'repeat');
}

/* ═══════════════════════════════════════════
   SCALE BAR
═══════════════════════════════════════════ */
function drawScaleBar(cw,sc){
  const sb=document.getElementById('scaleBar'),bw=Math.min(cw,280);
  sb.width=bw*DR;sb.height=22*DR;sb.style.width=bw+'px';sb.style.height='22px';sb.style.display='block';
  const c=sb.getContext('2d');c.scale(DR,DR);c.clearRect(0,0,bw,22);
  c.setLineDash([]); // FIX: önceki draw'dan kalan dashed line sızmasını önle
  const mm=Math.pow(10,Math.floor(Math.log10(SHEET_W*.3)));
  const px=Math.round(mm*sc),bx=10,by=11;
  c.strokeStyle='rgba(129,140,248,.65)';c.lineWidth=1.5;c.lineCap='round';
  c.beginPath();c.moveTo(bx,by);c.lineTo(bx+px,by);c.moveTo(bx,by-5);c.lineTo(bx,by+5);c.moveTo(bx+px,by-5);c.lineTo(bx+px,by+5);c.stroke();
  c.fillStyle='rgba(129,140,248,.60)';c.font=`500 9px 'DM Mono',monospace`;c.textAlign='center';c.textBaseline='top';c.fillText(mm+' mm',bx+px/2,by+6);
  document.getElementById('scaleBarWrap').className='scale-bar-wrap visible';
}

/* ═══════════════════════════════════════════
   LEGEND
═══════════════════════════════════════════ */
function buildLegend(placed,wArea,wpct){
  const el=document.getElementById('legend');
  const seen=new Set();
  let h=parts.map(p=>{
    if(seen.has(p.id))return'';seen.add(p.id);
    const pc=placed.filter(r=>r.id===p.id).length;
    return`<div class="legend-item"><div class="legend-color" style="background:${p.color}"></div><span>${esc(p.name)} — ${p.w}×${p.h} mm — ${pc}/${p.qty}</span></div>`;
  }).join('');
  h+=`<div class="legend-item" style="margin-left:auto;border-color:rgba(244,63,94,.25)"><div class="legend-color" style="background:repeating-linear-gradient(45deg,rgba(244,63,94,.5),rgba(244,63,94,.5) 3px,transparent 3px,transparent 8px)"></div><span style="color:var(--red)">Fire — ${wpct}%</span></div>`;
  el.innerHTML=h;
}

/* ═══════════════════════════════════════════
   CUTTING LIST
═══════════════════════════════════════════ */
function getSortedCutList(panels){
  const items=[];
  for(const panel of panels){
    for(const r of panel.placed){
      items.push({...r,panelNum:panel.panel});
    }
  }
  if(currentCutOrder==='lr-tb'){
    // Sol→Sağ, Yukarı→Aşağı (önce panel, sonra Y, sonra X)
    items.sort((a,b)=>a.panelNum-b.panelNum||a.y-b.y||a.x-b.x);
  } else if(currentCutOrder==='size'){
    // Büyükten küçüğe
    items.sort((a,b)=>(b.w*b.h)-(a.w*a.h));
  }
  // 'placement' sırası: zaten panels→placed sırasına göre
  return items;
}
function buildCutList(panels){
  const body=document.getElementById('cutListBody');
  const items=getSortedCutList(panels);
  var _gbpEl=document.getElementById('cutGroupByPanel');const groupByPanel=(_gbpEl?_gbpEl.checked:false)&&panels.length>1;

  /* ── Maliyet parametreleri (önbellekten) ── */
  if(!_cachedPrices.valid)refreshPriceCache();
  const sqmPrice=_cachedPrices.sqm, cutPrice=_cachedPrices.cut, sheetPrice=_cachedPrices.sheet;
  const hasCost=sqmPrice>0||cutPrice>0||sheetPrice>0;

  const totalPanelArea = SHEET_W * SHEET_H * panels.length;
  const totalPanelCost = sheetPrice * panels.length;
  const sym=getCurrencySymbol();
  const fmt=v=>v.toLocaleString(currentLang==='en'?'en-GB':'tr-TR',{minimumFractionDigits:2,maximumFractionDigits:2});

  function calcCutItemCost(r){
    if(!hasCost) return null;
    const areaM2=(r.w*r.h)/1e6;
    return areaM2*sqmPrice + cutPrice + (totalPanelArea>0?(r.w*r.h/totalPanelArea)*totalPanelCost:0);
  }

  let grandTotal=0, rows='', idx=0;

  if(groupByPanel){
    // Panel bazlı gruplandırma
    for(const panel of panels){
      const panelItems=items.filter(r=>r.panelNum===panel.panel);
      if(!panelItems.length)continue;
      let panelTotal=0;
      rows+=`<div style="padding:5px 10px;background:var(--surface2);border-top:1px solid var(--border2);border-bottom:1px solid var(--border2);font-family:var(--mono);font-size:.64rem;font-weight:700;color:var(--gold);display:flex;align-items:center;gap:8px">
        <span>📋 Panel ${panel.panel}</span>
        <span style="font-weight:400;color:var(--muted2)">${panelItems.length} kesim</span>
      </div>`;
      for(const r of panelItems){
        idx++;
        const unitCost=calcCutItemCost(r);
        if(unitCost!==null){grandTotal+=unitCost;panelTotal+=unitCost;}
        const panelColor=idx%2===0?'rgba(255,255,255,.012)':'transparent';
        const costHtml=hasCost?(unitCost>0?`<span class="cut-cost">${sym}${fmt(unitCost)}</span>`:`<span class="cut-cost no-cost">—</span>`):`<span class="cut-cost no-cost"></span>`;
        rows+=`<div class="cut-list-row" style="background:${panelColor}">
          <span class="cut-idx">${idx}</span>
          <span class="cut-name"><span style="display:inline-block;width:7px;height:7px;border-radius:2px;background:${r.color};margin-right:5px;vertical-align:middle"></span>${esc(r.name||r.label)}${(()=>{const p=parts.find(x=>x.id===r.id);return p&&p.note?` <span style="color:var(--muted2);font-size:.6rem;font-weight:400">${esc(p.note)}</span>`:''})()}</span>
          <span class="cut-dim">${r.w}×${r.h}</span>
          <span class="cut-pos">${r.x},${r.y}</span>
          <span class="cut-panel" style="color:var(--gold)">P${r.panelNum}</span>
          ${costHtml}
        </div>`;
      }
      if(hasCost&&panelTotal>0){
        rows+=`<div style="padding:4px 10px;text-align:right;font-family:var(--mono);font-size:.62rem;color:var(--muted2);border-bottom:1px solid var(--border)">Panel ${panel.panel} toplamı: <strong style="color:var(--text2)">${sym}${fmt(panelTotal)}</strong></div>`;
      }
    }
  } else {
    items.forEach((r,i)=>{
      const panelColor=(i+1)%2===0?'rgba(255,255,255,.012)':'transparent';
      const unitCost=calcCutItemCost(r);
      if(unitCost!==null) grandTotal+=unitCost;
      const costHtml=hasCost?(unitCost>0?`<span class="cut-cost">${sym}${fmt(unitCost)}</span>`:`<span class="cut-cost no-cost">—</span>`):`<span class="cut-cost no-cost"></span>`;
      rows+=`<div class="cut-list-row" style="background:${panelColor}">
        <span class="cut-idx">${i+1}</span>
        <span class="cut-name"><span style="display:inline-block;width:7px;height:7px;border-radius:2px;background:${r.color};margin-right:5px;vertical-align:middle"></span>${esc(r.name||r.label)}${(()=>{const p=parts.find(x=>x.id===r.id);return p&&p.note?` <span style="color:var(--muted2);font-size:.6rem;font-weight:400">${esc(p.note)}</span>`:''})()}</span>
        <span class="cut-dim">${r.w}×${r.h}</span>
        <span class="cut-pos">${r.x},${r.y}</span>
        <span class="cut-panel" style="color:var(--gold)">P${r.panelNum}</span>
        ${costHtml}
      </div>`;
    });
  }

  // Toplam footer
  const footerHtml=hasCost&&grandTotal>0?`
    <div class="cut-list-footer">
      <span class="cut-list-footer-label">
        ${items.length} parça · Fire dahil toplam maliyet
        ${sqmPrice>0&&sheetPrice>0?'<span style="font-size:.56rem;color:var(--muted);margin-left:4px">(panel payi fire maliyetini icerir)</span>':''}
      </span>
      <span class="cut-list-footer-val">${sym}${fmt(grandTotal)}</span>
    </div>`:'';

  body.innerHTML=(rows||`<div style="padding:12px;color:var(--muted);font-size:.7rem;text-align:center">${t('noResult')}</div>`)+footerHtml;
  // Maliyet sütununu göster/gizle
  const cutHead=document.getElementById('cutListHead');
  const cutCostH=document.getElementById('cutListCostHeader');
  if(cutHead){
    const cols=hasCost?'30px 1fr 70px 80px 50px 80px':'30px 1fr 70px 80px 50px';
    cutHead.style.gridTemplateColumns=cols;
    document.querySelectorAll('.cut-list-row').forEach(r=>r.style.gridTemplateColumns=cols);
    if(cutCostH)cutCostH.style.display=hasCost?'':'none';
  }
  document.getElementById('cutListWrap').style.display='block';

  // 10'dan fazla satır varsa "Tümünü Göster" butonu göster
  const expandBtn=document.getElementById('cutListExpandBtn');
  if(items.length>10){
    expandBtn.style.display='block';
    expandBtn.textContent=`${t('showAllLabel')} (${items.length} ${t('partUnit')})`;
    body.classList.remove('expanded');
  } else {
    expandBtn.style.display='none';
    body.classList.add('expanded');
  }
}
let _cutListExpanded=false;
function toggleCutListExpand(){
  _cutListExpanded=!_cutListExpanded;
  const body=document.getElementById('cutListBody');
  const btn=document.getElementById('cutListExpandBtn');
  if(_cutListExpanded){
    body.classList.add('expanded');
    btn.textContent='▲ Daralt';
  } else {
    body.classList.remove('expanded');
    const total=allPanelsResult.reduce((s,p)=>s+p.placed.length,0);
    btn.textContent=`${t('showAllLabel')} (${total} ${t('partUnit')})`;
  }
}
function exportCutListCSV(){
  if(!allPanelsResult.length){showMsg(t('msgCalcFirst'),'error');return}
  const sqmPrice  =parseFloat(document.getElementById('sqmPrice').value)||0;
  const cutPrice  =parseFloat(document.getElementById('cutPrice').value)||0;
  const sheetPrice=parseFloat(document.getElementById('sheetPrice').value)||0;
  const hasCost   =sqmPrice>0||cutPrice>0||sheetPrice>0;
  const totalPanelArea=SHEET_W*SHEET_H*allPanelsResult.length;
  const totalPanelCost=sheetPrice*allPanelsResult.length;

  const _csvSym=getCurrencySymbol();
  const header=hasCost
    ?`#,Ad,Genişlik,Yükseklik,X,Y,Panel,Not,Etiket,RenkKod,KenarBant,BirimMaliyet(${_csvSym})\n`
    :'#,Ad,Genişlik,Yükseklik,X,Y,Panel,Not,Etiket,RenkKod,KenarBant\n';

  let csv=header;
  const items=getSortedCutList(allPanelsResult);
  items.forEach((r,i)=>{
    const _np=parts.find(x=>x.id===r.id);
    const note=(_np&&_np.note?_np.note:'').replace(/"/g,'""');
    let line=`${i+1},"${(r.name||r.label||'').replace(/"/g,'""')}",${r.w},${r.h},${r.x},${r.y},${r.panelNum},"${note}","${((_np&&_np.labelCode)||'').replace(/"/g,'""')}","${((_np&&_np.materialCode)||'').replace(/"/g,'""')}","${((_np&&_np.edgeBand)||'').replace(/"/g,'""')}"`;
    if(hasCost){
      const areaM2=(r.w*r.h)/1e6;
      const cost=areaM2*sqmPrice+cutPrice+(totalPanelArea>0?(r.w*r.h/totalPanelArea)*totalPanelCost:0);
      line+=`,${cost.toFixed(2)}`;
    }
    csv+=line+'\n';
  });
  mobileDownload('data:text/csv;charset=utf-8,'+encodeURIComponent(csv), `kesim-listesi-${Date.now()}.csv`);
}

/* ═══════════════════════════════════════════
   EXCEL / XLSX EXPORT
═══════════════════════════════════════════ */
function exportCutListXLSX(){
  if(!allPanelsResult.length){showMsg(t('msgCalcFirst'),'error');return}
  if(!window.XLSX){showMsg(t('msgExcelLoading'),'info');return}

  const sym=getCurrencySymbol();
  if(!_cachedPrices.valid)refreshPriceCache();
  const sqmPrice=_cachedPrices.sqm, cutPrice=_cachedPrices.cut, sheetPrice=_cachedPrices.sheet;
  const hasCost=sqmPrice>0||cutPrice>0||sheetPrice>0;
  const totalPanelArea=SHEET_W*SHEET_H*allPanelsResult.length;
  const totalPanelCost=sheetPrice*allPanelsResult.length;

  const items=getSortedCutList(allPanelsResult);

  // Başlık satırı
  const headers=['#','Parça Adı','Etiket','Renk/Kod','Kenar Bant','Not','Genişlik (mm)','Yükseklik (mm)','X (mm)','Y (mm)','Panel'];
  if(hasCost)headers.push('Birim Maliyet ('+sym+')');

  const rows=[headers];

  items.forEach((r,i)=>{
    const prt=parts.find(x=>x.id===r.id);
    const row=[
      i+1,
      r.name||r.label||'',
      (prt&&prt.labelCode)||'',
      (prt&&prt.materialCode)||'',
      (prt&&prt.edgeBand)||'',
      (prt&&prt.note)||'',
      r.w, r.h, r.x, r.y,
      'Panel '+r.panelNum
    ];
    if(hasCost){
      const areaM2=(r.w*r.h)/1e6;
      const cost=areaM2*sqmPrice+cutPrice+(totalPanelArea>0?(r.w*r.h/totalPanelArea)*totalPanelCost:0);
      row.push(Math.round(cost*100)/100);
    }
    rows.push(row);
  });

  // Özet satırları
  rows.push([]);
  rows.push(['ÖZET']);
  rows.push(['Panel Boyutu',SHEET_W+'×'+SHEET_H+' mm']);
  rows.push(['Malzeme',currentMaterial||(document.getElementById('materialType').value||'—')]);
  rows.push(['Kalınlık',(currentThickness||18)+' mm']);
  rows.push(['Algoritma',currentAlgo.toUpperCase()]);
  rows.push(['Toplam Parça',items.length]);
  rows.push(['Panel Sayısı',allPanelsResult.length]);
  if(hasCost){
    const grandTotal=items.reduce((s,r)=>{
      const areaM2=(r.w*r.h)/1e6;
      return s+areaM2*sqmPrice+cutPrice+(totalPanelArea>0?(r.w*r.h/totalPanelArea)*totalPanelCost:0);
    },0);
    rows.push(['Toplam Maliyet (KDV Hariç)',sym+' '+grandTotal.toFixed(2)]);
    if(document.getElementById('kdvEnabled').checked){
      const rate=(parseInt(document.getElementById('kdvRate').value)||20)/100;
      rows.push(['KDV (%'+Math.round(rate*100)+')',sym+' '+(grandTotal*rate).toFixed(2)]);
      rows.push(['KDV Dahil Toplam',sym+' '+(grandTotal*(1+rate)).toFixed(2)]);
    }
  }
  rows.push(['Oluşturulma',new Date().toLocaleString('tr-TR')]);
  if(document.getElementById('customerName')&&document.getElementById('customerName').value){
    rows.push(['Müşteri',document.getElementById('customerName').value]);
  }

  const ws=XLSX.utils.aoa_to_sheet(rows);
  // Sütun genişlikleri
  ws['!cols']=[{wch:5},{wch:25},{wch:20},{wch:14},{wch:15},{wch:10},{wch:10},{wch:10},{wch:18}];

  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,'Kesim Listesi',ws);
  const fname='kesim-listesi-'+new Date().toISOString().slice(0,10)+'.xlsx';
  mobileDownloadXLSX(wb, fname);
  showMsg('Excel indirildi: '+fname,'success');
}
/* ═══════════════════════════════════════════
   DXF EXPORT
═══════════════════════════════════════════ */
function exportCutListDXF(){
  if(!allPanelsResult.length){showMsg(t('msgCalcFirst'),'error');return}
  // Basit DXF: her parça için LWPOLYLINE (kapalı dikdörtgen), her panel ayrı layer
  let dxf='0\nSECTION\n2\nHEADER\n9\n$ACADVER\n1\nAC1015\n0\nENDSEC\n';
  dxf+='0\nSECTION\n2\nENTITIES\n';
  for(const panel of allPanelsResult){
    for(const r of panel.placed){
      const x1=r.x,y1=r.y,x2=r.x+r.w,y2=r.y+r.h;
      // 66\n1 = vertices follow flag (CAM yazılımları için zorunlu)
      dxf+=`0\nLWPOLYLINE\n8\nPanel_${panel.panel}\n66\n1\n70\n1\n90\n4\n`;
      dxf+=`10\n${x1}\n20\n${y1}\n`;
      dxf+=`10\n${x2}\n20\n${y1}\n`;
      dxf+=`10\n${x2}\n20\n${y2}\n`;
      dxf+=`10\n${x1}\n20\n${y2}\n`;
      dxf+=`0\nTEXT\n8\nPanel_${panel.panel}_Labels\n10\n${x1+r.w/2}\n20\n${y1+r.h/2}\n30\n0\n40\n${Math.min(r.w,r.h)*0.1}\n1\n${(r.name||'').replace(/\n/g,' ')}\n`;
    }
    // Panel sınırı
    dxf+=`0\nLWPOLYLINE\n8\nPanel_${panel.panel}_Border\n66\n1\n70\n1\n90\n4\n`;
    dxf+=`10\n0\n20\n0\n10\n${SHEET_W}\n20\n0\n10\n${SHEET_W}\n20\n${SHEET_H}\n10\n0\n20\n${SHEET_H}\n`;
  }
  dxf+='0\nENDSEC\n0\nEOF\n';
  mobileDownload('data:application/dxf;charset=utf-8,'+encodeURIComponent(dxf), `panel-opt-${Date.now()}.dxf`);
  showMsg('DXF indirildi.','success');
}

/* ═══════════════════════════════════════════
   COST
═══════════════════════════════════════════ */
let costBreakdownVisible=false;

function updateCost(panelCount){
  refreshPriceCache();
  const sheetPrice=_cachedPrices.sheet, sqmPrice=_cachedPrices.sqm, cutPrice=_cachedPrices.cut;
  const cr = document.getElementById('costResult');
  if(!cr) return;

  if(!panelCount && !allPanelsResult.length){ cr.className='cost-result'; return; }
  const pc = panelCount || allPanelsResult.length;
  if(!pc){ cr.className='cost-result'; return; }

  const allPlaced       = allPanelsResult.flatMap(p => p.placed);
  const totalPlacedCount= allPlaced.length;
  const totalPanelArea  = SHEET_W * SHEET_H * pc;   // mm² — fire DAHİL
  const totalPanelCost  = sheetPrice * pc;
  const fmt = v => v.toLocaleString(currentLang==='en'?'en-GB':'tr-TR',{minimumFractionDigits:2,maximumFractionDigits:2});
  const sym = getCurrencySymbol();

  /* ── 1. PANEL MALİYETİ ── */
  const panelCost = totalPanelCost;

  /* ── 2. PARÇA m² MALİYETİ ── */
  const usedArea    = allPlaced.reduce((s,r)=>s+r.w*r.h, 0); // mm²
  const partSqmCost = (usedArea/1e6) * sqmPrice;

  /* ── 3. KESİM MALİYETİ ── */
  const cuttingCost = cutPrice * totalPlacedCount;

  /* ── 4. FIRE MALİYETİ (bilgi — panel payı zaten fire dahil) ── */
  const wasteAreaM2 = (totalPanelArea - usedArea) / 1e6;
  const wasteCost   = wasteAreaM2 * sqmPrice;

  /* ── GENEL TOPLAM ── */
  const grandTotal = panelCost + partSqmCost + cuttingCost;

  // Panel satırı
  var _el_cr_panel_cost=document.getElementById('cr-panel-cost');if(_el_cr_panel_cost)_el_cr_panel_cost.textContent =
    sheetPrice > 0 ? `${pc} panel × ${sym}${fmt(sheetPrice)} = ${sym}${fmt(panelCost)}` : '—';

  // Parça m² satırı
  var _el_cr_part_cost=document.getElementById('cr-part-cost');if(_el_cr_part_cost)_el_cr_part_cost.textContent =
    sqmPrice > 0 ? `${(usedArea/1e6).toFixed(4)} m² × ${sym}${fmt(sqmPrice)} = ${sym}${fmt(partSqmCost)}` : '—';

  // Kesim satırı
  var _el_cr_cut_cost=document.getElementById('cr-cut-cost');if(_el_cr_cut_cost)_el_cr_cut_cost.textContent =
    cutPrice > 0 ? `${totalPlacedCount} kesim × ${sym}${fmt(cutPrice)} = ${sym}${fmt(cuttingCost)}` : '—';

  // Fire (bilgi)
  var _el_cr_waste_cost=document.getElementById('cr-waste-cost');if(_el_cr_waste_cost)_el_cr_waste_cost.textContent =
    sqmPrice > 0 ? `${wasteAreaM2.toFixed(4)} m² × ${sym}${fmt(sqmPrice)} = ${sym}${fmt(wasteCost)}` : '—';

  // Toplam
  var _el_cr_total=document.getElementById('cr-total');if(_el_cr_total)_el_cr_total.textContent = grandTotal > 0 ? `${sym}${fmt(grandTotal)}` : '—';

  // KDV
  var _kdvEl=document.getElementById('kdvEnabled');const kdvEnabled=_kdvEl?_kdvEl.checked:false;
  var _kdvRateEl=document.getElementById('kdvRate');const kdvRate=(parseInt(_kdvRateEl?_kdvRateEl.value:'20')||20)/100;
  var kdvRow=document.getElementById('cr-kdv-row');
  var kdvTotalRow=document.getElementById('cr-total-kdv-row');
  if(kdvEnabled&&grandTotal>0){
    const kdvAmount=grandTotal*kdvRate;
    const grandWithKdv=grandTotal+kdvAmount;
    var _kdvrl=document.getElementById('cr-kdv-rate-lbl');if(_kdvrl)_kdvrl.textContent=Math.round(kdvRate*100);
    var _kdvc=document.getElementById('cr-kdv');if(_kdvc)_kdvc.textContent=`${sym}${fmt(kdvAmount)}`;
    var _kdvtc=document.getElementById('cr-total-kdv');if(_kdvtc)_kdvtc.textContent=`${sym}${fmt(grandWithKdv)}`;
    if(kdvRow)kdvRow.style.display='';if(kdvTotalRow)kdvTotalRow.style.display='';
  } else {
    if(kdvRow)kdvRow.style.display='none';if(kdvTotalRow)kdvTotalRow.style.display='none';
  }

  /* ── PARÇA DETAY BREAKDOWN (fire dahil panel payı ile tutarlı) ── */
  var bd = document.getElementById('cr-parts-breakdown');
  const hasData = (sqmPrice>0||sheetPrice>0||cutPrice>0) && allPlaced.length>0;
  var _ctb=document.getElementById('cr-toggle-btn');if(_ctb)_ctb.style.display=hasData?'block':'none';

  if(hasData){
    if(bd)bd.innerHTML = parts.map(p => {
      const placedQty = allPlaced.filter(r=>r.id===p.id).length;
      if(!placedQty) return '';
      const areaM2    = (p.w*p.h)/1e6;
      const sqmPart   = areaM2 * sqmPrice;
      const panelPart = totalPanelArea>0 ? (p.w*p.h/totalPanelArea)*totalPanelCost : 0;
      const unitCost  = sqmPart + panelPart + cutPrice;
      const totalCost = unitCost * placedQty;
      return `
        <div style="display:flex;justify-content:space-between;align-items:flex-start;padding:5px 0;font-size:.63rem;border-bottom:1px solid rgba(255,255,255,.04)">
          <span style="color:var(--muted2);display:flex;align-items:center;gap:4px;flex:1">
            <span style="display:inline-block;width:6px;height:6px;border-radius:2px;background:${p.color};flex-shrink:0"></span>
            <span>${esc(p.name)} <span style="opacity:.55">${p.w}×${p.h}mm</span> × ${placedQty} adet</span>
          </span>
          <span style="color:var(--gold);font-family:var(--mono);white-space:nowrap;margin-left:8px">
            ${sym}${fmt(unitCost)}/adet → ${sym}${fmt(totalCost)}
          </span>
        </div>`;
    }).join('');
  }

  cr.className = 'cost-result visible';
}

function toggleCostBreakdown(){
  costBreakdownVisible = !costBreakdownVisible;
  var _cbr=document.getElementById('cr-parts-breakdown-row');if(_cbr)_cbr.style.display=costBreakdownVisible?'flex':'none';
  var _el_cr_toggle_btn=document.getElementById('cr-toggle-btn');if(_el_cr_toggle_btn)_el_cr_toggle_btn.textContent =
    (costBreakdownVisible ? '▼' : '▶') + ' Parça detayı ' + (costBreakdownVisible ? 'gizle' : 'göster');
}

/* ═══════════════════════════════════════════
   STORAGE KULLANIM GÖSTERGESİ
═══════════════════════════════════════════ */
function updateStorageInfo(){
  try{
    const el=document.getElementById('storageInfo');
    if(!el)return;
    // Tüm pop9_ key'lerinin toplam boyutunu hesapla
    let total=0;
    for(let i=0;i<localStorage.length;i++){
      const k=localStorage.key(i);
      if(k&&k.startsWith('pop9_'))total+=((localStorage.getItem(k)||'').length*2);
    }
    const kb=(total/1024).toFixed(1);
    const pct=Math.min(100,(total/(5*1024*1024)*100)).toFixed(0);
    const projects=JSON.parse(localStorage.getItem('pop9_projects')||'[]').length;
    el.innerHTML=`<span>${projects} proje · ${kb}KB</span><span style="color:${pct>80?'var(--red)':pct>50?'var(--orange)':'var(--muted)'}">${pct}% dolu</span>`;
  }catch(e){}
}
/* ═══════════════════════════════════════════
   PROJECT MANAGER
═══════════════════════════════════════════ */
function getProjects(){try{return JSON.parse(localStorage.getItem('pop9_projects')||'[]')}catch(e){return[]}}
function setProjects(p){localStorage.setItem('pop9_projects',JSON.stringify(p))}

function saveProject(){
  const name=document.getElementById('projectName').value.trim()||`Proje ${new Date().toLocaleDateString(currentLang==='en'?'en-GB':'tr-TR')}`;
  const projects=getProjects();
  const proj={
    id:Date.now(),name,
    parts:JSON.parse(JSON.stringify(parts)),
    colorIdx,
    sheetW:parseInt(document.getElementById('sheetW').value)||1830,
    sheetH:parseInt(document.getElementById('sheetH').value)||3660,
    algo:currentAlgo,
    grain:currentGrain,
    cutOrder:currentCutOrder,
    marginEdge:parseInt(document.getElementById('marginEdge').value)||0,
    marginCut:parseInt(document.getElementById('marginCut').value)||0,
    price:parseFloat(document.getElementById('sheetPrice').value)||0,
    sqmPrice:parseFloat(document.getElementById('sqmPrice').value)||0,
    cutPrice:parseFloat(document.getElementById('cutPrice').value)||0,
    material:currentMaterial,
    thickness:currentThickness,
    currency:currentCurrency,
    customer:(document.getElementById('customerName')||{value:''}).value.trim(),
    orderNo:(document.getElementById('orderNo')||{value:''}).value||'',
    delivery:(document.getElementById('projectDelivery')||{value:''}).value.trim(),
    projectNote:(document.getElementById('projectNote')||{value:''}).value.trim(),
    operator:(document.getElementById('operatorName')||{value:''}).value.trim(),
    workshopNote:(document.getElementById('workshopNote')||{value:''}).value.trim(),
    kdvEnabled:document.getElementById('kdvEnabled').checked,
    kdvRate:parseInt(document.getElementById('kdvRate').value)||20,
    symmetric:document.getElementById('symmetricMode').checked,
    stock:parseInt(document.getElementById('sheetStock').value)||0,
    ts:Date.now()
  };
  projects.unshift(proj);
  if(projects.length>20)projects.pop();
  setProjects(projects);
  renderProjectList();
  showMsg('"'+esc(name)+'" kaydedildi.','success');
  document.getElementById('projectName').value='';
}

function loadProject(id){
  const proj=getProjects().find(p=>p.id===id);
  if(!proj)return;
  parts=proj.parts;colorIdx=proj.colorIdx||parts.length;
  document.getElementById('sheetW').value=proj.sheetW;
  document.getElementById('sheetH').value=proj.sheetH;
  document.getElementById('marginEdge').value=proj.marginEdge||0;
  document.getElementById('marginCut').value=proj.marginCut||0;
  document.getElementById('sheetPrice').value=proj.price||0;
  document.getElementById('sqmPrice').value=proj.sqmPrice||0;
  document.getElementById('cutPrice').value=proj.cutPrice||0;
  SHEET_W=proj.sheetW;SHEET_H=proj.sheetH;
  setAlgo(proj.algo||'maxrects');
  if(proj.grain)setGrain(proj.grain);
  if(proj.cutOrder)setCutOrder(proj.cutOrder);
  if(proj.material){currentMaterial=proj.material;const ms=document.getElementById('materialType');if(ms)ms.value=proj.material;}
  if(proj.thickness){currentThickness=proj.thickness;const ts=document.getElementById('sheetThickness');if(ts)ts.value=proj.thickness;}
  const cn=document.getElementById('customerName');if(cn)cn.value=proj.customer||'';
  const on=document.getElementById('orderNo');if(on)on.value=proj.orderNo||'';
  [['projectDelivery','delivery'],['projectNote','projectNote'],['operatorName','operator'],['workshopNote','workshopNote']].forEach(([id,key])=>{const el=document.getElementById(id);if(el)el.value=proj[key]||'';});
  if(proj.currency)setCurrency(proj.currency);
  if(proj.kdvEnabled!==undefined){const ke=document.getElementById('kdvEnabled');if(ke)ke.checked=proj.kdvEnabled;}
  if(proj.kdvRate){const kr=document.getElementById('kdvRate');if(kr)kr.value=proj.kdvRate;}
  if(proj.symmetric!==undefined){const sm=document.getElementById('symmetricMode');if(sm)sm.checked=proj.symmetric;}
  if(proj.stock!==undefined){const sk=document.getElementById('sheetStock');if(sk)sk.value=proj.stock||'';}
  _cachedPrices.valid=false;
  _rpCache={};
  renderParts();updateSubtitle();save();
  showMsg('"'+esc(proj.name)+'"'+t('msgLoaded'),'success');
}

function deleteProject(id){
  const projects=getProjects().filter(p=>p.id!==id);
  setProjects(projects);renderProjectList();
}

function renderProjectList(){
  const el=document.getElementById('projectList');
  const projects=getProjects();
  if(!projects.length){el.innerHTML=`<div class="project-empty">${t('emptyProjects')}</div>`;return}
  updateStorageInfo();
  el.innerHTML=projects.map(p=>`
    <div class="project-item" onclick="loadProject(${p.id})">
      <div class="project-item-name">${esc(p.name)}</div>
      <div class="project-item-meta">${p.parts.length} ${currentLang==='en'?'part(s)':'parça'} · ${new Date(p.ts).toLocaleDateString(currentLang==='en'?'en-GB':'tr-TR')}</div>
      <button class="project-item-del" onclick="event.stopPropagation();deleteProject(${p.id})">✕</button>
    </div>`).join('');
}

function exportProject(){
  const proj={
    name:document.getElementById('projectName').value.trim()||'Panel Optimizasyonu Projesi',
    parts:JSON.parse(JSON.stringify(parts)),colorIdx,
    sheetW:parseInt(document.getElementById('sheetW').value)||1830,
    sheetH:parseInt(document.getElementById('sheetH').value)||3660,
    algo:currentAlgo,
    grain:currentGrain,
    cutOrder:currentCutOrder,
    marginEdge:parseInt(document.getElementById('marginEdge').value)||0,
    marginCut:parseInt(document.getElementById('marginCut').value)||0,
    price:parseFloat(document.getElementById('sheetPrice').value)||0,
    sqmPrice:parseFloat(document.getElementById('sqmPrice').value)||0,
    cutPrice:parseFloat(document.getElementById('cutPrice').value)||0,
    material:currentMaterial,
    thickness:currentThickness,
    currency:currentCurrency,
    customer:(document.getElementById('customerName')||{value:''}).value.trim(),
    orderNo:(document.getElementById('orderNo')||{value:''}).value.trim(),
    delivery:(document.getElementById('projectDelivery')||{value:''}).value.trim(),
    projectNote:(document.getElementById('projectNote')||{value:''}).value.trim(),
    operator:(document.getElementById('operatorName')||{value:''}).value.trim(),
    workshopNote:(document.getElementById('workshopNote')||{value:''}).value.trim(),
    kdvEnabled:document.getElementById('kdvEnabled').checked,
    kdvRate:parseInt(document.getElementById('kdvRate').value)||20,
    symmetric:document.getElementById('symmetricMode').checked,
    stock:parseInt(document.getElementById('sheetStock').value)||0,
    exportedAt:new Date().toISOString()
  };
  mobileDownload('data:application/json;charset=utf-8,'+encodeURIComponent(JSON.stringify(proj,null,2)), `panel-project-${Date.now()}.json`);
}
function importProject(input){
  const f=input.files[0];if(!f)return;
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const proj=JSON.parse(e.target.result);
      parts=proj.parts||[];colorIdx=proj.colorIdx||parts.length;
      document.getElementById('sheetW').value=proj.sheetW||1830;
      document.getElementById('sheetH').value=proj.sheetH||3660;
      document.getElementById('marginEdge').value=proj.marginEdge||0;
      document.getElementById('marginCut').value=proj.marginCut||0;
      document.getElementById('sheetPrice').value=proj.price||0;
      document.getElementById('sqmPrice').value=proj.sqmPrice||0;
      document.getElementById('cutPrice').value=proj.cutPrice||0;
      SHEET_W=proj.sheetW||1830;SHEET_H=proj.sheetH||3660;
      setAlgo(proj.algo||'maxrects');
      if(proj.grain)setGrain(proj.grain);
      if(proj.cutOrder)setCutOrder(proj.cutOrder);
      if(proj.material){currentMaterial=proj.material;const ms=document.getElementById('materialType');if(ms)ms.value=proj.material;}
      if(proj.thickness){currentThickness=proj.thickness;const ts=document.getElementById('sheetThickness');if(ts)ts.value=proj.thickness;}
      if(proj.currency)setCurrency(proj.currency);
      const _cn=document.getElementById('customerName');if(_cn&&proj.customer!==undefined)_cn.value=proj.customer||'';
      const _on=document.getElementById('orderNo');if(_on&&proj.orderNo!==undefined)_on.value=proj.orderNo||'';
      [['projectDelivery','delivery'],['projectNote','projectNote'],['operatorName','operator'],['workshopNote','workshopNote']].forEach(([id,key])=>{const el=document.getElementById(id);if(el&&proj[key]!==undefined)el.value=proj[key]||'';});
      if(proj.kdvEnabled!==undefined){const ke=document.getElementById('kdvEnabled');if(ke)ke.checked=proj.kdvEnabled;}
      if(proj.kdvRate){const kr=document.getElementById('kdvRate');if(kr)kr.value=proj.kdvRate;}
      if(proj.symmetric!==undefined){const sm=document.getElementById('symmetricMode');if(sm)sm.checked=proj.symmetric;}
      if(proj.stock!==undefined){const sk=document.getElementById('sheetStock');if(sk)sk.value=proj.stock||'';}
      _cachedPrices.valid=false;
      renderParts();updateSubtitle();save();
      showMsg(esc(proj.name)+t('msgLoaded'),'success');
    }catch(err){showMsg(t('msgInvalidJSON'),'error')}
  };
  reader.readAsText(f);
  input.value='';
}

/* ═══════════════════════════════════════════
   HISTORY
═══════════════════════════════════════════ */
function addToHistory(entry){
  calcHistory.unshift(entry);
  if(calcHistory.length>15)calcHistory.pop();
  try{localStorage.setItem('pop9_history',JSON.stringify(calcHistory))}catch(e){}
  renderHistory();
}
function renderHistory(){
  // ✅ Düzeltildi: localStorage'dan okuma sadece load() sırasında yapılır
  // Bu fonksiyon sadece mevcut calcHistory state'ini render eder
  const el=document.getElementById('historyList');
  if(!calcHistory.length){el.innerHTML=`<div class="project-empty">${t('emptyHistory')}</div>`;return}
  el.innerHTML=calcHistory.map((h,i)=>`
    <div class="history-item">
      <div class="history-item-info">
        <div class="history-item-title">${h.placed}/${h.total} ${currentLang==='en'?'part(s)':'parça'} · ${h.panels} ${currentLang==='en'?'panel(s)':'panel'} · ${esc(String(h.algo||'').toUpperCase())}</div>
        <div class="history-item-meta">
          ${new Date(h.ts).toLocaleString(currentLang==='en'?'en-GB':'tr-TR',{hour:'2-digit',minute:'2-digit',day:'2-digit',month:'2-digit'})}
          ${h.cost>0?`<span style="color:var(--gold);font-family:var(--mono);margin-left:6px">₺${h.cost.toLocaleString(currentLang==='en'?'en-GB':'tr-TR',{minimumFractionDigits:2,maximumFractionDigits:2})}</span>`:''}
        </div>
        <div class="history-bar" style="width:${h.usage}%"></div>
      </div>
      <span class="history-badge">${h.usage}%</span>
    </div>`).join('');
}
function clearHistory(){
  calcHistory=[];
  localStorage.removeItem('pop9_history');
  renderHistory();
}

/* ═══════════════════════════════════════════
   SHARE URL
═══════════════════════════════════════════ */
function generateShareURL(){
  const state={
    parts,colorIdx,
    sheetW:parseInt(document.getElementById('sheetW').value)||1830,
    sheetH:parseInt(document.getElementById('sheetH').value)||3660,
    algo:currentAlgo,
    grain:currentGrain,
    cutOrder:currentCutOrder,
    marginEdge:parseInt(document.getElementById('marginEdge').value)||0,
    marginCut:parseInt(document.getElementById('marginCut').value)||0,
    price:parseFloat(document.getElementById('sheetPrice').value)||0,
    sqmPrice:parseFloat(document.getElementById('sqmPrice').value)||0,
    cutPrice:parseFloat(document.getElementById('cutPrice').value)||0,
    currency:currentCurrency,
    material:currentMaterial,
    thickness:currentThickness,
    customer:(document.getElementById('customerName')||{value:''}).value.trim(),
    orderNo:(document.getElementById('orderNo')||{value:''}).value.trim(),
    delivery:(document.getElementById('projectDelivery')||{value:''}).value.trim(),
    projectNote:(document.getElementById('projectNote')||{value:''}).value.trim(),
    operator:(document.getElementById('operatorName')||{value:''}).value.trim(),
    workshopNote:(document.getElementById('workshopNote')||{value:''}).value.trim(),
    kdvEnabled:document.getElementById('kdvEnabled').checked,
    kdvRate:parseInt(document.getElementById('kdvRate').value)||20,
    symmetric:document.getElementById('symmetricMode').checked,
    stock:parseInt(document.getElementById('sheetStock').value)||0
  };
  const encoded=btoa(unescape(encodeURIComponent(JSON.stringify(state))));
  const url=window.location.href.split('?')[0]+'?state='+encoded;
  document.getElementById('shareUrlInput').value=url;
  document.getElementById('shareUrlBox').style.display='flex';
}
function copyShareURL(){
  const inp=document.getElementById('shareUrlInput');
  if(navigator.clipboard&&navigator.clipboard.writeText){
    navigator.clipboard.writeText(inp.value).then(()=>showMsg(t('msgLinkCopied'),'success')).catch(()=>{
      inp.select();document.execCommand('copy');showMsg(t('msgLinkCopied'),'success');
    });
  } else {
    inp.select();document.execCommand('copy');showMsg(t('msgLinkCopied'),'success');
  }
}
function loadShareFromURL(){
  const params=new URLSearchParams(window.location.search);
  const stateParam=params.get('state');
  if(!stateParam)return;
  try{
    const state=JSON.parse(decodeURIComponent(escape(atob(stateParam))));
    parts=state.parts||[];colorIdx=state.colorIdx||0;
    document.getElementById('sheetW').value=state.sheetW||1830;
    document.getElementById('sheetH').value=state.sheetH||3660;
    document.getElementById('marginEdge').value=state.marginEdge||0;
    document.getElementById('marginCut').value=state.marginCut||0;
    document.getElementById('sheetPrice').value=state.price||0;
    document.getElementById('sqmPrice').value=state.sqmPrice||0;
    document.getElementById('cutPrice').value=state.cutPrice||0;
    SHEET_W=state.sheetW||1830;SHEET_H=state.sheetH||3660;
    setAlgo(state.algo||'maxrects');
    if(state.grain)setGrain(state.grain);
    if(state.cutOrder)setCutOrder(state.cutOrder);
    if(state.currency)setCurrency(state.currency);
    if(state.material){currentMaterial=state.material;const ms=document.getElementById('materialType');if(ms)ms.value=state.material;}
    if(state.thickness){currentThickness=state.thickness;const ts=document.getElementById('sheetThickness');if(ts)ts.value=state.thickness;}
    const _scn=document.getElementById('customerName');if(_scn&&state.customer)_scn.value=state.customer;
    const _son=document.getElementById('orderNo');if(_son&&state.orderNo)_son.value=state.orderNo;
    if(state.kdvEnabled!==undefined){const ke=document.getElementById('kdvEnabled');if(ke)ke.checked=state.kdvEnabled;}
    if(state.kdvRate){const kr=document.getElementById('kdvRate');if(kr)kr.value=state.kdvRate;}
    if(state.symmetric!==undefined){const sm=document.getElementById('symmetricMode');if(sm)sm.checked=state.symmetric;}
    if(state.stock!==undefined){const sk=document.getElementById('sheetStock');if(sk)sk.value=state.stock||'';}
    _cachedPrices.valid=false;
    renderParts();updateSubtitle();
    showMsg(t('msgSharedLoaded'),'success');
    window.history.replaceState({},'',window.location.pathname);
  }catch(e){showMsg(t('msgInvalidLink'),'error')}
}

/* ═══════════════════════════════════════════
   ZOOM
═══════════════════════════════════════════ */
function applyZoom(anim=true){
  const inner=document.getElementById('canvasInner');
  if(!anim){inner.style.transition='none';setTimeout(()=>inner.style.transition='',20)}
  inner.style.transform=`scale(${zoomLevel})`;inner.style.transformOrigin='top left';
  document.getElementById('zoomLevelLbl').textContent=Math.round(zoomLevel*100)+'%';
  // Canvas'ı yeniden çiz — zoom'a göre font boyutları güncellenir
  if(allPanelsResult.length){
    const panel=allPanelsResult[currentPanelIdx];
    if(!panel)return;
    const canvas=document.getElementById('mainCanvas');
    const ctx=canvas.getContext('2d');
    const cw=canvas.width/DR, ch=canvas.height/DR;
    const waste=getFreeAsWaste(panel.freeRects||[]);
    // FIX: setTransform ile mevcut birikmiş scale'i sıfırla, sonra DR uygula
    ctx.setTransform(DR,0,0,DR,0,0);
    draw(ctx,cw,ch,fitScale,panel.placed,waste,false,RULER_W,RULER_H);
  }
}
function zoomIn(){zoomLevel=Math.min(ZOOM_MAX,+(zoomLevel+ZOOM_STEP).toFixed(2));applyZoom()}
function zoomOut(){zoomLevel=Math.max(ZOOM_MIN,+(zoomLevel-ZOOM_STEP).toFixed(2));applyZoom()}
function zoomReset(){zoomLevel=1;applyZoom();document.getElementById('canvasScroll').scrollTo({top:0,left:0,behavior:'smooth'})}

/* WHEEL ZOOM */
document.getElementById('canvasWrap').addEventListener('wheel',function(e){
  if(document.getElementById('mainCanvas').style.display==='none')return;
  e.preventDefault();
  const d=e.deltaY>0?-ZOOM_STEP:ZOOM_STEP;
  const nz=Math.min(ZOOM_MAX,Math.max(ZOOM_MIN,+(zoomLevel+d).toFixed(2)));
  if(nz===zoomLevel)return;
  const scr=document.getElementById('canvasScroll'),rect=scr.getBoundingClientRect();
  const mx=e.clientX-rect.left+scr.scrollLeft,my=e.clientY-rect.top+scr.scrollTop;
  const ratio=nz/zoomLevel;zoomLevel=nz;applyZoom(false);
  scr.scrollLeft=mx*ratio-(e.clientX-rect.left);
  scr.scrollTop=my*ratio-(e.clientY-rect.top);
},{passive:false});

/* KEYBOARD */
document.addEventListener('keydown',e=>{
  if(e.ctrlKey&&e.key==='='){e.preventDefault();zoomIn()}
  if(e.ctrlKey&&e.key==='-'){e.preventDefault();zoomOut()}
  if(e.ctrlKey&&e.key==='0'){e.preventDefault();zoomReset()}
  if(e.ctrlKey&&e.key==='z'){e.preventDefault();undoLast()}
  if(e.ctrlKey&&e.key==='s'){e.preventDefault();saveProject();showMsg(t('msgSaved'),'success')}
  if(e.ctrlKey&&e.key==='Enter'){e.preventDefault();calculate()}
  if(e.ctrlKey&&(e.key==='y'||e.key==='Y')){e.preventDefault();redoLast()}
  if(e.key==='?'&&!e.ctrlKey&&!e.altKey&&document.activeElement.tagName!=='INPUT'&&document.activeElement.tagName!=='TEXTAREA'){showKbModal()}
  if(e.key==='f'&&!e.ctrlKey&&!e.altKey&&document.activeElement.tagName!=='INPUT'&&document.activeElement.tagName!=='TEXTAREA'){toggleFullscreen()}
  if(e.ctrlKey&&(e.key==='d'||e.key==='D')){
    e.preventDefault();
    // Aktif düzenleme varsa onu kopyala, yoksa son parçayı kopyala
    const targetId=activeEdit||(parts.length?parts[parts.length-1].id:null);
    if(targetId)duplicatePart({stopPropagation:()=>{}},targetId);
  }
  if(e.key==='Escape'){hideKbModal();cancelEdit(activeEdit)}
});

/* TOUCH PINCH — CSS transform sırasında, redraw sonunda */
let ltd=null, _pinchRaf=null;
document.getElementById('canvasWrap').addEventListener('touchstart',e=>{
  if(e.touches.length===2){
    ltd=Math.hypot(e.touches[0].clientX-e.touches[1].clientX,e.touches[0].clientY-e.touches[1].clientY);
    const scr=document.getElementById('canvasScroll');
    scr.classList.add('pinching');
    scr.style.overflow='hidden';
  }
},{passive:true});
document.getElementById('canvasWrap').addEventListener('touchmove',e=>{
  if(e.touches.length===2&&ltd){
    const d=Math.hypot(e.touches[0].clientX-e.touches[1].clientX,e.touches[0].clientY-e.touches[1].clientY);
    const newZoom=Math.min(ZOOM_MAX,Math.max(ZOOM_MIN,+(zoomLevel*(d/ltd)).toFixed(2)));
    ltd=d;
    if(newZoom===zoomLevel)return;
    zoomLevel=newZoom;
    // FIX: gesture sırasında sadece CSS transform — canvas yeniden çizilmez
    const inner=document.getElementById('canvasInner');
    inner.style.transition='none';
    inner.style.transform=`scale(${zoomLevel})`;
    inner.style.transformOrigin='top left';
    document.getElementById('zoomLevelLbl').textContent=Math.round(zoomLevel*100)+'%';
    // Throttle: son hareketten 80ms sonra canvas'ı çiz
    cancelAnimationFrame(_pinchRaf);
    _pinchRaf=requestAnimationFrame(()=>{
      if(allPanelsResult.length){
        const panel=allPanelsResult[currentPanelIdx];
        if(panel){
          const canvas=document.getElementById('mainCanvas');
          const ctx=canvas.getContext('2d');
          ctx.setTransform(DR,0,0,DR,0,0);
          draw(ctx,canvas.width/DR,canvas.height/DR,fitScale,panel.placed,getFreeAsWaste(panel.freeRects||[]));
        }
      }
    });
  }
},{passive:true});
document.getElementById('canvasWrap').addEventListener('touchend',e=>{
  if(e.touches.length<2){
    ltd=null;
    const scr=document.getElementById('canvasScroll');
    scr.classList.remove('pinching');
    scr.style.overflow='';
    // Gesture bitti — temiz bir redraw
    applyZoom(false);
  }
},{passive:true});

/* ORIENTATION & RESIZE — canvas'ı yeniden render et */
let _resizeTimer=null;
window.addEventListener('resize',()=>{
  clearTimeout(_resizeTimer);
  _resizeTimer=setTimeout(()=>{
    if(allPanelsResult.length)renderPanel(currentPanelIdx);
  },200);
});
window.addEventListener('orientationchange',()=>{
  setTimeout(()=>{
    if(allPanelsResult.length)renderPanel(currentPanelIdx);
  },500); // eski Android için daha güvenli gecikme
});
// visualViewport API — iOS Safari'de klavye açılması vs. için daha güvenilir
if(window.visualViewport){
  let _vvTimer=null;
  window.visualViewport.addEventListener('resize',()=>{
    clearTimeout(_vvTimer);
    _vvTimer=setTimeout(()=>{
      if(allPanelsResult.length)renderPanel(currentPanelIdx);
    },200);
  });
}

/* ENTER KEY */
['partName','partNote','partLabelCode','partMaterialCode','partEdgeBand','pw','ph','pq'].forEach(id=>document.getElementById(id).addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();addPart()}}));

/* ═══════════════════════════════════════════
   EXCEL / XLSX EXPORT
═══════════════════════════════════════════ */
function exportXLSX(){
  if(!allPanelsResult.length){showMsg(t('msgCalcFirst'),'error');return}
  ensureXLSX(function(err){
    if(err){showMsg(t('msgExcelLoading'),'error');return;}
    _doExportXLSX();
  });
}
function _doExportXLSX(){

  const sym=getCurrencySymbol();
  const now=new Date();
  const customerName=document.getElementById('customerName').value.trim()||'';
  const orderNo=document.getElementById('orderNo').value.trim()||'';
  const allPlaced=allPanelsResult.flatMap(p=>p.placed);
  const totalPanelArea=SHEET_W*SHEET_H*allPanelsResult.length;
  const totalPanelCost=(_cachedPrices.sheet||0)*allPanelsResult.length;
  const usedGroups={};
  allPlaced.forEach(r=>{if(!usedGroups[r.id])usedGroups[r.id]=0;usedGroups[r.id]++;});

  const wb=XLSX.utils.book_new();

  // ── SAYFA 1: KESİM LİSTESİ ──
  const cutRows=[['#','Parça Adı','Not','Genişlik (mm)','Yükseklik (mm)','X (mm)','Y (mm)','Panel','Birim Maliyet ('+sym+')']];
  const items=getSortedCutList(allPanelsResult);
  items.forEach((r,i)=>{
    const p=parts.find(x=>x.id===r.id);
    const areaM2=(r.w*r.h)/1e6;
    const unitCost=areaM2*(_cachedPrices.sqm||0)+(_cachedPrices.cut||0)+(totalPanelArea>0?(r.w*r.h/totalPanelArea)*totalPanelCost:0);
    cutRows.push([i+1,r.name||r.label||'',(p&&p.note)||'',r.w,r.h,r.x,r.y,'P'+r.panelNum,unitCost>0?parseFloat(unitCost.toFixed(2)):0]);
  });
  const wsCut=XLSX.utils.aoa_to_sheet(cutRows);
  // Sütun genişlikleri
  wsCut['!cols']=[{wch:5},{wch:20},{wch:18},{wch:12},{wch:12},{wch:8},{wch:8},{wch:7},{wch:16}];
  XLSX.utils.book_append_sheet(wb,wsCut,'Kesim Listesi');

  // ── SAYFA 2: PARÇA ÖZETİ ──
  const partRows=[['Parça Adı','Etiket','Renk/Kod','Kenar Bant','Not','G (mm)','Y (mm)','Talep','Yerleşen','Alan/adet (m²)','Toplam Alan (m²)','Birim Maliyet ('+sym+')','Toplam Maliyet ('+sym+')']];
  parts.forEach(p=>{
    const qty=usedGroups[p.id]||0;
    const areaM2=(p.w*p.h)/1e6;
    const unitCost=areaM2*(_cachedPrices.sqm||0)+(_cachedPrices.cut||0)+(totalPanelArea>0?(p.w*p.h/totalPanelArea)*totalPanelCost:0);
    partRows.push([p.name||'',p.labelCode||'',p.materialCode||'',p.edgeBand||'',p.note||'',p.w,p.h,p.qty,qty,parseFloat(areaM2.toFixed(6)),parseFloat((areaM2*qty).toFixed(6)),parseFloat(unitCost.toFixed(2)),parseFloat((unitCost*qty).toFixed(2))]);
  });
  const wsPart=XLSX.utils.aoa_to_sheet(partRows);
  wsPart['!cols']=[{wch:20},{wch:18},{wch:8},{wch:8},{wch:7},{wch:8},{wch:14},{wch:14},{wch:18},{wch:18}];
  XLSX.utils.book_append_sheet(wb,wsPart,'Parça Özeti');

  // ── SAYFA 3: MALİYET ──
  const usedArea=allPlaced.reduce((s,r)=>s+r.w*r.h,0);
  const panelCost=_cachedPrices.sheet*allPanelsResult.length;
  const sqmCost=(usedArea/1e6)*_cachedPrices.sqm;
  const cutCost=allPlaced.length*_cachedPrices.cut;
  const total=panelCost+sqmCost+cutCost;
  var _kdvEl=document.getElementById('kdvEnabled');const kdvEnabled=_kdvEl?_kdvEl.checked:false;
  const kdvRate=parseInt(document.getElementById('kdvRate').value)||20;
  const costRows=[
    ['Maliyet Özeti',''],
    ['Müşteri',customerName],
    ['Teklif No',orderNo],
    ['Teslim Tarihi',(document.getElementById('projectDelivery')||{value:''}).value.trim()],
    ['Operatör',(document.getElementById('operatorName')||{value:''}).value.trim()],
    ['Proje Notu',(document.getElementById('projectNote')||{value:''}).value.trim()],
    ['Atölye Notu',(document.getElementById('workshopNote')||{value:''}).value.trim()],
    ['Proje',document.getElementById('projectName').value.trim()||''],
    ['Tarih',now.toLocaleDateString(currentLang==='en'?'en-GB':'tr-TR')],
    ['',''],
    ['Panel Malzemesi ('+sym+')',parseFloat(panelCost.toFixed(2))],
    ['Malzeme m² ('+sym+')',parseFloat(sqmCost.toFixed(2))],
    ['Kesim İşçiliği ('+sym+')',parseFloat(cutCost.toFixed(2))],
    ['GENEL TOPLAM ('+sym+')',parseFloat(total.toFixed(2))],
  ];
  if(kdvEnabled){
    const kdvAmt=total*kdvRate/100;
    costRows.push(['KDV %'+kdvRate+' ('+sym+')',parseFloat(kdvAmt.toFixed(2))]);
    costRows.push(['KDV DAHİL TOPLAM ('+sym+')',parseFloat((total+kdvAmt).toFixed(2))]);
  }
  const wsCost=XLSX.utils.aoa_to_sheet(costRows);
  wsCost['!cols']=[{wch:28},{wch:20}];
  XLSX.utils.book_append_sheet(wb,wsCost,'Maliyet');

  const fname=`panel-opt-${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}.xlsx`;
  mobileDownloadXLSX(wb, fname);
  showMsg(t('msgExcelDone'),'success');
}
/* ═══════════════════════════════════════════
   TEKLİF PDF
═══════════════════════════════════════════ */
function exportQuotePDF(){
  if(!allPanelsResult.length){showMsg(t('msgCalcFirst'),'error');return}
  ensurePDF(function(err){
    if(err){showMsg(t('msgPdfLoading'),'error');return;}
    _doExportQuotePDF();
  });
}
function _doExportQuotePDF(){
  var _jspdf=window.jspdf;
  if(!_jspdf){showMsg(t('msgPdfLoading'),'error');return;}
  // Fiyat cache'ini tazele — stale değerle PDF üretilmesin
  if(!_cachedPrices.valid)refreshPriceCache();
  const{jsPDF}=_jspdf;
  const pdf=new jsPDF({orientation:'portrait',unit:'mm',format:'a4'});
  const pw=210,ph=297,mg=18;
  const now=new Date();
  const customerName=document.getElementById('customerName').value.trim()||'—';
  const orderNo=document.getElementById('orderNo').value.trim()||'—';
  const projectName=document.getElementById('projectName').value.trim()||'—';
  const sym=getCurrencySymbol();
  const fmt=v=>v.toLocaleString(currentLang==='en'?'en-GB':'tr-TR',{minimumFractionDigits:2,maximumFractionDigits:2});
  const allPlaced=allPanelsResult.flatMap(p=>p.placed);
  const totalPanelArea=SHEET_W*SHEET_H*allPanelsResult.length;
  const totalPanelCost=(_cachedPrices.sheet||0)*allPanelsResult.length;

  // — BAŞLIK ŞERIDI —
  pdf.setFillColor(7,11,18);pdf.rect(0,0,pw,28,'F');
  pdf.setFontSize(14);pdf.setFont('helvetica','bold');pdf.setTextColor(45,156,219);
  pdf.text('Panel Optimizasyonu',mg,12);
  pdf.setFontSize(8);pdf.setFont('helvetica','normal');pdf.setTextColor(150,163,184);
  pdf.text('Kesim & Maliyet Teklifi',mg,18);
  pdf.setTextColor(90,111,138);pdf.setFontSize(7);
  pdf.text(`Tarih: ${now.toLocaleDateString(currentLang==='en'?'en-GB':'tr-TR')} ${now.toLocaleTimeString('tr-TR',{hour:'2-digit',minute:'2-digit'})}`,pw-mg,12,{align:'right'});
  pdf.text(`Teklif No: ${orderNo}`,pw-mg,18,{align:'right'});

  // — MÜŞTERİ & PROJE BİLGİLERİ —
  let y=36;
  pdf.setFontSize(9);pdf.setFont('helvetica','bold');pdf.setTextColor(30,40,60);
  pdf.text('MÜŞTERİ BİLGİLERİ',mg,y);
  pdf.setDrawColor(220,214,232);pdf.setLineWidth(.3);pdf.line(mg,y+2,pw-mg,y+2);
  y+=8;
  const infoRows=[
    ['Müşteri',customerName],['Proje',projectName],
    ['Teslim',(document.getElementById('projectDelivery')||{value:'—'}).value.trim()||'—'],
    ['Operatör',(document.getElementById('operatorName')||{value:'—'}).value.trim()||'—'],
    ['Panel',`${SHEET_W}×${SHEET_H}mm · ${currentMaterial||'—'} · ${currentThickness}mm`],
    ['Algoritma',currentAlgo.toUpperCase()],['Panel Sayısı',allPanelsResult.length+' adet'],
  ];
  infoRows.forEach(([label,val])=>{
    pdf.setFont('helvetica','normal');pdf.setFontSize(8);pdf.setTextColor(90,111,138);pdf.text(label+':',mg,y);
    pdf.setTextColor(30,40,60);pdf.text(val,mg+30,y);y+=6;
  });

  // — PARÇA LİSTESİ —
  y+=4;
  pdf.setFontSize(9);pdf.setFont('helvetica','bold');pdf.setTextColor(30,40,60);
  pdf.text('PARÇA LİSTESİ',mg,y);
  pdf.line(mg,y+2,pw-mg,y+2);y+=8;

  // Tablo başlığı
  const cols=[mg,mg+40,mg+65,mg+85,mg+105,mg+130,mg+150];
  pdf.setFillColor(240,242,247);pdf.rect(mg,y-5,pw-mg*2,7,'F');
  pdf.setFontSize(7);pdf.setFont('helvetica','bold');pdf.setTextColor(60,80,100);
  ['#','Parça Adı','G×Y (mm)','Adet','Alan (m²)','Birim Maliyet','Not'].forEach((h,i)=>pdf.text(h,cols[i]||mg,y));
  y+=5;pdf.setLineWidth(.2);pdf.setDrawColor(200,210,225);pdf.line(mg,y,pw-mg,y);y+=4;

  const usedGroups={};
  allPlaced.forEach(r=>{
    if(!usedGroups[r.id])usedGroups[r.id]=0;
    usedGroups[r.id]++;
  });

  let grandCost=0;let rowNum=0;
  parts.forEach(p=>{
    const qty=usedGroups[p.id]||0;if(!qty)return;
    const areaM2=(p.w*p.h/1e6);
    const unitCost=areaM2*(_cachedPrices.sqm||0)+(_cachedPrices.cut||0)+(totalPanelArea>0?(p.w*p.h/totalPanelArea)*totalPanelCost:0);
    const totalCost=unitCost*qty;
    grandCost+=totalCost;
    if(y>ph-30){pdf.addPage();y=20;}
    if(rowNum%2===0){pdf.setFillColor(248,250,252);pdf.rect(mg,y-4,pw-mg*2,6,'F');}
    pdf.setFont('helvetica','normal');pdf.setFontSize(7);pdf.setTextColor(30,40,60);
    pdf.text(String(++rowNum),cols[0],y);
    pdf.text((p.name||'').substring(0,18),cols[1],y);
    pdf.text(`${p.w}×${p.h}`,cols[2],y);
    pdf.text(String(qty),cols[3],y);
    pdf.text((areaM2*qty).toFixed(4),cols[4],y);
    pdf.text(unitCost>0?`${sym}${fmt(unitCost)}`:'—',cols[5],y);
    pdf.text(((p.labelCode?('['+p.labelCode+'] '):'')+(p.edgeBand?('Bant:'+p.edgeBand+' '):'')+(p.note||'')).substring(0,18),cols[6]||mg,y);
    y+=6;
  });

  // — MALİYET ÖZETİ —
  y+=6;if(y>ph-50){pdf.addPage();y=20;}
  pdf.setFontSize(9);pdf.setFont('helvetica','bold');pdf.setTextColor(30,40,60);
  pdf.text('MALİYET ÖZETİ',mg,y);
  pdf.line(mg,y+2,pw-mg,y+2);y+=10;

  const costRows=[];
  if(_cachedPrices.sheet>0)costRows.push(['Panel Malzemesi',`${allPanelsResult.length} panel × ${sym}${fmt(_cachedPrices.sheet)}`,sym+fmt(_cachedPrices.sheet*allPanelsResult.length)]);
  const usedArea=allPlaced.reduce((s,r)=>s+r.w*r.h,0);
  if(_cachedPrices.sqm>0)costRows.push(['Malzeme (m²)',`${(usedArea/1e6).toFixed(4)} m² × ${sym}${fmt(_cachedPrices.sqm)}`,sym+fmt((usedArea/1e6)*_cachedPrices.sqm)]);
  if(_cachedPrices.cut>0)costRows.push(['Kesim İşçiliği',`${allPlaced.length} kesim × ${sym}${fmt(_cachedPrices.cut)}`,sym+fmt(allPlaced.length*_cachedPrices.cut)]);

  const xL=mg, xV=pw-mg-35;
  costRows.forEach(([label,detail,val])=>{
    pdf.setFont('helvetica','normal');pdf.setFontSize(8);pdf.setTextColor(90,111,138);pdf.text(label,xL,y);
    pdf.setTextColor(100,120,140);pdf.setFontSize(7);pdf.text(detail,xL+35,y);
    pdf.setFont('helvetica','bold');pdf.setTextColor(30,40,60);pdf.setFontSize(8);pdf.text(val,xV,y,{align:'right'});
    y+=7;
  });

  // Toplam
  y+=2;pdf.setLineWidth(.5);pdf.setDrawColor(200,210,225);pdf.line(xV-30,y,pw-mg,y);y+=5;
  pdf.setFont('helvetica','bold');pdf.setFontSize(10);pdf.setTextColor(30,40,60);pdf.text('GENEL TOPLAM',xL,y);
  pdf.setTextColor(26,158,80);pdf.text(`${sym}${fmt(grandCost)}`,pw-mg,y,{align:'right'});

  // KDV
  if(document.getElementById('kdvEnabled').checked){
    const rate=parseInt(document.getElementById('kdvRate').value)||20;
    const kdvAmt=grandCost*rate/100;
    y+=7;
    pdf.setFont('helvetica','normal');pdf.setFontSize(8);pdf.setTextColor(90,111,138);
    pdf.text(`KDV (%${rate})`,xL,y);
    pdf.setFont('helvetica','bold');pdf.setTextColor(251,140,60);
    pdf.text(`${sym}${fmt(kdvAmt)}`,pw-mg,y,{align:'right'});
    y+=7;
    pdf.setFont('helvetica','bold');pdf.setFontSize(11);pdf.setTextColor(30,40,60);
    pdf.text('KDV DAHİL TOPLAM',xL,y);
    pdf.setTextColor(26,158,80);
    pdf.text(`${sym}${fmt(grandCost+kdvAmt)}`,pw-mg,y,{align:'right'});
  }

  // — FOOTER —
  pdf.setFontSize(6.5);pdf.setFont('helvetica','normal');pdf.setTextColor(150,163,184);
  pdf.text(`Panel Optimizasyonu · ${now.toLocaleDateString(currentLang==='en'?'en-GB':'tr-TR')} · ${currentAlgo.toUpperCase()} · Alan kullanımı: ${((usedArea/(SHEET_W*SHEET_H*allPanelsResult.length))*100).toFixed(1)}%`,pw/2,ph-11,{align:'center'});
  pdf.setFontSize(6);pdf.setTextColor(120,130,150);
  pdf.text('Hazırlayan: Mücahit Bilgin',pw/2,ph-6,{align:'center'});

  const fname=`teklif-${(orderNo||'').replace(/[^a-zA-Z0-9]/g,'-')||Date.now()}.pdf`;
  mobileDownloadPDF(pdf, fname);
  showMsg(t('msgSaved'),'success');
}
/* ═══════════════════════════════════════════
   SVG EXPORT (vektörel)
═══════════════════════════════════════════ */
function exportSVG(){
  if(!allPanelsResult.length){showMsg(t('msgCalcFirst'),'error');return}
  const panel=allPanelsResult[currentPanelIdx];
  const waste=getFreeAsWaste(panel.freeRects||[]);
  const{edge}=getMargins();
  const W=SHEET_W, H=SHEET_H;
  const now=new Date();
  const tzShort=now.toLocaleTimeString('tr-TR',{timeZoneName:'short'}).split(' ').pop();

  let svg=`<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">
  <title>Panel Optimizasyonu — Panel ${currentPanelIdx+1}/${allPanelsResult.length}</title>
  <desc>Oluşturma: ${now.toLocaleDateString(currentLang==='en'?'en-GB':'tr-TR')} ${now.toLocaleTimeString('tr-TR',{hour:'2-digit',minute:'2-digit'})} ${tzShort} · ${currentAlgo.toUpperCase()} · ${W}×${H}mm</desc>
  <!-- Arkaplan -->
  <rect width="${W}" height="${H}" fill="#f8f8f8" stroke="#ccc" stroke-width="1"/>
`;

  // Edge margin
  if(edge>0){
    svg+=`  <!-- Kenar payı -->
  <rect x="${edge}" y="${edge}" width="${W-edge*2}" height="${H-edge*2}" fill="none" stroke="#f0c040" stroke-width="1" stroke-dasharray="6,4" opacity="0.6"/>
`;
  }

  // Fire alanları
  svg+=`  <!-- Fire alanları -->
`;
  for(const ws of waste){
    if(ws.w<=0||ws.h<=0)continue;
    svg+=`  <rect x="${ws.x}" y="${ws.y}" width="${ws.w}" height="${ws.h}" fill="#f06060" fill-opacity="0.08" stroke="#f06060" stroke-width="0.5" stroke-dasharray="4,4"/>
`;
  }

  // Parçalar
  svg+=`  <!-- Parçalar -->
`;
  for(const r of panel.placed){
    const p=parts.find(x=>x.id===r.id);
    const col=r.color||'#888';
    const lbl=esc(r.name||r.label||'');
    const note=p&&p.note?esc(p.note):'';
    const fs=Math.min(14,Math.max(7,Math.min(r.w,r.h)*0.12));
    svg+=`  <rect x="${r.x}" y="${r.y}" width="${r.w}" height="${r.h}" fill="${col}" fill-opacity="0.7" stroke="${col}" stroke-width="1"/>
  <text x="${r.x+r.w/2}" y="${r.y+r.h/2-(note?fs*.6:0)}" text-anchor="middle" dominant-baseline="middle" font-family="monospace" font-size="${fs}" font-weight="bold" fill="#fff" opacity="0.9">${lbl}</text>
${note?`  <text x="${r.x+r.w/2}" y="${r.y+r.h/2+fs*.8}" text-anchor="middle" dominant-baseline="middle" font-family="monospace" font-size="${Math.max(6,fs*.72)}" fill="#fff" opacity="0.65">${note}</text>`:''}
  <text x="${r.x+2}" y="${r.y+r.h-3}" font-family="monospace" font-size="${Math.max(5,fs*.55)}" fill="${col}" opacity="0.7">${r.w}×${r.h}</text>
`;
  }

  // Panel sınırı
  svg+=`  <!-- Panel sınırı -->
  <rect width="${W}" height="${H}" fill="none" stroke="#333" stroke-width="2"/>
`;

  svg+=`</svg>`;

  const fname=`panel-${currentPanelIdx+1}-${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}.svg`;
  mobileDownload('data:image/svg+xml;charset=utf-8,'+encodeURIComponent(svg), fname);
  showMsg('SVG indirildi.','success');
}
/* ═══════════════════════════════════════════
   DOWNLOAD PNG
═══════════════════════════════════════════ */
function downloadPNG(){
  if(!allPanelsResult.length){showMsg(t('msgCalcFirst'),'error');return}
  const panelNum=currentPanelIdx+1;
  const panel=allPanelsResult[currentPanelIdx];
  if(!panel){showMsg(t('msgPanelNotFound'),'error');return}

  // Zaman damgası bilgisi
  const now=new Date();
  const tz=Intl.DateTimeFormat().resolvedOptions().timeZone;
  const tzShort=now.toLocaleTimeString('tr-TR',{timeZoneName:'short'}).split(' ').pop();
  const dateStr=now.toLocaleDateString(currentLang==='en'?'en-GB':'tr-TR');
  const timeStr=now.toLocaleTimeString('tr-TR',{hour:'2-digit',minute:'2-digit',second:'2-digit'});

  // Yüksek kalite offscreen render — ekran canvas'ını kopyalamak yerine yeniden çiz
  const EXPORT_PX=3000;
  const sc=Math.min(EXPORT_PX/SHEET_W,EXPORT_PX/SHEET_H);
  const lcw=Math.round(SHEET_W*sc),lch=Math.round(SHEET_H*sc);
  const SS=2;
  const waste=getFreeAsWaste(panel.freeRects||[]);
  const oc=document.createElement('canvas');oc.width=lcw*SS;oc.height=lch*SS;
  const octx=oc.getContext('2d');octx.scale(SS,SS);
  draw(octx,lcw,lch,sc,panel.placed,waste,true);

  // Watermark şeridi — SS koordinatları
  const bh=Math.round(28*SS);
  octx.setTransform(1,0,0,1,0,0);
  octx.fillStyle='rgba(17,18,20,.82)';
  octx.fillRect(0,oc.height-bh,oc.width,bh);
  octx.font=`${Math.round(8*SS)}px 'DM Mono',monospace`;
  octx.fillStyle='rgba(90,111,138,.9)';
  octx.textAlign='left';octx.textBaseline='middle';
  octx.fillText(`${currentLang==='en'?'Panel Optimizer':'Panel Optimizasyonu'} · Panel ${panelNum}/${allPanelsResult.length} · ${SHEET_W}×${SHEET_H}mm · ${currentAlgo.toUpperCase()}`,Math.round(8*SS),oc.height-bh*0.65);
  octx.font=`${Math.round(6.5*SS)}px 'DM Mono',monospace`;
  octx.fillStyle='rgba(70,90,115,.75)';
  octx.fillText('Hazırlayan: Mücahit Bilgin',Math.round(8*SS),oc.height-bh*0.28);
  octx.textAlign='right';
  octx.fillStyle='rgba(129,140,248,.85)';
  octx.fillText(`${dateStr} ${timeStr} ${tzShort}`,oc.width-Math.round(8*SS),oc.height-bh*0.65);

  const fname=`panel-${panelNum}-${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}-${String(now.getHours()).padStart(2,'0')}${String(now.getMinutes()).padStart(2,'0')}.png`;
  mobileDownload(oc.toDataURL('image/png',1), fname);
}

/* ═══════════════════════════════════════════
   DOWNLOAD ALL PANELS AS PNG (sıralı)
═══════════════════════════════════════════ */
function downloadAllPNG(){
  if(!allPanelsResult.length){showMsg(t('msgCalcFirst'),'error');return}
  const btn=document.getElementById('pngAllBtn');
  btn.textContent=t('msgPreparing');btn.disabled=true;

  const now=new Date();
  const tzShort=now.toLocaleTimeString('tr-TR',{timeZoneName:'short'}).split(' ').pop();
  const dateStr=now.toLocaleDateString(currentLang==='en'?'en-GB':'tr-TR');
  const timeStr=now.toLocaleTimeString('tr-TR',{hour:'2-digit',minute:'2-digit',second:'2-digit'});
  const datePfx=`${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}-${String(now.getHours()).padStart(2,'0')}${String(now.getMinutes()).padStart(2,'0')}`;
  const DR2=window.devicePixelRatio||2;

  // Her paneli offscreen canvas'a çizip sırayla indir
  let delay=0;
  allPanelsResult.forEach((panel,pi)=>{
    setTimeout(()=>{
      const waste=getFreeAsWaste(panel.freeRects||[]);
      // Yüksek kalite export — uzun kenar 3000px hedefi
      const EXPORT_PX=3000;
      const sc=Math.min(EXPORT_PX/SHEET_W,EXPORT_PX/SHEET_H);
      const lcw=Math.round(SHEET_W*sc),lch=Math.round(SHEET_H*sc);
      const SS=2; // supersampling
      const oc=document.createElement('canvas');oc.width=lcw*SS;oc.height=lch*SS;
      const octx=oc.getContext('2d');octx.scale(SS,SS);
      draw(octx,lcw,lch,sc,panel.placed,waste,true);

      // Watermark şeridi — SS ile ölçeklenmiş koordinatlar
      const bh=Math.round(28*SS);
      octx.setTransform(1,0,0,1,0,0); // scale sıfırla
      octx.fillStyle='rgba(17,18,20,.82)';
      octx.fillRect(0,oc.height-bh,oc.width,bh);
      octx.font=`${Math.round(8*SS)}px 'DM Mono',monospace`;
      octx.fillStyle='rgba(90,111,138,.9)';
      octx.textAlign='left';octx.textBaseline='middle';
      octx.fillText(`${currentLang==='en'?'Panel Optimizer':'Panel Optimizasyonu'} · Panel ${pi+1}/${allPanelsResult.length} · ${SHEET_W}×${SHEET_H}mm · ${currentAlgo.toUpperCase()}`,Math.round(8*SS),oc.height-bh*0.65);
      octx.font=`${Math.round(6.5*SS)}px 'DM Mono',monospace`;
      octx.fillStyle='rgba(70,90,115,.75)';
      octx.fillText('Hazırlayan: Mücahit Bilgin',Math.round(8*SS),oc.height-bh*0.28);
      octx.font=`${Math.round(8*SS)}px 'DM Mono',monospace`;
      octx.textAlign='right';
      octx.fillStyle='rgba(129,140,248,.85)';
      octx.fillText(`${dateStr} ${timeStr} ${tzShort}`,oc.width-Math.round(8*SS),oc.height-bh*0.65);

      mobileDownload(oc.toDataURL('image/png',1), `panel-${pi+1}-${datePfx}.png`);

      // Son panel: butonu sıfırla
      if(pi===allPanelsResult.length-1){
        setTimeout(()=>{
          btn.innerHTML=`<svg width="11" height="11" viewBox="0 0 24 24" fill="none"><path d="M12 15V3M12 15L8 11M12 15L16 11" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/><path d="M3 17v2a2 2 0 002 2h14a2 2 0 002-2v-2" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"/></svg> PNG ×${allPanelsResult.length}`;
          btn.disabled=false;
          showMsg(`${allPanelsResult.length} panel PNG olarak indirildi.`,'success');
        },300);
      }
    },delay);
    delay+=400; // Tarayıcı birden fazla indirmeyi engellemsin diye kısa gecikme
  });
}

/* ═══════════════════════════════════════════
   DOWNLOAD PDF (ALL PANELS)
═══════════════════════════════════════════ */
function downloadPDF(){
  if(!allPanelsResult.length){showMsg(t('msgCalcFirst'),'error');return}
  const btn=document.getElementById('pdfBtn');
  btn.textContent=t('msgPreparing');btn.disabled=true;
  setTimeout(()=>{
    if(!window.jspdf||!window.jspdf.jsPDF){
      showMsg(t('msgPdfLoading'),'error');
      btn.innerHTML=`<svg width="11" height="11" viewBox="0 0 24 24" fill="none"><path d="M12 15V3M12 15L8 11M12 15L16 11" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/><path d="M3 17v2a2 2 0 002 2h14a2 2 0 002-2v-2" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"/></svg> PDF`;
      btn.disabled=false;return;
    }
    const{jsPDF}=window.jspdf;
    const ratio=SHEET_W/SHEET_H;
    let pw,ph,ori;
    if(ratio>1){ori='landscape';pw=297;ph=210}else{ori='portrait';pw=210;ph=297}
    const pdf=new jsPDF({orientation:ori,unit:'mm',format:'a4'});
    const mg=10,aw=pw-mg*2,ah=ph-mg*2-28;

    // Zaman bilgisi — PDF oluşturma anı (tüm sayfalarda aynı)
    const now=new Date();
    const tz=Intl.DateTimeFormat().resolvedOptions().timeZone;
    const tzShort=now.toLocaleTimeString('tr-TR',{timeZoneName:'short'}).split(' ').pop();
    const dateStr=now.toLocaleDateString('tr-TR',{day:'2-digit',month:'2-digit',year:'numeric'});
    const timeStr=now.toLocaleTimeString('tr-TR',{hour:'2-digit',minute:'2-digit',second:'2-digit'});
    const utcOffset=now.toLocaleTimeString('en',{timeZoneName:'longOffset'}).split(' ').pop().replace('GMT','UTC');
    const fullTimeLabel=`${dateStr}  ${timeStr}  ${tzShort} (${utcOffset})`;

    for(let pi=0;pi<allPanelsResult.length;pi++){
      if(pi>0)pdf.addPage();
      const panel=allPanelsResult[pi];
      const waste=getFreeAsWaste(panel.freeRects||[]);

      // Draw to offscreen canvas — yüksek çözünürlük
      // PDF'deki alan mm cinsinden; piksele çevirmek için 96dpi / 25.4mm * supersampling
      // Hedef: en az 2400×2400px fiziksel canvas (baskı kalitesi)
      const EXPORT_PX = 3000; // uzun kenar piksel hedefi
      const sc = Math.min(EXPORT_PX / SHEET_W, EXPORT_PX / SHEET_H);
      const lcw = Math.round(SHEET_W * sc); // logical width (draw fonksiyonuna verilecek)
      const lch = Math.round(SHEET_H * sc); // logical height
      const SS = 2; // supersampling çarpanı — anti-aliasing için
      const oc = document.createElement('canvas');
      oc.width = lcw * SS; oc.height = lch * SS;
      const octx = oc.getContext('2d');
      octx.scale(SS, SS);
      draw(octx, lcw, lch, sc, panel.placed, waste, true);
      const img = oc.toDataURL('image/png', 1);

      let dw=aw,dh=aw*(lch/lcw);if(dh>ah){dh=ah;dw=ah*(lcw/lch)}
      const ox=mg+(aw-dw)/2,oy=mg+22;

      // Arkaplan
      pdf.setFillColor(7,11,18);pdf.rect(0,0,pw,ph,'F');
      // Üst başlık şeridi (daha geniş — zaman için yer)
      pdf.setFillColor(11,16,24);pdf.rect(0,0,pw,20,'F');

      // Sol: Başlık
      const _cust=document.getElementById('customerName').value.trim();
      const _ordn=document.getElementById('orderNo').value.trim();
      pdf.setFont('helvetica','bold');pdf.setFontSize(10);pdf.setTextColor(45,156,219);
      pdf.text(`Panel Optimizasyonu — Panel ${pi+1}/${allPanelsResult.length}`,mg,8);
      if(_cust||_ordn){
        pdf.setFontSize(7);pdf.setFont('helvetica','normal');pdf.setTextColor(45,156,219);
        const custLine=[_cust,_ordn].filter(Boolean).join('  |  No: ');
        pdf.text(custLine,mg,15);
      }

      // Sağ üst: Panel + algo bilgisi
      const grainInfo=currentGrain!=='any'?` · Tahıl:${currentGrain==='horizontal'?'Yatay':'Dikey'}`:'';
      pdf.setFontSize(7);pdf.setFont('helvetica','normal');pdf.setTextColor(90,111,138);
      pdf.text(`${SHEET_W}×${SHEET_H}mm · ${currentAlgo.toUpperCase()}${grainInfo}`,pw-mg,8,{align:'right'});

      // Sağ alt (başlık şeridinin alt satırı): Zaman bilgisi — mor/purple renk
      pdf.setFontSize(6.5);pdf.setTextColor(45,156,219);
      pdf.text(`🕐 ${fullTimeLabel}`,pw-mg,15,{align:'right'});

      // Sol alt: Zaman dilimi tam adı
      pdf.setFontSize(6);pdf.setTextColor(70,90,115);
      pdf.text(`TZ: ${tz}`,mg,15);

      pdf.setDrawColor(29,39,64);pdf.setLineWidth(.3);pdf.line(mg,20,pw-mg,20);
      pdf.addImage(img,'PNG',ox,oy,dw,dh);
      pdf.setDrawColor(29,39,64);pdf.rect(ox,oy,dw,dh);

      // Stats footer
      const sy=oy+dh+5;
      if(sy<ph-10){
        const totalPlaced=panel.placed.length;
        const used=panel.placed.reduce((s,r)=>s+r.w*r.h,0);
        const uPct=(used/(SHEET_W*SHEET_H)*100).toFixed(1);
        const si=[
          {l:'Yerleştirilen',v:totalPlaced+' adet',c:[79,209,122]},
          {l:'Alan Kull.',v:uPct+'%',c:[65,196,240]},
          {l:'Fire',v:(100-uPct).toFixed(1)+'%',c:[240,96,96]},
          {l:'Algoritma',v:currentAlgo.toUpperCase(),c:[240,192,64]},
          {l:'Malzeme',v:(currentMaterial||(document.getElementById('materialType').value||'—'))+' '+(currentThickness||18)+'mm',c:[185,124,248]}
        ];
        const cw2=(pw-mg*2)/Math.min(si.length,5);
        si.forEach((s,i)=>{
          const cx=mg+cw2*i+cw2/2;
          pdf.setTextColor(...s.c);pdf.setFontSize(9);pdf.setFont('helvetica','bold');pdf.text(s.v,cx,sy+5,{align:'center'});
          pdf.setFontSize(6);pdf.setFont('helvetica','normal');pdf.setTextColor(90,111,138);pdf.text(s.l,cx,sy+9,{align:'center'});
        });

        // Footer alt çizgi + zaman damgası (sayfa altı)
        if(sy+14<ph-4){
          pdf.setDrawColor(29,39,64);pdf.setLineWidth(.2);pdf.line(mg,sy+13,pw-mg,sy+13);
          pdf.setFontSize(5.5);pdf.setFont('helvetica','normal');pdf.setTextColor(50,70,100);
          pdf.text(`İndirilme zamanı: ${fullTimeLabel}  ·  ${tz}`,pw/2,sy+17,{align:'center'});
          pdf.setFontSize(5);pdf.setTextColor(90,100,120);
          pdf.text('Hazırlayan: Mücahit Bilgin',pw/2,sy+21,{align:'center'});
        }
      }
    }

    const fname=`panel-opt-${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}-${String(now.getHours()).padStart(2,'0')}${String(now.getMinutes()).padStart(2,'0')}.pdf`;
    mobileDownloadPDF(pdf, fname);
    btn.innerHTML=`<svg width="11" height="11" viewBox="0 0 24 24" fill="none"><path d="M12 15V3M12 15L8 11M12 15L16 11" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/><path d="M3 17v2a2 2 0 002 2h14a2 2 0 002-2v-2" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"/></svg> PDF`;
    btn.disabled=false;
  },50);
}

/* ═══════════════════════════════════════════
   OTOMATİK YEDEK
   Her hesap sonrası son 5 durumu saklar.
   Boyut optimizasyonu: parça verisini tam kopyalamak
   yerine o anki pop9_parts değerini referans olarak
   saklıyoruz. Bu sayede 5 yedek yerine sadece 1 kopya
   + 5 adet küçük ayar objesi tutulur (~%80 tasarruf).
═══════════════════════════════════════════ */
function autoBackup(){
  try{
    // Parça verisini o anki localStorage değerinden al (save() zaten yazdı)
    const partsRef=localStorage.getItem('pop9_parts')||'[]';
    // Boyut tahmini — uyarı için
    const estBytes=partsRef.length*2;
    if(estBytes>200000){
      console.warn('autoBackup: parça verisi büyük (~'+Math.round(estBytes/1024)+'KB), yedek atlandı');
      return;
    }
    const snapshot={
      ts:Date.now(),
      partsJSON:partsRef,          // tam kopya değil, string referans
      colorIdx,
      sheetW:SHEET_W,sheetH:SHEET_H,
      algo:currentAlgo,grain:currentGrain,
      cutOrder:currentCutOrder||'placement',
      marginEdge:parseInt(document.getElementById('marginEdge').value)||0,
      marginCut:parseInt(document.getElementById('marginCut').value)||0,
      price:parseFloat(document.getElementById('sheetPrice').value)||0,
      sqmPrice:parseFloat(document.getElementById('sqmPrice').value)||0,
      cutPrice:parseFloat(document.getElementById('cutPrice').value)||0,
      currency:currentCurrency,
      material:currentMaterial,
      thickness:currentThickness,
      customer:(document.getElementById('customerName')||{value:''}).value.trim(),
      orderNo:(document.getElementById('orderNo')||{value:''}).value.trim(),
    delivery:(document.getElementById('projectDelivery')||{value:''}).value.trim(),
    projectNote:(document.getElementById('projectNote')||{value:''}).value.trim(),
    operator:(document.getElementById('operatorName')||{value:''}).value.trim(),
    workshopNote:(document.getElementById('workshopNote')||{value:''}).value.trim(),
      kdvEnabled:document.getElementById('kdvEnabled').checked,
      kdvRate:parseInt(document.getElementById('kdvRate').value)||20,
      symmetric:document.getElementById('symmetricMode').checked,
      stock:parseInt(document.getElementById('sheetStock').value)||0
    };
    let backups=[];
    try{backups=JSON.parse(localStorage.getItem('pop9_backups')||'[]');}catch(e){}
    // Aynı dakika içinde arka arkaya hesap yapıldıysa öncekini değiştir
    if(backups.length&&Date.now()-backups[0].ts<60000){backups[0]=snapshot;}
    else{
      backups.unshift(snapshot);
      if(backups.length>5)backups=backups.slice(0,5);
    }
    localStorage.setItem('pop9_backups',JSON.stringify(backups));
  }catch(e){
    if(e.name==='QuotaExceededError')console.warn('Yedek: depolama dolu');
  }
}
function getBackups(){
  try{return JSON.parse(localStorage.getItem('pop9_backups')||'[]');}catch(e){return[];}
}
function restoreBackup(idx){
  const backups=getBackups();
  const b=backups[idx];
  if(!b)return;
  pushUndo(parts,'Yedekten geri yükle');
  // Geriye dönük uyumluluk: eski format parts[], yeni format partsJSON string
  try{parts=b.partsJSON?JSON.parse(b.partsJSON):(b.parts||[]);}catch(e){parts=[];}
  colorIdx=b.colorIdx||parts.length;
  SHEET_W=b.sheetW;SHEET_H=b.sheetH;
  document.getElementById('sheetW').value=b.sheetW;
  document.getElementById('sheetH').value=b.sheetH;
  document.getElementById('marginEdge').value=b.marginEdge||0;
  document.getElementById('marginCut').value=b.marginCut||0;
  document.getElementById('sheetPrice').value=b.price||0;
  document.getElementById('sqmPrice').value=b.sqmPrice||0;
  document.getElementById('cutPrice').value=b.cutPrice||0;
  setAlgo(b.algo||'maxrects');
  if(b.grain)setGrain(b.grain);
  if(b.currency)setCurrency(b.currency);
  if(b.material){currentMaterial=b.material;const ms=document.getElementById('materialType');if(ms)ms.value=b.material;}
  if(b.thickness){currentThickness=b.thickness;const ts=document.getElementById('sheetThickness');if(ts)ts.value=b.thickness;}
  const _bcn=document.getElementById('customerName');if(_bcn&&b.customer!==undefined)_bcn.value=b.customer||'';
  const _bon=document.getElementById('orderNo');if(_bon&&b.orderNo!==undefined)_bon.value=b.orderNo||'';
  if(b.kdvEnabled!==undefined){const ke=document.getElementById('kdvEnabled');if(ke)ke.checked=b.kdvEnabled;}
  if(b.kdvRate){const kr=document.getElementById('kdvRate');if(kr)kr.value=b.kdvRate;}
  if(b.symmetric!==undefined){const sm=document.getElementById('symmetricMode');if(sm)sm.checked=b.symmetric;}
  if(b.stock!==undefined){const sk=document.getElementById('sheetStock');if(sk)sk.value=b.stock||'';}
  _cachedPrices.valid=false;
  renderParts();updateSubtitle();save();
  showMsg('Yedek geri yüklendi: '+new Date(b.ts).toLocaleString('tr-TR'),'success');
  renderBackupList();
}
function renderBackupList(){
  const el=document.getElementById('backupList');
  if(!el)return;
  const backups=getBackups();
  if(!backups.length){el.innerHTML=`<div class="project-empty">${t('emptyBackups')}</div>`;return;}
  el.innerHTML=backups.map((b,i)=>{
    let pCount=0;
    try{pCount=b.partsJSON?JSON.parse(b.partsJSON).length:(b.parts?b.parts.length:0);}catch(e){}
    return `
    <div class="project-item" onclick="restoreBackup(${i})" title="Geri yükle">
      <div class="project-item-name">${pCount} parça · ${b.sheetW}×${b.sheetH}</div>
      <div class="project-item-meta">${new Date(b.ts).toLocaleString(currentLang==='en'?'en-GB':'tr-TR',{day:'2-digit',month:'2-digit',hour:'2-digit',minute:'2-digit'})}</div>
      <span style="font-size:.62rem;color:var(--cyan)">↩</span>
    </div>`;
  }).join('');
}
/* ═══════════════════════════════════════════
   JSON DRAG-DROP PROJE YÜKLEME
═══════════════════════════════════════════ */
(function setupProjectDrop(){
  // Sidebar geneli ve drop zone için dinle
  ['sidebar','projectDropZone'].forEach(id=>{
    const el=document.getElementById(id);
    if(!el)return;
    el.addEventListener('dragover',e=>{
      // Sadece dosya ise kabul et
      if(e.dataTransfer.types.includes('Files')){
        e.preventDefault();
        const dz=document.getElementById('projectDropZone');
        if(dz){dz.style.borderColor='var(--gold)';dz.style.color='var(--gold)';dz.style.background='var(--gold-dim)';}
      }
    });
    el.addEventListener('dragleave',e=>{
      const dz=document.getElementById('projectDropZone');
      if(dz&&!el.contains(e.relatedTarget)){
        dz.style.borderColor='';dz.style.color='';dz.style.background='';
      }
    });
    el.addEventListener('drop',e=>{
      e.preventDefault();
      const dz=document.getElementById('projectDropZone');
      if(dz){dz.style.borderColor='';dz.style.color='';dz.style.background='';}
      const files=e.dataTransfer.files;
      if(!files.length)return;
      const file=files[0];
      if(!file.name.endsWith('.json')){showMsg(t('msgJsonOnly'),'error');return;}
      const reader=new FileReader();
      reader.onload=ev=>{
        try{
          const proj=JSON.parse(ev.target.result);
          parts=proj.parts||[];colorIdx=proj.colorIdx||parts.length;
          document.getElementById('sheetW').value=proj.sheetW||1830;
          document.getElementById('sheetH').value=proj.sheetH||3660;
          document.getElementById('marginEdge').value=proj.marginEdge||0;
          document.getElementById('marginCut').value=proj.marginCut||0;
          document.getElementById('sheetPrice').value=proj.price||0;
          document.getElementById('sqmPrice').value=proj.sqmPrice||0;
          document.getElementById('cutPrice').value=proj.cutPrice||0;
          SHEET_W=proj.sheetW||1830;SHEET_H=proj.sheetH||3660;
          setAlgo(proj.algo||'maxrects');
          if(proj.grain)setGrain(proj.grain);
          if(proj.cutOrder)setCutOrder(proj.cutOrder);
          if(proj.currency)setCurrency(proj.currency);
          if(proj.material){currentMaterial=proj.material;const ms=document.getElementById('materialType');if(ms)ms.value=proj.material;}
          if(proj.thickness){currentThickness=proj.thickness;const ts=document.getElementById('sheetThickness');if(ts)ts.value=proj.thickness;}
          const _cn=document.getElementById('customerName');if(_cn&&proj.customer!==undefined)_cn.value=proj.customer||'';
          const _on=document.getElementById('orderNo');if(_on&&proj.orderNo!==undefined)_on.value=proj.orderNo||'';
      [['projectDelivery','delivery'],['projectNote','projectNote'],['operatorName','operator'],['workshopNote','workshopNote']].forEach(([id,key])=>{const el=document.getElementById(id);if(el&&proj[key]!==undefined)el.value=proj[key]||'';});
          if(proj.kdvEnabled!==undefined){const ke=document.getElementById('kdvEnabled');if(ke)ke.checked=proj.kdvEnabled;}
          if(proj.kdvRate){const kr=document.getElementById('kdvRate');if(kr)kr.value=proj.kdvRate;}
          if(proj.symmetric!==undefined){const sm=document.getElementById('symmetricMode');if(sm)sm.checked=proj.symmetric;}
          if(proj.stock!==undefined){const sk=document.getElementById('sheetStock');if(sk)sk.value=proj.stock||'';}
          _cachedPrices.valid=false;
          renderParts();updateSubtitle();save();
          showMsg(t('msgDropLoaded')+esc(proj.name||''),'success');
        }catch(err){showMsg(t('msgInvalidJSON'),'error');}
      };
      reader.readAsText(file);
    });
  });
})();
/* ═══════════════════════════════════════════
   INIT
═══════════════════════════════════════════ */
load();

/* ═══════════════════════════════════════════
   MOBİL ONAY MODALI (window.confirm yerine)
═══════════════════════════════════════════ */
let _confirmResolve=null;
function mobileConfirm(msg){
  return new Promise(resolve=>{
    document.getElementById('confirmMsg').textContent=msg;
    const modal=document.getElementById('confirmModal');
    modal.classList.add('visible');
    _confirmResolve=function(ok){
      modal.classList.remove('visible');
      _confirmResolve=null;
      resolve(ok);
    };
  });
}
document.getElementById('confirmModal').addEventListener('click',function(e){
  if(e.target===this&&_confirmResolve)_confirmResolve(false);
});

/* ═══════════════════════════════════════════
   CANVAS HOVER TOOLTIP + DOKUNMATIK DESTEK
═══════════════════════════════════════════ */
(function setupCanvasTooltip(){
  const wrap=document.getElementById('canvasWrap');
  const scr=document.getElementById('canvasScroll');
  const tip=document.getElementById('canvasTooltip');
  const fmt=v=>v.toLocaleString(currentLang==='en'?'en-GB':'tr-TR',{minimumFractionDigits:2,maximumFractionDigits:2});
  let _touchTipTimer=null;

  function showTooltipAt(clientX,clientY){
    const canvas=document.getElementById('mainCanvas');
    if(canvas.style.display==='none'||!allPanelsResult.length){tip.classList.remove('visible');return}
    const panel=allPanelsResult[currentPanelIdx];
    if(!panel){tip.classList.remove('visible');return}
    if(!fitScale||fitScale<=0){tip.classList.remove('visible');return}

    const inner=document.getElementById('canvasInner');
    const innerRect=inner.getBoundingClientRect();
    const px=(clientX-innerRect.left)/zoomLevel;
    const py=(clientY-innerRect.top)/zoomLevel;
    const mx=px/fitScale, my=py/fitScale;

    let hit=null;
    for(const r of panel.placed){
      if(mx>=r.x&&mx<=r.x+r.w&&my>=r.y&&my<=r.y+r.h){hit=r;break}
    }
    if(!hit){tip.classList.remove('visible');return}

    const sqmPrice=parseFloat(document.getElementById('sqmPrice').value)||0;
    const cutPrice=parseFloat(document.getElementById('cutPrice').value)||0;
    const sheetPrice=parseFloat(document.getElementById('sheetPrice').value)||0;
    const totalPanelArea=SHEET_W*SHEET_H*allPanelsResult.length;
    const totalPanelCost=sheetPrice*allPanelsResult.length;
    const areaM2=(hit.w*hit.h)/1e6;
    const unitCost=areaM2*sqmPrice+cutPrice+(totalPanelArea>0?(hit.w*hit.h/totalPanelArea)*totalPanelCost:0);

    document.getElementById('ct-name').textContent=hit.name||(hit.label||'');
    document.getElementById('ct-dim').textContent=hit.w+'×'+hit.h+' mm'+(hit.rotated?` (${t('rotatedLabel')})`:'');
    document.getElementById('ct-pos').textContent='X:'+hit.x+' Y:'+hit.y+' mm';
    const costEl=document.getElementById('ct-cost');
    if(unitCost>0){costEl.textContent=getCurrencySymbol()+fmt(unitCost)+'/adet';costEl.style.display='block';}
    else{costEl.style.display='none';}

    const wRect=wrap.getBoundingClientRect();
    let tx=clientX-wRect.left+14;
    let ty=clientY-wRect.top+14;
    tip.classList.add('visible');
    const tw=tip.offsetWidth,th=tip.offsetHeight;
    if(tx+tw>wRect.width-8)tx=clientX-wRect.left-tw-10;
    if(ty+th>wRect.height-8)ty=clientY-wRect.top-th-10;
    tip.style.left=tx+'px';tip.style.top=ty+'px';
    return true;
  }

  // Mouse (masaüstü)
  wrap.addEventListener('mousemove',e=>showTooltipAt(e.clientX,e.clientY));
  wrap.addEventListener('mouseleave',()=>tip.classList.remove('visible'));
  scr.addEventListener('scroll',()=>tip.classList.remove('visible'));

  // Touch (mobil) — tek parmak dokunuş sonrası 2 sn göster
  wrap.addEventListener('touchend',e=>{
    if(e.touches.length>0)return; // çok parmak yok say
    if(ltd!==null)return; // pinch bitti, tooltip gösterme
    const t=e.changedTouches[0];
    clearTimeout(_touchTipTimer);
    const shown=showTooltipAt(t.clientX,t.clientY);
    if(shown){
      _touchTipTimer=setTimeout(()=>tip.classList.remove('visible'),2500);
    }
  },{passive:true});
})();

/* ═══════════════════════════════════════════
   CANLI SAAT
═══════════════════════════════════════════ */
(function startClock(){
  const timeEl=document.getElementById('clockTime');
  const tzEl=document.getElementById('clockTZ');

  // Zaman dilimi kısa adını bir kez al (değişmez)
  const tzFull=Intl.DateTimeFormat().resolvedOptions().timeZone;
  // Örn: "Europe/Istanbul" → kısa offset "UTC+3"
  const tzShort=(()=>{
    try{
      const s=new Date().toLocaleTimeString('en',{timeZoneName:'short'});
      return s.split(' ').pop(); // "GMT+3" veya "EST" gibi
    }catch(e){return tzFull.split('/').pop()}
  })();
  tzEl.textContent=tzShort;
  tzEl.title=`Tam zaman dilimi: ${tzFull}`;

  function tick(){
    const now=new Date();
    timeEl.textContent=now.toLocaleTimeString('tr-TR',{
      hour:'2-digit',minute:'2-digit',second:'2-digit',hour12:false
    });
  }
  tick();
  setInterval(tick,1000);
})();
