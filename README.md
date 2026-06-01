# PanelOptimizer v30 Modular Edition

Bu paket iOS Safari ve statik hosting için modüler hale getirildi.

## Yapı
- `index.html`: Ana sayfa
- `css/style.css`: Tüm stiller
- `js/early-globals.js`: Erken global değişkenler
- `js/phase1-mode.js`: Basit/Pro mod motoru
- `js/app-core.js`: Ana optimizasyon uygulaması
- `js/lazy-loader.js`: XLSX/PDF gibi opsiyonel kütüphane yükleyiciler
- `manifest.json`: PWA manifest
- `service-worker.js`: Offline cache
- `icons/`: iOS/PWA ikonları

## iPhone'da en stabil kullanım
1. Dosyaları bir hostinge yükle: Cloudflare Pages, GitHub Pages, Vercel veya kendi hostingin.
2. Safari'de linki aç.
3. Paylaş > Ana Ekrana Ekle.

## Yerelde test
Terminalde klasör içinde:

```bash
python3 -m http.server 8000
```

Sonra:

```txt
http://localhost:8000
```

Not: Service worker `file://` üzerinde çalışmaz; yayın URL'si veya localhost gerekir.
