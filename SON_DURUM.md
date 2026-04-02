# Son Durum ve Gelistirme Rehberi (Fur Agent)

Bu dosya, projede degisiklik yaparken once neyin korunacagini netlestirmek ve
"bir yeri duzeltirken baska bir yeri bozma" riskini azaltmak icin tutulur.

## 1) Guncel Durum
- Aktif branch: `funnel_main`
- `main` ve `funnel_main` GitHub ile senkron.
- Son ortak commit: `e8815c1`

## 2) Kritik Dosya Politikasi

### Sablon (Git'te kalacak tek Excel)
- `Hole_Pool_Otomatik_Analiz_v1_US_20260325-20260331_20260401_212238.xlsx`

### Git'te kalmayacaklar
- Diger tum `*.xlsx`
- Tum `veri-*.csv`
- `debug-fb19ec.log`
- `__pycache__/`, `*.py[cod]`

## 3) Kodda Korunmasi Gereken Kurallar

### Sablon secimi
- `process_data.py` icinde sabit sablon adi:
  - `WORKBOOK_TEMPLATE_XLSX`
- Script bu dosyayi `SCRIPT_DIR` icinde bulamazsa hata verip cikar.

### And / IOS hesap mantigi
- `T` kolonu: rewarded / complete
- `U` kolonu: owned coin / complete
- `U1`: sadece baslik varsa yazilir (otomatik X/24 olusturma yok)
- `Owned Coin` ve `Avg. Coin` icin 1. satira ortalama yazilir.

### ARPU / Av.Rw
- Referans sayfalar kopyalanir (grafiklerle birlikte).
- U/V/W kolonlari max level'e gore formullenir.
- ARPU/Av.Rw icin ana analizden ayri level limiti sorulur.
- Grafik serileri secilen limite gore uzatilir.

## 4) Is Akisi (Bozmadan Ilerleme)
Her degisiklikte bu sirayla ilerle:

1. Degisiklik amacini tek cümle yaz.
2. Etkileyecegin dosyalari net listelerle sinirla.
3. Once `process_data.py` kritik kurallarini (yukaridaki bolum) kontrol et.
4. Degisikligi yap.
5. En az su kontrolleri yap:
   - Script syntax (compile) gecerli mi?
   - And/IOS T-U-U1 mantigi degismis mi?
   - ARPU/Av.Rw U-V-W ve chart range mantigi bozulmus mu?
6. Commit message'da "neden" bilgisini yaz.
7. Bu dosyada "Degisiklik Kaydi" bolumunu guncelle.

## 5) Degisiklik Kaydi (Template)
Her is bittiginde bu formati ekle:

---
- Tarih:
- Branch:
- Amac:
- Etkilenen dosyalar:
- Davranis degisimi:
- Dogrulama:
- Risk/Not:
- Commit:
---

## 6) Son Kayit
- Tarih: 2026-04-02
- Branch: `funnel_main`
- Amac: Sablonu sabitlemek, Excel/CSV gundelik dosyalarini Git disinda tutmak,
  ARPU/Av.Rw surecini ve T-U-U1 kurallarini stabilize etmek.
- Etkilenen dosyalar: `process_data.py`, `.gitignore`, bu dosya.
- Davranis degisimi:
  - Sablon tek dosya adindan okunuyor.
  - Diger Excel dosyalari gitte tutulmuyor.
  - ARPU/Av.Rw icin ayri level limiti var.
- Dogrulama: Git branch sync, syntax kontrolu, chart/formul akisinin korunmasi.
- Risk/Not: Sablon dosya adi degisirse `WORKBOOK_TEMPLATE_XLSX` da guncellenmeli.
- Commit: `e8815c1`
