# Son Durum ve Gelistirme Rehberi (Fur Agent)

Bu dosya, projede degisiklik yaparken once neyin korunacagini netlestirmek ve
"bir yeri duzeltirken baska bir yeri bozma" riskini azaltmak icin tutulur.

## 1) Guncel Durum
- Yerel proje klasoru: `C:\Users\erdem\Desktop\cursor projects\fur-agent`
- GitHub repo: `git@github.com-unico:erdemunico/fur-agent.git`
- Aktif branch: `funnel_main`
- `main` ve `funnel_main` GitHub ile senkron.
- Son ortak commit: `d8480a7`

## 2) Kritik Dosya Politikasi

### Sablon (Git'te kalacak tek Excel)
- `Hole_Pool_Otomatik_Analiz_v1_US_20260325-20260331_20260401_212238.xlsx`

### Git'te kalmayacaklar
- Diger tum `*.xlsx`
- Tum `veri-*.csv`
- `*detailed_level_table*.csv` (Google / iOS purchase export)
- `debug-fb19ec.log`
- `__pycache__/`, `*.py[cod]`

## 3) Kodda Korunmasi Gereken Kurallar

### Sablon secimi
- `process_data.py` icinde sabit sablon adi:
  - `WORKBOOK_TEMPLATE_XLSX`
- Script bu dosyayi `SCRIPT_DIR` icinde bulamazsa hata verip cikar.
- Cikti dosya adi artik oyuna gore dinamik uretilebilir:
  - `<Game>_Otomatik_Analiz_v1_<Country>_<DateRange>.xlsx`

### Metaveri ve oyun adi
- Veri taramasinda once metaveri kontrol edilir.
- `# <game-name>` etiketi yakalanirsa oyun adi buradan okunur.
- Oyun adinda birden fazla farkli etiket varsa en sik gecen etiket secilir.
- Metaveri yoksa fallback ad: `Hole_Pool`.
- Oyun adi dosya-adina uygunlasir:
  - Bosluk/ozel karakterler `_` ile degisir (`Farm Block Escape` -> `Farm_Block_Escape`).

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

### Purchase export isimlendirme (analytics ekibi)
- `detailed_level_table` CSV icinde platform/store kolonu yok.
- Google ve iOS ayri export oldugunda **dosya adina platform yazilmali**:
  - `google_detailed_level_table_<tarih>.csv` -> And (N sutunu)
  - `ios_detailed_level_table_<tarih>.csv` -> IOS (N sutunu)
- Etiketsiz dosya adlari (yalnizca timestamp) script tarafindan otomatik ayirt edilemez.
- Script etiketsiz dosyalarda konsola uyari ve analytics ekibine iletilecek metin basar.
- Iki etiketsiz dosya varsa gecici olarak kullaniciya sorulur; 3+ etiketsiz dosyada gelir yazilmaz.

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

---
- Tarih: 2026-04-28
- Branch: `funnel_main`
- Amac: Cikti dosya adini oyun metaverisine gore dinamiklestirmek.
- Etkilenen dosyalar: `process_data.py`, `SON_DURUM.md`
- Davranis degisimi:
  - Veri taramasinda metaveri (`# oyun-adi`) artik once kontrol ediliyor.
  - Oyun adi metaveriden okunup normalize edilerek output adina yansiyor.
  - Output algisi `*_otomatik_analiz*` desenine genel hale getirildi.
- Dogrulama: `python -m py_compile process_data.py`, linter temiz.
- Risk/Not:
  - Metaveri formati beklenenden farkliysa oyun adi fallback `Hole_Pool` olur.
  - Cok farkli metaveri yazimlari varsa regex genisletmesi gerekebilir.
- Commit: (bu guncellemeden sonra eklenecek)

---
- Tarih: 2026-06-02
- Branch: `funnel_main`
- Amac: Google ve iOS purchase verisini ayri `detailed_level_table` CSV dosyalarindan okumak.
- Etkilenen dosyalar: `process_data.py`, `.gitignore`, `SON_DURUM.md`
- Davranis degisimi:
  - `*detailed_level_table*.csv` ve `*.xlsx` destekleniyor (`Level`, `Revenue ($)`).
  - Level sirasi dosyada karisik olsa bile level'a gore toplanip siralaniyor; And/IOS N sutununa sirali yaziliyor.
  - Dosya adinda `google`/`android` -> And, `ios`/`appstore` -> IOS.
  - Etiketsiz dosyada konsol uyari + analytics ekibine iletilecek metin.
  - Iki etiketsiz dosya: gecici secim sorusu; 3+ etiketsiz: gelir atlanir.
  - Yedek kaynak: `veri-*.csv` (Store Name ile ayrim).
- Dogrulama: `python -m py_compile process_data.py`; ornek CSV ile And/IOS ayri yukleme testi.
- Risk/Not: Export dosya adinda platform yoksa analytics ekibinden `google_` / `ios_` on eki istenmeli.
- Commit: (asagidaki 2026-06-30 kaydinda)

---
- Tarih: 2026-06-30
- Branch: `funnel_main`
- Amac: Yerel proje klasorunu Desktop `cursor projects` altina tasimak ve GitHub'i guncellemek.
- Etkilenen dosyalar: `SON_DURUM.md`
- Davranis degisimi:
  - Kanonik yerel yol: `C:\Users\erdem\Desktop\cursor projects\fur-agent`
  - Eski kopyalar (`Downloads\fur agent`, `.cursor\projects\fur-agent`) kaldirildi.
- Dogrulama: `git status`, `git push` funnel_main.
- Risk/Not: Cursor'da yeni klasoru ac; script `SCRIPT_DIR` kullandigi icin kod degisikligi gerekmez.
- Commit: (bu guncellemeden sonra eklenecek)
