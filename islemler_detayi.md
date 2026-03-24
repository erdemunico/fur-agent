# `process_data.py` İşlem Detayları

Bu belge, `process_data.py` scriptinde adım adım gerçekleştirilen işlemleri detaylı bir şekilde açıklamaktadır. 

Scriptin asıl amacı, oyun analitik verilerini (Excel formatında) ve gelir verilerini (CSV formatında) okuyarak işlemek, ardından bunları hazır bir şablon Excel dosyasına ("Copy of Hole_Pool...") Android ve iOS ayrımlarıyla yazdırmaktır.

## 1. Hazırlık ve Hata Önleme (Monkeypatch)
Script ilk olarak `pandas`, `openpyxl` ve diğer gerekli kütüphaneleri dahil eder. 
- Excel dosyasının şablonunda yer alan "Koşullu Biçimlendirme" (Conditional Formatting) kaynaklı "invalid XML" hatalarını es geçebilmek (bypass) amacıyla `openpyxl` kütüphanesinin arka planına ufak bir yama (monkeypatch) yapılır. 

## 2. Sabit Değişkenler ve Sütun Haritaları
Scriptte, hesaplanan değerlerin şablon (template) Excel'de hangi sütun numaralarına yazılacağını belirten iki adet ana sözlük (dictionary) tanımlanmıştır:
- **`GENEL_COL_MAP`:** `And` (Android) ve `IOS` sekmelerine yazılacak genel oyuncu değerleri (Active users, avarage_attempt, avg_thinktime, vb.).
- **`BOOSTER_COL_MAP`:** `Boosters_And` ve `Boosters_IOS` sekmelerine yazılacak booster (güçlendirici) kullanımı ve sahipliğiyle ilgili değerler.

## 3. Kullanıcı Girdisi Alma
Script çalıştırıldığında, konsoldan kullanıcıya kaçıncı seviyeye (level) kadar analiz yapılacağı sorulur (Örn: 300). Kullanıcı herhangi bir sayı girmezse varsayılan olarak 300 atanır.

## 4. Kaynak Veri Dosyalarının Taranması
Dosya sistemindeki tüm `.xlsx` dosyaları taranır. İçerisinde `# hole-pool` ifadesi geçen geçerli veri dosyaları belirlenir. Bu işlem sırasında şablon dosyaları (Örn: "Copy of Hole_Pool...") ve daha önce kaydedilmiş çıktı dosyaları atlanır.

## 5. Verileri Okuma, Analiz Etme ve Parçalama
Dosyalar iki temel platform (Android ve iOS) için ayrıştırılır:
- **`detect_funnel_type`:** Okunan her Excel tablosunun meta verisine bakarak verinin türünü ('Start', 'Complete', 'Fail', 'Tryagain', 'ADs') belirler.
- **`read_platform_data`:** "Operating system" (İşletim sistemi) ve "Level" etiketlerini arayarak tablo yapısını anlar. Düz liste şeklindeki yapılar veya Android/iOS olarak yan yana gruplanmış çoklu tablo yapıları dinamik olarak uygun veri çerçevelerine (DataFrame) çevrilir.

## 6. Verilerin Birleştirilmesi ve Seviyelere Göre Doldurulması
Huni türlerine göre toplanan veriler birleştirilir.
- **`safe_concat`:** Aynı adla tekrar eden sütunlar varsa (örn: Start_1, Start_2), sütunları çakışmayacak biçimde formatlar ve DataFrameleri birleştirir.
- **`format_and_fill_levels`:** Elde edilen DataFrameler'deki verileri seviye numaralarına (Level) göre gruplar, metrikleri toplar (sum). 1. seviyeden, `max_level` değişkeni ile belirlenen hedefe (Örn: 300) kadar olası boşluklar 0 değeriyle doldurulur.

## 7. CSV Dosyalarından Gelir (In-App Purchase) Verisinin Alınması
Tüm `veri-*.csv` dosyaları taranır. 
- Bu CSV dosyalarından **Level**, **USD Price** ve **Store Name** satırları eşleştirilir. 
- 'Google' ve 'AppStore' mağazalarına göre ayrılarak her seviyenin gelir toplamları elde edilir.

## 8. Verilerin Şablon Excel'e Yazılması
Daha önceden belirlenen şablon dosya ("Copy of Hole_Pool*.xlsx") açılır ve Android ile iOS verileri hesaplanan haritalamalara göre sayfalara yazdırılır:
- Orijinal şablondaki eski veriler temizlenir (Ancak formüllerin tutulduğu veya değiştirilmemesi istenen K (11.) sütun atlanır).
- 'And', 'IOS', 'Boosters_And', 'Boosters_IOS' sayfalarındaki ilgili sütunlara işlenen toplam metrikler ve toplam "Total $" gelir verileri basılır.

## 9. Formül Kaydırma (Formula Shifting)
Şablon dosyasının 3. satırındaki otomatik formüller bulunur. Analiz yapılan diğer seviyelerdeki (4. satır, 5. satır vb.) hücrelere doğrudan formül adı yazılırken, `shift_formula` fonksiyonu sayesinde içindeki hücre referansları (Örn: N3 -> N4 olarak) dinamik olarak birer birer kaydırılır (Regex ile). Böylece tüm döküm boyunca Excel içi hesaplamalar korunur.

## 10. Final
İşlemler tamamlandıktan sonra sonuç, `Hole_Pool_Otomatik_Analiz_v1.xlsx` adlı yepyeni bir Excel dosyası olarak kaydedilir.
