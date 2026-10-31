# Mimari

## Çalışma akışı

`ThisWorkbook.Workbook_Open → ExcelOpsAc → frmMain → RunOperation → ilgili motor`

`frmSettings` ayarları uygulama içindeki Ayarlar sayfasına yazar. `ButtonHandler` dinamik MSForms düğmelerinin Click olaylarını tutar. `FormUI` ortak kontrol görünümünü oluşturur. VBA projesine programatik erişim veya kendini değiştiren kod yoktur.

| Bileşen | Sorumluluk |
|---|---|
| ExcelOps.bas | İşlem yaşam döngüsü, kaynak açma, ortak hata/iptal temizliği |
| Utilities.bas | Excel durumunu saklama/geri yükleme, dosya adı/yolu, yerel biçimde kopya, log, tarih/tutar doğrulama |
| Settings.bas | Kitap içi yapılandırma ve dosya/klasör seçiciler |
| FileSplitter.bas | Tek geçişte grup→satır indeksleri; tamamlayıcı satır aralıklarını aşağıdan silme; çıktı manifesti |
| Highlighter.bas | Grup→renk sözlüğü, ardışık aynı grupları toplu boyama |
| Reporting.bas | Tek geçişte adet/toplam/ortalama/maksimum/tarih aralığı |
| Reconciliation.bas | İki anahtar indeksi, tam dış birleşim, adet/tutar farkları, normalleştirilmiş adaylar |
| DataQuality.bas | Mükerrer indeksi ve satır başına birleştirilmiş kalite bulguları |
| SmartFilter.bas | Üç koşuldan dizi maskesi üretme ve yardımcı sütun üzerinden AutoFilter |
| SelfTests.bas | Microsoft Excel'de çalıştırılan sentetik kabul testleri |
| ThisWorkbook / Sheet*.cls | Kitap ve sayfa belge modülleri |

## UserForm ekranları

Ana ekran: başlık, kaynak dosya özeti, Kaynak seç / Ayarlar / Kayıt konumu; iki satırda altı işlem; Durdur / Kılavuz / Kapat. Ana menü modeless çalışır; `Busy` yeniden girişleri engeller. İptal düğmesi `DoEvents` denetim noktalarında işlenir.

Ayarlar: kaydırılabilir etiket + giriş alanları; seçenek alanlarında ComboBox; üstte ve altta Uygula. İkinci dosya ayrı dosya seçiciyle belirlenir. Sütun/sayı/aralık doğrulaması işlem başlamadan veya ilgili motor veri okurken yapılır. Form kaynaklarının `.frm` dosyaları kontrol düzenini çalışma anında oluşturur; klasik designer denetimleri içeren `.frx` dosyasına ihtiyaç yoktur. Dağıtılan kitapta iki boş designer CFB deposu zaten gömülüdür.

## Veri bütünlüğü

Kaynak `ReadOnly=True`, `UpdateLinks=0`, `AutomationSecurity=3` ile açılır; makroları çalıştırılmaz. Yeni dosya `.SaveCopyAs` ile doğru yerel uzantıda kopyalanır. Biçim dönüşümü yalnız çıktı üzerinde SaveAs ile yapılır. Kaynak hiçbir zaman Save/SaveAs ile kaydedilmez. Açık kaynak dosyaları reddedilerek kaydedilmemiş değişikliklerle belirsizlik önlenir.

Kaynak ActiveSheet/ThisWorkbook karışıklığı yoktur: ThisWorkbook yalnız uygulama; veri kitabı açıkça seçilen dosyadır. Kaynak kitap imzaları, dış veri bağlantılarının otomatik davranışı ve formül sonuçlarının değişmezliği garanti edilmez. Seçili sayfada satır silmek Excel referanslarını doğal olarak değiştirir.

`Application` durumları tek bir yapı içinde tutulur ve eski değerlerine geri döndürülür; hesaplama koşulsuz Automatic yapılmaz. Hata sırasında açık geçici kopya ve kaynak kapatılır. Log'da kod, açıklama, zaman ve işlem bulunur. Tamamlanmış çıktılar geri alınmaz.

## Performans ve sınırlar

Grup/tutar/tarih sütunları Value2 dizilerine alınır. Ayırma indeksi O(n) kayıt tutar; sözlük ve Collection Windows/VBA içinde çalışır. Her grup için tamamlayıcı satır aralıkları hesaplanır, 128 aralığa kadar Union ile aşağıdan silinir. Bu, kitap kopyalama ve Excel satır silme maliyetini ortadan kaldırmaz. Dağınık gruplarda silme sayısı çok yüksek olabilir. Büyük veri performansı ölçülmemiştir.

Raporlama ve mutabakat indeksleri O(n+m); kalite kontrolü tüm gerekli veri dikdörtgenini belleğe alır. Çok geniş 100.000 satırlık veri, özellikle 32 bit Excel'de bellek sınırına ulaşabilir. Finansal tutarlar Excel'in Double sayı modelini kullanır; mutabakat için kullanıcı toleransı uygulanır. Sayısal hassasiyeti 15 basamağı aşan kimlikleri metin saklamak gerekir.

## Paketleme

`tools/build.cjs` ExcelJS ile sayfaları, MS-OVBA kaynak kapsayıcısını ve boş UserForm depolarını oluşturur. Modüller Windows-1254 olarak gömülür; kaynak dosyaları UTF-8 tutulur. Kaynak modül ofsetleri sıfırdır; eski p-code ve SRP önbellekleri dahil edilmez. Excel projeyi yüklerken derlemelidir. Bu nedenle kaynak ve kapsayıcı testlerine ek olarak gerçek Excel kabul testi zorunludur.

Son kullanıcının bilgisayarına yalnız ExcelOps.xlsm gider. Node.js ve geliştirici bağımlılıkları sadece bu depoyu geliştirenler içindir. Kullanıcı isteğine uygun olarak uygulamada Python, Power Query veya harici servis kullanılmaz.
