# ExcelOps

Excel 2024 LTSC (Windows masaüstü) için tamamen VBA tabanlı portföy, masraf ve veri operasyonları uygulaması. Bu proje mevcut `Blacksidemre/excel-data-segmenter` deposunda geliştirilir.

## Yalnızca bir dosya indirin

**[ExcelOps.xlsm dosyasını indir](https://github.com/Blacksidemre/excel-data-segmenter/raw/refs/heads/main/ExcelOps.xlsm)**

Bilgisayarınıza yalnızca bu dosyayı aktarın. `src`, `forms`, örnekler veya geliştirici araçları gerekli değildir. Python, Power Query, eklenti, ayrı kurulum veya VBA editörüne kod yapıştırma gerekmez. Menü, ayarlar, yardım, günlük ve VBA modülleri dosyanın içindedir.

1. Dosyayı Windows masaüstü Excel ile açın. İnternetten indirildiği için engellenmişse kapatıp dosyanın **Özellikler → Engellemeyi kaldır** seçeneğini kullanın. Yalnızca güvendiğiniz bu dosyada makrolara izin verin; genel Excel güvenliğini düşürmeyin. Kurum politikaları engelliyorsa BT yöneticinizle görüşün.
2. Menü otomatik açılır. Açılmazsa **Alt+F8 → ExcelOpsAc → Çalıştır**.
3. **Kaynak seç**, **Kayıt konumu**, **Ayarlar** adımlarını tamamlayın. Kaynak dosya diğer Excel pencerelerinde kapalı olmalıdır.
4. İstediğiniz işlemi başlatın. Ayrıntılı yardım dosyanın **Kılavuz** sayfasında bulunur.

## İşlemler

| İşlem | Sonuç |
|---|---|
| Dosya Ayır | Seçilen sütunun gruplarına göre tam çalışma kitabı kopyaları; ayrı isim sütunu ve ad şablonu |
| Renklendir | Grup başına benzersiz renk; varsayılan olarak dosya oluşturmadan salt okunur önizleme |
| Özet Çıkar | Adet, toplam, ortalama, en büyük tutar, ilk ve son tarih |
| Mutabakat Yap | İki dosyada anahtar başına adet/tutar farkı, eksik/fazla kayıt, mükerrer ve olası eşleşme önerileri |
| Veri Kontrolü | Boş alan, mükerrer, sayısal alan, negatif tutar, tarih ve beklenen değer denetimi |
| Akıllı Filtre | Dahil/hariç listeleri, tarih/tutar aralıkları; en fazla üç kural |

Kaynak salt okunur açılır; değişiklikler kaynak dosyaya kaydedilmez. Her işlem ayrı bir tarih-saat klasörü kullanır. Mevcut çıktı dosyaları üzerine yazılmaz. Ayırma, renklendirme ve veri kontrolünde diğer sayfalar, gizli sayfalar, formüller ve biçimler tam kopyada korunur. **Diğer sayfalardaki kayıtlar ayrıştırılmaz.** Çıktı paylaşılmadan önce bunlar incelenmelidir.

Varsayılan çıktı biçimi kaynağın biçimidir. `xlsx` seçildiğinde önce doğru uzantıyla geçici kopya oluşturulur, sonra Excel `SaveAs` ile gerçek dönüşüm yapılır; VBA içeriği `.xlsx` çıktıda yer almaz.

## Doğrulama durumu

Bu sürüm **Excel üzerinde kabul testi bekleyen bir sürümdür**. Linux ortamında paket yapısı, gömülü VBA/kaynak eşitliği, örnek dosyalar ve bağımsız Visual Basic grameri ile sözdizimi kontrol edilmiştir. Bunlar Microsoft Excel derleme/çalıştırma testinin yerine geçmez. UserForm açılışı, MSForms başvurularının çözülmesi, dosya işlemleri ve 100.000 satır / 1.000 grup performansı Excel 2024 üzerinde henüz doğrulanmamıştır.

Dosyanın içindeki **Alt+F8 → ExcelOpsSelfTest** makrosu Excel'de kabul testlerini çalıştırır; sonuçlar Log sayfasındadır. İlk kullanımda örnek veriyle doğrulayın. [Test kapsamı](docs/Validation.md).

## Kaynak ve geliştirici belgeleri

- [Kurulum](docs/Installation.md) · [Kullanım](docs/UserGuide.md) · [Mimari ve sınırlar](docs/Architecture.md)
- `src/`: VBA iş mantığı, olay sınıfları ve gömülü kabul testleri
- `forms/`: ana menü ve ayar UserForm kaynakları
- `examples/`: ayırma ve mutabakat örnekleri (isteğe bağlı)
- `tools/`: yalnız geliştiriciler için tekrar üretilebilir `.xlsm` paketleyici

Geliştirici derlemesi: Node.js 24 ile `cd tools && npm ci && npm run build && npm test`. Çalışan uygulama yalnızca VBA kullanır; Node.js son kullanıcıya gerekmez. Form denetimleri VBA ile çalışma anında oluşturulur; çalışmak için VBProject erişim izni istemez. Bu Linux paketleyicisi VBA p-code derlemez; Excel kaynak kodu yüklerken derler.

MIT lisansı. Dosya biçimi örneğinden türetilen boş UserForm metaverisinin atıfları [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md) dosyasındadır.
