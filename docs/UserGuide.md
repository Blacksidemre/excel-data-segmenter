# Kullanım

## Kaynak ve ortak ayarlar

Kaynak seç ile Excel dosyasını, Kayıt konumu ile mevcut ana klasörü seçin. Ayarlar ekranında veri sayfasının tam adını ve başlık satırını girin. Sütunlar `M`, `AA` veya `13` şeklinde girilebilir. Varsayılan veri sayfası `Veri`, ayırma/ad sütunu `M`, tutar `N`, tarih `O`, kayıt anahtarı `A`'dır. Bunları kendi kaynağınıza göre değiştirin.

Başlık altındaki **tüm kullanılan satırlar** veri sayılır. Dip toplam, ara başlık ve birleşik veri hücresi kullanmayın. Boş satırlar kalite kontrolünde hata sayılabilir. Koruma, parola ve kaynak dış bağlantıları içeren dosyalarda önce sade bir kopyayla doğrulayın. Kaynak `.xlsx`, `.xlsm`, `.xlsb`, `.xls` olabilir.

Anahtarlar baş/son boşluk ve kesilemeyen boşluklardan temizlenir; büyük/küçük harf farkı Windows yerel karşılaştırmasına göre yok sayılır. Sayısal `1` ve metin `1` aynı grup olabilir. Kimlikleri baştaki sıfırları koruyacak şekilde gerçek metin olarak saklayın. `*`, `?`, `~` ayırmada joker karakter değildir.

## Dosya Ayır

Ayırma sütunu grupları belirler; ad sütunu aynı grupta tek bir ada karşılık gelmelidir. Farklı adlar varsa işlem durur. Boş grup anahtarları dışarıda bırakılır. Her grubun başlık dışı diğer satırları tam kopyadan silinir; kaynak değiştirilmez. Mevcut filtreler ve gizli satırlar gruplama kapsamını daraltmaz.

Dosya formatı `{ad}`, `{grup}`, `{sira}`, `{tarih}` alanlarını destekler. Örnekler: `{ad}`, `Masraf_{sira}`, `{grup}_{tarih}`. Ön/son ek ayrı alanlardır. Geçersiz dosya karakterleri `_` olur; temizlemeden sonra çakışan adlara sayı eklenir. Klasör adı sizin seçiminizdir. Altına her işlem için benzersiz zaman klasörü açılır.

Varsayılan `Kaynak` biçimi kaynak türünü korur. `xlsx` gerçek SaveAs dönüşümüdür; makrolar çıktıdan kaldırılır. Makro tabanlı kullanıcı işlevlerine bağlı formüller bu çıktıda çalışmaz.

Diğer sayfalardaki **tüm bilgiler ve gizli sayfalar** kopyalanır. Seçili sayfadaki satır silmeleri Excel'in formül referanslarını günceller; silinen hücrelere bağlı formüllerde `#REF!` oluşabilir. Formüllerin kaynak metninin ve sonuçlarının hiç değişmeyeceği vaat edilmez. Var olan filtre ölçütleri korunur; farklı grup çıktısında bu ölçütler bütün satırları gizleyebilir. Görünür kayıt yoksa filtreyi temizleyin.

## Renklendir

`RenkDosya=Hayir` varsayılanıdır: dosya oluşturmadan salt okunur açık kaynakta renk önizlemesi sağlar; kalıcılaştırmak isterseniz Farklı Kaydet kullanın. Sonraki işlemden önce bu önizleme kitabını kapatın. `Evet` ayrı çıktı oluşturur. Pastel ve Canli paletleri grup başına farklı RGB üretir. Benzersiz renkler insan gözüyle ayırt edilebilirlik garantisi değildir; koşullu biçimlendirme normal dolguyu örtebilir.

## Özet Çıkar

Grup ve tutar zorunludur. Tarih sütunu boş bırakılırsa tarih özeti üretilmez. Doluysa tüm kayıtlar gerçek Excel seri tarihi içermelidir. 1904 tarih sistemi normalize edilir. Metin olarak yazılan tarihler reddedilir. Özet adet, toplam, ortalama, en büyük kayıt ve ilk/son tarih üretir. Hatalı tutar/anahtarlar sessizce atlanmaz.

## Mutabakat Yap

A kaynak dosya, B Ayarlar → İkinci dosya seç ile seçilen dosyadır. İki sayfa/başlık/anahtar/tutar sütunları ayrı ayarlanır. Aynı anahtarın kayıtları toplanır; mükerrer sayısı korunur ve durum sütununda belirtilir. Sonuç farkları A−B yönündedir. Adet farkı sıfır ve tutar farkının mutlak değeri tolerans dahilindeyse eşleşir. Tolerans Excel'in bölgesel sayı biçiminde girilir.

Eksik A anahtarları için B'deki boşluk, nokta, alt çizgi, eğik çizgi ve tireleri kaldırılmış anahtarlar önerilebilir. Birden fazla aday ayrıca belirtilir. Öneriler gerçek eşleşme sayılmaz. Karşılaştırma anahtar bazında özet mutabakattır; aynı anahtara ait tek tek satırlar birbirine eşleştirilmez. Ayrı Kontrol Toplamlari sayfası adet/tutar toplamlarını gösterir.

## Veri Kontrolü

Zorunlu sütunları `A;M` biçiminde girin. Kayıt anahtarı mükerrer/boş kontrolü için kullanılır. Tutar ve tarih sütunlarını boş bırakarak ilgili kontrolleri devre dışı bırakabilirsiniz. Beklenen sütun ve `Garanti;Akbank;Denizbank` listesi birlikte tanımlanır. Tarih kontrolü gerçek Excel seri değerini denetler; sayısal bir tutarın yanlışlıkla tarih sütunu seçilmesini anlamsal olarak ayırt edemez.

Ayrı kopyada hatalı satırlar işaretlenir; ExcelOps_Kalite sayfası kaynak satır numarası ve tüm bulguları içerir. Kaynak çalışma kitabı değiştirilmez.

## Akıllı Filtre

Üç isteğe bağlı kural vardır. Kurallar VE; Dahil listesindeki değerler VEYA ilişkisiyle çalışır. Dahil/Haric tam değer eşleştirir; joker karakter uygulamaz. Boş liste değeri boş hücreyi ifade eder. Tutar aralığı `100;500`, tarih aralığı `2026-01-01;2026-12-31` biçimindedir. Son tarih gününün tüm saatleri dahildir.

Kopyada eski filtre/manuel gizli veri satırları sıfırlanır. Gizli ExcelOps_Filtre sütununa maskeler yazılır ve AutoFilter uygulanır; diğer kayıtlar silinmez. Excel tablo nesneleri bu işlemde desteklenmez; normal aralık kullanın. Ayırma, özet ve kalite özellikleri filtrelenmiş görünüm yerine bütün veri aralığını işler.

## İptal, hata ve performans

Durdur veya Esc iş akışının denetim noktalarında işlemi keser. Excel'in devam eden tek bir SaveAs/Delete çağrısı bitmeden iptal gerçekleşmeyebilir. Tamamlanmış dosyalar saklanır, açık geçici kopya kapatılır, geçici dosya kaldırılır. Başarıdaki dosya manifesti `_Dosya_Listesi.xlsx` olur; yarım kalan işlerde Log'u ve işlem klasörünü kontrol edin.

Uygulama ekran güncelleme, uyarı, olay, hesaplama, makro açma güvenliği ve iptal ayarlarının önceki değerlerini geri yükler. Log başarısızlıkları sessizce işlemi başarılı göstermez; günlük yazılamazsa VBA Immediate penceresine bildirilir.

Diziler ve Dictionary Excel hücre erişimini azaltır. Ancak 1.000 grup 1.000 tam çalışma kitabı kopyasıdır. Dağınık grup satırlarının silinmesi, formüller ve biçimler iş süresini belirler; disk alanı yaklaşık grup sayısı × kaynak büyüklüğü kadar olabilir. Performans hedefleri gerçek Excel'de henüz ölçülmemiştir.
