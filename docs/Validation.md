# Doğrulama kaydı — 31 Ekim 2026

## Bu ortamda geçen kontroller

- Node.js 24 ile kilit dosyasından temiz geliştirici bağımlılığı kurulumu: `npm ci --prefix tools --ignore-scripts`.
- `.xlsm` tekrar üretimi: `npm run build --prefix tools`.
- `npm test --prefix tools`: **33 paket kontrolü geçti**. Makrolu içerik türü, VBA ilişki kaydı, SHA-256, MS-OVBA dizin kayıtları, 19 bileşenin gömülü kaynakla birebir eşitliği, iki UserForm deposu, sayfa kod adları, ayar anahtarları ve örneklerin tarih/formül/gizli sayfa içeriği kontrol edildi.
- ANTLR 4.13.2 ve antlr/grammars-v4 VisualBasic6 grameri ile 19 kaynak bileşeninde **0 sözdizimi hatası**. Designer başlıkları ve Attribute metaverisi bu kontrolden çıkarıldı. Bu gramer VBA ile ortak Visual Basic sözdizimini denetler; Excel nesne modelindeki üyeleri ve başvuruları çözmez.
- Makrolar kapalı, ayrı LibreOffice profiliyle `.xlsm` yüklenip geçici `.xlsx` kopyası oluşturuldu. Bu sadece çalışma sayfası paketi içe aktarma kontrolüdür; dağıtılan `.xlsm` bu dönüşümle değiştirilmedi.

Paket kontrolleri `tests/package.test.cjs` içinde depoya dahil edilmiştir. Kaynak tutarlılığı kontrolleri derlenmiş VBA yerine kaynak kapsayıcısını doğrular. Dosya SHA-256 özeti `ExcelOps.xlsm.sha256` içindedir.

## Henüz çalıştırılmayan gerekli Excel kabul testleri

**Microsoft Excel 2024 LTSC bu Linux ortamında yoktur.** Aşağıdakilerin geçtiği iddia edilmez:

- VBA projesinin Excel'de yüklenmesi/derlenmesi ve MSForms başvurularının çözülmesi.
- Ana menü/ayar UserForm'larının açılması, düğmelerin işlevi.
- Windows dosya sistemiyle SaveCopyAs/SaveAs/filtre/satır silme davranışı.
- Kaynak özelliklerinin gerçek Excel dosyalarında korunması.
- 32/64 bit Excel uyumluluğu ve 100.000 satır / 1.000 grup performansı.

## Excel içindeki kabul testi

Alt+F8 → **ExcelOpsSelfTest** → Çalıştır. Ek dosya gerektirmez. Sentetik kaynaklar Windows TEMP altında oluşturulur; mevcut ayarlar saklanıp geri yüklenir. Log sayfasına PASS/FAIL kaydı yazılır. Test çıktıları inceleme için korunur.

16 doğrulama: boşluk temizlenerek ayırma, çok gizli sayfa, formül, sütun genişliği, filtre yapısı, yıldız karakterli grup, doğru gruptaki kayıtlar, kaynak satırlarının korunması, iki özet toplamı, aynı/farklı grup renkleri, mükerrer/negatif tutar denetimi, tutar aralığı sınırları, mutabakat tutar farkı ve mükerrer anahtar işaretlemesi.

## Elle sınanacak ek durumlar

1. `.xlsm` ve `.xlsb` kaynakta yerel biçim; ardından gerçek `.xlsx` dönüşümü.
2. `A/B` ve `A:B` gibi aynı ada temizlenen gruplar; mevcut dosyalara dokunulmadığı.
3. Boş/hatalı ayırma anahtarları, tek satır veri, tümü boş anahtar, farklı başlık satırı, alternatif ad sütunu.
4. Mükerrer satırlar, ikinci dosyada eksik/fazla anahtar, sıfır/negatif tutar, eşik toleransı, olası eşleşme çakışmaları.
5. 1900/1904 tarih sistemi, son gün içindeki saat, ters tarih aralığı, sayısal olmayan tutar.
6. Korumalı/çok geniş sayfa, birleşik hücre, kaydetme izni olmayan klasör, disk yetersizliği ve kullanıcı iptali.
7. İşlem öncesi manuel/otomatik hesaplama ve açık/kapalı uyarı/olay durumlarının geri gelmesi.
8. 100.000 satır, 1.000 grup için hem sıralı hem dağınık gruplarla süre, bellek ve disk ölçümü. Renklerin benzersizliği ile gözle ayırt edilebilirliği ayrı değerlendirilmelidir.
