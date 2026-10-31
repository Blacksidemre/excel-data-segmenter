# Tek dosyalık kurulum

Gerekli: Windows üzerinde Microsoft Excel 2024 LTSC masaüstü, bu dosya için VBA makrolarının çalışmasına izin veren güvenlik politikası. 32 ve 64 bit uyumluluğu hedeflenir; Windows API bildirimi kullanılmaz. Excel Web ve Mac desteklenmez.

GitHub'daki `ExcelOps.xlsm` dosyasını indirin. Başka bir dosya kopyalamayın. İnternet kaynaklı makro engeli varsa dosyayı kapatın; Windows dosya Özellikleri içinden, mevcutsa, Engellemeyi kaldır seçeneğini uygulayın. Kurumsal engelleri aşmaya çalışmayın. Dosya dijital olarak imzalanmamıştır.

Menü otomatik açılır; alternatif giriş Alt+F8 → ExcelOpsAc. Kurulum programı, ActiveX kaydı, COM eklentisi, Python veya Power Query kurulmaz. Microsoft Forms ve Scripting.Dictionary Windows/Office bileşenleridir. **“VBA proje nesne modeline güven” izni gerekmez**: uygulama kendi kodunu değiştirmez.

Ayarlar → Uygula değerleri bu `.xlsm` dosyasına kaydeder. Dosya salt okunursa ayarlar yalnız oturumda tutulur. Log sayfasını kalıcılaştırmak için uygulama kitabını kaydedin. Kaynak veri dosyası ayrı seçilir; diskteki kaynak üzerine hiçbir işlem kaydedilmez.

Bu dağıtımın Excel 2024 üzerinde gerçek çalışma testi henüz yapılmadı. Alt+F8 → ExcelOpsSelfTest, Excel'de test verilerini `%TEMP%\ExcelOps_KabulTesti` altında oluşturur. Test sonuçlarını Log sayfasında kontrol edin.
