ORANLAB Android - çevrimdışı kaynak proje

ÖNEMLİ: Bu ZIP tek başına APK değildir. APK'yı bulut ortamında derlemek için GitHub Actions iş akışı eklendi.

Proje hedefleri:
- İnternet/API isteği olmadan yerel analiz
- XLSB, XLSX ve CSV dosyalarını okuma hedefi
- Maç ve bahisle ilgili sütunları otomatik tanıma
- Metin, !hariç, >, >=, <, <= ve 1.5..2.0 aralığı filtreleri
- MS/İY 1-X-2, KG, 2.5 Alt/Üst ve en sık skor yüzdeleri

APK üretmek için (bilgisayarda Android SDK kurmadan):
1. github.com adresinde hesabınla giriş yap.
2. Yeni bir repository oluştur (ör. oranlab-android). Public veya Private olabilir.
3. ZIP'i bilgisayarında bir klasöre çıkar.
4. Çıkarılan klasörün içindeki tüm dosya ve klasörleri (özellikle .github klasörünü) GitHub deposuna yükle. GitHub web arayüzünde Add file > Upload files seçeneğini kullanabilirsin.
5. GitHub deposunda Actions sekmesini aç.
6. Soldan "Build ORANLAB APK" iş akışını seç ve Run workflow düğmesine bas.
7. İşlem başarıyla biterse tamamlanan çalışmayı aç, Artifacts bölümünden ORANLAB-Android-APK dosyasını indir.
8. ZIP artifact'ini aç; içindeki APK'yı Android telefona gönderip kur.

Derleme ilk seferde birkaç dakika sürebilir ve SDK/NDK bileşenlerini indirir. Bu iş akışı APK derlemeyi dener; başarılı olacağı garanti değildir.

Test durumu ve sınırlar:
- Bu çalışma ortamında Android SDK/NDK ve Buildozer bulunmadığı için APK burada derlenmedi.
- 486 sütunlu XLSB'nin başlıkları görüldü; tam dosyanın Android'de baştan sona açılması ve performansı henüz doğrulanmadı.
- Android dosya seçici izinleri, XLSB okuma ve analiz sonuçları gerçek cihazda test edilmelidir.
- .xls doğrudan desteklenmez; .xlsx veya .csv olarak dışa aktarılmalıdır.
