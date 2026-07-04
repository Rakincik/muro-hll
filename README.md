# MURO - Mevzuat Adam (MVZ) Veri Aktarım Scripti

Bu depo, Mevzuat Adam (mvz.muro.click) sistemi için Okinar'dan çekilen kullanıcı ve ders kayıtlarının MURO LMS veritabanına otomatik olarak aktarılması için gerekli scriptleri barındırır.

## Dosyalar

* `okinar_scraper_v2.js`: Okinar panelinde tarayıcı konsolunda (F12) çalıştırılacak veri kazıyıcı script. Alt grup / üst grup hiyerarşisini destekler.
* `mevzuatadam.xlsx`: Scraper'dan alınan verilerin yapıştırılacağı Excel şablonu (Sayfa1 ve Sayfa2 sekmeleri içerir).
* `import_data.py`: Verileri Excel'den okuyup, grupları, dersleri ve öğrencileri otomatik olarak MURO LMS veritabanına aktaran Python scripti.
* `check_courses.py`: Aktarılan verilerin durumunu hızlıca kontrol eden araç.

## Nasıl Çalıştırılır?

1. Gerekli kütüphaneleri yükleyin:
   ```bash
   pip install -r requirements.txt
   ```

2. Aktarım scriptini `Dry-Run` (Sadece Okuma ve Test) modunda çalıştırın:
   ```bash
   python import_data.py --file mevzuatadam.xlsx
   ```
   *(Hata alırsanız `--host 172.19.0.3` gibi bir IP ekleyerek test edebilirsiniz).*

3. Her şey düzgün görünüyorsa, veritabanına yazmak için `--execute` bayrağı ile çalıştırın:
   ```bash
   python import_data.py --file mevzuatadam.xlsx --execute
   ```
