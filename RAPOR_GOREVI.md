Görevin: Bu repodaki, transkripti hazır ama raporu yazılmamış her ses kaydı için kullanıcının örnek raporlarına
uygun bir Word raporu hazırlamak ve Drive'a yüklemek. Çalışma dizini repo köküdür. Tüm komutları buradan çalıştır.

1) Bekleyenleri al: `python pipeline.py pending`
   Çıktı bir JSON listesidir. Her öğede key, source, category, txt, duration, language ve suggested_docx bulunur.
   Liste boşsa "Bekleyen rapor yok." yaz ve bitir.

2) `RAPOR_FORMATLARI.md` dosyasını oku. Örnek raporların metinleri `ornekler/*.txt` dosyalarındadır
   (satır biçimi "stil adı | metin"). Kayda uygun türün örneğini oku. Rapor türleri:
   - İş: Toplantı Raporu, Görüşme Özeti ya da (İngilizce kayıtsa) Meeting Minutes
   - Özel: Tefsir Dersi Raporu ya da içeriğe uygun bir ders/sohbet raporu
   - Doktora ve diğerleri: rehberin 5. bölümü
   Bölüm kurgusunu, başlık üslubunu, ayrıntı düzeyini ve uzunluğu örneğe göre ayarla.

3) Listedeki HER öğe için sırayla:
   a) "txt" dosyasının TAMAMINI oku. Uzunsa parça parça oku, hiçbir kısmı atlama.
   b) Raporu kaydın dilinde yaz. "language" "en" ise İngilizce yaz ve JSON'a `"lang": "en"` koy.
   c) Doğruluk kuralları:
      - Transkriptte olmayan bilgi, isim, rakam ve tarih ekleme.
      - Transkript konuşmacıları ayırmaz. Sorumlusu belli olmayan aksiyona "Belirtilmedi" yaz.
      - Tarihi dosya adındaki YYMMDD'den al (ör. "Ses 261014_132520" → 14 Ekim 2026). Göreli tarihleri buna göre hesapla.
      - Bariz ses tanıma hatalarını düzelt ve bunu "note" alanında belirt. Her raporda "note" bulunsun.
   d) İçerik JSON'unu `report_builder.py` dosyasının başındaki şemaya göre `tmp/<key>.json` dosyasına yaz.
      Sonra şunu çalıştır:
      `python report_builder.py tmp/<key>.json "<suggested_docx>"`
   e) Raporu doğrula. Şu komutun yazdığı sayı 0'dan büyük olmalı:
      `python -c "import docx,sys; print(len(docx.Document(sys.argv[1]).paragraphs))" "<suggested_docx>"`
   f) Raporu işaretle ve yükle:
      `python pipeline.py mark-reported "<key>" "<suggested_docx>"`
      Çıktıdaki "upload_error" doluysa not al.

KISITLAR
- `pipeline.py run`, `git` ve ağ komutları çalıştırma. state.json, config.json ve .py dosyalarını düzenleme.
  Durum yalnızca `mark-reported` ile değişir.
- Daha önce raporlanmış bir kaydı yeniden işleme.

Son olarak kısa bir Türkçe özet yaz: yazılan raporlar (kategori/dosya), hatalar.
