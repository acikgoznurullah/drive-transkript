# Drive ses kayıtları → transkript + Word rapor (GitHub Actions)

Bu repo, herkese açık Google Drive klasöründeki kategori klasörlerine (İş, Özel, Doktora …) ait `Sesler/`
alt klasörlerine yüklenen yeni ses kayıtlarını her gün **09:00 ve 22:00'de (Türkiye saati)** işler:

1. Kaydı indirir, faster-whisper (small, CPU) ile Türkçe/İngilizce transkript eder (`.txt` + `.srt`).
2. Raporu, örnek raporlara (`ornekler/`) ve `RAPOR_FORMATLARI.md` kurallarına göre Claude Code ile yazar (`.docx`).
   Claude Code, kullanıcının Claude aboneliğiyle çalışır, API ücreti yoktur.
3. Çıktıları Apps Script web uygulaması üzerinden aynı kategorinin `Transkriptler/` ve `Raporlar/` klasörlerine yükler.

Çalışma GitHub'ın sunucularında olur; kişisel bilgisayarın açık olması gerekmez.

## Dosyalar

| Dosya | Görev |
|---|---|
| `.github/workflows/transkript.yml` | Zamanlama (cron) ve çalışma adımları |
| `pipeline.py` | Drive tarama, indirme, transkript, yükleme; durum `state.json` |
| `transcribe_fast.py` | faster-whisper transkripsiyonu, tr/en dil seçimi |
| `RAPOR_GOREVI.md` | Claude Code'a verilen rapor yazma görevi |
| `report_builder.py` | JSON → Word, sabit tasarım |
| `RAPOR_FORMATLARI.md`, `ornekler/` | Rapor türleri ve örnek raporlar (`.docx` + metin `.txt`) |
| `config.json` | Drive klasörü, Apps Script adresi, kategoriler (gizli değer içermez) |
| `state.json` | İşlenen kayıtlar (bot günceller, elle düzenlemeyin) |
| `outputs/` | Üretilen txt/srt/docx kopyaları (bot günceller) |

## Gizli değerler (Settings → Secrets and variables → Actions)

| Ad | Değer |
|---|---|
| `APPS_SCRIPT_TOKEN` | Apps Script kodundaki token |
| `CLAUDE_CODE_OAUTH_TOKEN` | `claude setup-token` komutunun verdiği abonelik token'ı |

## Elle çalıştırma

**Actions** sekmesi → **Drive transkript ve rapor** → **Run workflow**. Her çalışmanın sayfasında, işlenen
kayıtları, Drive bağlantılarını ve hataları gösteren bir özet bulunur.

## Notlar

- Bir çalışma en fazla 5,5 saat sürer. 4,5 saatten sonra yeni kayıt başlatılmaz, kalanlar sonraki tura kalır.
  Her kayıttan sonra ilerleme repoya yazılır.
- Bozuk ya da okunamayan dosyalar iki denemeden sonra `failed` olarak işaretlenir ve çalışma özetinde görünür.
- Daha önce raporlanmış bir kaydın yeniden işlenmesi için `state.json` içindeki satırı silmek yeterlidir.
- Ücretsiz planda özel (private) repolar için ayda 2.000 Actions dakikası vardır. Her tur yaklaşık 2–3 dakika,
  1 saatlik bir kayıt ise yaklaşık 1,5–2 saat sürer. Kalan dakikalar Settings → Billing → Usage sayfasında görülür.
- GitHub, 60 gün boyunca hiç commit olmayan repolarda zamanlanmış çalışmaları durdurur. Uzun süre kayıt
  yüklenmezse Actions sekmesinden iş akışını yeniden etkinleştirmek gerekir.
