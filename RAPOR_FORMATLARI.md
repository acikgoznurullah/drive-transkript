# Rapor formatları

Raporlar, kullanıcının Drive'a koyduğu örnek raporlara göre yazılır. Örnekler `ornekler/` klasöründedir;
`report_ai.py` hepsinin metnini modele birlikte verir.

Görsel tasarım (başlık bloğu, renkler, altın çizgiler, tablolar, altbilgi) `report_builder.py` içinde
sabittir. Senin işin yalnızca **içerik JSON'u** üretmek. Şema `report_builder.py` dosyasının başında.

## Genel kurallar (tüm türler)

- **Dil:** Rapor kaydın dilinde yazılır. `pending_reports[].language` alanı "en" ise rapor İngilizce
  (`"lang": "en"`), aksi halde Türkçe olur.
- **Ton:** Düzenlenmiş, akıcı bir özet. Transkripti satır satır aktarma; konuşmayı konulara göre topla.
  Her bölümde önce 1-2 cümlelik bağlam paragrafı, ardından maddeler gelir.
- **Uydurma yok:** Transkriptte olmayan bilgi, isim, rakam ve tarih eklenmez. Belirsizlik açıkça yazılır,
  ör. "yaklaşık", "kayıtta net değil".
- **İsimler:** Ses tanıma hataları bariz ise düzeltilebilir. Örneğin aynı kişi farklı yazılmışsa tek
  biçime getir. Kimin söylediği belli değilse "bir katılımcı" denir. Transkript konuşmacıları ayırt etmez.
- **Sorumlu:** Sorumlusu kayıttan anlaşılmayan aksiyonlara "Belirtilmedi" yazılır.
- **Tarih:** Dosya adındaki YYMMDD (ör. `Ses 261014_132520` → 14 Ekim 2026). Yoksa transkripsiyon günü.
  "Yarın", "pazartesi" gibi ifadeler bu tarihe göre hesaplanır.
- **Vurgu:** Önemli kavramlar paragraf içinde `**kalın**` yazılabilir, ama ölçülü kullanılır.
- **Not:** Her raporun sonuna `note` eklenir. Not raporun ses kaydının dökümünden düzenlendiğini
  söyler ve gerekiyorsa isim ve tarih varsayımlarını belirtir.
- **Dosya adı:** `pending_reports[].suggested_docx` kullanılır.
- **Uzunluk:** 45-60 dakikalık bir kayıt için 3-6 sayfa. Örnekler ölçü alınır.

## 1. Toplantı Raporu (İş, Türkçe). Örnek: `Is_Toplanti_Raporu_DASK.docx`

- `title`: "Toplantı Raporu"
- `subtitle`: Toplantının konusu, tek satır.
- `meta`: [tarih, "Taraf A → Taraf B" veya toplantı türü]
- `footer`: "<Kısa konu> — Toplantı Raporu"
- `info` (Alan/Bilgi): "Tarih / süre", "Konu", varsa "Sunan", varsa "Muhataplar", "Katılımcılar"
- `info_note`: "İsimler kayıttan aktarılmıştır; bazı isimlerde yazım farkı olabilir."
- Bölümler:
  1. Toplantının Amacı (paragraf)
  2. … Konu bölümleri (her ana başlık bir H1; alt başlıklar H2; rakamlı karşılaştırmalar tablo)
  3. Müzakere ve Değerlendirmeler (madde madde, kimin ne dediği biliniyorsa belirtilir)
  4. Kararlar ve Aksiyon Maddeleri: tablo `["#", "Aksiyon", "Sorumlu"]`, widths `[1.0, 10.5, 4.5]`.
     Termin konuşulduysa aksiyon metnine eklenir.
- `note`: "Not: Bu rapor, toplantı ses kaydının dökümünden düzenlenmiştir. …"

## 2. Görüşme Özeti (İş, Türkçe, iki-üç kişilik değerlendirme ve tanıtım görüşmesi). Örnek: `Is_Gorusme_Ozeti_EminYahya.docx`

- `title`: "Görüşme Özeti"
- `subtitle`: "<Kişi> ile gerçekleştirilen <konu> görüşmesi"
- `meta`: [tarih]
- `footer`: "Görüşme Özeti — <Kişi>"
- Bölümler:
  1. Görüşmenin Amacı ve Kapsamı
  2. … İçerik bölümleri (ör. ürünün yapısı, kullanılan veriler, yöntem). Listeler maddeli.
  3. Teknik Değerlendirmeler ve Geliştirme Önerileri (H2 alt başlıklar: "7.1. …", "7.2. …")
  4. Kararlaştırılan Sonraki Adımlar (maddeler)
  5. Genel Değerlendirme (paragraf)
- `note` eklenir.

## 3. Meeting Minutes (İş, İngilizce kayıt). Örnek: `Is_Meeting_Minutes_OASIS_EN.docx`

- `"lang": "en"`
- `title`: "<Org A> – <Org B> Meeting Minutes"
- `subtitle`: Kısa açıklama, ör. "Detailed technical and commercial discussion"
- `meta`: [date, "Participants: <Org A>: names • <Org B>: names"]
- `footer`: "<Org A> × <Org B> | Meeting Minutes | <date>"
- `intro`: "Source: English meeting transcript, approximately N minutes. Names and technical terms are
  standardized where identifiable. …"
- Bölümler:
  1. Executive Summary (bullets, 4-6 madde)
  2. … Konu bölümleri (2.1, 2.2 … H2 alt başlıklar; bullets)
  3. Seçenekler karşılaştırıldıysa "Comparative evaluation" tablosu
  4. Decisions, Agreements, and Confirmed Clarifications (`numbered`)
  5. Action Items and Follow-Up Plan (`numbered`, her madde `{"bold": "<Owner> — ", "text": "<task>. Timing/status: …"}`)
  6. Open Questions / Items Requiring Resolution (bullets)
  7. Closing and Timeline (bullets)
- `note`: "Editorial note: The transcript is machine-generated …"

## 4. Tefsir Dersi Raporu (Özel, tefsir/ders/sohbet). Örnekler: `Ozel_Tefsir_Dersi_Raporu_67-72.docx`, `…_72-74.docx`

- `title`: "Tefsir Dersi Raporu"
- `subtitle`: "<Sûre> Sûresi, <x–y>. Âyetler — <Ana tema>"
- `meta`: ["Ders tarihi: <gün Ay yıl>", varsa "<N>. Hafta", müzakere yapıldıysa "Ders müzakeresiyle birlikte"]
- `footer`: "Tefsir Dersi Raporu — <Sûre> <x–y>"
- Bilgi tablosu kullanılmaz. Yerine `intro` paragrafı yazılır: "Bu rapor, <sûre> sûresinin <x–y>.
  âyetleri üzerine yapılan tefsir dersinin ve dersi izleyen müzakerenin düzenlenmiş bir özetidir. …"
- Bölümler dersin akışını izler. Örnek kurgu:
  1. Âyetlerin Özet Meâli (âyet numaralı maddeler, ör. "… (68).")
  2. Dersin işlediği ana konular, her biri bir H1: anlatının hikmeti, kavram açıklamaları (Arapça
     kelimeler transliterasyonla, ör. "yuhyî"), rivayetler, örnek âyetler (sûre ve âyet numarasıyla)
  3. Alt noktalar H2 olarak yazılır: "a) …", "b) …" ya da kısa başlıklar.
  4. Müzakere: <konu> (katılımcı görüşleri. İsim verilmez, "bir görüşe göre", "bir katılımcı" denir.)
  5. Ana Çıkarımlar (bullets)
  6. Ders Sonu — İdari Notlar (sonraki ders, ara, kaynaklar; konuşulduysa)
- Dinî terimlerde örneklerdeki yazım korunur: Âyet, Sûre, Hz. Mûsâ, İsrâiloğulları, itikadî, ahlâkî.
- `note`: "Not: Bu rapor, ders kaydının dökümünden hareketle düzenlenmiş bir özettir. Tefsirî
  değerlendirmeler ve müzakere bölümündeki görüşler, derste dile getirilen değerlendirmeleri yansıtır."

## 5. Doktora ve diğer kayıtlar

Örnek yoksa en yakın türü uyarla:
- Akademik danışman görüşmesi → Görüşme Özeti. Bölümler: Amaç, Görüşülen Konular, Geri Bildirimler,
  Sonraki Adımlar.
- Ders veya seminer → Tefsir raporunun ders kurgusu. Başlık "Ders Raporu"; bilgi tablosu yerine
  `intro` paragrafı.
- Toplantı → Toplantı Raporu.

Başlık ve kurgu, kayıt içeriğine en uygun olacak şekilde seçilir.
