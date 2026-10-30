"""
transcribe_fast.py – faster-whisper (CTranslate2) ile hızlı, ilerleme gösteren
CPU transkripsiyon. openai-whisper'a göre çok daha hızlı; sessizlikleri VAD ile atlar.

Çalıştırma:
    python transcribe_fast.py <ses dosyası>
ffmpeg/ffprobe PATH'ten bulunur; iş parçacığı sayısı WHISPER_THREADS ya da işlemci sayısıdır.
"""
from __future__ import annotations
import sys, time, subprocess, tempfile, os, shutil
from pathlib import Path
from faster_whisper import WhisperModel

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

# ===== KONFİG =====
# Dosya adı komut satırından verilebilir: python transcribe_fast.py <dosya>
AUDIO_FILE = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("Tefsir_190626.mp4")
MODEL_SIZE = "small"        # tiny|base|small|medium|large-v3
# Dil: WHISPER_LANG ortam değişkeniyle sabitlenebilir (ör. "tr", "en"). Boşsa kaydın farklı
# yerlerinden alınan örneklerle YALNIZCA aşağıdaki adaylar arasından seçilir. Tefsir kayıtlarındaki
# Arapça tilavet ya da sessiz girişler, serbest algılamada dili yanlış (ör. "ar") gösterebiliyor.
LANGUAGE   = os.environ.get("WHISPER_LANG") or None
CANDIDATE_LANGS = ("tr", "en")
BEAM_SIZE  = 5
CPU_THREADS = int(os.environ.get("WHISPER_THREADS") or os.cpu_count() or 4)
FFMPEG     = shutil.which("ffmpeg") or r"C:\ffmpeg\bin\ffmpeg.exe"
FFPROBE    = shutil.which("ffprobe") or r"C:\ffmpeg\bin\ffprobe.exe"

def fmt(sec: float) -> str:
    m, s = divmod(int(sec), 60); h, m = divmod(m, 60)
    return f"{h:d}:{m:02d}:{s:02d}"

def probe_duration(path: Path) -> float:
    """ffprobe ile güvenilir süre (saniye). Bulamazsa 0."""
    try:
        out = subprocess.run(
            [FFPROBE, "-v", "error", "-show_entries", "format=duration",
             "-of", "default=noprint_wrappers=1:nokey=1", str(path)],
            capture_output=True, text=True, check=True).stdout.strip()
        return float(out)
    except Exception:
        return 0.0

def to_wav(src: Path) -> Path:
    """Girdiyi 16kHz mono WAV'a çevirir — bazı m4a konteynerlerini faster-whisper
    doğrudan okuyamıyor (süre 0, boş çıktı). ffmpeg her formatı güvenle çözer."""
    fd, tmp = tempfile.mkstemp(suffix=".wav"); os.close(fd)
    subprocess.run([FFMPEG, "-y", "-i", str(src), "-ar", "16000", "-ac", "1",
                    "-vn", tmp], capture_output=True, check=True)
    return Path(tmp)

def pick_language(model, wav: Path) -> tuple[str, dict]:
    """Kayıt boyunca 6 noktadan 30 sn'lik örnek alır, aday diller için olasılıkları toplar."""
    from faster_whisper import decode_audio
    audio = decode_audio(str(wav), sampling_rate=16000)
    n, win = len(audio), 30 * 16000
    starts = [int(n * f) for f in (0.10, 0.25, 0.40, 0.55, 0.70, 0.85)] if n > 2 * win else [0]
    scores = {c: 0.0 for c in CANDIDATE_LANGS}
    for st in starts:
        chunk = audio[st:st + win]
        if len(chunk) < 16000:
            continue
        _, _, probs = model.detect_language(audio=chunk)
        for lang, pr in probs:
            if lang in scores:
                scores[lang] += pr
    best = max(scores, key=scores.get) if any(scores.values()) else CANDIDATE_LANGS[0]
    return best, scores


def main():
    if not AUDIO_FILE.is_file():
        sys.exit(f"Ses dosyası bulunamadı: {AUDIO_FILE}")

    print(f"▶️  Model yükleniyor: {MODEL_SIZE} (CPU, int8, {CPU_THREADS} thread)…", flush=True)
    t0 = time.perf_counter()
    model = WhisperModel(MODEL_SIZE, device="cpu", compute_type="int8", cpu_threads=CPU_THREADS)
    print(f"⏱  Model yüklendi: {time.perf_counter()-t0:.1f}s", flush=True)

    # Girdiyi her zaman temiz WAV'a çevir (m4a konteyner uyumsuzluklarına karşı garanti)
    print(f"▶️  Ses hazırlanıyor (16kHz mono WAV)…", flush=True)
    wav = to_wav(AUDIO_FILE)

    lang = LANGUAGE
    if not lang:
        lang, scores = pick_language(model, wav)
        print("ℹ️  Dil seçimi: " + ", ".join(f"{k}={v:.2f}" for k, v in scores.items()) + f" → {lang}", flush=True)

    print(f"▶️  Transkripsiyon başlıyor → {AUDIO_FILE.name}", flush=True)
    t1 = time.perf_counter()
    segments, info = model.transcribe(
        str(wav), language=lang, beam_size=BEAM_SIZE,
        vad_filter=True, vad_parameters=dict(min_silence_duration_ms=500),
    )
    # Süre: faster-whisper bazen 0 döndürüyor → ffprobe'a düş
    total = info.duration or probe_duration(AUDIO_FILE)
    print(f"ℹ️  Ses süresi: {fmt(total)}  | dil: {info.language} ({info.language_probability:.2f})", flush=True)

    txt_path = AUDIO_FILE.with_suffix(".txt")
    srt_path = AUDIO_FILE.with_suffix(".srt")
    parts = []
    with open(txt_path, "w", encoding="utf-8") as ftxt, open(srt_path, "w", encoding="utf-8") as fsrt:
        for i, seg in enumerate(segments, 1):
            ftxt.write(seg.text.strip() + " "); ftxt.flush()
            fsrt.write(f"{i}\n{_srt_t(seg.start)} --> {_srt_t(seg.end)}\n{seg.text.strip()}\n\n"); fsrt.flush()
            parts.append(seg.text)
            pct = seg.end / total * 100 if total else 0
            print(f"  [{fmt(seg.end)}/{fmt(total)}  %{pct:4.1f}]  {seg.text.strip()[:70]}", flush=True)

    try:
        wav.unlink()  # geçici WAV'ı sil
    except OSError:
        pass

    dt = time.perf_counter() - t1
    rtf = f"{dt/total:.2f}" if total else "n/a"
    print(f"\n✅  Metin kaydedildi → {txt_path}", flush=True)
    print(f"✅  Altyazı kaydedildi → {srt_path}", flush=True)
    print(f"⏱  Transkripsiyon süresi: {fmt(dt)}  (RTF {rtf})", flush=True)

def _srt_t(s: float) -> str:
    ms = int((s - int(s)) * 1000); m, sec = divmod(int(s), 60); h, m = divmod(m, 60)
    return f"{h:02d}:{m:02d}:{sec:02d},{ms:03d}"

if __name__ == "__main__":
    main()
