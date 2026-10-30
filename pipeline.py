"""
pipeline.py – Bulut (GitHub Actions) sürümü. Herkese açık Google Drive klasöründeki yeni ses kayıtlarını
indirir, transkript eder ve sonuçları Apps Script üzerinden Drive'a yükler. Raporları iş akışının sonraki
adımında Claude Code yazar (RAPOR_GOREVI.md) ve "mark-reported" ile yükler.

Komutlar (python pipeline.py <komut>):
    run                         Yeni kayıtları indir + transkript et + txt/srt yükle (kayıt kayıt)
    pending                     Raporu yazılacak kayıtları JSON olarak listele
    mark-reported <key> <docx>  Raporu transkriptin yanına koy, Drive'a yükle, "rapor edildi" işaretle
    status                      Durumu, rapor ve yükleme bekleyenleri göster
    sync                        Sadece bekleyen yüklemeleri dene

Gizli değerler ortam değişkenlerinden okunur (GitHub → Settings → Secrets → Actions):
    APPS_SCRIPT_TOKEN   Apps Script web uygulamasının token'ı
    APPS_SCRIPT_URL     (isteğe bağlı) config.json'daki adresi geçersiz kılar

Durum: state.json (anahtar = Drive dosya kimliği). Yollar repo köküne göre görelidir. GIT_PERSIST=1 iken
her kayıttan sonra state.json ve outputs/ commit edilip push'lanır; iş yarıda kesilse de ilerleme kaybolmaz.

Klasör yapısı: <ana klasör>/<kategori>/{Sesler, Transkriptler, Raporlar}. Kayıt hangi kategorinin Sesler/
klasöründeyse çıktıları da o kategoriye gider.
"""
from __future__ import annotations
import base64, json, os, shutil, subprocess, sys, time, traceback
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests

HERE = Path(__file__).resolve().parent
TRANSCRIBER = HERE / "transcribe_fast.py"
CONFIG = HERE / "config.json"
STATE = HERE / "state.json"
WORK = HERE / "work"                                 # indirilen ses dosyaları (geçici, commit edilmez)
OUTPUTS = HERE / "outputs"                           # <kategori>/ txt, srt, docx
LAST_RUN = HERE / "last_run.json"

AUDIO_EXT = {".m4a", ".mp4", ".mp3", ".wav", ".aac", ".ogg", ".opus", ".flac", ".webm", ".amr", ".3gp", ".mov"}
MAX_ATTEMPTS = 2            # geçici hatalarda en fazla bu kadar dene
TIME_BUDGET_H = float(os.environ.get("TIME_BUDGET_HOURS", "4.5"))   # bu süreden sonra yeni kayıt başlatma
MIME = {".txt": "text/plain", ".srt": "application/x-subrip",
        ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"}
TR = timezone(timedelta(hours=3))                    # Türkiye saati (kayıtlar ve günlük bu saatle)


# ---------------------------------------------------------------- yardımcılar
def now() -> datetime:
    return datetime.now(TR).replace(tzinfo=None)


def log(msg: str):
    print(f"[{now():%Y-%m-%d %H:%M:%S}] {msg}", file=sys.stderr, flush=True)


def emit(obj):
    print(json.dumps(obj, ensure_ascii=False, indent=2))


def rel(p: Path) -> str:
    return p.resolve().relative_to(HERE).as_posix()


def load_state() -> dict:
    try:
        return json.loads(STATE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save_state(state: dict):
    tmp = STATE.with_suffix(".tmp")
    tmp.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(STATE)


def persist(msg: str):
    """state.json ve outputs/ değişikliklerini repoya yazar (yalnızca GIT_PERSIST=1 iken)."""
    if os.environ.get("GIT_PERSIST") != "1":
        return
    try:
        subprocess.run(["git", "add", "state.json", "outputs"], cwd=HERE, check=True)
        if subprocess.run(["git", "diff", "--cached", "--quiet"], cwd=HERE).returncode == 0:
            return
        subprocess.run(["git", "commit", "-q", "-m", msg], cwd=HERE, check=True)
        for _ in range(3):
            if subprocess.run(["git", "push", "-q"], cwd=HERE).returncode == 0:
                return
            subprocess.run(["git", "pull", "-q", "--rebase"], cwd=HERE)
        log("UYARI: git push başarısız; değişiklikler iş sonunda tekrar denenecek.")
    except subprocess.CalledProcessError as e:
        log(f"UYARI: git commit başarısız: {e}")


def load_config() -> tuple[dict, str | None]:
    try:
        cfg = json.loads(CONFIG.read_text(encoding="utf-8"))
    except (OSError, ValueError) as e:
        return {}, f"config.json okunamadı: {e}"
    cfg["apps_script_url"] = os.environ.get("APPS_SCRIPT_URL") or cfg.get("apps_script_url")
    cfg["apps_script_token"] = os.environ.get("APPS_SCRIPT_TOKEN")
    if not cfg.get("drive_folder_id"):
        return cfg, "config.json içinde drive_folder_id boş."
    if not (cfg["apps_script_url"] and cfg["apps_script_token"]):
        return cfg, "APPS_SCRIPT_TOKEN gizli değeri tanımlı değil (GitHub → Settings → Secrets and variables → Actions)."
    return cfg, None


def script_call(cfg, payload: dict, timeout=120, attempts=3) -> dict:
    """Apps Script'e POST eder. Google ara sıra geçici 404/5xx döndürdüğü için birkaç kez dener."""
    data = None
    for i in range(attempts):
        try:
            r = requests.post(cfg["apps_script_url"], json={**payload, "token": cfg["apps_script_token"]},
                              timeout=timeout)
            r.raise_for_status()
            data = r.json()
            break
        except (requests.RequestException, ValueError) as e:
            if i == attempts - 1:
                raise RuntimeError(f"Apps Script'e ulaşılamadı ({e}). Dağıtım erişimi 'Herkes' mi, URL doğru mu?")
            time.sleep(5 * (i + 1))
    if not data.get("ok"):
        raise RuntimeError(f"Apps Script hatası: {data.get('error')}")
    return data


# ---------------------------------------------------------------- Drive
def list_remote(cfg) -> list[dict]:
    """Kategorilerin Sesler/ klasörlerindeki ses dosyaları: [{id, name, category}]."""
    try:
        files = script_call(cfg, {"action": "list"})["files"]
        return [{"id": f["id"], "name": f["name"], "category": f["category"]} for f in files
                if Path(f["name"]).suffix.lower() in AUDIO_EXT]
    except Exception as e:  # noqa: BLE001 – listeleme için gdown'a düş
        log(f"Apps Script listeleme başarısız, gdown ile denenecek: {e}")
    import gdown
    audio_sub = cfg.get("audio_subfolder", "Sesler")
    url = f"https://drive.google.com/drive/folders/{cfg['drive_folder_id']}"
    out = []
    for f in gdown.download_folder(url=url, skip_download=True, quiet=True) or []:
        parts = f.path.replace(os.sep, "/").split("/")          # ör. "İş/Sesler/kayit.m4a"
        if len(parts) == 3 and parts[1] == audio_sub and Path(parts[2]).suffix.lower() in AUDIO_EXT:
            out.append({"id": f.id, "name": parts[2], "category": parts[0]})
    return out


def download(file_id: str, name: str) -> Path:
    import gdown
    WORK.mkdir(exist_ok=True)
    dest = WORK / name
    gdown.download(id=file_id, output=str(dest), quiet=True)
    if not dest.is_file() or dest.stat().st_size == 0:
        raise RuntimeError("İndirme başarısız ya da dosya boş.")
    return dest


def upload(cfg, category: str, subfolder: str, path: Path) -> str:
    data = base64.b64encode(path.read_bytes()).decode("ascii")
    res = script_call(cfg, {"action": "upload", "category": category, "subfolder": subfolder,
                            "name": path.name, "mimeType": MIME.get(path.suffix.lower(), "application/octet-stream"),
                            "data": data}, timeout=300)
    return res.get("url", "")


def sync_uploads(cfg, state) -> tuple[list, str | None]:
    """Yüklenmemiş txt/srt/docx dosyalarını ilgili kategorinin klasörlerine gönderir."""
    done, err = [], None
    for v in state.values():
        if v.get("status") not in ("transcribed", "reported") or not v.get("category"):
            continue
        txt = HERE / v["txt"]
        jobs = [(ext, cfg.get("transcript_subfolder", "Transkriptler"), txt.with_suffix(ext))
                for ext in (".txt", ".srt") if not v.get(f"uploaded{ext}")]
        if v.get("status") == "reported" and not v.get("uploaded.docx"):
            jobs.append((".docx", cfg.get("report_subfolder", "Raporlar"), HERE / v["report"]))
        for ext, sub, p in jobs:
            if not p.is_file():
                continue
            try:
                url = upload(cfg, v["category"], sub, p)
            except Exception as e:  # noqa: BLE001 – bir sonraki turda tekrar denenir
                err = f"Yükleme başarısız ({p.name}): {e}"
                log(err)
                save_state(state)
                return done, err
            v[f"uploaded{ext}"] = True
            if ext == ".docx":
                v["report_url"] = url
            done.append({"file": p.name, "folder": f"{v['category']}/{sub}", "url": url})
            log(f"Yüklendi: {v['category']}/{sub}/{p.name}")
    save_state(state)
    return done, err


def pending_uploads(state):
    out = []
    for v in state.values():
        if v.get("status") in ("transcribed", "reported"):
            out += [Path(v["txt"]).stem + e for e in (".txt", ".srt") if not v.get(f"uploaded{e}")]
        if v.get("status") == "reported" and not v.get("uploaded.docx"):
            out.append(Path(v["report"]).name)
    return out


# ---------------------------------------------------------------- transkripsiyon ve rapor
def transcribe_one(local: Path, out_dir: Path) -> dict:
    out_dir.mkdir(parents=True, exist_ok=True)
    t0 = time.time()
    proc = subprocess.run([sys.executable, str(TRANSCRIBER), str(local)], cwd=str(HERE),
                          capture_output=True, text=True, encoding="utf-8", errors="replace")
    txt, srt = local.with_suffix(".txt"), local.with_suffix(".srt")
    info_line = next((l for l in proc.stdout.splitlines() if "Ses süresi:" in l), "")
    duration = info_line.split("Ses süresi:")[1].split("|")[0].strip() if info_line else None
    language = info_line.split("dil:")[1].split()[0] if "dil:" in info_line else None
    try:
        if proc.returncode != 0:
            tail = (proc.stderr or proc.stdout)[-600:]
            return {"ok": False, "retry": True, "error": f"transcribe_fast.py hata kodu {proc.returncode}: {tail}"}
        if not txt.exists() or txt.stat().st_size == 0:
            return {"ok": False, "retry": False,
                    "error": "Ses çözülemedi, transkript boş (dosya bozuk veya yarıda kesilmiş olabilir)."}
        shutil.copy2(txt, out_dir / txt.name)
        if srt.exists():
            shutil.copy2(srt, out_dir / srt.name)
        return {"ok": True, "txt": rel(out_dir / txt.name), "duration": duration, "language": language,
                "seconds": round(time.time() - t0)}
    finally:
        for f in (local, txt, srt):
            try:
                f.unlink()
            except OSError:
                pass


def pending_reports(state):
    return [{"key": k, "source": v["source"], "category": v.get("category"), "txt": v["txt"],
             "duration": v.get("duration"), "language": v.get("language"),
             "suggested_docx": Path(v["txt"]).with_name(Path(v["source"]).stem + "_Rapor.docx").as_posix()}
            for k, v in state.items() if v.get("status") == "transcribed"]


def set_output(name: str, value):
    """GitHub Actions adım çıktısı (sonraki adımlar steps.<id>.outputs.<name> ile okur)."""
    path = os.environ.get("GITHUB_OUTPUT")
    if path:
        with open(path, "a", encoding="utf-8") as f:
            f.write(f"{name}={value}\n")


# ---------------------------------------------------------------- komutlar
def cmd_run():
    t_start = time.time()
    cfg, err = load_config()
    state = load_state()
    result = {"time": now().isoformat(timespec="seconds"), "error": err, "new_transcribed": [],
              "failed": [], "uploaded": [], "upload_error": None, "deferred": [],
              "pending_uploads": [], "pending_reports": []}
    if err:
        log(err); emit(result); sys.exit(1)

    # 1) Önceki turlardan kalan yüklemeler
    up, result["upload_error"] = sync_uploads(cfg, state)
    result["uploaded"] += up
    persist("Bekleyen yüklemeler")

    # 2) Yeni kayıtlar
    try:
        remote = list_remote(cfg)
    except Exception as e:  # noqa: BLE001
        result["error"] = f"Drive klasörü listelenemedi: {e}"
        log(result["error"]); remote = []
    for f in remote:
        entry = state.get(f["id"], {})
        if entry.get("status") in ("transcribed", "reported", "baseline", "failed"):
            continue
        if (time.time() - t_start) / 3600 > TIME_BUDGET_H:
            result["deferred"].append(f"{f['category']}/{f['name']}")
            continue
        attempts = entry.get("attempts", 0) + 1
        log(f"İşleniyor: {f['category']}/{f['name']}")
        try:
            r = transcribe_one(download(f["id"], f["name"]), OUTPUTS / f["category"])
        except Exception as e:  # noqa: BLE001 – indirme hatası: geçici say
            r = {"ok": False, "retry": True, "error": f"İndirme hatası: {e}"}
        if r["ok"]:
            state[f["id"]] = {"status": "transcribed", "source": f["name"], "category": f["category"],
                              "txt": r["txt"], "duration": r["duration"], "language": r["language"],
                              "transcribed_at": now().isoformat(timespec="seconds")}
            result["new_transcribed"].append({"source": f["name"], "category": f["category"],
                                              "duration": r["duration"], "language": r["language"],
                                              "seconds": r["seconds"]})
            log(f"Tamam: {f['name']} ({r['duration']}, {r['seconds']} sn)")
        else:
            final = not r["retry"] or attempts >= MAX_ATTEMPTS
            state[f["id"]] = {"status": "failed" if final else "retry", "source": f["name"],
                              "category": f["category"], "attempts": attempts, "error": r["error"]}
            result["failed"].append({"source": f["name"], "category": f["category"],
                                     "error": r["error"], "will_retry": not final})
            log(f"HATA: {f['name']}: {r['error']}")
        save_state(state)
        up, result["upload_error"] = sync_uploads(cfg, state)
        result["uploaded"] += up
        persist(f"İşlendi: {f['category']}/{f['name']}")

    result["pending_uploads"] = pending_uploads(state)
    result["pending_reports"] = [p["source"] for p in pending_reports(state)]
    result["finished"] = now().isoformat(timespec="seconds")
    LAST_RUN.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    set_output("pending_reports", len(result["pending_reports"]))
    emit(result)
    write_summary(result)


def write_summary(result):
    """GitHub Actions çalıştırma sayfasına okunur bir özet yazar."""
    path = os.environ.get("GITHUB_STEP_SUMMARY")
    if not path:
        return
    lines = [f"## Drive transkript turu — {result['time']}", ""]
    for t in result["new_transcribed"]:
        lines.append(f"- 🎙️ **{t['category']}/{t['source']}** — {t['duration']}, dil: {t['language']}")
    for u in result["uploaded"]:
        lines.append(f"- ⬆️ [{u['folder']}/{u['file']}]({u['url']})")
    for f in result["failed"]:
        lines.append(f"- ❌ {f['category']}/{f['source']}: {f['error'][:300]}"
                     + (" (tekrar denenecek)" if f["will_retry"] else " (dosyayı kontrol edin)"))
    for p in result["pending_reports"]:
        lines.append(f"- 📝 Raporu yazılacak: {p}")
    for d in result["deferred"]:
        lines.append(f"- ⏳ Süre doldu, sonraki tura kaldı: {d}")
    if result["error"]:
        lines.append(f"- ❗ {result['error']}")
    if result["upload_error"]:
        lines.append(f"- ❗ {result['upload_error']}")
    if len(lines) == 2:
        lines.append("Yeni kayıt yok.")
    with open(path, "a", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


def cmd_pending():
    emit(pending_reports(load_state()))


def cmd_mark_reported(key: str, docx: str):
    cfg, err = load_config()
    state = load_state()
    if key not in state:
        sys.exit(f"Bilinmeyen anahtar: {key}")
    src = (HERE / docx).resolve()
    if not src.is_file():
        sys.exit(f"Rapor dosyası yok: {docx}")
    dest = (HERE / state[key]["txt"]).with_name(src.name)     # raporu transkriptin yanına koy
    if src != dest.resolve():
        shutil.copy2(src, dest)
    state[key].update(status="reported", report=rel(dest), reported_at=now().isoformat(timespec="seconds"))
    save_state(state)
    uploaded, up_err = sync_uploads(cfg, state) if not err else ([], err)
    persist(f"Rapor: {state[key].get('category')}/{dest.name}")
    with open(os.environ.get("GITHUB_STEP_SUMMARY") or os.devnull, "a", encoding="utf-8") as f:
        for u in uploaded:
            f.write(f"- 📄 [{u['folder']}/{u['file']}]({u['url']})\n")
        if up_err:
            f.write(f"- ❗ {up_err}\n")
    emit({"report": rel(dest), "uploaded": uploaded, "upload_error": up_err})


def cmd_sync():
    cfg, err = load_config()
    if err:
        sys.exit(err)
    state = load_state()
    uploaded, up_err = sync_uploads(cfg, state)
    persist("Bekleyen yüklemeler")
    emit({"uploaded": uploaded, "upload_error": up_err, "pending_uploads": pending_uploads(state)})


def cmd_status():
    cfg, err = load_config()
    state = load_state()
    counts = {}
    for v in state.values():
        counts[v.get("status")] = counts.get(v.get("status"), 0) + 1
    emit({"config_error": err, "counts": counts,
          "pending_reports": [v["source"] for v in state.values() if v.get("status") == "transcribed"],
          "pending_uploads": pending_uploads(state),
          "failed": [{"source": v["source"], "category": v.get("category"), "error": v.get("error")}
                     for v in state.values() if v.get("status") == "failed"]})


if __name__ == "__main__":
    cmds = {"run": cmd_run, "status": cmd_status, "sync": cmd_sync, "pending": cmd_pending}
    args = sys.argv[1:]
    if len(args) == 1 and args[0] in cmds:
        try:
            cmds[args[0]]()
        except Exception:  # noqa: BLE001
            log("BEKLENMEDİK HATA:\n" + traceback.format_exc())
            persist("Hata öncesi durum")
            raise
    elif args[:1] == ["mark-reported"] and len(args) == 3:
        cmd_mark_reported(args[1], args[2])
    else:
        sys.exit(__doc__)
