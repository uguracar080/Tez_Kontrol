from fastapi import FastAPI, UploadFile, File, Request
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pathlib import Path
import uuid
import shutil
import yaml
from datetime import datetime, timedelta

import Tez_Kontrol as tez

app = FastAPI()

BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads_tmp"
UPLOAD_DIR.mkdir(exist_ok=True)

REPORTS_DIR = BASE_DIR / "reports_tmp"
REPORTS_DIR.mkdir(exist_ok=True)

ALLOWED_EXT = {".docx"}

import os
import json

def get_build_info() -> dict:
    """
    HTML'nin çağırdığı /version endpoint'i ile AYNI kaynak.
    Sürüm bilgisi .py içine gömülmez.
    Öncelik sırası:
      1) ENV (APP_VERSION, GIT_SHA, BUILD_TIME)
      2) version.json (container / repo içinde)
    """
    v = (os.getenv("APP_VERSION") or "").strip()
    sha = (os.getenv("GIT_SHA") or "").strip()
    bt = (os.getenv("BUILD_TIME") or "").strip()

    if v or sha or bt:
        return {
            "version": v or "dev",
            "sha": sha,
            "build_time": bt,
        }

    base_dir = Path(__file__).parent
    vp = base_dir / "version.json"

    if vp.exists():
        try:
            data = json.loads(vp.read_text(encoding="utf-8"))
            return {
                "version": str(data.get("version") or "dev"),
                "sha": str(data.get("sha") or ""),
                "build_time": str(data.get("build_time") or ""),
            }
        except Exception:
            pass

    return {"version": "dev", "sha": "", "build_time": ""}


# job_id -> indirirken görünecek pdf adı
REPORT_DOWNLOAD_NAMES: dict[str, str] = {}

def _safe_filename_part(s: str) -> str:
    """
    Dosya adında sorun çıkarabilecek karakterleri temizler.
    """
    s = (s or "").strip()
    if not s:
        return "OGRENCI"
    s = s.replace(" ", "_")
    out = []
    for ch in s:
        if ch.isalnum() or ch in ("_", "-", "."):
            out.append(ch)
        else:
            out.append("_")
    cleaned = "".join(out)
    while "__" in cleaned:
        cleaned = cleaned.replace("__", "_")
    return cleaned.strip("_") or "OGRENCI"


app.mount("/static", StaticFiles(directory="static", html=True), name="static")


@app.get("/")
async def home():
    return FileResponse(str(BASE_DIR / "static" / "index.html"))


def _find_col_idx(cols, candidates):
    cols_lower = [str(c).strip().lower() for c in cols]
    for cand in candidates:
        cand_l = str(cand).lower()
        if cand_l in cols_lower:
            return cols_lower.index(cand_l)
    return None


def _is_ok(v) -> bool:
    s = str(v).strip()
    return s in ("✔", "✓", "E", "EVET", "True", "1") or (isinstance(v, bool) and v is True)


def _is_fail(v) -> bool:
    s = str(v).strip()
    return s in ("✘", "✗", "H", "HAYIR", "False", "0") or (isinstance(v, bool) and v is False)


def compute_summary(results_by_section: dict, report_config: dict):
    """
    - Genel uyum yüzdesi (AĞIRLIKLI):
        Ön Sayfalar %15 + Tez Metni %80 + Arka Sayfalar %5
      Not: Ön Sayfalar grubuna abstract_tr DAHİL
    - Bölüm bazında uyum yüzdesi (mevcut mantık korunur)
    """
    report_cfg = (report_config or {}).get("report", {}) or {}
    order = report_cfg.get("section_order", report_cfg.get("order", [])) or []
    titles = report_cfg.get("section_labels", report_cfg.get("section_titles", {})) or {}
    table_cols = report_cfg.get("table_columns", []) or []

    e_idx = _find_col_idx(table_cols, ["Evet", "E", "Yes"])
    h_idx = _find_col_idx(table_cols, ["Hayır", "Hayir", "H", "No"])

    total_checked = 0
    total_ok = 0
    total_fail = 0

    sections = []

    # Bölüm bazlı sayımlar
    per_key_stats = {}  # key -> {"total":..., "ok":..., "fail":...}

    for section_key in order:
        section_results = results_by_section.get(section_key, []) or []
        ok_count = 0
        fail_count = 0

        for row in section_results:
            total_checked += 1
            row = list(row)

            if e_idx is not None and e_idx < len(row) and _is_ok(row[e_idx]):
                total_ok += 1
                ok_count += 1

            if h_idx is not None and h_idx < len(row) and _is_fail(row[h_idx]):
                total_fail += 1
                fail_count += 1

        total_rules = len(section_results)
        pct = (ok_count / total_rules * 100.0) if total_rules > 0 else 0.0

        per_key_stats[section_key] = {
            "total": total_rules,
            "ok": ok_count,
            "fail": fail_count,
        }

        sections.append({
            "key": section_key,
            "label": titles.get(section_key, section_key.upper()),
            "total": total_rules,
            "ok": ok_count,
            "fail": fail_count,
            "pct": round(pct, 1),
        })

    # --- KATEGORİ TANIMLARI (index.html ile uyumlu) ---
    # Ön Sayfalar: abstract_tr DAHİL
    front_keys = {
        "general",
        "inner_cover",
        "approval",
        "ethics",
        "abstract_tr",      # ✅ eklendi
        "abstract_en",
        "acknowledgements",
        "toc",
        "list_of_tables",
        "list_of_figures",
        "symbols_abbreviations",
    }
    back_keys = {"references", "appendices", "cv"}

    def _group_pct(keys: set[str]) -> float:
        group_total = 0
        group_ok = 0
        for k in keys:
            st = per_key_stats.get(k)
            if not st:
                continue
            group_total += int(st.get("total", 0) or 0)
            group_ok += int(st.get("ok", 0) or 0)
        return (group_ok / group_total * 100.0) if group_total > 0 else 0.0

    # Body (Tez Metni): order içinden front/back dışında kalan her şey
    body_keys = {k for k in order if (k not in front_keys and k not in back_keys)}

    front_pct = _group_pct(front_keys)
    body_pct = _group_pct(body_keys)
    back_pct = _group_pct(back_keys)

    # ✅ AĞIRLIKLI GENEL YÜZDE
    weighted_overall_pct = 0.15 * front_pct + 0.80 * body_pct + 0.05 * back_pct

    return {
        "overall": {
            "pct": round(weighted_overall_pct, 1),  # ✅ artık ağırlıklı
            "total": total_checked,
            "ok": total_ok,
            "fail": total_fail,
            # İstersen debug için bunları da gönderebilirsin (UI kullanmazsa sorun olmaz):
            "front_pct": round(front_pct, 1),
            "body_pct": round(body_pct, 1),
            "back_pct": round(back_pct, 1),
        },
        "sections": sections,
        "order": order,
        "e_idx": e_idx,
        "h_idx": h_idx,
    }



def extract_violations(results_by_section: dict, order: list, e_idx: int | None, h_idx: int | None):
    """
    JSON'a bölüm bazında ihlaller:
    violations = {
      "section_key": [ {"no":1,"title":"...","explanation":"..."} , ... ],
      ...
    }
    Not: Yalnızca FAIL (✘ / Hayır) olanları alır.
    """
    violations = {}

    for section_key in order or []:
        section_results = results_by_section.get(section_key, []) or []
        items = []

        for row in section_results:
            row = list(row)

            # beklenen satır: [rule_no, rule_title, yes, no, explanation]
            rule_no = row[0] if len(row) > 0 else None
            rule_title = row[1] if len(row) > 1 else ""
            no_val = row[h_idx] if (h_idx is not None and h_idx < len(row)) else (row[3] if len(row) > 3 else "")
            exp = row[4] if len(row) > 4 else ""

            if _is_fail(no_val):
                items.append({
                    "no": int(rule_no) if str(rule_no).isdigit() else rule_no,
                    "title": str(rule_title),
                    "explanation": str(exp).strip(),
                })

        violations[section_key] = items

    return violations


def cleanup_old_reports(older_than_minutes: int = 60):
    try:
        cutoff = datetime.now() - timedelta(minutes=older_than_minutes)
        for p in REPORTS_DIR.glob("*.pdf"):
            try:
                mtime = datetime.fromtimestamp(p.stat().st_mtime)
                if mtime < cutoff:
                    p.unlink(missing_ok=True)
                    # mapping temizliği (job_id = dosya adı gövdesi)
                    job_id = p.stem
                    REPORT_DOWNLOAD_NAMES.pop(job_id, None)

            except:
                pass
    except:
        pass


@app.get("/api/analyze")
async def analyze_get(request: Request):
    return JSONResponse(
        status_code=405,
        content={
            "ok": False,
            "error": "Bu endpoint POST ister. Network’te /api/analyze Request Method POST olmalı.",
            "got_method": request.method,
            "got_url": str(request.url),
        },
    )


@app.get("/api/report/{job_id}.pdf")
async def get_report(job_id: str):
    pdf_path = REPORTS_DIR / f"{job_id}.pdf"
    if not pdf_path.exists():
        return JSONResponse(status_code=404, content={"ok": False, "error": "Rapor bulunamadı veya süresi doldu."})
    
    download_name = REPORT_DOWNLOAD_NAMES.get(job_id, pdf_path.name)
    
    return FileResponse(
        path=str(pdf_path),
        media_type="application/pdf",
        filename=download_name,
    )


@app.post("/api/analyze")
async def analyze(file: UploadFile = File(...)):
    cleanup_old_reports(older_than_minutes=90)

    ext = Path(file.filename).suffix.lower()
    if ext not in ALLOWED_EXT:
        return JSONResponse(status_code=400, content={"ok": False, "error": "Sadece .docx kabul ediliyor."})

    job_id = str(uuid.uuid4())
    tmp_path = UPLOAD_DIR / f"{job_id}{ext}"

    try:
        with open(tmp_path, "wb") as f:
            shutil.copyfileobj(file.file, f)



        rules_file = Path(tez.__file__).parent / "rules.yaml"
        report_file = Path(tez.__file__).parent / "report.yaml"

        # --- sürüm bilgisini al (HTML ile AYNI kaynak) ---
        bi = get_build_info()   # /version ile aynı mantık
        ver = bi.get("version") or "dev"
        sha = bi.get("sha") or ""

        # UI’de görünen formatla birebir
        app_version_text = f"{ver} ({sha})" if sha else ver

        pdf_path, results_by_section, student_name = tez.run_thesis_check(tmp_path, rules_file, report_file,app_version_text=app_version_text)
        # ✅ İndirme adı: RAPOR_OGRENCI_ADI_SOYADI_Tarih_Saat.pdf
        ts = datetime.now().strftime("%d.%m.%Y_%H-%M")
        name_part = _safe_filename_part(student_name)
        suggested_name = f"RAPOR_{name_part}_{ts}.pdf"
        REPORT_DOWNLOAD_NAMES[job_id] = suggested_name

        pdf_path = Path(pdf_path)
        if not pdf_path.exists():
            return JSONResponse(status_code=500, content={"ok": False, "error": "PDF raporu oluşturuldu ama bulunamadı."})

        cached_pdf = REPORTS_DIR / f"{job_id}.pdf"
        shutil.copyfile(pdf_path, cached_pdf)

        report_cfg = {}
        try:
            with open(report_file, "r", encoding="utf-8") as rf:
                report_cfg = yaml.safe_load(rf) or {}
        except:
            report_cfg = {}

        summary = compute_summary(results_by_section or {}, report_cfg)
        violations = extract_violations(
            results_by_section or {},
            summary.get("order", []),
            summary.get("e_idx"),
            summary.get("h_idx"),
        )

        return JSONResponse(
            status_code=200,
            content={
                "ok": True,
                "job_id": job_id,
                "student_name": student_name,
                "pdf_url": f"/api/report/{job_id}.pdf",
                "overall": summary["overall"],
                "sections": summary["sections"],
                "violations": violations,  # ✅ yeni
            },
        )

    except Exception as e:
        return JSONResponse(status_code=500, content={"ok": False, "error": str(e)})

    finally:
        try:
            if tmp_path.exists():
                tmp_path.unlink()
        except:
            pass
