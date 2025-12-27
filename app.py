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
    - Genel uyum yüzdesi
    - Bölüm bazında uyum yüzdesi
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

        sections.append({
            "key": section_key,
            "label": titles.get(section_key, section_key.upper()),
            "total": total_rules,
            "ok": ok_count,
            "fail": fail_count,
            "pct": round(pct, 1),
        })

    overall_pct = (total_ok / total_checked * 100.0) if total_checked > 0 else 0.0

    return {
        "overall": {
            "pct": round(overall_pct, 1),
            "total": total_checked,
            "ok": total_ok,
            "fail": total_fail,
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

    return FileResponse(
        path=str(pdf_path),
        media_type="application/pdf",
        filename=pdf_path.name,
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

        pdf_path, results_by_section, student_name = tez.run_thesis_check(tmp_path, rules_file, report_file)

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
