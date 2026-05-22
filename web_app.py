# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import shutil
import time
import traceback
import zipfile
from pathlib import Path
from uuid import uuid4

from flask import Flask, abort, redirect, render_template, request, send_file, url_for

from app import app_base_dir, convert_one_pdf, parse_hoc_ba_pdf, safe_filename


BASE_DIR = app_base_dir()
DEFAULT_TEMPLATE = BASE_DIR / "templates" / "hoc_ba_mau (2).xlsx"
if not DEFAULT_TEMPLATE.exists():
    DEFAULT_TEMPLATE = BASE_DIR / "templates" / "hoc_ba_mau.xlsx"
STORAGE_DIR = Path(os.environ.get("HOCBA_WEB_STORAGE", BASE_DIR / "web_storage"))
MAX_CONTENT_LENGTH = int(os.environ.get("HOCBA_MAX_UPLOAD_MB", "200")) * 1024 * 1024
JOB_TTL_SECONDS = int(os.environ.get("HOCBA_JOB_TTL_SECONDS", str(24 * 60 * 60)))

web = Flask(__name__)
web.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH
web.config["SECRET_KEY"] = os.environ.get("HOCBA_SECRET_KEY", "change-me-in-production")


def job_dir(job_id: str) -> Path:
    if not job_id or any(ch not in "0123456789abcdef-" for ch in job_id.lower()):
        abort(404)
    path = STORAGE_DIR / job_id
    if not path.exists():
        abort(404)
    return path


def unique_path(folder: Path, filename: str) -> Path:
    stem = safe_filename(Path(filename).stem) or "file"
    suffix = Path(filename).suffix.lower()
    candidate = folder / f"{stem}{suffix}"
    count = 2
    while candidate.exists():
        candidate = folder / f"{stem}_{count}{suffix}"
        count += 1
    return candidate


def cleanup_old_jobs() -> None:
    STORAGE_DIR.mkdir(parents=True, exist_ok=True)
    now = time.time()
    for child in STORAGE_DIR.iterdir():
        try:
            if child.is_dir() and now - child.stat().st_mtime > JOB_TTL_SECONDS:
                shutil.rmtree(child, ignore_errors=True)
        except OSError:
            continue


def parse_pdf_summary(pdf_path: Path) -> dict:
    try:
        data = parse_hoc_ba_pdf(pdf_path)
        return {
            "file": pdf_path.name,
            "student": data.student_name,
            "class_name": data.class_name,
            "status": "Sẵn sàng",
            "error": "",
        }
    except Exception as exc:
        return {
            "file": pdf_path.name,
            "student": "",
            "class_name": "",
            "status": "Vẫn có thể thử chuyển",
            "error": str(exc),
        }


@web.get("/")
def index():
    cleanup_old_jobs()
    return render_template("web_index.html", default_template_exists=DEFAULT_TEMPLATE.exists())


@web.post("/upload")
def upload():
    cleanup_old_jobs()
    uploaded_pdfs = [f for f in request.files.getlist("pdf_files") if f and f.filename]
    template_file = request.files.get("template_file")
    if not uploaded_pdfs:
        return render_template(
            "web_index.html",
            default_template_exists=DEFAULT_TEMPLATE.exists(),
            error="Vui lòng chọn ít nhất một file PDF.",
        ), 400

    job_id = str(uuid4())
    root = STORAGE_DIR / job_id
    input_dir = root / "input"
    input_dir.mkdir(parents=True, exist_ok=True)

    template_path = root / "template.xlsx"
    using_default_template = False
    if template_file and template_file.filename:
        if not template_file.filename.lower().endswith(".xlsx"):
            shutil.rmtree(root, ignore_errors=True)
            return render_template(
                "web_index.html",
                default_template_exists=DEFAULT_TEMPLATE.exists(),
                error="File mẫu phải là Excel .xlsx.",
            ), 400
        template_file.save(template_path)
    elif DEFAULT_TEMPLATE.exists():
        shutil.copy2(DEFAULT_TEMPLATE, template_path)
        using_default_template = True
    else:
        shutil.rmtree(root, ignore_errors=True)
        return render_template(
            "web_index.html",
            default_template_exists=False,
            error="Không tìm thấy file mẫu mặc định trong thư mục templates.",
        ), 500

    if using_default_template:
        for map_name in ("field_map.csv", "field_map.json"):
            source_map = DEFAULT_TEMPLATE.parent / map_name
            if source_map.exists():
                shutil.copy2(source_map, root / map_name)

    rows = []
    for file_storage in uploaded_pdfs:
        if not file_storage.filename.lower().endswith(".pdf"):
            continue
        pdf_path = unique_path(input_dir, file_storage.filename)
        file_storage.save(pdf_path)
        rows.append(parse_pdf_summary(pdf_path))

    if not rows:
        shutil.rmtree(root, ignore_errors=True)
        return render_template(
            "web_index.html",
            default_template_exists=DEFAULT_TEMPLATE.exists(),
            error="Không có file PDF hợp lệ.",
        ), 400

    return render_template("web_review.html", job_id=job_id, rows=rows)


@web.post("/convert/<job_id>")
def convert(job_id: str):
    root = job_dir(job_id)
    input_dir = root / "input"
    output_dir = root / "output"
    template_path = root / "template.xlsx"
    selected_files = request.form.getlist("selected_files")

    if not selected_files:
        rows = [parse_pdf_summary(path) for path in sorted(input_dir.glob("*.pdf"))]
        return render_template(
            "web_review.html",
            job_id=job_id,
            rows=rows,
            error="Vui lòng chọn ít nhất một file để chuyển.",
        ), 400

    output_dir.mkdir(parents=True, exist_ok=True)
    ok = 0
    fail = 0

    for filename in selected_files:
        pdf_path = input_dir / Path(filename).name
        if not pdf_path.exists() or pdf_path.suffix.lower() != ".pdf":
            continue
        try:
            convert_one_pdf(pdf_path, template_path, output_dir)
            ok += 1
        except Exception:
            fail += 1
            err_file = output_dir / f"LOI_{safe_filename(pdf_path.stem)}.txt"
            err_file.write_text(traceback.format_exc(), encoding="utf-8")

    zip_path = root / "ket_qua_hoc_ba.zip"
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(output_dir.iterdir()):
            if path.is_file():
                zf.write(path, arcname=path.name)

    if ok == 0 and fail == 0:
        abort(400, "Không có file hợp lệ để chuyển.")

    return send_file(
        zip_path,
        as_attachment=True,
        download_name=f"ket_qua_hoc_ba_ok_{ok}_loi_{fail}.zip",
        mimetype="application/zip",
    )


@web.post("/delete/<job_id>")
def delete_job(job_id: str):
    root = job_dir(job_id)
    shutil.rmtree(root, ignore_errors=True)
    return redirect(url_for("index"))


if __name__ == "__main__":
    host = os.environ.get("HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", "8000"))
    web.run(host=host, port=port, debug=os.environ.get("FLASK_DEBUG") == "1")
