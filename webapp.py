"""
Flask web UI for the Google Consolidated Investigation Report generator.

Endpoints:
    GET  /                                -> main UI page
    POST /upload                          -> accept ZIP upload, start processing job
    GET  /status/<job_id>                 -> poll job status + log lines + stats
    GET  /download/<job_id>/<file_type>   -> download one of the 7 generated outputs
"""

import os
import sys
import uuid
import shutil
import threading
import traceback
from datetime import datetime, timedelta

from flask import (
    Flask,
    jsonify,
    render_template,
    request,
    send_from_directory,
)
from werkzeug.utils import secure_filename

# Load IPINFO_TOKEN (and any other secrets) from a local .env file.
# Falls back silently if python-dotenv isn't installed; the OS env
# vars still take precedence.
try:
    from dotenv import load_dotenv
    _env_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), ".env"
    )
    if os.path.isfile(_env_path):
        load_dotenv(_env_path, override=False)
except ImportError:
    pass

# Make sure we can import the pipeline module from this directory.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from google_investigation import run_pipeline  # noqa: E402

# ---------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------

app = Flask(__name__)

UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
REPORTS_FOLDER = os.path.join(BASE_DIR, "reports")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORTS_FOLDER, exist_ok=True)

MAX_UPLOAD_MB = 2048  # 2 GB cap on uploaded ZIP
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024

ALLOWED_EXT = {".zip"}

JOB_TTL_HOURS = 24

# IP geolocation API token. When set, every job will enrich
# its extracted IPs through ipinfo.io and produce an extra
# IP_Geolocation_Lookup.xlsx file plus an "IP Geolocation"
# sheet inside the consolidated Excel.
IPINFO_TOKEN = os.environ.get("IPINFO_TOKEN", "").strip()

# In-memory job store. A real deployment would use Redis or similar;
# for this single-user investigative tool an in-process dict is fine.
jobs = {}
jobs_lock = threading.Lock()


# All outputs the pipeline produces, in the order they
# should appear on the result screen.
OUTPUT_FILES = [
    {
        "key": "excel_file",
        "type": "excel",
        "label": "Excel (.xlsx)",
        "subtitle": "All sheets: summary, subscriber, IP activity, "
                    "IP details, devices, phone associations",
        "icon": "spreadsheet",
    },
    {
        "key": "docx_file",
        "type": "docx",
        "label": "Word (.docx)",
        "subtitle": "Consolidated investigative narrative report",
        "icon": "document",
    },
    {
        "key": "csv_file",
        "type": "csv",
        "label": "IP Details CSV",
        "subtitle": "Deduplicated, formatted (IP, FROM DATE, TO DATE)",
        "icon": "table",
    },
    {
        "key": "ipv6_word_file",
        "type": "ipv6_word",
        "label": "2405 / 2409 IPv6 Word",
        "subtitle": "IPv6 IPs (2405: / 2409:) - IPV, FROM/TO DATE/TIME",
        "icon": "document",
    },
    {
        "key": "remaining_word_file",
        "type": "remaining_word",
        "label": "Remaining IPs Word",
        "subtitle": "Other IPs (2401: Airtel: DD/MMM/YYYY HH:MM:SS)",
        "icon": "document",
    },
    {
        "key": "ipv6_txt_file",
        "type": "ipv6_txt",
        "label": "2405 / 2409 IPv6 Text",
        "subtitle": "Tab-separated IPv6 IP list",
        "icon": "text",
    },
    {
        "key": "remaining_txt_file",
        "type": "remaining_txt",
        "label": "Remaining IPs Text",
        "subtitle": "Tab-separated remaining IP list",
        "icon": "text",
    },
    {
        "key": "geo_excel_file",
        "type": "ip_geo",
        "label": "IP Geolocation Lookup",
        "subtitle": "Per-IP city / region / country / ISP / coordinates",
        "icon": "spreadsheet",
        "optional": True,
    },
]

# Lookup: file_type -> output metadata
OUTPUT_BY_TYPE = {entry["type"]: entry for entry in OUTPUT_FILES}

# Disk filenames by file_type (must match what the pipeline writes).
FILENAMES_BY_TYPE = {
    "excel": "Consolidated_Google_Investigation_Report.xlsx",
    "docx": "Consolidated_Google_Investigation_Report.docx",
    "csv": "IP_Details_Formatted.csv",
    "ipv6_word": "2405_2409_IP_Details.docx",
    "remaining_word": "Remaining_IP_Details.docx",
    "ipv6_txt": "2405_2409_IP_Details.txt",
    "remaining_txt": "Remaining_IP_Details.txt",
    "ip_geo": "IP_Geolocation_Lookup.xlsx",
}


# ---------------------------------------------------------------
# Job helpers
# ---------------------------------------------------------------

def _new_job(upload_path):
    job_id = uuid.uuid4().hex
    output_dir = os.path.join(REPORTS_FOLDER, job_id)

    with jobs_lock:
        jobs[job_id] = {
            "status": "pending",
            "created": datetime.utcnow(),
            "upload_path": upload_path,
            "output_dir": output_dir,
            "logs": [],
            "result": None,
            "error": None,
            "stats": None,
        }
    return job_id


def _append_log(job_id, line):
    if not line:
        return
    with jobs_lock:
        if job_id in jobs:
            jobs[job_id]["logs"].append(line)
            if len(jobs[job_id]["logs"]) > 5000:
                jobs[job_id]["logs"] = jobs[job_id]["logs"][-5000:]


def _set_field(job_id, key, value):
    with jobs_lock:
        if job_id in jobs:
            jobs[job_id][key] = value


def _job_public_view(job):
    return {
        "status": job["status"],
        "created": job["created"].isoformat() + "Z",
        "logs": list(job["logs"]),
        "result": job["result"],
        "error": job["error"],
        "stats": job["stats"],
    }


def cleanup_old_jobs():
    """Remove jobs older than JOB_TTL_HOURS along with their files."""
    cutoff = datetime.utcnow() - timedelta(hours=JOB_TTL_HOURS)
    to_delete = []
    with jobs_lock:
        for job_id, job in jobs.items():
            if job["created"] < cutoff:
                to_delete.append(job_id)

    for job_id in to_delete:
        with jobs_lock:
            job = jobs.pop(job_id, None)
        if not job:
            continue
        if job.get("upload_path") and os.path.exists(job["upload_path"]):
            try:
                os.remove(job["upload_path"])
            except OSError:
                pass
        out = job.get("output_dir")
        if out and os.path.exists(out):
            shutil.rmtree(out, ignore_errors=True)


# ---------------------------------------------------------------
# Background worker
# ---------------------------------------------------------------

def _run_job(job_id, zip_path, output_dir):
    try:
        _set_field(job_id, "status", "processing")

        def progress(line):
            _append_log(job_id, line)

        result = run_pipeline(
            zip_path=zip_path,
            output_dir=output_dir,
            progress_callback=progress,
            ipinfo_token=IPINFO_TOKEN or None,
        )

        # Build a UI-friendly dict of all outputs with size info.
        result_payload = {}
        for entry in OUTPUT_FILES:
            path = result.get(entry["key"])
            if path and os.path.isfile(path):
                result_payload[entry["type"]] = {
                    "filename": os.path.basename(path),
                    "size": os.path.getsize(path),
                    "label": entry["label"],
                    "subtitle": entry["subtitle"],
                    "icon": entry["icon"],
                    "key": entry["key"],
                }

        _set_field(job_id, "result", result_payload)
        _set_field(job_id, "stats", result["stats"])
        _set_field(job_id, "status", "completed")

    except Exception as exc:
        _append_log(job_id, f"[ERROR] {exc}")
        _append_log(job_id, traceback.format_exc())
        _set_field(job_id, "error", str(exc))
        _set_field(job_id, "status", "failed")
    finally:
        # Free the temporary upload after the job is done.
        if os.path.exists(zip_path):
            try:
                os.remove(zip_path)
            except OSError:
                pass


# ---------------------------------------------------------------
# Routes
# ---------------------------------------------------------------

@app.route("/")
def index():
    return render_template(
        "index.html",
        max_upload_mb=MAX_UPLOAD_MB,
        output_files=OUTPUT_FILES,
        ipinfo_configured=bool(IPINFO_TOKEN),
    )


@app.route("/kyc-intelligence")
def kyc_intelligence():
    """Serve the Cyber North OSINT KYC intelligence page.

    The HTML is fully client-side (uses CDN libraries for ZIP / PDF /
    OCR / XLSX / jsPDF processing). Files never leave the user's
    browser, so no upload endpoint is required.
    """
    return render_template("kyc_intelligence.html")


@app.route("/health")
def health():
    return jsonify({"ok": True, "jobs": len(jobs)})


@app.route("/upload", methods=["POST"])
def upload():
    cleanup_old_jobs()

    if "zipfile" not in request.files:
        return jsonify({"error": "No ZIP file part in the request."}), 400

    file = request.files["zipfile"]
    if not file or file.filename == "":
        return jsonify({"error": "No ZIP file selected."}), 400

    original_name = secure_filename(file.filename) or "upload.zip"
    ext = os.path.splitext(original_name)[1].lower()
    if ext not in ALLOWED_EXT:
        return jsonify({
            "error": f"Invalid file type '{ext}'. Only .zip files are accepted."
        }), 400

    job_id = _new_job(upload_path="")
    saved_path = os.path.join(UPLOAD_FOLDER, f"{job_id}_{original_name}")
    file.save(saved_path)
    _set_field(job_id, "upload_path", saved_path)

    output_dir = os.path.join(REPORTS_FOLDER, job_id)
    os.makedirs(output_dir, exist_ok=True)
    _set_field(job_id, "output_dir", output_dir)

    thread = threading.Thread(
        target=_run_job,
        args=(job_id, saved_path, output_dir),
        daemon=True,
        name=f"job-{job_id}",
    )
    thread.start()

    return jsonify({"job_id": job_id})


@app.route("/status/<job_id>")
def status(job_id):
    with jobs_lock:
        job = jobs.get(job_id)
        if not job:
            return jsonify({"error": "Unknown job id."}), 404
        return jsonify(_job_public_view(job))


@app.route("/download/<job_id>/<file_type>")
def download(job_id, file_type):
    with jobs_lock:
        job = jobs.get(job_id)
        if not job:
            return jsonify({"error": "Unknown job id."}), 404
        if job["status"] != "completed":
            return jsonify({
                "error": f"Job is not completed (status: {job['status']})."
            }), 400
        output_dir = job["output_dir"]

    entry = OUTPUT_BY_TYPE.get(file_type)
    if not entry:
        return jsonify({"error": "Invalid file type."}), 400

    filename = FILENAMES_BY_TYPE.get(file_type)
    if not filename:
        return jsonify({"error": "Invalid file type."}), 400

    full_path = os.path.join(output_dir, filename)
    if not os.path.isfile(full_path):
        return jsonify({"error": "Output file not found."}), 404

    return send_from_directory(
        output_dir,
        filename,
        as_attachment=True,
        download_name=filename,
    )


@app.errorhandler(413)
def too_large(_e):
    return jsonify({
        "error": (
            f"Uploaded file exceeds the {MAX_UPLOAD_MB} MB limit."
        )
    }), 413


# ---------------------------------------------------------------
# Entrypoint
# ---------------------------------------------------------------

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False, threaded=True)