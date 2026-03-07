import io
import os
import re
import threading
import uuid
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import requests
from flask import Flask, jsonify, render_template, request, send_file

from generate_report import generate_report_pdf

BASE_DIR = Path(__file__).resolve().parent

INPI_BASE_URL = os.getenv("INPI_BASE_URL", "https://registre-national-entreprises.inpi.fr/api")
INPI_USERNAME = os.getenv("INPI_USERNAME", "")
INPI_PASSWORD = os.getenv("INPI_PASSWORD", "")
MISTRAL_BASE_URL = os.getenv("MISTRAL_BASE_URL", "https://api.mistral.ai/v1")
MISTRAL_API_KEY = os.getenv("MISTRAL_API_KEY", "")
MISTRAL_OCR_MODEL = os.getenv("MISTRAL_OCR_MODEL", "mistral-ocr-latest")
PORT = int(os.getenv("PORT", "5000"))

app = Flask(__name__, template_folder=str(BASE_DIR / "templates"))
http = requests.Session()


@dataclass
class Job:
    siren: str
    id: str = field(default_factory=lambda: str(uuid.uuid4())[:8])
    status: str = "pending"
    message: str = ""
    progress: int = 0
    created_at: str = field(default_factory=lambda: datetime.now().isoformat(timespec="seconds"))
    zip_bytes: Optional[bytes] = None
    error: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return {
            "id": self.id,
            "siren": self.siren,
            "status": self.status,
            "message": self.message,
            "progress": self.progress,
            "created_at": self.created_at,
            "error": self.error,
        }


jobs: Dict[str, Job] = {}
jobs_lock = threading.Lock()


def sanitize_name(value: str) -> str:
    cleaned = re.sub(r"[\\/:*?\"<>|]", "", value or "document")
    cleaned = re.sub(r"\s+", " ", cleaned).strip().rstrip(".")
    return cleaned or "document"


def _require_env(name: str, value: str) -> None:
    if not value.strip():
        raise ValueError(f"Variable d'environnement manquante: {name}")


def inpi_login() -> str:
    _require_env("INPI_USERNAME", INPI_USERNAME)
    _require_env("INPI_PASSWORD", INPI_PASSWORD)
    response = http.post(
        f"{INPI_BASE_URL}/sso/login",
        json={"username": INPI_USERNAME, "password": INPI_PASSWORD},
        timeout=30,
    )
    response.raise_for_status()
    token = response.json().get("token")
    if not token:
        raise ValueError("Token INPI introuvable")
    return token


def _walk_documents(node: Any) -> List[Dict[str, Any]]:
    found: List[Dict[str, Any]] = []
    if isinstance(node, dict):
        endpoint = node.get("endpoint") or node.get("downloadLink") or node.get("path")
        if isinstance(endpoint, str) and endpoint.strip():
            found.append(node)
        for value in node.values():
            found.extend(_walk_documents(value))
    elif isinstance(node, list):
        for item in node:
            found.extend(_walk_documents(item))
    return found


def fetch_company_documents(siren: str) -> Dict[str, List[Dict[str, Any]]]:
    token = inpi_login()
    response = http.get(
        f"{INPI_BASE_URL}/companies/{siren}/attachments",
        headers={"Authorization": f"Bearer {token}"},
        timeout=60,
    )
    response.raise_for_status()
    payload = response.json()

    bilans: List[Dict[str, Any]] = []
    actes: List[Dict[str, Any]] = []

    for doc in _walk_documents(payload):
        doc_type = str(doc.get("typeDocument") or doc.get("type") or "").lower()
        family = str(doc.get("famille") or doc.get("category") or "").lower()
        endpoint = doc.get("endpoint") or doc.get("downloadLink") or doc.get("path")
        filename = doc.get("filename") or doc.get("nomFichier") or doc.get("name") or "document.pdf"

        normalized = {
            "endpoint": endpoint,
            "filename": filename,
            "meta": doc,
        }
        if "acte" in doc_type or "acte" in family:
            actes.append(normalized)
        else:
            bilans.append(normalized)

    return {"bilans": bilans, "actes": actes}


def download_document(endpoint: str, token: str) -> bytes:
    if endpoint.startswith("http://") or endpoint.startswith("https://"):
        url = endpoint
    else:
        url = f"{INPI_BASE_URL}/{endpoint.lstrip('/')}"

    response = http.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=90)
    response.raise_for_status()
    return response.content


def mistral_upload_pdf(pdf_bytes: bytes, filename: str) -> str:
    _require_env("MISTRAL_API_KEY", MISTRAL_API_KEY)
    response = http.post(
        f"{MISTRAL_BASE_URL}/files",
        headers={"Authorization": f"Bearer {MISTRAL_API_KEY}"},
        data={"purpose": "ocr"},
        files={"file": (filename, pdf_bytes, "application/pdf")},
        timeout=120,
    )
    response.raise_for_status()
    file_id = response.json().get("id")
    if not file_id:
        raise ValueError("file_id Mistral introuvable")
    return file_id


def mistral_ocr_pdf_markdown(file_id: str) -> str:
    _require_env("MISTRAL_API_KEY", MISTRAL_API_KEY)
    response = http.post(
        f"{MISTRAL_BASE_URL}/ocr",
        headers={
            "Authorization": f"Bearer {MISTRAL_API_KEY}",
            "Content-Type": "application/json",
        },
        json={
            "model": MISTRAL_OCR_MODEL,
            "document": {"type": "file", "file_id": file_id},
            "include_image_base64": False,
        },
        timeout=180,
    )
    response.raise_for_status()
    payload = response.json()

    markdown_chunks: List[str] = []
    pages = payload.get("pages", [])
    for page in pages:
        text = page.get("markdown") or page.get("text") or ""
        if text:
            markdown_chunks.append(text)

    if markdown_chunks:
        return "\n\n".join(markdown_chunks)

    fallback = payload.get("markdown") or payload.get("text") or ""
    return str(fallback)


def process_one_document(document: Dict[str, Any], token: str) -> Dict[str, Any]:
    filename = sanitize_name(str(document.get("filename", "document.pdf")))
    filename_base = filename[:-4] if filename.lower().endswith(".pdf") else filename
    endpoint = str(document.get("endpoint", ""))

    pdf_bytes = download_document(endpoint, token)
    file_id = mistral_upload_pdf(pdf_bytes, f"{filename_base}.pdf")
    markdown = mistral_ocr_pdf_markdown(file_id)

    return {
        "filename_base": filename_base,
        "markdown": markdown,
        "meta": document.get("meta", {}),
        "pdf_bytes": pdf_bytes,
    }


def _build_zip(
    siren: str,
    denomination: str,
    bilans_results: List[Dict[str, Any]],
    actes_results: List[Dict[str, Any]],
    report_pdf: bytes,
) -> bytes:
    memory_file = io.BytesIO()
    with zipfile.ZipFile(memory_file, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for item in actes_results:
            name = sanitize_name(item["filename_base"])
            zf.writestr(f"ACTES/{name}.pdf", item["pdf_bytes"])
            zf.writestr(f"ACTES_MD/{name}.md", item["markdown"])

        for item in bilans_results:
            name = sanitize_name(item["filename_base"])
            zf.writestr(f"COMPTES_ANNUELS/{name}.pdf", item["pdf_bytes"])
            zf.writestr(f"COMPTES_ANNUELS_MD/{name}.md", item["markdown"])

        report_name = f"{datetime.now():%Y-%m-%d}_{siren}_{sanitize_name(denomination)}_Rapport_Analyse_RNE.pdf"
        zf.writestr(f"RAPPORT/{report_name}", report_pdf)

    return memory_file.getvalue()


def _run_collect(job_id: str) -> None:
    with jobs_lock:
        job = jobs[job_id]
        job.status = "running"
        job.message = "Connexion INPI"
        job.progress = 5

    try:
        token = inpi_login()
        company_response = http.get(
            f"{INPI_BASE_URL}/companies/{job.siren}",
            headers={"Authorization": f"Bearer {token}"},
            timeout=60,
        )
        company_response.raise_for_status()
        company_data = company_response.json()

        with jobs_lock:
            job.message = "Lecture des documents"
            job.progress = 15

        docs = fetch_company_documents(job.siren)
        bilans = docs["bilans"]
        actes = docs["actes"]
        total = max(len(bilans) + len(actes), 1)

        bilans_results: List[Dict[str, Any]] = []
        actes_results: List[Dict[str, Any]] = []
        done = 0

        for document in actes:
            result = process_one_document(document, token)
            actes_results.append(result)
            done += 1
            with jobs_lock:
                job.message = f"OCR acte {done}/{total}"
                job.progress = 15 + int((done / total) * 70)

        for document in bilans:
            result = process_one_document(document, token)
            bilans_results.append(result)
            done += 1
            with jobs_lock:
                job.message = f"OCR bilan {done}/{total}"
                job.progress = 15 + int((done / total) * 70)

        report_input_docs = []
        for result in actes_results:
            report_input_docs.append({"famille": "ACTE", "filename_base": result["filename_base"], "meta": result["meta"]})
        for result in bilans_results:
            report_input_docs.append({"famille": "COMPTES_ANNUELS", "filename_base": result["filename_base"], "meta": result["meta"]})

        denomination = (
            company_data.get("formality", {})
            .get("content", {})
            .get("personneMorale", {})
            .get("identite", {})
            .get("entreprise", {})
            .get("denomination")
            or f"SIREN_{job.siren}"
        )

        with jobs_lock:
            job.message = "Génération du rapport PDF"
            job.progress = 90

        report_pdf = generate_report_pdf(
            siren=job.siren,
            denomination=denomination,
            rne_data=company_data,
            doc_results=report_input_docs,
            run_date=datetime.now().strftime("%Y-%m-%d"),
        )

        zip_bytes = _build_zip(job.siren, denomination, bilans_results, actes_results, report_pdf)

        with jobs_lock:
            job.zip_bytes = zip_bytes
            job.status = "done"
            job.message = "Terminé"
            job.progress = 100

    except Exception as exc:
        with jobs_lock:
            job.status = "failed"
            job.error = str(exc)
            job.message = "Erreur"


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/collect")
@app.post("/api/collect")
def collect():
    payload = request.get_json(silent=True) or {}
    siren = str(payload.get("siren", "")).strip()
    if not re.fullmatch(r"\d{9}", siren):
        return jsonify({"error": "SIREN invalide"}), 400

    job = Job(siren=siren)
    with jobs_lock:
        jobs[job.id] = job

    threading.Thread(target=_run_collect, args=(job.id,), daemon=True).start()
    return jsonify({"job_id": job.id, "status": job.status})


@app.get("/status/<job_id>")
@app.get("/api/status/<job_id>")
def status(job_id: str):
    with jobs_lock:
        job = jobs.get(job_id)
        if not job:
            return jsonify({"error": "Job introuvable"}), 404
        return jsonify(job.to_dict())


@app.get("/download/<job_id>")
@app.get("/api/download/<job_id>")
def download(job_id: str):
    with jobs_lock:
        job = jobs.get(job_id)
        if not job:
            return jsonify({"error": "Job introuvable"}), 404
        if job.status != "done" or not job.zip_bytes:
            return jsonify({"error": "ZIP non prêt"}), 409
        zip_bytes = job.zip_bytes
        siren = job.siren

    return send_file(
        io.BytesIO(zip_bytes),
        mimetype="application/zip",
        as_attachment=True,
        download_name=f"RNE_{siren}_{datetime.now():%Y%m%d_%H%M}.zip",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=PORT, debug=True)
