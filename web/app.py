"""Stateless web front-end for the incident-report parser.

Design:
  * /parse  - accept an uploaded report PDF, parse it in a per-request temp dir
              (0700, deleted in a finally), and return the .xlsx bytes. Nothing is
              persisted: the browser holds the only copy of the result.
  * /print  - accept the .xlsx the browser is holding, render it to PDF with
              LibreOffice headless and send it to the network printer over IPP.
              Again a per-request temp dir, deleted afterwards.
  * /        - the single-page UI.
  * /healthz - liveness/readiness for Kubernetes.

Privacy: this handles sensitive personal data. We never write outside a private
per-request temp dir, never log file *contents*, and strip the parser's internal
`source_path` so no server-side path leaks into the .xlsx hyperlinks. Access
control (identity) is enforced *in front* by an auth proxy — not here.
"""
from __future__ import annotations

import logging
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from fastapi import FastAPI, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, Response

# The parser lives at the repo root, one level above this file's directory.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import parse_incidents as parser  # noqa: E402

# ---- configuration (all via env so the image is config-free) -----------------
MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", "25"))
MAX_UPLOAD_BYTES = MAX_UPLOAD_MB * 1024 * 1024
# Direct IPP Everywhere endpoint of the network printer, e.g.
# "ipp://<printer-host>/ipp/print". Printing straight to the printer means we don't
# depend on any intermediate CUPS host being powered on.
PRINTER_URI = os.environ.get("PRINTER_URI", "")
IPP_TEMPLATE = Path(__file__).parent / "ipp" / "print-job.test"
SOFFICE_BIN = os.environ.get("SOFFICE_BIN", "soffice")
SOFFICE_TIMEOUT = int(os.environ.get("SOFFICE_TIMEOUT", "120"))
PRINT_TIMEOUT = int(os.environ.get("PRINT_TIMEOUT", "60"))

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("file-parser-web")

app = FastAPI(title="Incident report parser", docs_url=None, redoc_url=None)

_INDEX_HTML = (Path(__file__).parent / "templates" / "index.html").read_text()


def _who(request: Request) -> str:
    """Identity for audit logging — set by Cloudflare Access in front of us."""
    return request.headers.get("Cf-Access-Authenticated-User-Email", "anonymous")


def _private_tmpdir() -> str:
    """A 0700 temp dir for one request's files; caller must rmtree it."""
    return tempfile.mkdtemp(prefix="fp-")


PRINT_SHEET = "Incidents"


def _incidents_only(src: Path) -> Path:
    """Return a copy of the workbook containing only the `Incidents` sheet, so the
    printout omits Summary/Legend/Discrepancies/Skipped Pages. If that sheet isn't
    present (unexpected), fall back to printing the workbook unchanged."""
    from openpyxl import load_workbook

    wb = load_workbook(src)
    if PRINT_SHEET not in wb.sheetnames:
        return src
    for name in [n for n in wb.sheetnames if n != PRINT_SHEET]:
        del wb[name]
    out = src.with_name("print.xlsx")
    wb.save(out)
    return out


async def _save_upload(upload: UploadFile, dest: Path) -> int:
    """Stream an upload to `dest`, enforcing the size cap. Returns bytes written."""
    written = 0
    with dest.open("wb") as fh:
        while chunk := await upload.read(1024 * 1024):
            written += len(chunk)
            if written > MAX_UPLOAD_BYTES:
                raise HTTPException(413, f"File exceeds {MAX_UPLOAD_MB} MB limit.")
            fh.write(chunk)
    return written


@app.get("/", response_class=HTMLResponse)
async def index() -> HTMLResponse:
    return HTMLResponse(_INDEX_HTML)


@app.get("/healthz")
async def healthz() -> JSONResponse:
    return JSONResponse({"status": "ok"})


@app.post("/parse")
async def parse(request: Request, file: UploadFile) -> Response:
    name = file.filename or "report.pdf"
    if not name.lower().endswith(".pdf"):
        raise HTTPException(400, "Please upload a .pdf export of the report.")
    work = _private_tmpdir()
    try:
        pdf_path = Path(work) / "input.pdf"
        size = await _save_upload(file, pdf_path)

        rows, skipped = parser.parse_files([pdf_path])
        if not rows:
            raise HTTPException(
                422,
                "No incidents were extracted. If the PDF looks full, it is probably "
                "a print-to-PDF or scan with no text layer — re-export it from the "
                "report system (or OCR it first).",
            )
        # Strip the parser's internal absolute path so it can't leak into the
        # workbook's Page hyperlinks (those point at this server's temp dir, which
        # the recipient can't open and shouldn't see). write_xlsx skips the
        # hyperlink when source_path is absent.
        for r in rows:
            r.pop("source_path", None)
        for s in skipped:
            s.pop("source_path", None)

        out_path = Path(work) / "incidents.xlsx"
        parser.write_xlsx(rows, out_path, skipped)
        data = out_path.read_bytes()

        log.info("parse user=%s in=%dB rows=%d skipped=%d out=%dB",
                 _who(request), size, len(rows), len(skipped), len(data))

        download_name = Path(name).stem + "-incidents.xlsx"
        return Response(
            content=data,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f'attachment; filename="{download_name}"',
                "X-Incident-Rows": str(len(rows)),
                "X-Skipped-Pages": str(len(skipped)),
            },
        )
    finally:
        shutil.rmtree(work, ignore_errors=True)


@app.post("/print")
async def print_xlsx(request: Request, file: UploadFile) -> JSONResponse:
    if not PRINTER_URI:
        raise HTTPException(503, "Printing is not configured on this server.")
    work = _private_tmpdir()
    try:
        xlsx_path = Path(work) / "incidents.xlsx"
        await _save_upload(file, xlsx_path)

        # Print the "Incidents" sheet only — LibreOffice renders every sheet, so
        # drop the rest from a throwaway copy first (Summary/Legend/Discrepancies/
        # Skipped Pages are for on-screen review, not the printout). The user's
        # downloaded workbook is untouched; this copy exists only to print.
        xlsx_path = _incidents_only(xlsx_path)

        # Render the print-first workbook to PDF. A throwaway LibreOffice profile
        # per call keeps concurrent requests from fighting over one user profile.
        profile = Path(work) / "lo-profile"
        conv = subprocess.run(
            [SOFFICE_BIN, "--headless", "--norestore",
             f"-env:UserInstallation=file://{profile}",
             "--convert-to", "pdf", "--outdir", work, str(xlsx_path)],
            capture_output=True, text=True, timeout=SOFFICE_TIMEOUT,
            env={**os.environ, "HOME": work},
        )
        pdf_path = xlsx_path.with_suffix(".pdf")
        if conv.returncode != 0 or not pdf_path.exists():
            log.error("soffice failed: %s", conv.stderr.strip())
            raise HTTPException(500, "Could not render the workbook for printing.")

        # Submit the PDF straight to the printer's IPP endpoint. ipptool returns
        # non-zero unless the Print-Job's STATUS expectation (successful-ok) is met.
        pr = subprocess.run(
            ["ipptool", "-tv", "-f", str(pdf_path), PRINTER_URI, str(IPP_TEMPLATE)],
            capture_output=True, text=True, timeout=PRINT_TIMEOUT,
        )
        if pr.returncode != 0:
            log.error("ipptool failed: %s", (pr.stderr or pr.stdout).strip())
            raise HTTPException(502, "The printer rejected the job.")

        log.info("print user=%s printer=%s ok", _who(request), PRINTER_URI)
        return JSONResponse({"status": "queued", "detail": "Sent to the printer."})
    finally:
        shutil.rmtree(work, ignore_errors=True)
