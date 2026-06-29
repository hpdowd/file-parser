# Incident report parser — web service image.
#
# Bundles LibreOffice (to render the print-first .xlsx to PDF) and the CUPS
# client (to send that PDF to the office printer). LibreOffice makes this image
# large (~700MB+); that's the price of faithful printing.
FROM python:3.12-slim

# soffice for xlsx->pdf, ipptool (cups-ipp-utils) to submit the PDF straight to
# the network printer over IPP, fonts so the PDF renders with real glyphs.
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-calc \
        cups-ipp-utils \
        fonts-dejavu \
        fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY web/requirements.txt /app/web/requirements.txt
RUN pip install --no-cache-dir -r /app/web/requirements.txt

# The parser (root of the repo) plus the web layer.
COPY parse_incidents.py /app/parse_incidents.py
COPY web /app/web

# Run unprivileged; /tmp is a writable emptyDir in k8s and holds both the
# per-request temp dirs and LibreOffice's throwaway profile + HOME.
RUN useradd --uid 1000 --create-home appuser
USER 1000
ENV HOME=/tmp \
    TMPDIR=/tmp \
    MAX_UPLOAD_MB=25

EXPOSE 8000
CMD ["uvicorn", "web.app:app", "--host", "0.0.0.0", "--port", "8000"]
