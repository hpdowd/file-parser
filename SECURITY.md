# Security

This document describes the security model of the incident-report parser and its
optional web service, the controls that are built in, and how to deploy it safely.

## Reporting a vulnerability

Please report suspected vulnerabilities privately via GitHub **Security → Report a
vulnerability** (private advisory) rather than a public issue. Do not include real
report data or personal information in a report — this tool exists precisely to keep
that data off third-party systems.

## Data handling

- The CLI never transmits anything: it reads a PDF and writes a spreadsheet locally.
- The web service is **stateless**. Each request uses a private `0700` temp
  directory that is deleted in a `finally`; the parsed workbook is returned to the
  browser and **not** persisted server-side. The parser's internal source path is
  stripped so no server path leaks into the workbook's page hyperlinks.
- Logs record only the caller identity and byte/row counts — **never file contents**.

## Built-in controls

### Spreadsheet / CSV formula injection
Values extracted from an uploaded PDF are treated as untrusted. A cell whose text
begins with a formula trigger (`= + - @` or a control character) is neutralised with
a leading apostrophe before being written (`parse_incidents._formula_safe`). This
matters because a leading `=` is otherwise stored as a **live formula** — dangerous
both when the workbook is opened in Excel/LibreOffice and, in the web service,
because the **Print** path renders the workbook through LibreOffice server-side
(e.g. `=WEBSERVICE(...)` would become server-side request forgery). Numbers are
unaffected, and legitimate incident data never begins with these characters.

### Upload limits and resource caps
- Uploads are streamed with a hard size cap (`MAX_UPLOAD_MB`, default 25).
- CPU/subprocess-heavy work (PDF parsing, LibreOffice rendering) is bounded by a
  concurrency semaphore (`MAX_CONCURRENCY`, default 2) so a burst of large or complex
  inputs cannot exhaust the process.
- External tools (LibreOffice, `ipptool`) run with timeouts, in their own process
  group, so a hang is killed as a whole tree rather than leaving orphans.

### Print is bound to server-produced output
The **Print** endpoint only accepts a workbook that this server produced. `/parse`
returns an HMAC over the workbook bytes; `/print` re-computes and compares it
(constant-time) **before** the file reaches openpyxl or LibreOffice, so an arbitrary
document cannot be pushed through the renderer. Set `PRINT_SIGNING_KEY` to a stable
random value for multi-replica deployments or to allow printing a download after a
restart; otherwise an ephemeral per-process key is used.

### Authentication (defence in depth)
The service performs **no identity checks by itself** — it is designed to run behind
an identity-aware reverse proxy (e.g. Cloudflare Access). As an additional layer, the
app can verify the proxy's signed assertion itself: set `ACCESS_TEAM_DOMAIN` and
`ACCESS_AUD` to enable cryptographic verification of the `Cf-Access-Jwt-Assertion`
JWT (signature, audience, issuer). When enabled, a request that reaches the pod
directly — bypassing the proxy — is rejected (`401`), and the audit-log identity is
taken from the **verified** token rather than a spoofable header. When disabled a
warning is logged; enabling it is strongly recommended in production.

### Response hardening
Every response carries `X-Content-Type-Options: nosniff`, `X-Frame-Options: DENY`,
`Referrer-Policy: no-referrer`, `Cross-Origin-Opener-Policy: same-origin`, a
`Permissions-Policy`, and a strict **Content-Security-Policy** with a per-request
nonce (no `unsafe-inline`). The interactive API docs are disabled.

### Region map at rest (CLI)
The optional Sec-19 region map can be encrypted at rest with AES-256-GCM keyed by a
scrypt-derived passphrase (`--encrypt-region-map`; passphrase read via `getpass`,
never a CLI argument). A wrong passphrase or tampered file fails loudly (GCM tag
check) instead of returning partial data. In the web service the map is supplied as a
mounted secret and read once at startup — no passphrase lives in the app.

## Configuration reference (web service)

| Env var | Default | Purpose |
|---|---|---|
| `MAX_UPLOAD_MB` | `25` | Upload size cap. |
| `MAX_CONCURRENCY` | `2` | Concurrent heavy requests. |
| `PRINT_SIGNING_KEY` | *(ephemeral)* | Stable key for the /print binding HMAC. |
| `ACCESS_TEAM_DOMAIN` | *(unset)* | Identity-proxy team domain; enables JWT verification with `ACCESS_AUD`. |
| `ACCESS_AUD` | *(unset)* | Expected audience (application) tag of the assertion. |
| `PRINTER_URI` | *(unset)* | IPP endpoint; `ipps://` preferred so output is encrypted on the wire. |
| `REGION_MAP_PATH` | `/etc/region-map/region_stations.md` | Optional region map (mounted secret). |
| `SOFFICE_TIMEOUT` / `PRINT_TIMEOUT` | `120` / `60` | Render / print timeouts (seconds). |

## Deployment hardening checklist

- [ ] Put the service behind an identity-aware proxy **and** enable in-app JWT
      verification (`ACCESS_TEAM_DOMAIN` + `ACCESS_AUD`).
- [ ] Restrict pod ingress with a NetworkPolicy so only the ingress controller can
      reach it (prevents direct in-cluster access that would bypass the proxy).
- [ ] Set CPU/memory `limits` on the pod; run non-root with a read-only root
      filesystem and a writable `emptyDir` for `/tmp` (memory-backed).
- [ ] Set a stable `PRINT_SIGNING_KEY` (secret) if running more than one replica.
- [ ] Use `ipps://` to the printer if it supports TLS, so rendered output is not sent
      in clear text over the network.
- [ ] Pin the base image by digest and install Python deps from a hash-pinned
      lockfile (`uv pip compile --generate-hashes` → `pip install --require-hashes`).
- [ ] Pin GitHub Actions to commit SHAs rather than mutable tags.

## Supply chain

Dependencies in `requirements.txt` / `web/requirements.txt` are declared as minimum
versions. For reproducible, tamper-evident builds, generate a hash-pinned lockfile
and pin the container base image by digest (see the checklist above). CI uses the
job's built-in token with least privilege (`packages: write` only) and builds only on
the canonical host.

## No secrets or data in the repository

All report data file types are gitignored; only source, docs, and a blank template
are tracked. Do not commit real report data, region maps, or infrastructure secrets.
