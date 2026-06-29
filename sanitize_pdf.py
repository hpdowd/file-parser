#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.11"
# dependencies = ["pdfplumber>=0.11", "reportlab>=4.0"]
# ///
"""
sanitize_pdf.py
===============
Make a *shareable* copy of a structured report PDF by REBUILDING
each page with the same layout — word positions, table lines, the black heading
bars, font sizes approximated — but with every data value MASKED to its shape:

    letters -> x       digits -> 0       punctuation / € / spacing -> kept

The fixed template text (section titles, column headers, and field labels such as
"Local Station:", "Station (Review Station)", "Narrative") is PRESERVED, because
those labels are the report's *structure*, not anyone's data — they are exactly
what makes the example useful for tuning the parser. Anything that is NOT part of
the blank template is masked.

Because each page is rebuilt from scratch (rather than the original with boxes
painted over it) there is NO hidden text layer left underneath — what you see in
the output is all there is. Open the result and eyeball it before sharing it.

Fidelity: word positions, table lines and the filled heading bars are reproduced
faithfully; text is redrawn in Helvetica at the original size, so glyph widths (and
thus inter-word gaps) drift slightly. The layout, labels and field shapes are intact
— which is all the parser's anchor matching needs — but raw text-extraction of the
output may show neighbouring words run together.

Usage
-----
    ./sanitize_pdf.py real_report.pdf                 # -> real_report.sanitized.pdf
    ./sanitize_pdf.py real_report.pdf -o example.pdf
    ./sanitize_pdf.py folder_of_pdfs/*.pdf            # each -> <name>.sanitized.pdf
    ./sanitize_pdf.py real_report.pdf --aggressive    # mask the labels too (wireframe)

How "label vs data" is decided
------------------------------
A word is kept verbatim only if it appears in the blank template (`template.pdf`
next to this script; override with --reference) AND contains no digit. So dates,
names, reference numbers, free-text and amounts — none of which are in the blank
template — are always masked. `--aggressive` masks the labels too, for a pure
layout wireframe.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Optional

import pdfplumber
from reportlab.pdfgen import canvas


# --------------------------------------------------------------------------- #
#  Masking                                                                     #
# --------------------------------------------------------------------------- #
def mask_token(t: str) -> str:
    """Shape-only mask: letters -> x, digits -> 0, everything else (punctuation,
    €, /, -, (), :, whitespace) kept so the layout and field shapes survive."""
    return "".join("x" if c.isalpha() else "0" if c.isdigit() else c for c in t)


def load_label_vocab(reference: Optional[Path]) -> set[str]:
    """Vocabulary of fixed template labels = every word in the blank template that
    has no digit (dates/placeholder numbers and the *** markers are excluded). Any
    real-page word NOT in here is treated as data and masked."""
    vocab: set[str] = set()
    if not reference or not reference.exists():
        print(f"warning: reference template {reference} not found — every word will "
              f"be masked (labels included).", file=sys.stderr)
        return vocab
    try:
        with pdfplumber.open(reference) as pdf:
            for page in pdf.pages:
                for w in page.extract_words():
                    t = w["text"]
                    if "*" in t or any(ch.isdigit() for ch in t):
                        continue
                    vocab.add(t)
    except Exception as exc:
        print(f"warning: could not read reference {reference}: {exc} "
              f"(labels will be masked too)", file=sys.stderr)
    return vocab


# --------------------------------------------------------------------------- #
#  Colour handling (so the black heading bars and any coloured text survive)    #
# --------------------------------------------------------------------------- #
def to_rgb(c) -> Optional[tuple[float, float, float]]:
    """Coerce a pdfplumber colour (gray float, RGB/CMYK tuple, or None) to RGB."""
    try:
        if c is None:
            return None
        if isinstance(c, (int, float)):
            g = min(1.0, max(0.0, float(c)))
            return (g, g, g)
        if isinstance(c, (list, tuple)) and c:
            v = [min(1.0, max(0.0, float(x))) for x in c]
            if len(v) == 1:
                return (v[0], v[0], v[0])
            if len(v) == 3:
                return (v[0], v[1], v[2])
            if len(v) == 4:  # CMYK -> RGB
                cy, mg, yl, k = v
                return (1 - min(1.0, cy + k), 1 - min(1.0, mg + k), 1 - min(1.0, yl + k))
    except Exception:
        return None
    return None


def sample_text_color(chars: list, w: dict) -> Optional[tuple[float, float, float]]:
    """Colour of a word = colour of the page char nearest its top-left corner.
    Keeps white-on-black banner text white without special-casing the bars."""
    wx, wt = w["x0"], w["top"]
    best, best_d = None, 1e9
    for ch in chars:
        d = abs(ch["x0"] - wx) + abs(ch["top"] - wt)
        if d < best_d:
            best, best_d = ch, d
            if d < 0.5:
                break
    return to_rgb(best.get("non_stroking_color")) if best else None


# --------------------------------------------------------------------------- #
#  Page rebuild                                                                #
# --------------------------------------------------------------------------- #
def force_mask_bands(page) -> list[tuple[float, float]]:
    """Vertical bands whose words are ALWAYS masked, even if they match a template
    label. The Narrative is free prose (the most sensitive field), so a word that
    happens to equal a column header — "Force", "Case", … — must not slip through.
    Band = from just below the "Narrative" label down to "Quality Control Tests"
    (the next section) or the page bottom."""
    words = page.extract_words()
    nar = next((w for w in words if w["text"] == "Narrative"), None)
    if not nar:
        return []
    qct = next((w for w in words if w["text"] == "Quality" and w["top"] > nar["top"]), None)
    return [(nar["bottom"] + 1, (qct["top"] - 1) if qct else float(page.height))]


def render_page(c: canvas.Canvas, page, vocab: set[str], aggressive: bool,
                stats: dict) -> None:
    W, H = float(page.width), float(page.height)
    c.setPageSize((W, H))
    bands = force_mask_bands(page)

    def y(top: float) -> float:  # pdfplumber top-origin -> reportlab bottom-origin
        return H - top

    # 1) Filled rectangles first — this is what draws the black heading bars and
    #    any shaded table cells. Layout, not data, so reproduced faithfully.
    for r in page.rects:
        x0, x1, top, bottom = r["x0"], r["x1"], r["top"], r["bottom"]
        w, h = x1 - x0, bottom - top
        fill = to_rgb(r.get("non_stroking_color")) if r.get("fill") else None
        if fill is not None:
            c.setFillColorRGB(*fill)
            c.rect(x0, y(bottom), w, h, stroke=0, fill=1)
        edge = to_rgb(r.get("stroking_color")) if r.get("stroke") else None
        if edge is not None:
            c.setStrokeColorRGB(*edge)
            c.setLineWidth(float(r.get("linewidth") or 0.5))
            c.rect(x0, y(bottom), w, h, stroke=1, fill=0)

    # 2) Lines (table borders).
    c.setStrokeColorRGB(0.0, 0.0, 0.0)
    for ln in page.lines:
        c.setLineWidth(float(ln.get("linewidth") or 0.5))
        c.line(ln["x0"], y(ln["top"]), ln["x1"], y(ln["bottom"]))

    # 3) Images are never copied (a scanned signature could be personal data) —
    #    just outline where one sat so the layout is still legible.
    for im in page.images:
        c.setStrokeColorRGB(0.6, 0.6, 0.6)
        c.setLineWidth(0.5)
        c.rect(im["x0"], y(im["bottom"]), im["x1"] - im["x0"],
               im["bottom"] - im["top"], stroke=1, fill=0)

    # 4) Text: keep template labels, mask everything else.
    chars = page.chars
    for w in page.extract_words(extra_attrs=["size", "fontname"]):
        t = w["text"]
        in_prose = any(lo <= w["top"] <= hi for lo, hi in bands)
        keep = (not aggressive and not in_prose
                and t in vocab and not any(ch.isdigit() for ch in t))
        out = t if keep else mask_token(t)
        stats["kept" if keep else "masked"] += 1

        rgb = sample_text_color(chars, w) or (0.0, 0.0, 0.0)
        c.setFillColorRGB(*rgb)
        size = float(w.get("size") or (w["bottom"] - w["top"]))
        c.setFont("Helvetica", max(4.0, size))
        baseline = w["bottom"] - (w["bottom"] - w["top"]) * 0.18
        try:
            c.drawString(w["x0"], y(baseline), out)
        except Exception:  # unencodable glyph -> drop to ASCII rather than crash
            c.drawString(w["x0"], y(baseline), out.encode("ascii", "replace").decode())

    c.showPage()


def sanitize(src: Path, dst: Path, vocab: set[str], aggressive: bool) -> dict:
    stats = {"kept": 0, "masked": 0, "pages": 0}
    c = canvas.Canvas(str(dst))
    with pdfplumber.open(src) as pdf:
        for page in pdf.pages:
            render_page(c, page, vocab, aggressive, stats)
            stats["pages"] += 1
    c.save()
    return stats


# --------------------------------------------------------------------------- #
#  CLI                                                                         #
# --------------------------------------------------------------------------- #
def main(argv: Optional[list[str]] = None) -> int:
    ap = argparse.ArgumentParser(
        description="Rebuild report PDFs with the layout kept but all data "
        "masked to its shape, so a full-page example is safe to share.")
    ap.add_argument("inputs", nargs="+", help="PDF file(s) to sanitize")
    ap.add_argument("-o", "--output", help="output path (only with a single input)")
    ap.add_argument("--reference", help="blank template used to recognise labels "
                    "(default: template.pdf next to this script)")
    ap.add_argument("--aggressive", action="store_true",
                    help="mask the template labels too (pure layout wireframe)")
    args = ap.parse_args(argv)

    paths = [Path(p) for p in args.inputs]
    missing = [p for p in paths if not (p.is_file() and p.suffix.lower() == ".pdf")]
    if missing:
        print(f"error: not a PDF file: {', '.join(map(str, missing))}", file=sys.stderr)
        return 1
    if args.output and len(paths) > 1:
        ap.error("-o/--output only works with a single input PDF")

    ref = Path(args.reference) if args.reference else Path(__file__).with_name("template.pdf")
    vocab = load_label_vocab(None if args.aggressive else ref)
    if not args.aggressive and not vocab:
        print("note: no template labels loaded — output will mask everything "
              "(use --reference to point at the blank template).")

    rc = 0
    for src in paths:
        dst = Path(args.output) if args.output else src.with_name(src.stem + ".sanitized.pdf")
        try:
            st = sanitize(src, dst, vocab, args.aggressive)
        except Exception as exc:
            print(f"error: failed on {src}: {exc}", file=sys.stderr)
            rc = 1
            continue
        total = st["kept"] + st["masked"]
        print(f"{src.name}: {st['pages']} page(s), {st['masked']}/{total} words masked, "
              f"{st['kept']} template labels kept -> {dst}")
    print("\nReview the output (it is text-based — open it, or run "
          "`./parse_incidents.py <out> --inspect`) and confirm no real data remains "
          "before sharing.")
    return rc


if __name__ == "__main__":
    raise SystemExit(main())
