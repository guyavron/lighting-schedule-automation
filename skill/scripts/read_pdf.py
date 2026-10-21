#!/usr/bin/env python3
"""
Extract per-page text and images from a manufacturer lighting-fixture PDF,
so Claude can read the text and decide field values itself.

Usage:
    python3 read_pdf.py <pdf_path> [--out <dir>]

Prints per-page text to stdout, and writes rendered page images + any
embedded raster images to <dir> (default: ./pdf_extract/<pdf-stem>/).
"""
import argparse
import subprocess
import sys
import zipfile
import json
from pathlib import Path

# Hebrew output on Windows too, when stdout is a pipe (the default there is cp1252)
for _s in (sys.stdout, sys.stderr):
    _s.reconfigure(encoding="utf-8", errors="replace")

MIN_TEXT_CHARS = 40   # less than this on all pages = no usable text layer

# exit codes, so a caller (Claude, or a script) can tell the cases apart
EXIT_NOT_FOUND, EXIT_NOT_PDF, EXIT_BROKEN, EXIT_NO_TEXT = 2, 3, 4, 5


def is_manifest_zip(pdf_path: Path) -> bool:
    try:
        with zipfile.ZipFile(pdf_path) as z:
            return "manifest.json" in z.namelist()
    except zipfile.BadZipFile:
        return False


def _fail(code, message):
    print(f"\nERROR: {message}", file=sys.stderr)
    return code


def _render_pages(pdf_path: Path, out_dir: Path):
    """rendered page images (for visual confirmation of the right product photo)"""
    subprocess.run(["pdftoppm", "-jpeg", "-r", "120", str(pdf_path), str(out_dir / "page")],
                   check=False, capture_output=True)


def handle_manifest_zip(pdf_path: Path, out_dir: Path):
    """Some project-attached files are stored as a page-per-file archive
    (manifest.json + N.txt + N.jpeg) rather than raw PDF bytes."""
    out_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(pdf_path) as z:
        z.extractall(out_dir)
    manifest = json.loads((out_dir / "manifest.json").read_text(encoding="utf-8"))
    for page in manifest["pages"]:
        n = page["page_number"]
        txt_path = out_dir / page["text"]["path"] if isinstance(page["text"], dict) else out_dir / f"{n}.txt"
        text = txt_path.read_text(encoding="utf-8", errors="replace")
        img_path = out_dir / page["image"]["path"]
        print(f"\n===== PAGE {n} (image: {img_path}) =====")
        print(text)


def handle_real_pdf(pdf_path: Path, out_dir: Path):
    with open(pdf_path, "rb") as f:
        head = f.read(1024)
    if b"%PDF-" not in head:
        sys.exit(_fail(EXIT_NOT_PDF, f"{pdf_path.name} is not a PDF (no %PDF header) -- "
                       "maybe a saved web page or an image renamed to .pdf. Nothing was extracted."))
    out_dir.mkdir(parents=True, exist_ok=True)
    # per-page text
    txt_out = out_dir / "text.txt"
    r = subprocess.run(["pdftotext", "-layout", str(pdf_path), str(txt_out)],
                       capture_output=True, text=True, errors="replace")
    if r.returncode != 0 or not txt_out.exists():
        sys.exit(_fail(EXIT_BROKEN, f"{pdf_path.name} is damaged or incomplete (pdftotext: "
                       f"{(r.stderr or '').strip()[:200]}). Ask for a fresh copy of the spec sheet; "
                       "do not fill any field from it."))
    text = txt_out.read_text(encoding="utf-8", errors="replace")
    print(text)
    if len("".join(text.split())) < MIN_TEXT_CHARS:
        # still render the pages (below) so a human can look -- but say it loudly
        _render_pages(pdf_path, out_dir)
        sys.exit(_fail(EXIT_NO_TEXT, f"{pdf_path.name} has no text layer (scanned or image-only). "
                       f"Page images were written to {out_dir} for a human to read. "
                       "Every field read from an image must be flagged as uncertain -- never typed in as confirmed."))

    _render_pages(pdf_path, out_dir)
    # embedded raster images (often higher quality than the page render)
    img_prefix = out_dir / "embedded"
    subprocess.run(
        ["pdfimages", "-png", str(pdf_path), str(img_prefix)],
        check=False,
    )
    print(f"\n[page images and embedded images written to {out_dir}]")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf_path")
    ap.add_argument("--out", default=None)
    args = ap.parse_args()

    pdf_path = Path(args.pdf_path)
    if not pdf_path.exists():
        sys.exit(_fail(EXIT_NOT_FOUND, f"File not found: {pdf_path}"))
    out_dir = Path(args.out) if args.out else Path("pdf_extract") / pdf_path.stem

    if is_manifest_zip(pdf_path):
        handle_manifest_zip(pdf_path, out_dir)
    else:
        handle_real_pdf(pdf_path, out_dir)


if __name__ == "__main__":
    main()
