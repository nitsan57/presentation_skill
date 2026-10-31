#!/usr/bin/env python3
"""Optional visual QA helper.

Converts a .pptx to PDF (and, if pdftoppm is available, to PNG keyframes) using
LibreOffice. Skips gracefully when the tooling is absent.
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser(description="Render a .pptx preview (optional).")
    parser.add_argument("deck", help="path to the .pptx file")
    parser.add_argument("-o", "--outdir", required=True, help="output directory")
    parser.add_argument("--dpi", type=int, default=110, help="PNG resolution")
    args = parser.parse_args()

    deck = Path(args.deck)
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    if not deck.is_file():
        print(f"error: deck not found: {deck}", file=sys.stderr)
        return 2

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("skipped: LibreOffice (soffice) not available; preview not rendered.")
        return 0

    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(deck)],
        check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    pdf = outdir / (deck.stem + ".pdf")
    if not pdf.is_file():
        print("skipped: PDF conversion failed.")
        return 0
    print(f"Rendered PDF: {pdf}")

    pdftoppm = shutil.which("pdftoppm")
    if pdftoppm:
        subprocess.run(
            [pdftoppm, "-png", "-r", str(args.dpi), str(pdf), str(outdir / "slide")],
            check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        print(f"Rendered PNG keyframes in: {outdir}")
    else:
        print("skipped: pdftoppm not available; PNG keyframes not rendered.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
