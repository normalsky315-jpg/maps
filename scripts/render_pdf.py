#!/usr/bin/env python3
"""Render the report HTML to a single-page A3-landscape PDF (+ PNG preview).

However many rows/columns the table ends up with, #content is measured
and scaled down (never up) to fit exactly inside the printable page box,
so the output is guaranteed to be one page.

Usage:
    python3 scripts/render_pdf.py [html_path] [pdf_out] [png_out]
"""
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parent.parent
HTML_PATH = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "output" / "guoyang_report.html"
PDF_OUT = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "output" / "guoyang_report.pdf"
PNG_OUT = Path(sys.argv[3]) if len(sys.argv) > 3 else ROOT / "output" / "guoyang_report.png"

PAGE_WIDTH_MM, PAGE_HEIGHT_MM = 420, 297
PAGE_MARGIN_MM = 10
CONTENT_WIDTH_MM = PAGE_WIDTH_MM - 2 * PAGE_MARGIN_MM
CONTENT_HEIGHT_MM = PAGE_HEIGHT_MM - 2 * PAGE_MARGIN_MM
MM_TO_PX = 96 / 25.4  # Chromium print uses 96 CSS px per inch


def main():
    box_w_px = CONTENT_WIDTH_MM * MM_TO_PX
    box_h_px = CONTENT_HEIGHT_MM * MM_TO_PX

    with sync_playwright() as p:
        browser = p.chromium.launch(executable_path="/opt/pw-browsers/chromium")
        page = browser.new_page(viewport={"width": int(box_w_px) + 40, "height": int(box_h_px) + 40})
        page.goto(f"file://{HTML_PATH.resolve()}")
        page.emulate_media(media="print")

        natural = page.evaluate(
            "() => { const c = document.getElementById('content');"
            " return {w: c.scrollWidth, h: c.scrollHeight}; }"
        )
        scale = min(1.0, box_w_px / natural["w"], box_h_px / natural["h"])
        page.evaluate(
            "(scale) => { document.getElementById('content').style.transform = "
            "`scale(${scale})`; }",
            scale,
        )

        PDF_OUT.parent.mkdir(parents=True, exist_ok=True)
        page.pdf(
            path=str(PDF_OUT),
            width=f"{PAGE_WIDTH_MM}mm",
            height=f"{PAGE_HEIGHT_MM}mm",
            print_background=True,
            margin={"top": f"{PAGE_MARGIN_MM}mm", "bottom": f"{PAGE_MARGIN_MM}mm",
                    "left": f"{PAGE_MARGIN_MM}mm", "right": f"{PAGE_MARGIN_MM}mm"},
        )
        page.screenshot(path=str(PNG_OUT), full_page=False)
        browser.close()

    print(f"scale={scale:.3f} natural={natural} -> {PDF_OUT}")


if __name__ == "__main__":
    main()
