"""Shared page-geometry constants for generate_report.py and render_pdf.py."""

ORIENTATION = "landscape"  # "portrait" or "landscape"

_A3_LONG_MM, _A3_SHORT_MM = 420, 297
if ORIENTATION == "portrait":
    PAGE_WIDTH_MM, PAGE_HEIGHT_MM = _A3_SHORT_MM, _A3_LONG_MM
else:
    PAGE_WIDTH_MM, PAGE_HEIGHT_MM = _A3_LONG_MM, _A3_SHORT_MM

PAGE_MARGIN_MM = 10
PAGE_CONTENT_WIDTH_MM = PAGE_WIDTH_MM - 2 * PAGE_MARGIN_MM
PAGE_CONTENT_HEIGHT_MM = PAGE_HEIGHT_MM - 2 * PAGE_MARGIN_MM
