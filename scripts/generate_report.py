#!/usr/bin/env python3
"""Generate the 國揚鉑御 樓層/單價 實價登錄 對照表 (HTML, single A3 page) from the
raw 實價登錄 Excel export. Page orientation is set in report_config.py.

Usage:
    python3 scripts/generate_report.py [xlsx_path] [html_out]
"""
import re
import sys
from datetime import date
from pathlib import Path

import openpyxl

from report_config import (
    ORIENTATION,
    PAGE_CONTENT_HEIGHT_MM,
    PAGE_CONTENT_WIDTH_MM,
    PAGE_MARGIN_MM,
)

ROOT = Path(__file__).resolve().parent.parent
XLSX_PATH = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "data" / "guoyang_boyu_實價登錄.xlsx"
HTML_OUT = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "output" / "guoyang_report.html"

BUILDING_NAME = "國揚鉑御"
CITY_TITLE = "高雄市 國揚鉑御社區實價控表"
FLOOR_MIN, FLOOR_MAX = 2, 20

# #content has a fixed intrinsic design size (roughly matching the target
# page's aspect ratio) and every floor row gets the same explicit height,
# so the grid is perfectly uniform. render_pdf.py then measures this
# intrinsic box and scales it (up or down) to exactly fill one printable
# page, however many rows/columns end up in the table.
CONTENT_WIDTH_PX = 820
ROW_HEIGHT_PX = 52
THEAD_HEIGHT_PX = 44

TIER_HIGH = (70.0, float("inf"), "#fde3e0", "#c0392b", "70\n萬以上")
TIER_MID = (65.0, 69.9999, "#fbe6cf", "#c07a1e", "65-\n69.99萬")
TIER_LOW = (60.0, 64.9999, "#e2ecfa", "#1f5fa8", "60-\n64.99萬")
TIERS = [TIER_HIGH, TIER_MID, TIER_LOW]


def parse_rows(xlsx_path: Path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    records = []
    for r in rows[2:]:
        if not r or not r[1]:
            continue
        m = re.match(r"([AB])棟(\d+)號", str(r[1]))
        if not m:
            continue
        col = f"{m.group(1)}{m.group(2)}"
        fm = re.search(r"(\d+)\s*/\s*(\d+)", str(r[9]))
        floor = int(fm.group(1))
        total_price = int(r[4])
        unit_price = float(r[5])
        area = float(r[6])
        parking_price = r[13]
        parking_price = int(parking_price) if parking_price not in (None, "") else None
        trade_date = str(r[3]).strip()
        records.append(
            dict(
                col=col,
                floor=floor,
                total_price=total_price,
                unit_price=unit_price,
                area=area,
                parking_price=parking_price,
                trade_date=trade_date,
            )
        )
    return records


def column_sort_key(col: str):
    m = re.match(r"([AB])(\d+)", col)
    return (m.group(1), int(m.group(2)))


def tier_for(unit_price: float):
    for lo, hi, bg, fg, _label in TIERS:
        if lo <= unit_price <= hi:
            return bg, fg
    return "#eef1f4", "#33475b"


def build_html(records):
    columns = sorted({r["col"] for r in records}, key=column_sort_key)
    grid = {(r["col"], r["floor"]): r for r in records}

    count = len(records)
    max_price = max(r["unit_price"] for r in records)
    min_price = min(r["unit_price"] for r in records)
    avg_price = sum(r["unit_price"] for r in records) / count
    avg_area = sum(r["area"] for r in records) / count

    update_month = date.today().strftime("%Y/%m")

    # A unit column whose every transaction shares the same 坪數 gets that
    # figure hoisted up into the column header (once) instead of repeated
    # in every cell's footnote.
    areas_by_col = {}
    for r in records:
        areas_by_col.setdefault(r["col"], set()).add(round(r["area"], 2))
    fixed_area = {col: next(iter(areas)) for col, areas in areas_by_col.items() if len(areas) == 1}

    head_cells = "".join(
        f'<th>{c}<span class="col-area">{fixed_area[c]:.2f}坪</span></th>' if c in fixed_area else f"<th>{c}</th>"
        for c in columns
    )

    body_rows = []
    for floor in range(FLOOR_MAX, FLOOR_MIN - 1, -1):
        cells = []
        for col in columns:
            rec = grid.get((col, floor))
            if not rec:
                cells.append('<td class="empty">—</td>')
                continue
            bg, fg = tier_for(rec["unit_price"])
            parking = f'{rec["parking_price"]}萬' if rec["parking_price"] is not None else "無"
            note = f"車{parking}" if col in fixed_area else f'{rec["area"]:.2f}坪/車{parking}'
            cells.append(
                f'<td style="background:{bg};color:{fg};">'
                f'<div class="price">{rec["total_price"]:,}萬</div>'
                f'<div class="unitprice">{rec["unit_price"]:.2f}萬</div>'
                f'<div class="note">{note}</div>'
                f"</td>"
            )
        body_rows.append(f'<tr><th class="floor">{floor}F</th>{"".join(cells)}</tr>')

    legend_swatches = "".join(
        f'<span class="swatch" style="background:{bg}"></span><span class="legend-label">{label}</span>'
        for _, _, bg, _fg, label in TIERS
    )

    orientation_label = "直式" if ORIENTATION == "portrait" else "橫式"

    return f"""<title>{CITY_TITLE}</title>
<style>
  * {{ box-sizing: border-box; }}
  html, body {{
    margin: 0;
    background: #ffffff;
  }}
  body {{
    font-family: "Noto Sans TC", "PingFang TC", "Microsoft JhengHei", sans-serif;
    color: #1c2733;
  }}
  /* Fixed to the printable page box (A3 {ORIENTATION} minus @page margin).
     #content has its own fixed intrinsic size below; render_pdf.py measures
     it and scales it (up or down, preserving aspect ratio) to fill this
     box exactly, so the report always lands on a single page no matter
     how many rows/columns it ends up with. */
  #sheet {{
    width: {PAGE_CONTENT_WIDTH_MM}mm;
    height: {PAGE_CONTENT_HEIGHT_MM}mm;
    overflow: hidden;
    position: relative;
  }}
  #content {{
    position: absolute;
    top: 0;
    left: 0;
    width: {CONTENT_WIDTH_PX}px;
    padding: 16px 20px;
    transform-origin: top left;
  }}
  h1 {{
    font-size: 22px;
    margin: 0 0 2px;
    text-align: center;
    letter-spacing: 1.5px;
  }}
  .subtitle {{
    text-align: center;
    color: #6b7684;
    font-size: 9px;
    margin-bottom: 10px;
  }}
  .stats {{
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    justify-content: space-between;
    gap: 6px 14px;
    background: #f4f6f8;
    border-radius: 8px;
    padding: 8px 14px;
    margin-bottom: 10px;
  }}
  .stat {{ text-align: left; }}
  .stat .label {{ font-size: 8px; color: #6b7684; }}
  .stat .value {{ font-size: 14px; font-weight: 700; color: #1c2733; }}
  .stat .value.accent {{ color: #c0392b; }}
  .legend {{ display: flex; align-items: center; gap: 4px; font-size: 8px; color: #33475b; }}
  .swatch {{ display: inline-block; width: 10px; height: 10px; border-radius: 2px; margin-left: 6px; }}
  .legend-label {{ white-space: pre-line; line-height: 1.05; }}
  table {{
    width: 100%;
    border-collapse: collapse;
    table-layout: fixed;
  }}
  th, td {{
    border: 1px solid #dbe1e8;
    text-align: center;
    padding: 1px;
  }}
  thead th {{
    background: #24344a;
    color: #ffffff;
    font-size: 10px;
    height: {THEAD_HEIGHT_PX}px;
    line-height: 1.3;
  }}
  .col-area {{
    display: block;
    font-size: 7px;
    font-weight: 400;
    color: #a9b4c2;
  }}
  th.floor {{
    background: #24344a;
    color: #ffffff;
    width: 42px;
    font-size: 9px;
  }}
  /* Every floor row and every cell in it shares the same explicit height,
     so the grid reads as a perfectly even set of boxes regardless of
     whether a cell has 3 lines of data or is empty. */
  tbody tr {{ height: {ROW_HEIGHT_PX}px; }}
  tbody td, tbody th.floor {{ height: {ROW_HEIGHT_PX}px; vertical-align: middle; }}
  td.empty {{ color: #c3cad2; background: #fafbfc; font-size: 10px; }}
  td .price {{ font-size: 10px; font-weight: 700; white-space: nowrap; line-height: 1.2; }}
  td .unitprice {{ font-size: 10px; font-weight: 700; white-space: nowrap; line-height: 1.2; }}
  td .note {{ font-size: 6.5px; opacity: 0.85; white-space: nowrap; line-height: 1.2; }}
  @page {{ size: A3 {ORIENTATION}; margin: {PAGE_MARGIN_MM}mm; }}
</style>
<div id="sheet"><div id="content">
<h1>{CITY_TITLE}</h1>
<div class="subtitle">更新日期：{update_month} | 單位：萬元/坪 | 規格：A3單頁滿版・列高一致 ({orientation_label}，{FLOOR_MIN}-{FLOOR_MAX}F全樓層版)</div>
<div class="stats">
  <div class="stat"><div class="label">成交戶數</div><div class="value accent">{count} 戶</div></div>
  <div class="stat"><div class="label">最高單價</div><div class="value accent">{max_price:.2f} 萬/坪</div></div>
  <div class="stat"><div class="label">最低單價</div><div class="value accent">{min_price:.2f} 萬/坪</div></div>
  <div class="stat"><div class="label">平均單價</div><div class="value">{avg_price:.2f} 萬/坪</div></div>
  <div class="stat"><div class="label">平均坪數</div><div class="value">{avg_area:.2f} 坪</div></div>
  <div class="legend">單價級距：{legend_swatches}</div>
</div>
<table>
  <thead><tr><th>樓層</th>{head_cells}</tr></thead>
  <tbody>
    {''.join(body_rows)}
  </tbody>
</table>
</div></div>
"""


def main():
    records = parse_rows(XLSX_PATH)
    html = build_html(records)
    HTML_OUT.parent.mkdir(parents=True, exist_ok=True)
    HTML_OUT.write_text(html, encoding="utf-8")
    print(f"parsed {len(records)} records -> {HTML_OUT}")


if __name__ == "__main__":
    main()
