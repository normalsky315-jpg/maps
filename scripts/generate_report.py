#!/usr/bin/env python3
"""Generate the 國揚鉑御 樓層/單價 實價登錄 對照表 (HTML, A3 landscape) from the
raw 實價登錄 Excel export.

Usage:
    python3 scripts/generate_report.py [xlsx_path] [html_out]
"""
import re
import sys
from datetime import date
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
XLSX_PATH = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "data" / "guoyang_boyu_實價登錄.xlsx"
HTML_OUT = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "output" / "guoyang_report.html"

BUILDING_NAME = "國揚鉑御"
CITY_TITLE = "高雄市 國揚鉑御社區實價控表"
FLOOR_MIN, FLOOR_MAX = 2, 20

# A3 landscape = 420 x 297mm. #content is scaled to fit inside this box
# (see render_pdf.py) so the report is always exactly one page.
PAGE_WIDTH_MM, PAGE_HEIGHT_MM = 420, 297
PAGE_MARGIN_MM = 10
PAGE_CONTENT_WIDTH_MM = PAGE_WIDTH_MM - 2 * PAGE_MARGIN_MM
PAGE_CONTENT_HEIGHT_MM = PAGE_HEIGHT_MM - 2 * PAGE_MARGIN_MM

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

    head_cells = "".join(f"<th>{c}</th>" for c in columns)

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
            cells.append(
                f'<td style="background:{bg};color:{fg};">'
                f'<div class="price">{rec["total_price"]:,}萬</div>'
                f'<div class="unitprice">{rec["unit_price"]:.2f}萬</div>'
                f'<div class="note">{rec["area"]:.2f}坪/車{parking}</div>'
                f"</td>"
            )
        body_rows.append(f'<tr><th class="floor">{floor}F</th>{"".join(cells)}</tr>')

    legend_swatches = "".join(
        f'<span class="swatch" style="background:{bg}"></span><span class="legend-label">{label}</span>'
        for _, _, bg, _fg, label in TIERS
    )

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
  /* Fixed to the printable content box (A3 landscape minus @page margin).
     #content is scaled down (never up) by render_pdf.py so the whole
     report always lands on a single page, however many rows/columns. */
  #sheet {{
    width: {PAGE_CONTENT_WIDTH_MM}mm;
    height: {PAGE_CONTENT_HEIGHT_MM}mm;
    overflow: hidden;
    position: relative;
  }}
  #content {{
    padding: 20px 28px;
    transform-origin: top left;
  }}
  h1 {{
    font-size: 30px;
    margin: 0 0 4px;
    text-align: center;
    letter-spacing: 2px;
  }}
  .subtitle {{
    text-align: center;
    color: #6b7684;
    font-size: 13px;
    margin-bottom: 18px;
  }}
  .stats {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    background: #f4f6f8;
    border-radius: 10px;
    padding: 14px 24px;
    margin-bottom: 18px;
  }}
  .stat {{ text-align: left; }}
  .stat .label {{ font-size: 12px; color: #6b7684; }}
  .stat .value {{ font-size: 20px; font-weight: 700; color: #1c2733; }}
  .stat .value.accent {{ color: #c0392b; }}
  .legend {{ display: flex; align-items: center; gap: 8px; font-size: 12px; color: #33475b; }}
  .swatch {{ display: inline-block; width: 14px; height: 14px; border-radius: 3px; margin-left: 10px; }}
  .legend-label {{ white-space: pre-line; line-height: 1.1; }}
  table {{
    width: 100%;
    border-collapse: collapse;
    table-layout: fixed;
  }}
  th, td {{
    border: 1px solid #dbe1e8;
    text-align: center;
    padding: 4px 2px;
  }}
  thead th {{
    background: #24344a;
    color: #ffffff;
    font-size: 14px;
    padding: 8px 2px;
  }}
  th.floor {{
    background: #24344a;
    color: #ffffff;
    width: 52px;
    font-size: 13px;
  }}
  td.empty {{ color: #c3cad2; background: #fafbfc; }}
  td .price {{ font-size: 13px; font-weight: 700; }}
  td .unitprice {{ font-size: 13px; font-weight: 700; }}
  td .note {{ font-size: 9px; opacity: 0.85; }}
  @page {{ size: A3 landscape; margin: {PAGE_MARGIN_MM}mm; }}
</style>
<div id="sheet"><div id="content">
<h1>{CITY_TITLE}</h1>
<div class="subtitle">更新日期：{update_month} | 單位：萬元/坪 | 規格：A3單頁滿版大字對齊制 ({FLOOR_MIN}-{FLOOR_MAX}F全樓層版)</div>
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
