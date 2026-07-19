#!/usr/bin/env python3
"""Generate the 國揚鉑御 樓層/單價 實價登錄 對照表 (HTML, single A3 page) from the
raw 實價登錄 Excel export. Page orientation is set in report_config.py.

Usage:
    python3 scripts/generate_report.py [xlsx_path] [html_out]
"""
import re
import sys
from collections import Counter
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

# Full per-floor unit roster (from the 鼓山區國揚鉑御 戶型配置圖: A棟 skips A4/A14,
# B棟 skips B4 — no such units exist). Every column appears even if it has
# zero transactions, so unsold units (e.g. B2) still show up as an all-"—"
# column instead of being silently omitted.
FULL_UNITS = [
    "A1", "A2", "A3", "A5", "A6", "A7", "A8", "A9", "A10", "A11", "A12", "A13", "A15", "A16",
    "B1", "B2", "B3", "B5", "B6", "B7", "B8", "B9",
]

# 從建案官方戶型規劃圖(FLOOR PLANNING)抄錄的標準坪數，只用來補「從未成交、
# 實價登錄沒有任何紀錄」的戶別 —— 已有成交紀錄的戶別一律以 area_stats() 算出
# 的實際登錄坪數為準（兩者基準可能不同：規劃圖坪數通常不含車位面積，登錄
# 坪數則是該筆交易「房地+車位」的總面積，若有配車位通常會大個幾坪)。
# A棟3R戶型（A1、A7）官方表僅標示「41-42.5坪」區間、圖上又只清楚標出角間
# A8的42.5坪，A1/A7無法判斷是41或42.5哪一個，因此原樣保留區間文字。
SPEC_AREA = {
    "A2": 23.00, "A3": 23.00, "A5": 23.00, "A6": 23.00, "A11": 23.00,  # 1+1R，圖上標示23P
    "A1": "41-42.5", "A7": "41-42.5",  # 3R，區間無法細分至單一戶別
    "B2": 21.50, "B3": 21.50, "B5": 21.50,  # 1+1R，圖上標示21.5P
    "B8": 37.00,  # 3R，四個角間圖上均標示37P
}

# #content has a fixed intrinsic design size (roughly matching the target
# page's aspect ratio) and every floor row gets the same explicit height,
# so the grid is perfectly uniform. render_pdf.py then measures this
# intrinsic box and scales it (up or down) to exactly fill one printable
# page, however many rows/columns end up in the table.
CONTENT_WIDTH_PX = 1970
ROW_HEIGHT_PX = 63
THEAD_HEIGHT_PX = 34
AREA_ROW_HEIGHT_PX = 26

# A unit column's 坪數 row shows its single most common (mode) 總面積.
# Columns whose recorded area actually spreads by more than this get
# called out separately (see main()) instead of silently averaged away.
AREA_SPREAD_TOLERANCE = 0.1

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


def area_stats(records):
    """Per-column 標準坪數 (mode, ties broken toward the smaller value) plus
    the full spread (max-min) so callers can flag columns that don't
    actually share one consistent 坪數."""
    by_col = {}
    for r in records:
        by_col.setdefault(r["col"], []).append(round(r["area"], 2))
    stats = {}
    for col, areas in by_col.items():
        counts = Counter(areas)
        top = max(counts.values())
        standard = min(v for v, c in counts.items() if c == top)
        stats[col] = {
            "standard": standard,
            "spread": max(areas) - min(areas),
            "values": sorted(set(areas)),
        }
    return stats


def area_cell_html(col, col_area):
    """The 坪數 header cell for one column: actual registry data takes
    priority; a never-sold column falls back to the official spec sheet
    (SPEC_AREA), and only shows 未成交 if neither is available."""
    if col in col_area:
        flagged = col_area[col]["spread"] > AREA_SPREAD_TOLERANCE
        cls = "area-cell area-flag" if flagged else "area-cell"
        return f'<th class="{cls}">{col_area[col]["standard"]:.2f}坪</th>'
    spec = SPEC_AREA.get(col)
    if isinstance(spec, float):
        return f'<th class="area-cell area-spec">{spec:.2f}坪</th>'
    if isinstance(spec, str):
        return f'<th class="area-cell area-spec">{spec}坪</th>'
    return '<th class="area-cell area-unsold">未成交</th>'


def build_html(records):
    columns = sorted(set(FULL_UNITS) | {r["col"] for r in records}, key=column_sort_key)
    grid = {(r["col"], r["floor"]): r for r in records}
    col_area = area_stats(records)

    count = len(records)
    max_price = max(r["unit_price"] for r in records)
    min_price = min(r["unit_price"] for r in records)
    avg_price = sum(r["unit_price"] for r in records) / count
    avg_area = sum(r["area"] for r in records) / count

    update_month = date.today().strftime("%Y/%m")

    head_cells = "".join(f"<th>{c}</th>" for c in columns)
    area_cells = "".join(area_cell_html(c, col_area) for c in columns)

    body_rows = []
    for floor in range(FLOOR_MAX, FLOOR_MIN - 1, -1):
        cells = []
        for col in columns:
            rec = grid.get((col, floor))
            if not rec:
                cells.append('<td class="empty">—</td>')
                continue
            bg, fg = tier_for(rec["unit_price"])
            parking = f'車{rec["parking_price"]}萬' if rec["parking_price"] is not None else "車無"
            cells.append(
                f'<td style="background:{bg};color:{fg};">'
                f'<div class="price">{rec["total_price"]:,}萬</div>'
                f'<div class="unitprice">{rec["unit_price"]:.2f}萬</div>'
                f'<div class="note">{parking}</div>'
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
    font-size: 12px;
    height: {THEAD_HEIGHT_PX}px;
  }}
  th.area-cell {{
    background: #3a4c66;
    color: #cfd8e3;
    font-size: 10px;
    font-weight: 400;
    height: {AREA_ROW_HEIGHT_PX}px;
  }}
  /* Flags a column whose recorded 總面積 actually varies by more than
     AREA_SPREAD_TOLERANCE — the 標準坪數 shown is only the mode, so this
     calls out that it's not the whole story. */
  th.area-flag {{
    background: #5c2430;
    color: #ff8a80;
    font-weight: 700;
  }}
  th.area-unsold {{
    color: #7c8998;
    font-style: italic;
    font-weight: 400;
  }}
  /* Never-sold column, filled from the developer's official 戶型規劃圖
     spec sheet rather than 實價登錄 — visually distinct (teal, italic)
     from the white mode-of-actual-transactions cells. */
  th.area-spec {{
    color: #8fd6c9;
    font-style: italic;
    font-weight: 400;
  }}
  th.corner {{
    background: #24344a;
    color: #ffffff;
    width: 46px;
  }}
  th.floor {{
    background: #24344a;
    color: #ffffff;
    width: 46px;
    font-size: 11px;
  }}
  /* Every floor row and every cell in it shares the same explicit height,
     so the grid reads as a perfectly even set of boxes regardless of
     whether a cell has 3 lines of data or is empty. */
  tbody tr {{ height: {ROW_HEIGHT_PX}px; }}
  tbody td, tbody th.floor {{ height: {ROW_HEIGHT_PX}px; vertical-align: middle; }}
  td.empty {{ color: #c3cad2; background: #fafbfc; font-size: 13px; }}
  td .price {{ font-size: 15px; font-weight: 700; white-space: nowrap; line-height: 1.25; }}
  td .unitprice {{ font-size: 15px; font-weight: 700; white-space: nowrap; line-height: 1.25; }}
  td .note {{ font-size: 15px; opacity: 0.85; white-space: nowrap; line-height: 1.25; }}
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
  <thead>
    <tr><th class="corner">樓層</th>{head_cells}</tr>
    <tr><th class="corner">坪數</th>{area_cells}</tr>
  </thead>
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

    flagged = {
        col: s for col, s in area_stats(records).items() if s["spread"] > AREA_SPREAD_TOLERANCE
    }
    if flagged:
        print(f"\n坪數誤差超過 {AREA_SPREAD_TOLERANCE} 坪的戶別（標準坪數欄採眾數，最小值為準）：")
        for col in sorted(flagged, key=column_sort_key):
            s = flagged[col]
            print(f"  {col}: 標準={s['standard']:.2f}坪，實際出現值={s['values']}，最大差距={s['spread']:.2f}坪")


if __name__ == "__main__":
    main()
