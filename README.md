# maps

實價登錄樓層對照表產生工具。

`scripts/generate_report.py` 讀取 `data/` 內的實價登錄 Excel 匯出檔，依「棟號 x 樓層」樞紐整理成
單頁滿版對照表（成交戶數、最高/最低/平均單價、平均坪數，並依單價級距上色，所有樓層列高一致），
輸出至 `output/guoyang_report.html`。頁面尺寸／方向（A3 直式或橫式）由 `scripts/report_config.py` 設定。

```
python3 scripts/generate_report.py [xlsx路徑] [html輸出路徑]
python3 scripts/render_pdf.py [html路徑] [pdf輸出路徑] [png輸出路徑]
```

`render_pdf.py` 用 Playwright(Chromium) 將 HTML 轉成 PDF/PNG：內容區塊會依比例縮放（可放大也可縮小）
剛好填滿單一 A3 頁面，因此不論資料筆數或欄位多寡，輸出永遠是一頁。
