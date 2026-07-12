# maps

實價登錄樓層對照表產生工具。

`scripts/generate_report.py` 讀取 `data/` 內的實價登錄 Excel 匯出檔，依「棟號 x 樓層」樞紐整理成
單頁滿版對照表（成交戶數、最高/最低/平均單價、平均坪數，並依單價級距上色），輸出至 `output/guoyang_report.html`。

```
python3 scripts/generate_report.py [xlsx路徑] [html輸出路徑]
```

再用瀏覽器（或 Playwright／`chromium --headless --print-to-pdf`）將 HTML 轉存為 A3 橫向 PDF 即為最終報表。
