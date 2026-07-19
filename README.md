# maps

這個 repo 目前收納兩個各自獨立的專案：

## 1. 國揚鉑御實價登錄樓層對照表

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

## 2. 「友眾16青砥」預售屋合約審閱本

客戶掃描 QR Code 即可線上閱讀完整合約條款。

### 內容

- `docs/index.html` — 審閱本首頁（頁面載入時以 JS 動態產生 QR Code，內容為目前網站網址）
- `docs/land-contract.html` — 預定土地買賣契約書全文
- `docs/house-contract.html` — 預定房屋買賣契約書全文（含全部附件）
- `docs/assets/qrcode.png` — 頁面內嵌用 QR Code
- `docs/assets/qrcode-print.png` — 可直接列印張貼的 QR Code（含標題與網址）

合約內容為契約範本原文逐條轉錄，未填寫之空白欄位（金額、日期、坪數等）維持原樣，供實際簽約時手寫填入；正式簽約仍以紙本契約書為準。

### 線上瀏覽（GitHub Pages）

本站已透過 GitHub Pages 部署（Settings → Pages → Deploy from a branch → `main` / `/docs`），每次 push 到 `main` 都會自動重新部署，不依賴第三方平台的 Git 連線設定。

網址：**https://normalsky315-jpg.github.io/maps/**

`docs/index.html` 首頁的 QR Code 是頁面載入時以 JS 動態產生，內容固定為「目前網站網址」，換網域也不用手動更新。`docs/assets/qrcode.png` 與 `qrcode-print.png` 則是給列印或單獨分享用的靜態圖檔，已指向上述網址。

（先前曾嘗試 Cloudflare Pages，但該專案的 GitHub 連線多次中斷導致部署卡在舊版本，故改回 GitHub Pages。）
