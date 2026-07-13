# maps

「友眾16青砥」預售屋合約審閱本 — 客戶掃描 QR Code 即可線上閱讀完整合約條款。

## 內容

- `docs/index.html` — 審閱本首頁（含 QR Code）
- `docs/land-contract.html` — 預定土地買賣契約書全文
- `docs/house-contract.html` — 預定房屋買賣契約書全文（含全部附件）
- `docs/assets/qrcode.png` — 頁面內嵌用 QR Code
- `docs/assets/qrcode-print.png` — 可直接列印張貼的 QR Code（含標題與網址）

合約內容為契約範本原文逐條轉錄，未填寫之空白欄位（金額、日期、坪數等）維持原樣，供實際簽約時手寫填入；正式簽約仍以紙本契約書為準。

## 線上瀏覽（Cloudflare Pages）

本站已透過 Cloudflare Pages 部署，Production branch 為 `main`，Build output directory 為 `docs`，每次 push 到 `main` 都會自動重新部署。

網址：**https://youzhong16-contract.pages.dev/**

`docs/assets/qrcode.png` 與 `qrcode-print.png` 皆已指向這個網址，客人掃描即可直接開啟合約。若之後改用自己的網域，重新產生 QR Code 指向新網址即可。
