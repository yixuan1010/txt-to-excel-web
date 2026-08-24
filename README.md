# Laser Scan TXT 轉 Excel 網頁版

這是原本 `txt_to_excel.py` 的純前端網頁版本，可直接部署到 GitHub Pages。

## 功能

- 手動選擇多個 TXT 檔
- 手動選擇整個資料夾，包含子資料夾內的 TXT
- 支援拖曳 TXT 檔
- 支援 UTF-8、UTF-8 BOM、Big5/CP950 與 UTF-16 Log
- 擷取相鄰的 Move / Scan Laser 動作
- 保留原本的無條件捨去、排序及 Excel 欄位規則
- 手動下載 `result.xlsx`
- 所有 TXT 內容只在使用者瀏覽器內處理，不會傳到伺服器

## 部署到 GitHub Pages

1. 在 GitHub 建立新的 Repository，例如 `txt-to-excel-web`。
2. 將本壓縮檔中的 `index.html`、`styles.css`、`app.js` 和 `README.md` 上傳到 Repository 根目錄。
3. 打開 Repository 的 **Settings**。
4. 左側選擇 **Pages**。
5. 在 **Build and deployment** 中，Source 選擇 **Deploy from a branch**。
6. Branch 選擇 `main`，資料夾選擇 `/(root)`，按 **Save**。
7. 等待 GitHub 完成部署後，即可從 Pages 顯示的網址開啟工具。

## 使用方式

1. 開啟網頁。
2. 按「選擇 TXT 檔」或「選擇資料夾」。選擇資料夾功能建議使用 Edge 或 Chrome。
3. 按「轉換並下載 Excel」。
4. 瀏覽器會下載 `result.xlsx`。

## 技術說明

- 網頁本身不需要 Python 或後端伺服器。
- Excel 產生使用 ExcelJS 4.4.0 CDN，因此第一次開啟與轉換時需要網路連線。
- GitHub Pages 是靜態網站託管，非常適合此純前端版本。
