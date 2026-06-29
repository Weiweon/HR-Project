# HR-Project 使用說明

此資料夾是正式版本，也是可攜帶版本。要帶去公司或其他電腦使用時，請整個複製 `HR-Project` 資料夾。

## 開啟方式

1. 進入 `HR-Project` 資料夾。
2. 直接雙擊 `index.html`。
3. 在網頁中匯入公司系統匯出的 Excel。

## 檔案說明

- `index.html`：網頁入口。
- `app.js`：Excel 匯入、加班費估算與異常判斷邏輯。
- `styles.css`：畫面樣式。
- `vendor/xlsx.full.min.js`：本機 Excel 解析函式庫。

## 維護原則

- 後續程式修改都以 `HR-Project` 內檔案為準。
- 外層資料夾主要保留 Git、需求文件、變更紀錄與 Excel 原始資料。

## 注意事項

- `http://127.0.0.1:8765/index.html` 只有在該台電腦已啟動本機網頁服務時才可使用。
- 若只是帶去公司使用，通常直接雙擊 `index.html` 即可。
- Excel 資料只在本機瀏覽器處理，不會上傳。
