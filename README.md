# ETI_js
## 簡述
Backend Chrome extension

## 執行步驟
1. 執行eti.js
* 搜尋並安裝Chrome的擴充元件——cjs(Custom JavaScript for websites)
* 前往 http://www.eti-tw.com/admin ，並登入管理員帳號
* 點選 創作者資料管理 -> 編輯作品 -> DB Excel，
如圖https://imgur.com/2jme7Au
* 在這個頁面開啟cjs，並將檔案eti.js的程式碼複製到cjs
* 選擇jquery版本
* 按"save"按鈕，網頁會重新載入並執行cjs
    
2. 登入Google
* 執行成功後，將會跳出登入Google帳號的彈跳視窗
```注意彈跳視窗有沒有被封鎖```
* 登入的Google帳號需據有google sheet api授權
    
3. 選擇檔案
* 按"選擇檔案"。會自動觸發handFile這個function
* 選取本地端(電腦中)的excel檔
* 選好檔案後，JavaScript將開始比對雲端上的excel檔和剛剛選取的本地端excel檔
    
4. 完成

* "影音作品 DB Excel 資料"欄位會自動填入本地端那份excel的資料，並在Chrome DevTools(在Chrome按F12開啟)的Console可以看到執行結果
