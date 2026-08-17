# 📱 個人視覺化工作儀表板 (Personal Workspace Dashboard) `v4.6`

一個基於 **Google Apps Script (GAS)** 與 **Google 試算表** 打造的現代化個人工作儀表板 Web App。全面整合 Google 日曆、外部 iCal、專案進度追蹤、週期性重複任務與即時待辦管理。

---

## ✨ 核心特色

* 📊 **首屏一覽無遺 (Above-the-Fold)**：Banner 搭配 4 大 KPI 統計卡（總專案數、待辦任務、今日行程、完成率）與緊湊雙欄佈局，徹底消除長捲軸滑動摩擦。
* 🔄 **週期性重複任務 (Recurring Tasks)**：支援「每週 (+7天)」、「每月 (+1月)」、「每年 (+1年)」設定。點擊勾選完成當期任務時，系統自動推算截止日並展延建立下一期待辦！
* 📑 **頁籤式工作台 (Tabs)**：
  * **今日待辦事項**：聚焦當日需完成的任務，支援一鍵勾選完成。
  * **今日日曆行程**：整合 Google 主要日曆與外部 iCal 行程，當前進行中的行程自帶呼吸發光動畫。
  * **未來待辦清單**：內建 **專案下拉篩選器**，並支援點擊上方進度條自動跳轉過濾。
  * **已完成任務清單**：清楚記錄歷史成就，支援誤按一鍵復原為進行中。
* 🎨 **專案進度與色盤系統**：自動統計各專案完成百分比，支援 6 種主題色票自訂專案代表色。
* 🚀 **1-Click 試算表自愈**：試算表內建自訂選單，一鍵自動修復分頁結構、欄位驗證與 ArrayFormula 統計公式。

---

## 🚀 3 分鐘快速啟用教學

1. **建立試算表副本**：點擊 [👉 取得試算表範本（一鍵複製）](https://docs.google.com/spreadsheets/d/1NE927knXdJSZ6Q2-SRjgDXvVxDcRP2C4NmhKjaQhsHU/copy)
2. **初始化分頁與公式**：
   * 等待試算表上方工具列出現 **`📱 工作儀表板 App`** 選單。
   * 點選 **`🚀 1-Click 一鍵最佳化分頁與公式`** 並完成 Google 帳號安全性授權。
3. **部署專屬 Web App**：
   * 點選上方選單 **「擴充功能」 ➔ 「Apps Script」**。
   * 點擊右上角 **「部署」 ➔ 「新增部署作業」**。
   * 類型選擇 **「網頁應用程式 (Web App)」**。
   * **執行身分**：選擇 `我`。
   * **誰可以存取**：選擇 `僅限我自己` *(保護您的個人資料隱私)*。
   * 點擊「部署」，並複製產生的 **網頁應用程式網址** 即可開始使用！

---

## 🛠️ 技術架構
* **前端 (Frontend)**：HTML5, Bootstrap 5.3, Bootstrap Icons, Vanilla JavaScript (RWD & Tabs)
* **後端 (Backend)**：Google Apps Script (GAS)
* **資料庫 (Database)**：Google Sheets (試算表)
* **服務整合 (Integration)**：Google Calendar API, RFC 5545 iCal Regex Parser

---

## 📄 開源授權
本專案採用 [MIT License](LICENSE) 授權開源。
