# 花蓮蔬食地圖 (Hualien Veggie & Friendly Map)

一款專為花蓮在地居民、東華與慈濟大專院校師生、以及造訪花蓮的旅客打造的**輕量化互動式蔬食地圖 Web App**。

![專案特色](https://img.shields.io/badge/地圖架構-Leaflet.js%20%2B%20OpenStreetMap-10B981)
![無金鑰規範](https://img.shields.io/badge/API%20Key-100%25%20Free%20%26%20No%20Key-blue)
![授權](https://img.shields.io/badge/License-MIT-orange)

---

## 🌟 核心功能亮點

1. **互動式動態地圖**：
   - 採用 OpenStreetMap 與 Leaflet.js，免申請 Google API 金鑰，零額度限制、零外洩風險。
   - 支援花蓮市、吉安鄉、壽豐鄉、新城鄉等地區之特色蔬食店家經緯度視覺化標註。
2. **多維度飲食篩選系統**：
   - 依飲食需求快速切換：**全素/純素**、**蛋奶素**、**蔬食友善**。
   - 地區行政區即時過濾（花蓮市、吉安、壽豐、新城）。
   - 全文即時搜尋（店名、地址、招牌菜色如「蒲燒鰻」、「臭豆腐」、「蔬食拌飯」）。
3. **餐廳詳細資訊與一鍵外連**：
   - 營業時間、公休提示、價位區間、特色招牌菜標籤。
   - **一鍵通話 (CALL)**：點擊直接啟動手機撥號聯絡預約。
   - **一鍵 Google 導航 (DIRECTIONS)**：自動帶入店家座標，開啟 Google 地圖進行即時路線規劃。
4. **極致輕量與跨平台響應**：
   - 純原生前端架構（HTML5 + CSS3 + Vanilla JavaScript），免安裝 Python 或 Node.js 後端環境。
   - 雙擊 `index.html` 即可於任何瀏覽器本機秒開執行，亦可一鍵部署至 GitHub Pages 支援手機掃碼瀏覽。

---

## 🚀 快速開始 (Quick Start)

### 方式一：本機直接執行（最簡單）
直接使用任何瀏覽器（Chrome, Edge, Safari, Firefox）雙擊打開 `index.html` 即可立即體驗完整功能。

### 方式二：使用本機靜態伺服器（推薦）
```bash
# 使用 Python 內建伺服器
python -m http.server 8000
# 瀏覽器打開 http://localhost:8000
```

---

## 📂 專案檔案結構

```text
├── index.html       # 應用程式首頁骨架與介面佈局
├── data.js          # 花蓮在地蔬食精選店家資料庫 (JSON 格式陣列)
├── css/
│   └── style.css    # 現代化清新風格響應式樣式表
├── js/
│   └── app.js       # Leaflet 地圖串接、標記生成、過濾搜尋與聯動邏輯
├── README.md        # 專案說明文件
└── .gitignore       # Git 版本控制忽略設定 (確保無敏感資訊)
```

---

## 🔒 隱私與安全聲明
- 本專案完全使用開源地圖服務（OpenStreetMap），**程式碼與設定檔中絕無包含任何 API 金鑰、Token 或個人敏感資訊**。
