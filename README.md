# 🏃‍♂️ 臺北馬拉松賽事爬蟲

一個以 **Streamlit** 製作的互動式小工具，從「臺北馬拉松賽事列表」網站爬取資料，提供 **年份 / 行政區 / 賽事類型** 篩選、表格檢視、地圖嵌入預覽，並可一鍵匯出 **A4 橫向 PDF**（含 **賽事連結** 與 **地點 Google Maps 連結** 的可點擊超連結）。

> 專案語言：**Python 3.10+** ｜ 前端框架：**Streamlit** ｜ PDF 引擎：**ReportLab**

---

## ✨ 功能亮點（Features）

* 🔎 **即時篩選**：支援「全部年份 / 目前賽事 / 指定年份」、「北中南東其他」行政區與多種賽事類型。
* 🔗 **完整連結**：自動擷取 **賽事名稱原始連結**，PDF 內也保留可點擊超連結。
* 🗺️ **地圖預覽**：從表格選地點後，右側即時顯示 **Google Maps** 內嵌地圖。
* 🧾 **PDF 匯出**：自動調整欄寬與字級，輸出 **A4 橫向** 報表；表頭灰底、置中、含連結。
* 🈶 **中文字型**：預設使用 **微軟正黑體（MSJH.TTC）** 以避免 PDF 亂碼。
* ⚠️ **健壯錯誤處理**：偵測 ASP.NET PostBack 與網站改版情境，提供明確錯誤訊息。

---

## 📦 安裝（Installation）

1. **克隆專案**

```bash
git clone https://github.com/<YOUR_ORG>/<YOUR_REPO>.git
cd <YOUR_REPO>
```

2. **建立虛擬環境並安裝套件**

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\\Scripts\\activate
pip install -r requirements.txt
```

> 若尚未建立 `requirements.txt`，可參考：
>
> ```txt
> streamlit
> requests
> beautifulsoup4
> pandas
> reportlab
> ```

3. **放置中文字型**（重要）

* 將 **`MSJH.TTC`**（微軟正黑體）放在 **專案根目錄**（與 `scraper.py` 同層）。
* 專案會以以下程式碼自動註冊：

  ```python
  FONT_PATH = os.path.join(os.path.dirname(__file__), "MSJH.TTC")
  pdfmetrics.registerFont(TTFont("MSJH-Regular", FONT_PATH, subfontIndex=0))
  pdfmetrics.registerFont(TTFont("MSJH-Bold",    FONT_PATH, subfontIndex=1))
  pdfmetrics.registerFontFamily("MSJH", normal="MSJH-Regular", bold="MSJH-Bold",
                                italic="MSJH-Regular", boldItalic="MSJH-Bold")
  ```
* 若未放置字型，系統會在 UI 顯示警告並嘗試以預設字型生成 PDF（可能出現亂碼）。

> 版權注意：請確認你擁有該字型之授權。若需替代，可改用 [Noto Sans CJK](https://www.google.com/get/noto/) 並調整註冊程式碼。

---

## ▶️ 執行（Run）

```bash
streamlit run scraper.py
```

啟動後依側欄選擇：**年份 / 行政區 / 賽事類型**，可另外輸入 **關鍵字** 篩選（如「臺南」）。

* 主畫面：顯示查詢結果表格（已依日期排序）。
* 地圖：由下拉選單選擇地點後，右側即時顯示 Google Maps 內嵌地圖。
* PDF：點擊「**下載查詢結果 (PDF)**」輸出報表；檔名會依條件自動命名，如：

  ```
  allYears_北_public_臺南.pdf
  ```

---

## 🧠 主要設計（How it works）

* **ASP.NET PostBack 模擬**：

  * 第一次 `GET` 取得頁面與 `<select>` 欄位名稱與 `VIEWSTATE` 等隱藏欄位。
  * 依序對「年份 / 行政區 / 類型」送出模擬 `POST`，以 `__EVENTTARGET` 觸發回傳更新後的列表。
* **表格解析**：透過 `BeautifulSoup` 解析 `table#ctl00_ContentPlaceHolder1_GridView1`，萃取「賽事名稱 / 日期 / 地點」並保留 **賽事超連結**。
* **日期推斷**：若 `日期` 欄無年份資訊，會嘗試由「賽事名稱」前綴推斷年份。
* **地圖連結**：同時產出 `地點連結`（Google Maps 搜尋）與 `地點嵌入URL`（iframe 預覽）。
* **PDF 排版**：

  * A4 橫向，多欄寬度自動分配；欄寬過窄時自動降低字級（10 → 8 → 6）。
  * 表頭灰底、置中；**賽事連結 / 地點連結** 以 `<link>` 建立可點擊超連結。

---

## 🐳 Docker（可選）

> 若你偏好容器化，以下為最小可行的 Dockerfile 範例。

```dockerfile
FROM python:3.11-slim
WORKDIR /app
COPY . /app
# 安裝套件
RUN pip install --no-cache-dir -r requirements.txt \
    && apt-get update && apt-get install -y fonts-noto-cjk --no-install-recommends \
    && rm -rf /var/lib/apt/lists/*
# 若你有 MSJH.TTC，請自行 COPY 進來或改用 Noto Sans CJK
# COPY MSJH.TTC /app/MSJH.TTC
EXPOSE 8501
CMD ["streamlit", "run", "scraper.py", "--server.port=8501", "--server.address=0.0.0.0"]
```

執行：

```bash
docker build -t tpe-marathon-scraper .
docker run --rm -p 8501:8501 tpe-marathon-scraper
```

---

## 🗂️ 專案結構（Project Structure）

```
<repo>
├─ scraper.py           # 主程式（Streamlit + 爬蟲 + PDF）
├─ MSJH.TTC             # 微軟正黑體（可選，建議放）
├─ requirements.txt     # 相依套件（建議建立）
└─ README.md            # 本檔案
```

---

## 🧪 開發 & 限制（Notes & Limitations）

* 本專案依賴目標網站（ASP.NET）之 DOM 結構與 PostBack 行為；**若網站改版**，請更新選擇器與欄位名稱。
* 網站可能有流量保護或封鎖策略，請 **避免過度頻繁請求**，遵守使用規範。
* 「日期年份推斷」採簡化規則，少數賽事名稱若未含年份可能導致排序落在最後。

---

## 🛠️ 疑難排解（Troubleshooting）

* **PDF 中文亂碼 / 變方塊**

  * 確認 `MSJH.TTC` 是否放在與 `scraper.py` 同層。
  * 或改裝 `Noto Sans CJK`，並修改程式中的字型註冊。
* **顯示「找不到 `<select id='Year'>`」或 PostBack 相關錯誤**

  * 目標網站可能改版，請檢查 `get_filtered_soup()` 內的選擇器與欄位名稱。
* **Streamlit 頁面空白 / PDF 下載按鈕不出現**

  * 需先完成「開始查詢」並確保有資料；若無資料會顯示警告訊息。
* **Docker 內字型失效**

  * 請額外 COPY 字型或改用 `fonts-noto-cjk`。

---

## 📜 授權（License）

建議使用 **MIT License**，你也可以依需求替換。請在根目錄加入 `LICENSE`。

---

## 🙏 致謝（Acknowledgements）

* 臺北馬拉松賽事列表
* Streamlit / ReportLab / BeautifulSoup / pandas

---

## 🗓️ Changelog（可選）

* 2025‑08‑26：初版 README 草稿。
