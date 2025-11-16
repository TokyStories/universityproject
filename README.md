這是一套基於 FastAPI (後端) 與 Web (前端) 的自動化處理系統，旨在將手動的 ATE 測試報告流程（從 PDF 規格書比對到 Excel 數據填寫）完全自動化。

系統架構
後端: FastAPI (Python)

前端: HTML, CSS, JavaScript 

資料流 (Pipeline):

P1: (PDF -> Excel) PDF 規格書解析

P2: (Excel + TXT -> CSV) 規格與 ATE 報告比對

P3: (CSV -> CSV) 參數名稱正規化

P4: (CSV -> Excel) 動態 Excel 模板生成

P5: (CSV -> JSON) 資料映射字典生成

P6: (Excel + JSON + TXTs -> Excel) 報表自動填寫

🛠️ 安裝指南
步驟 1：前置作業 (Prerequisites)
在安裝 Python 套件前，你的電腦必須安裝以下兩個軟體：

Python 3.8 (或更高版本)

Tesseract-OCR

(P1 階段需要)

下載與安裝：Tesseract 官方 Windows 安裝程式

重要： 安裝時，請勾選「繁體中文」(Traditional Chinese) 語言包。

步驟 2：環境設定 (Configuration)
本專案依賴兩個外部設定：

Google Gemini API 金鑰

(P1 階段需要)

請在專案根目錄下建立一個名為 .env 的檔案。

在 .env 檔案中加入以下內容 (替換為你自己的金鑰)：

GEMINI_API_KEY=AIzaSy...YOUR_API_KEY...
Tesseract-OCR 路徑

階段一.py 檔案寫死了 Tesseract 路徑。

【必須修改】 請開啟 階段一.py，找到第 16 行：

Python

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
請將 r"C:\Program Files\..." 修改為你電腦上 tesseract.exe 的實際安裝路徑。

步驟 3：安裝 Python 套件
# (建議) 建立並啟用虛擬環境
python -m venv venv
venv\Scripts\activate

# 執行此指令，一鍵安裝所有必要套件
pip install -r requirements.txt


啟動伺服器
當所有套件安裝完畢、.env 和 Tesseract 路徑都設定好後，執行以下指令來啟動 FastAPI 伺服器：
# uvicorn main:app --reload

