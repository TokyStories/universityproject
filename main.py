"""
FastAPI 伺服器主程式 (main.py)
--------------------------------
負責提供 API 路由，並協調後端處理模組 (P1, P2, P3, P4)。
"""

# --- 基礎函式庫 ---
import os
import tempfile
import asyncio
from typing import List, Optional

# --- FastAPI 核心 ---
from fastapi import FastAPI, UploadFile, File, BackgroundTasks, HTTPException
from fastapi.responses import StreamingResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
from contextlib import asynccontextmanager

# --- 類型提示 ---
from starlette.concurrency import run_in_threadpool

# --- 匯入核心處理模組 ---
import 階段一 as P1
import report_processor as P2
import parameter_normalizer as P3 # 【新增】匯入新的正規化模組
import DataTransformer as P4
import 轉換Dict功能 as P5 
import 自動輸出Excel報表功能 as P6 # 【【新增】】
# --- 全域設定 ---
# 建立一個 set 來追蹤正在被寫入的暫存檔案
# (防止多個請求同時寫入同一個檔案，雖然在 tempfile 中機率很低)
# (此範例暫時未使用，但可作為未來擴充)
# processing_files = set()

# ----------------------------------------------------
# 應用程式生命週期 (Lifecycle)
# ----------------------------------------------------
@asynccontextmanager
async def lifespan(app: FastAPI):
    # 應用程式啟動時
    print("--- 伺服器啟動 ---")
    print("FastAPI 正在啟動...")
    print("將在 http://127.0.0.1:8000 啟動服務")
    yield
    # 應用程式關閉時
    print("--- 伺服器關閉 ---")

# ----------------------------------------------------
# 應用程式實例 (App Instance)
# ----------------------------------------------------
app = FastAPI(
    title="測試報告自動化處理 API",
    description="""
這是一個用於處理多階段測試報告的 API 伺服器。

- **Phase 1**: (PDF -> Excel) 處理 PDF 規格書，產出 Excel 測試參數。
- **Phase 2**: (Excel + TXT -> CSV) 比對 Excel 參數與 TXT 報告，產出 CSV。
- **Phase 3**: (CSV -> CSV) 正規化 CSV 中的 'parameter' 欄位名稱。
- **Phase 4**: (CSV -> Excel) 根據 CSV 產出最終的 Excel 報告模板。
    """,
    version="1.0.0",
    lifespan=lifespan
)

# ----------------------------------------------------
# 中介軟體 (Middleware)
# ----------------------------------------------------
# 允許所有來源 (CORS)
# 注意：在生產環境中，應限制 origins
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], # 允許所有來源
    allow_credentials=True,
    allow_methods=["*"], # 允許所有 HTTP 方法
    allow_headers=["*"], # 允許所有標頭
)

# ----------------------------------------------------
# 靜態路由 (Static Routes)
# ----------------------------------------------------
@app.get("/", include_in_schema=False)
async def read_index():
    """
    提供根路由，用於顯示 index.html。
    """
    if not os.path.exists("index.html"):
        raise HTTPException(status_code=404, detail="index.html not found.")
    return FileResponse("index.html")

# --- 【【新增以下兩個路由】】 ---

@app.get("/style.css", include_in_schema=False)
async def get_css():
    """
    提供 style.css 檔案。
    """
    if not os.path.exists("style.css"):
        raise HTTPException(status_code=404, detail="style.css not found.")
    return FileResponse("style.css", media_type="text/css")

@app.get("/script.js", include_in_schema=False)
async def get_js():
    """
    提供 script.js 檔案。
    """
    if not os.path.exists("script.js"):
        raise HTTPException(status_code=404, detail="script.js not found.")
    return FileResponse("script.js", media_type="application/javascript")

# ----------------------------------------------------
# 路由 1: 階段一 (PDF -> Excel)
# ----------------------------------------------------
@app.post("/process-and-download/", tags=["Phase 1 - PDF Processing"])
async def process_pdf_and_download_excel(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(..., description="要上傳的 PDF 規格書")
):
    """
    接收 PDF 檔案，呼叫 `階段一.py` (P1) 模組處理。
    - P1 模組負責 OCR 和 Gemini API 分析。
    - 成功後，回傳處理完的 `stage1_output.xlsx` 檔案。
    """
    
    # 檢查檔案類型
    if file.content_type != 'application/pdf':
        raise HTTPException(status_code=400, detail="無效的檔案類型，請上傳 PDF。")
        
    temp_pdf = None
    temp_excel = None

    try:
        # 1. 建立 PDF 暫存檔案
        temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        with temp_pdf as f:
            f.write(await file.read())
            
        # 2. 準備 Excel 輸出路徑 (P1 函式會幫我們建立它)
        temp_excel_path = temp_pdf.name.replace(".pdf", ".xlsx")

        print(f"【非同步 階段一】: (輸入) {temp_pdf.name}, (輸出) {temp_excel_path}")

        # 3. 將 I/O 密集的處理任務 (包含 P1 的 API 呼叫) 丟到背景執行緒
        success = await run_in_threadpool(
            P1.process_pdf_to_excel, 
            temp_pdf.name, 
            temp_excel_path
        )
        
        # 4. (重要) 將「兩個」檔案都加入背景清理任務
        # 這裡我們傳入檔案「路徑」，因為 temp_excel 物件我們沒有
        background_tasks.add_task(cleanup_by_path, [temp_pdf.name, temp_excel_path])

        # 5. 檢查 P1 函式的執行結果
        if not success:
            print("【非同步 階段一】處理失敗，P1.process_pdf_to_excel 回傳 False。")
            raise HTTPException(status_code=500, detail="[Phase 1] 伺服器在處理 PDF 時發生錯誤 (Gemini API or OCR)。")

        print(f"【非同步 階段一】處理完成，結果: {success}")

        # 6. 回傳 Excel 檔案
        return StreamingResponse(
            open(temp_excel_path, "rb"),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename=stage1_output.xlsx"}
        )
        
    except Exception as e:
        print(f"❌【非同步 階段一】發生未預期的錯誤: {e}")
        # 確保即使發生錯誤，也嘗試清理檔案
        cleanup_paths = []
        if temp_pdf:
            cleanup_paths.append(temp_pdf.name)
        if 'temp_excel_path' in locals() and os.path.exists(temp_excel_path):
            cleanup_paths.append(temp_excel_path)
        
        if cleanup_paths:
            background_tasks.add_task(cleanup_by_path, cleanup_paths)
            
        raise HTTPException(status_code=500, detail=f"[Phase 1] 伺服器錯誤: {str(e)}")

# ----------------------------------------------------
# 路由 2: 階段二、三 (Excel + TXT -> CSV)
# ----------------------------------------------------
@app.post("/process-stage-two/", tags=["Phase 2 - Report Comparison"])
async def process_stage_two_files(
    background_tasks: BackgroundTasks,
    excel_file: UploadFile = File(..., description="階段一產出的 .xlsx 檔案"),
    txt_file: UploadFile = File(..., description="ATE 測試報告 .txt 檔案")
):
    """
    接收 Excel 和 TXT 檔案，呼叫 `report_processor.py` (P2) 模組處理。
    - P2 模組負責比對兩份檔案的測試項目。
    - 成功後，回傳比對結果 `stage2_output.csv` 檔案。
    """
    
    # 1. 建立三個暫存檔案 (Excel, TXT, CSV)
    temp_excel = None
    temp_txt = None
    temp_csv = None

    try:
        # 建立 Excel 暫存檔案
        temp_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        with temp_excel as f:
            f.write(await excel_file.read())

        # 建立 TXT 暫存檔案
        temp_txt = tempfile.NamedTemporaryFile(delete=False, suffix=".txt")
        with temp_txt as f:
            f.write(await txt_file.read())
            
        # 建立 CSV 輸出檔案 (P2 函式會寫入此檔案)
        temp_csv = tempfile.NamedTemporaryFile(delete=False, suffix=".csv")
        temp_csv.close() # 關閉檔案，讓 P2 函式可以寫入

        print(f"【非同步 階段二】: (Excel) {temp_excel.name}, (TXT) {temp_txt.name}, (CSV) {temp_csv.name}")

        # 2. 將 I/O 密集的 P2 處理任務丟到背景執行緒
        success = await run_in_threadpool(
            P2.process_report_to_csv, 
            temp_excel.name, 
            temp_txt.name, 
            temp_csv.name
        )
        
        # 3. (重要) 將所有暫存檔案加入背景清理任務
        background_tasks.add_task(cleanup, [temp_excel, temp_txt, temp_csv])

        # 4. 檢查 P2 函式的執行結果
        if not success:
            print("【非同步 階段二】處理失敗，P2.process_report_to_csv 回傳 False。")
            raise HTTPException(status_code=500, detail="[Phase 2] 伺服器在比對 Excel 和 TXT 時發生錯誤。")

        print(f"【非同步 階段二】處理完成，結果: {success}")

        # 5. 回傳 CSV 檔案
        return StreamingResponse(
            open(temp_csv.name, "rb"),
            media_type="text/csv",
            headers={"Content-Disposition": f"attachment; filename=stage2_output.csv"}
        )
        
    except Exception as e:
        print(f"❌【非同步 階段二】發生未預期的錯誤: {e}")
        # 確保即使發生錯誤，也清理所有已建立的檔案
        background_tasks.add_task(cleanup, [temp_excel, temp_txt, temp_csv])
        raise HTTPException(status_code=500, detail=f"[Phase 2] 伺服器錯誤: {str(e)}")


# ----------------------------------------------------
# 路由 3: (新) 階段 2.5 - 正規化參數名稱
# ----------------------------------------------------
@app.post("/normalize-parameters/", tags=["Phase 3 - Normalization"])
async def normalize_csv_parameters(
    background_tasks: BackgroundTasks,
    csv_file: UploadFile = File(..., description="從階段二產出的 stage2_output.csv")
):
    """
    接收 stage2_output.csv，
    執行 DataTransformer 中的決策樹邏輯來「正規化」parameter 欄位名稱。
    (例如 MIN. LOAD ON/OFF TEST -> MIN. LOAD ON_OFF TEST_1)
    
    完成後，回傳「已被正規化」的 CSV 檔案。
    """
    
    # 1. 建立兩個暫存檔案 (一個讀取，一個寫入)
    temp_in_csv = None
    temp_out_csv = None
    temp_out_csv_path = "" # 確保變數存在
    
    try:
        # 建立輸入檔案
        temp_in_csv = tempfile.NamedTemporaryFile(delete=False, suffix=".csv")
        with temp_in_csv as f:
            f.write(await csv_file.read())
        
        # 建立輸出檔案的路徑 (P3 函式會幫我們建立它)
        temp_out_csv_path = temp_in_csv.name.replace(".csv", "_normalized.csv")

        print(f"【非同步 階段三】: (輸入) {temp_in_csv.name}, (輸出) {temp_out_csv_path}")

        # 2. 將 I/O 密集的處理任務丟到背景執行緒
        success = await run_in_threadpool(
            P3.normalize_csv,  # 呼叫我們的新函式
            temp_in_csv.name,  # 傳入輸入路徑
            temp_out_csv_path  # 傳入輸出路徑
        )
        
        # 3. (重要) 將「兩個」檔案都加入背景清理任務
        # 這裡我們傳入檔案「路徑」，因為 temp_out_csv 物件我們沒有
        background_tasks.add_task(cleanup_by_path, [temp_in_csv.name, temp_out_csv_path])

        # 4. 檢查 P3 函式的執行結果
        if not success:
            print("【非同步 階段三】處理失敗，P3.normalize_csv 回傳 False。")
            raise HTTPException(status_code=500, detail="[Phase 3] 伺服器在正規化 CSV 參數時發生錯誤。")
        
        print(f"【非同步 階段三】處理完成，結果: {success}")

        # 5. 回傳「處理過的」CSV 檔案
        return StreamingResponse(
            open(temp_out_csv_path, "rb"),
            media_type="text/csv",
            headers={"Content-Disposition": f"attachment; filename=stage2_normalized.csv"}
        )

    except Exception as e:
        print(f"❌【非同步 階段三】發生未預期的錯誤: {e}")
        # 確保即使發生錯誤，也嘗試清理檔案
        cleanup_paths = []
        if temp_in_csv:
            cleanup_paths.append(temp_in_csv.name)
        if temp_out_csv_path and os.path.exists(temp_out_csv_path):
            cleanup_paths.append(temp_out_csv_path)
        
        if cleanup_paths:
            background_tasks.add_task(cleanup_by_path, cleanup_paths)
            
        raise HTTPException(status_code=500, detail=f"[Phase 3] 伺服器錯誤: {str(e)}")


# ----------------------------------------------------
# 路由 4: 階段四 (CSV -> Excel 模板)
# ----------------------------------------------------
@app.post("/create-report-template/", tags=["Phase 4 - Report Generation"])
async def create_report_template(
    background_tasks: BackgroundTasks,
    csv_file: UploadFile = File(..., description="從階段二 (或三) 產出的 .csv 檔案")
):
    """
    接收 stage2_output.csv 檔案，呼叫 `DataTransformer.py` (P4) 模組處理。
    - P4 模組負責動態生成多工作表的 Excel 報告模板。
    - 成功後，回傳最終的 `output_report.xlsx` 檔案。
    """
    
    # 1. 建立兩個暫存檔案 (CSV, Excel)
    temp_csv = None
    temp_excel = None

    try:
        # 建立 CSV 暫存檔案
        temp_csv = tempfile.NamedTemporaryFile(delete=False, suffix=".csv")
        with temp_csv as f:
            f.write(await csv_file.read())
            
        # 建立 Excel 輸出檔案 (P4 函式會寫入此檔案)
        temp_excel_path = temp_csv.name.replace(".csv", ".xlsx")

        print(f"【非同步 階段四】: (輸入) {temp_csv.name}, (輸出) {temp_excel_path}")

        # 2. 將 I/O 密集的 P4 處理任務丟到背景執行緒
        success = await run_in_threadpool(
            P4.create_report_template, 
            temp_csv.name, 
            temp_excel_path # 【修正】: 補上 P4 函式需要的第二個參數
        )
        
        # 3. (重要) 將「兩個」檔案都加入背景清理任務
        background_tasks.add_task(cleanup_by_path, [temp_csv.name, temp_excel_path])

        # 4. 檢查 P4 函式的執行結果
        if not success:
            print("【非同步 階段四】處理失敗，P4.create_report_template 回傳 False。")
            raise HTTPException(status_code=500, detail="[Phase 4] 伺服器在生成 Excel 報告模板時發生錯誤。")

        print(f"【非同步 階段四】處理完成，結果: {success}")

        # 5. 回傳 Excel 檔案
        return StreamingResponse(
            open(temp_excel_path, "rb"),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename=output_report.xlsx"}
        )
        
    except Exception as e:
        print(f"❌【非同步 階段四】發生未預期的錯誤: {e}")
        # 確保即使發生錯誤，也嘗試清理檔案
        cleanup_paths = []
        if temp_csv:
            cleanup_paths.append(temp_csv.name)
        if 'temp_excel_path' in locals() and os.path.exists(temp_excel_path):
            cleanup_paths.append(temp_excel_path)
            
        if cleanup_paths:
            background_tasks.add_task(cleanup_by_path, cleanup_paths)
            
        raise HTTPException(status_code=500, detail=f"[Phase 4] 伺服器錯誤: {str(e)}")

# ----------------------------------------------------
# 路由 5: (新) 階段五 (CSV -> JSON 映射)
# ----------------------------------------------------
@app.post("/create-data-mapping/", tags=["Phase 5 - Data Mapping (JSON)"])
async def create_data_mapping_json(
    background_tasks: BackgroundTasks,
    csv_file: UploadFile = File(..., description="從階段二 (或三) 產出的 .csv 檔案")
):
    """
    接收 stage2_output.csv 檔案，呼叫 `轉換Dict功能.py` (P5) 模組處理。
    - P5 模組負責將 CSV 的 step, line, list 欄位轉換為 JSON 格式的資料映射字典。
    - 成功後，回傳最終的 `data_mapping.json` 檔案。
    """
    
    # 1. 建立兩個暫存檔案 (CSV 輸入, JSON 輸出)
    temp_csv = None
    temp_json_path = "" # 確保變數存在

    try:
        # 建立 CSV 暫存檔案
        temp_csv = tempfile.NamedTemporaryFile(delete=False, suffix=".csv")
        with temp_csv as f:
            f.write(await csv_file.read())
            
        # 建立 JSON 輸出檔案的路徑 (P5 函式會寫入此檔案)
        temp_json_path = temp_csv.name.replace(".csv", ".json")

        print(f"【非同步 階段五】: (輸入) {temp_csv.name}, (輸出) {temp_json_path}")

        # 2. 將 I/O 密集的 P5 處理任務丟到背景執行緒
        success = await run_in_threadpool(
            P5.data_mapping_csv_to_dict, # 【【注意】】 呼叫 P5 的函式
            temp_csv.name, 
            temp_json_path
        )
        
        # 3. (重要) 將「兩個」檔案都加入背景清理任務
        background_tasks.add_task(cleanup_by_path, [temp_csv.name, temp_json_path])

        # 4. 檢查 P5 函式的執行結果
        if not success:
            print("【非同步 階段五】處理失敗，P5.data_mapping_csv_to_dict 回傳 False。")
            raise HTTPException(status_code=500, detail="[Phase 5] 伺服器在生成 JSON 映射時發生錯誤。")

        print(f"【非同步 階段五】處理完成，結果: {success}")

        # 5. 回傳 JSON 檔案
        return StreamingResponse(
            open(temp_json_path, "rb"),
            media_type="application/json", # 【【注意】】 媒體類型改為 JSON
            headers={"Content-Disposition": f"attachment; filename=data_mapping.json"}
        )
        
    except Exception as e:
        print(f"❌【非同步 階段五】發生未預期的錯誤: {e}")
        # 確保即使發生錯誤，也嘗試清理檔案
        cleanup_paths = []
        if temp_csv:
            cleanup_paths.append(temp_csv.name)
        if 'temp_json_path' in locals() and os.path.exists(temp_json_path):
            cleanup_paths.append(temp_json_path)
            
        if cleanup_paths:
            background_tasks.add_task(cleanup_by_path, cleanup_paths)
            
        raise HTTPException(status_code=500, detail=f"[Phase 5] 伺服器錯誤: {str(e)}")

# ----------------------------------------------------
# 路由 6: (新) 階段六 (TXTs + Excel + JSON -> 完整報表)
# ----------------------------------------------------
@app.post("/fill-report-data/", tags=["Phase 6 - Report Filling"])
async def fill_report_data(
    background_tasks: BackgroundTasks,
    excel_template: UploadFile = File(..., description="P4 產出的 Excel 報告模板 (.xlsx)"),
    json_mapping: UploadFile = File(..., description="P5 產出的 JSON 資料映射 (.json)"),
    txt_reports: List[UploadFile] = File(..., description="ATE 測試報告 TXT 檔案 (可多選)")
):
    """
    接收 P4 模板 (Excel), P5 映射 (JSON), 以及多個 TXT 報告。
    呼叫 `自動輸出Excel報表功能.py` (P6) 模組處理。
    - P6 模組會遍歷所有 TXT，根據 JSON 映射，將數據填入 Excel 模板。
    - 成功後，回傳最終的 `Product Test Report.xlsx` 檔案。
    """
    
    # P6 需要 4 個路徑：
    temp_excel_in = None          # (檔案) P4 模板
    temp_json_in = None           # (檔案) P5 映射
    temp_txt_dir_in = None        # (資料夾) 存放所有 TXT
    temp_excel_dir_out = None     # (資料夾) P6 寫入報表的地方
    
    # 建立一個列表來追蹤所有需要清理的路徑
    cleanup_list = []

    try:
        # 1. 建立 P4 Excel 暫存檔案
        temp_excel_in = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        temp_excel_in.write(await excel_template.read())
        temp_excel_in.close() # 關閉檔案，讓 P6 可以讀取
        cleanup_list.append(temp_excel_in.name)

        # 2. 建立 P5 JSON 暫存檔案
        temp_json_in = tempfile.NamedTemporaryFile(delete=False, suffix=".json")
        temp_json_in.write(await json_mapping.read())
        temp_json_in.close()
        cleanup_list.append(temp_json_in.name)

        # 3. 建立「輸入 TXT 的暫存資料夾」
        temp_txt_dir_in = tempfile.mkdtemp()
        cleanup_list.append(temp_txt_dir_in)
        
        print(f"【非同步 階段六】: 建立 TXT 暫存資料夾於: {temp_txt_dir_in}")
        # 遍歷所有上傳的 TXT 檔案並存入該資料夾
        for txt_file in txt_reports:
            txt_save_path = os.path.join(temp_txt_dir_in, txt_file.filename)
            with open(txt_save_path, "wb") as f:
                f.write(await txt_file.read())
        
        # 4. 建立「輸出 Excel 的暫存資料夾」 (P6 腳本要求一個輸出 "目錄")
        temp_excel_dir_out = tempfile.mkdtemp()
        cleanup_list.append(temp_excel_dir_out)
        
        print(f"【非同步 階段六】: (Excel) {temp_excel_in.name}, (JSON) {temp_json_in.name}")
        print(f"【非同步 階段六】: (TXT Dir) {temp_txt_dir_in}, (Out Dir) {temp_excel_dir_out}")

        # 5. 將 I/O 密集的 P6 處理任務丟到背景執行緒
        success = await run_in_threadpool(
            P6.export_excel, 
            temp_excel_in.name,      # raw_excel_path
            temp_txt_dir_in,         # txt_folder_path
            temp_json_in.name,       # data_postion_map
            temp_excel_dir_out       # output_excel_path
        )
        
        # 6. (重要) 將「所有」暫存資源加入背景清理任務
        background_tasks.add_task(cleanup_dirs_and_files, cleanup_list)

        # 7. 檢查 P6 函式的執行結果
        if not success:
            print("【非同步 階段六】處理失敗，P6.export_excel 回傳 False。")
            raise HTTPException(status_code=500, detail="[Phase 6] 伺服器在填寫 Excel 報告時發生錯誤。")

        # 8. 準備 P6 產出的檔案路徑
        final_report_path = os.path.join(temp_excel_dir_out, "Product Test Report.xlsx")
        if not os.path.exists(final_report_path):
             print("【非同步 階段六】處理成功，但 P6 未產生 'Product Test Report.xlsx' 檔案。")
             raise HTTPException(status_code=500, detail="[Phase 6] 伺服器錯誤，P6 未產生預期的輸出檔案。")

        print(f"【非同步 階段六】處理完成，回傳檔案: {final_report_path}")

        # 9. 回傳最終填滿的 Excel 報告
        return StreamingResponse(
            open(final_report_path, "rb"),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename=Product_Test_Report_Filled.xlsx"}
        )
        
    except Exception as e:
        print(f"❌【非同步 階段六】發生未預期的錯誤: {e}")
        # 確保即使發生錯誤，也嘗試清理所有已建立的資源
        background_tasks.add_task(cleanup_dirs_and_files, cleanup_list)
        raise HTTPException(status_code=500, detail=f"[Phase 6] 伺服器錯誤: {str(e)}")


# ----------------------------------------------------
# 輔助工具 (Helpers)
# ----------------------------------------------------
# ( ... 接著是你原有的 cleanup_by_path 和 cleanup 函式 ...)
# ----------------------------------------------------
# 輔助工具 (Helpers)
# ----------------------------------------------------
# ( ... 接著是 cleanup_by_path 等函式 ...)
# ----------------------------------------------------
# 輔助工具 (Helpers)
# ----------------------------------------------------
def cleanup_dirs_and_files(paths: list[str]):
    """
    (背景任務) 清理檔案「和」資料夾。
    """
    print(f"【背景清理 (Dirs/Files)】: 準備清理 {len(paths)} 個資源...")
    for path in paths:
        if not path or not os.path.exists(path):
            print(f"  > 無需清理，路徑不存在: {path}")
            continue
        try:
            if os.path.isdir(path):
                shutil.rmtree(path) # 【【關鍵】】 使用 shutil 刪除資料夾
                print(f"  > 已清理 (Directory): {path}")
            else:
                os.unlink(path) # (維持原有功能)
                print(f"  > 已清理 (File): {path}")
        except Exception as e:
            print(f"  > 清理失敗 {path}: {e}")
# 【新增】(新函式) 透過路徑清理
def cleanup_by_path(file_paths: list[str]):
    """
    (背景任務) 根據檔案「路徑」刪除檔案。
    適用於無法取得 tempfile 物件時。
    """
    print(f"【背景清理 (By Path)】: 準備清理 {len(file_paths)} 個檔案...")
    for path in file_paths:
        if path and isinstance(path, str) and os.path.exists(path):
            try:
                os.unlink(path)
                print(f"  > 已清理 (By Path): {path}")
            except Exception as e:
                print(f"  > 清理失敗 (By Path) {path}: {e}")
        elif not os.path.exists(path):
             print(f"  > 無需清理 (By Path)，檔案已不存在: {path}")

# (原有函式) 透過物件清理
def cleanup(temp_files: list[Optional[tempfile._TemporaryFileWrapper]]):
    """
    (背景任務) 關閉並刪除暫存檔案。
    """
    print(f"【背景清理】: 準備清理 {len(temp_files)} 個暫存檔案物件...")
    for temp_file in temp_files:
        if temp_file:
            try:
                temp_file.close() # 關閉檔案
                if os.path.exists(temp_file.name):
                    os.unlink(temp_file.name) # 刪除檔案
                print(f"  > 已清理 (By Object): {temp_file.name}")
            except Exception as e:
                # 即使關閉/刪除失敗，也繼續處理下一個
                print(f"  > 清理失敗 (By Object) {temp_file.name}: {e}")

# ----------------------------------------------------
# 伺服器啟動 (Server Boot)
# ----------------------------------------------------
if __name__ == "__main__":
    import uvicorn
    # 執行 Uvicorn 伺服器
    # host="0.0.0.0" 讓區域網路內的其他裝置可以連線
    # reload=True 讓程式碼變動時自動重載 (僅限開發時使用)
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)