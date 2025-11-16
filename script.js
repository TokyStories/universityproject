// (所有 JS 程式碼合併於此檔案)

// ----------------------------------------------------
// 階段一
// ----------------------------------------------------
const fileInput = document.getElementById('pdf-file');
const fileNameDisplay = document.getElementById('file-name');
const uploadButton = document.getElementById('upload-button');
const loader = document.getElementById('loader');
const statusMessage = document.getElementById('status-message');

fileInput.addEventListener('change', () => {
    if (fileInput.files.length > 0) {
        const fileName = fileInput.files[0].name;
        fileNameDisplay.textContent = fileName;
        fileNameDisplay.classList.add('selected'); 
        uploadButton.disabled = false; 
        statusMessage.textContent = ''; 
    } else {
        fileNameDisplay.textContent = '尚未選擇檔案';
        fileNameDisplay.classList.remove('selected'); 
        uploadButton.disabled = true; 
    }
});

uploadButton.addEventListener('click', async () => {
    if (fileInput.files.length === 0) {
        statusMessage.textContent = '請先選擇一個 PDF 檔案。';
        return;
    }
    const file = fileInput.files[0];
    const formData = new FormData();
    formData.append('file', file);
    uploadButton.disabled = true;
    uploadButton.textContent = '處理中，請稍候...';
    loader.style.display = 'block';
    statusMessage.textContent = '';

    try {
        const response = await fetch('/process-and-download/', { 
            method: 'POST',
            body: formData,
        });
        if (response.ok) {
            const blob = await response.blob();
            const downloadUrl = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = downloadUrl;
            a.download = 'stage1_output.xlsx'; // 修正：對應 main.py 檔名
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(downloadUrl);
            statusMessage.textContent = '✅ 處理完成，已觸發下載！';
        } else {
            const errorData = await response.json(); 
            statusMessage.textContent = `❌ 處理失敗：${errorData.detail || response.statusText}`;
        }
    } catch (error) {
        console.error('上傳失敗 (階段一):', error);
        statusMessage.textContent = '❌ 網路錯誤，無法連線至伺服器。';
    } finally {
        uploadButton.disabled = false;
        uploadButton.textContent = '上傳並處理';
        loader.style.display = 'none';
        fileInput.value = ''; 
        fileNameDisplay.textContent = '尚未選擇檔案';
        fileNameDisplay.classList.remove('selected'); 
    }
});

// ----------------------------------------------------
// 階段二
// ----------------------------------------------------
const fileInput_excel_p2 = document.getElementById('file-excel-p2');
const fileInput_txt_p2 = document.getElementById('file-txt-p2');

const fileNameDisplay_excel_p2 = document.getElementById('file-name-excel-p2');
const fileNameDisplay_txt_p2 = document.getElementById('file-name-txt-p2');

const uploadButton_p2 = document.getElementById('upload-button-p2');
const loader_p2 = document.getElementById('loader-p2');
const statusMessage_p2 = document.getElementById('status-message-p2');

// 監聽 Excel 檔案輸入
fileInput_excel_p2.addEventListener('change', () => {
    if (fileInput_excel_p2.files.length > 0) {
        fileNameDisplay_excel_p2.textContent = fileInput_excel_p2.files[0].name;
        fileNameDisplay_excel_p2.classList.add('selected');
    } else {
        fileNameDisplay_excel_p2.textContent = '尚未選擇 Excel';
        fileNameDisplay_excel_p2.classList.remove('selected');
    }
    checkPhaseTwoFiles(); // 檢查是否兩個都選了
});

// 監聽 TXT 檔案輸入
fileInput_txt_p2.addEventListener('change', () => {
    if (fileInput_txt_p2.files.length > 0) {
        fileNameDisplay_txt_p2.textContent = fileInput_txt_p2.files[0].name;
        fileNameDisplay_txt_p2.classList.add('selected');
    } else {
        fileNameDisplay_txt_p2.textContent = '尚未選擇 TXT';
        fileNameDisplay_txt_p2.classList.remove('selected');
    }
    checkPhaseTwoFiles(); // 檢查是否兩個都選了
});

// 檢查函式：只有兩個檔案都選了，才啟用上傳按鈕
function checkPhaseTwoFiles() {
    statusMessage_p2.textContent = ''; // 清除舊錯誤
    if (fileInput_excel_p2.files.length > 0 && fileInput_txt_p2.files.length > 0) {
        uploadButton_p2.disabled = false;
    } else {
        uploadButton_p2.disabled = true;
    }
}

// 當使用者點擊 階段二 上傳按鈕時
uploadButton_p2.addEventListener('click', async () => {
    
    // 再次確認檔案
    if (fileInput_excel_p2.files.length === 0 || fileInput_txt_p2.files.length === 0) {
        statusMessage_p2.textContent = '錯誤：必須同時提供 Excel 和 TXT 檔案。';
        return;
    }

    const excelFile = fileInput_excel_p2.files[0];
    const txtFile = fileInput_txt_p2.files[0];
    
    // 準備 FormData (包含兩個檔案)
    const formData = new FormData();
    
    formData.append('excel_file', excelFile); // 對應 FastAPI 函式的 'excel_file'
    formData.append('txt_file', txtFile);   // 對應 FastAPI 函式的 'txt_file'

    // 更新 UI (進入讀取狀態)
    uploadButton_p2.disabled = true;
    uploadButton_p2.textContent = '比對中，請稍候...';
    loader_p2.style.display = 'block';
    statusMessage_p2.textContent = '';

    try {
        // 發送 fetch 請求
        const response = await fetch('/process-stage-two/', {
            method: 'POST',
            body: formData,
        });

        // 處理回應
        if (response.ok) {
            // --- 成功：觸發下載 CSV ---
            const blob = await response.blob();
            const downloadUrl = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = downloadUrl;
            
            a.download = 'stage2_output.csv'; 
            
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(downloadUrl);
            statusMessage_p2.textContent = '✅ 比對完成，已觸發 CSV 下載！';
            
        } else {
            // --- 失敗：顯示錯誤訊息 ---
            const errorData = await response.json();
            statusMessage_p2.textContent = `❌ 處理失敗：${errorData.detail || response.statusText}`;
        }

    } catch (error) {
        // --- 網路或其他錯誤 ---
        console.error('上傳失敗 (階段二):', error);
        statusMessage_p2.textContent = '❌ 網路錯誤，無法連線至伺服器。';
    } finally {
        // 恢復 UI
        uploadButton_p2.disabled = false;
        uploadButton_p2.textContent = '上傳並比對';
        loader_p2.style.display = 'none';
        
        // 清空檔案選擇
        fileInput_excel_p2.value = ''; 
        fileInput_txt_p2.value = '';
        
        // 重設顯示文字
        fileNameDisplay_excel_p2.textContent = '尚未選擇 Excel';
        fileNameDisplay_txt_p2.textContent = '尚未選擇 TXT';
        fileNameDisplay_excel_p2.classList.remove('selected');
        fileNameDisplay_txt_p2.classList.remove('selected');
        
        checkPhaseTwoFiles(); // 確保按鈕被禁用
    }
});

// ----------------------------------------------------
// 階段三 (複製 P4 邏輯並修改)
// ----------------------------------------------------
const fileInput_csv_p3 = document.getElementById('file-csv-p3');
const fileNameDisplay_csv_p3 = document.getElementById('file-name-csv-p3');
const uploadButton_p3 = document.getElementById('upload-button-p3');
const loader_p3 = document.getElementById('loader-p3');
const statusMessage_p3 = document.getElementById('status-message-p3');

// 監聽 CSV 檔案輸入
fileInput_csv_p3.addEventListener('change', () => {
    if (fileInput_csv_p3.files.length > 0) {
        fileNameDisplay_csv_p3.textContent = fileInput_csv_p3.files[0].name;
        fileNameDisplay_csv_p3.classList.add('selected');
        uploadButton_p3.disabled = false;
        statusMessage_p3.textContent = '';
    } else {
        fileNameDisplay_csv_p3.textContent = '尚未選擇 CSV';
        fileNameDisplay_csv_p3.classList.remove('selected');
        uploadButton_p3.disabled = true;
    }
});

// 當使用者點擊 階段三 上傳按鈕時
uploadButton_p3.addEventListener('click', async () => {
    
    // 再次確認檔案
    if (fileInput_csv_p3.files.length === 0) {
        statusMessage_p3.textContent = '錯誤：請先選擇一個 CSV 檔案。';
        return;
    }

    const csvFile = fileInput_csv_p3.files[0];
    
    // 準備 FormData
    const formData = new FormData();
    formData.append('csv_file', csvFile); // 對應 main.py 的 'csv_file'

    // 更新 UI (進入讀取狀態)
    uploadButton_p3.disabled = true;
    uploadButton_p3.textContent = '正規化中...';
    loader_p3.style.display = 'block';
    statusMessage_p3.textContent = '';

    try {
        // 發送 fetch 請求
        const response = await fetch('/normalize-parameters/', { // 對應 P3 路由
            method: 'POST',
            body: formData,
        });

        // 處理回應
        if (response.ok) {
            // --- 成功：觸發下載 Excel ---
            const blob = await response.blob();
            const downloadUrl = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = downloadUrl;
            
            a.download = 'stage3_normalized.csv'; // P3 下載檔名
            
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(downloadUrl);
            statusMessage_p3.textContent = '✅ 正規化完成，已觸發 CSV 下載！';
            
        } else {
            // --- 失敗：顯示錯誤訊息 ---
            const errorData = await response.json();
            statusMessage_p3.textContent = `❌ 處理失敗：${errorData.detail || response.statusText}`;
        }

    } catch (error) {
        // --- 網路或其他錯誤 ---
        console.error('上傳失敗 (階段三):', error);
        statusMessage_p3.textContent = '❌ 網路錯誤，無法連線至伺服器。';
    } finally {
        // 恢復 UI
        uploadButton_p3.disabled = false;
        uploadButton_p3.textContent = '上傳並正規化';
        loader_p3.style.display = 'none';
        
        // 清空檔案選擇
        fileInput_csv_p3.value = ''; 
        
        // 重設顯示文字
        fileNameDisplay_csv_p3.textContent = '尚未選擇 CSV';
        fileNameDisplay_csv_p3.classList.remove('selected');
        
        uploadButton_p3.disabled = true; // 確保按鈕被禁用
    }
});

// ----------------------------------------------------
// 階段四
// ----------------------------------------------------
const fileInput_csv_p4 = document.getElementById('file-csv-p4');
const fileNameDisplay_csv_p4 = document.getElementById('file-name-csv-p4');
const uploadButton_p4 = document.getElementById('upload-button-p4');
const loader_p4 = document.getElementById('loader-p4');
const statusMessage_p4 = document.getElementById('status-message-p4');

// 監聽 CSV 檔案輸入
fileInput_csv_p4.addEventListener('change', () => {
    if (fileInput_csv_p4.files.length > 0) {
        fileNameDisplay_csv_p4.textContent = fileInput_csv_p4.files[0].name;
        fileNameDisplay_csv_p4.classList.add('selected');
        uploadButton_p4.disabled = false;
        statusMessage_p4.textContent = '';
    } else {
        fileNameDisplay_csv_p4.textContent = '尚未選擇 CSV';
        fileNameDisplay_csv_p4.classList.remove('selected');
        uploadButton_p4.disabled = true;
    }
});

// 當使用者點擊 階段四 上傳按鈕時
uploadButton_p4.addEventListener('click', async () => {
    
    // 再次確認檔案
    if (fileInput_csv_p4.files.length === 0) {
        statusMessage_p4.textContent = '錯誤：請先選擇一個 CSV 檔案。';
        return;
    }

    const csvFile = fileInput_csv_p4.files[0];
    
    // 準備 FormData
    const formData = new FormData();
    formData.append('csv_file', csvFile); // 對應 main.py 的 'csv_file'

    // 更新 UI (進入讀取狀態)
    uploadButton_p4.disabled = true;
    uploadButton_p4.textContent = '生成模板中...';
    loader_p4.style.display = 'block';
    statusMessage_p4.textContent = '';

    try {
        // 發送 fetch 請求
        const response = await fetch('/create-report-template/', { // 對應 main.py 的新路由
            method: 'POST',
            body: formData,
        });

        // 處理回應
        if (response.ok) {
            // --- 成功：觸發下載 Excel ---
            const blob = await response.blob();
            const downloadUrl = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = downloadUrl;
            
            a.download = 'output_report.xlsx'; // 對應 main.py 的檔名
            
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(downloadUrl);
            statusMessage_p4.textContent = '✅ 模板生成完成，已觸發 Excel 下載！';
            
        } else {
            // --- 失敗：顯示錯誤訊息 ---
            const errorData = await response.json();
            statusMessage_p4.textContent = `❌ 處理失敗：${errorData.detail || response.statusText}`;
        }

    } catch (error) {
        // --- 網路或其他錯誤 ---
        console.error('上傳失敗 (階段四):', error);
        statusMessage_p4.textContent = '❌ 網路錯誤，無法連線至伺服器。';
    } finally {
        // 恢復 UI
        uploadButton_p4.disabled = false;
        uploadButton_p4.textContent = '上傳並生成模板';
        loader_p4.style.display = 'none';
        
        // 清空檔案選擇
        fileInput_csv_p4.value = ''; 
        
        // 重設顯示文字
        fileNameDisplay_csv_p4.textContent = '尚未選擇 CSV';
        fileNameDisplay_csv_p4.classList.remove('selected');
        
        uploadButton_p4.disabled = true; // 確保按鈕被禁用
    }
});
// ----------------------------------------------------
// 階段五 (複製 P4 邏輯並修改)
// ----------------------------------------------------
const fileInput_csv_p5 = document.getElementById('file-csv-p5');
const fileNameDisplay_csv_p5 = document.getElementById('file-name-csv-p5');
const uploadButton_p5 = document.getElementById('upload-button-p5');
const loader_p5 = document.getElementById('loader-p5');
const statusMessage_p5 = document.getElementById('status-message-p5');

// 監聽 CSV 檔案輸入
fileInput_csv_p5.addEventListener('change', () => {
    if (fileInput_csv_p5.files.length > 0) {
        fileNameDisplay_csv_p5.textContent = fileInput_csv_p5.files[0].name;
        fileNameDisplay_csv_p5.classList.add('selected');
        uploadButton_p5.disabled = false;
        statusMessage_p5.textContent = '';
    } else {
        fileNameDisplay_csv_p5.textContent = '尚未選擇 CSV';
        fileNameDisplay_csv_p5.classList.remove('selected');
        uploadButton_p5.disabled = true;
    }
});

// 當使用者點擊 階段五 上傳按鈕時
uploadButton_p5.addEventListener('click', async () => {
    
    // 再次確認檔案
    if (fileInput_csv_p5.files.length === 0) {
        statusMessage_p5.textContent = '錯誤：請先選擇一個 CSV 檔案。';
        return;
    }

    const csvFile = fileInput_csv_p5.files[0];
    
    // 準備 FormData
    const formData = new FormData();
    formData.append('csv_file', csvFile); // 對應 main.py 的 'csv_file'

    // 更新 UI (進入讀取狀態)
    uploadButton_p5.disabled = true;
    uploadButton_p5.textContent = '生成映射中...'; // 【修改】
    loader_p5.style.display = 'block';
    statusMessage_p5.textContent = '';

    try {
        // 發送 fetch 請求
        const response = await fetch('/create-data-mapping/', { // 【【關鍵修改】】 對應 P5 路由
            method: 'POST',
            body: formData,
        });

        // 處理回應
        if (response.ok) {
            // --- 成功：觸發下載 JSON ---
            const blob = await response.blob();
            const downloadUrl = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = downloadUrl;
            
            a.download = 'data_mapping.json'; // 【【關鍵修改】】 P5 下載檔名
            
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(downloadUrl);
            statusMessage_p5.textContent = '✅ 映射生成完成，已觸發 JSON 下載！'; // 【修改】
            
        } else {
            // --- 失敗：顯示錯誤訊息 ---
            const errorData = await response.json();
            statusMessage_p5.textContent = `❌ 處理失敗：${errorData.detail || response.statusText}`;
        }

    } catch (error) {
        // --- 網路或其他錯誤 ---
        console.error('上傳失敗 (階段五):', error); // 【修改】
        statusMessage_p5.textContent = '❌ 網路錯誤，無法連線至伺服器。';
    } finally {
        // 恢復 UI
        uploadButton_p5.disabled = false;
        uploadButton_p5.textContent = '上傳並生成映射'; // 【修改】
        loader_p5.style.display = 'none';
        
        // 清空檔案選擇
        fileInput_csv_p5.value = ''; 
        
        // 重設顯示文字
        fileNameDisplay_csv_p5.textContent = '尚未選擇 CSV';
        fileNameDisplay_csv_p5.classList.remove('selected');
        
        uploadButton_p5.disabled = true; // 確保按鈕被禁用
    }
    
});
// ----------------------------------------------------
// 階段六 (複製 P2 邏輯並修改為三輸入)
// ----------------------------------------------------
const fileInput_excel_p6 = document.getElementById('file-excel-p6');
const fileInput_json_p6 = document.getElementById('file-json-p6');
const fileInput_txt_p6 = document.getElementById('file-txt-p6');

const fileNameDisplay_excel_p6 = document.getElementById('file-name-excel-p6');
const fileNameDisplay_json_p6 = document.getElementById('file-name-json-p6');
const fileNameDisplay_txt_p6 = document.getElementById('file-name-txt-p6');

const uploadButton_p6 = document.getElementById('upload-button-p6');
const loader_p6 = document.getElementById('loader-p6');
const statusMessage_p6 = document.getElementById('status-message-p6');

// 監聽 Excel 檔案輸入
fileInput_excel_p6.addEventListener('change', () => {
    if (fileInput_excel_p6.files.length > 0) {
        fileNameDisplay_excel_p6.textContent = fileInput_excel_p6.files[0].name;
        fileNameDisplay_excel_p6.classList.add('selected');
    } else {
        fileNameDisplay_excel_p6.textContent = '尚未選擇 Excel 模板';
        fileNameDisplay_excel_p6.classList.remove('selected');
    }
    checkPhaseSixFiles(); // 檢查是否三個都選了
});

// 監聽 JSON 檔案輸入
fileInput_json_p6.addEventListener('change', () => {
    if (fileInput_json_p6.files.length > 0) {
        fileNameDisplay_json_p6.textContent = fileInput_json_p6.files[0].name;
        fileNameDisplay_json_p6.classList.add('selected');
    } else {
        fileNameDisplay_json_p6.textContent = '尚未選擇 JSON 映射';
        fileNameDisplay_json_p6.classList.remove('selected');
    }
    checkPhaseSixFiles(); // 檢查是否三個都選了
});

// 監聽 TXT 檔案輸入 (多選)
fileInput_txt_p6.addEventListener('change', () => {
    if (fileInput_txt_p6.files.length > 0) {
        // 【【修改】】 顯示選擇的檔案數量
        fileNameDisplay_txt_p6.textContent = `已選擇 ${fileInput_txt_p6.files.length} 個 TXT 檔案`;
        fileNameDisplay_txt_p6.classList.add('selected');
    } else {
        fileNameDisplay_txt_p6.textContent = '尚未選擇 TXT 報告';
        fileNameDisplay_txt_p6.classList.remove('selected');
    }
    checkPhaseSixFiles(); // 檢查是否三個都選了
});


// 檢查函式：只有三個檔案都選了，才啟用上傳按鈕
function checkPhaseSixFiles() {
    statusMessage_p6.textContent = ''; // 清除舊錯誤
    if (fileInput_excel_p6.files.length > 0 && 
        fileInput_json_p6.files.length > 0 &&
        fileInput_txt_p6.files.length > 0) {
        uploadButton_p6.disabled = false;
    } else {
        uploadButton_p6.disabled = true;
    }
}

// 當使用者點擊 階段六 上傳按鈕時
uploadButton_p6.addEventListener('click', async () => {
    
    // 再次確認檔案
    if (fileInput_excel_p6.files.length === 0 || 
        fileInput_json_p6.files.length === 0 ||
        fileInput_txt_p6.files.length === 0) {
        statusMessage_p6.textContent = '錯誤：必須同時提供 Excel 模板、JSON 映射和 TXT 報告。';
        return;
    }

    const excelFile = fileInput_excel_p6.files[0];
    const jsonFile = fileInput_json_p6.files[0];
    const txtFiles = fileInput_txt_p6.files; // 
    
    // 準備 FormData (包含所有檔案)
    const formData = new FormData();
    
    formData.append('excel_template', excelFile); // 對應 main.py 的 'excel_template'
    formData.append('json_mapping', jsonFile);   // 對應 main.py 的 'json_mapping'
    
    // 【【關鍵】】 遍歷 FileList 並將所有 TXT 檔案加入
    for (let i = 0; i < txtFiles.length; i++) {
        formData.append('txt_reports', txtFiles[i]); // 對應 main.py 的 'txt_reports'
    }

    // 更新 UI (進入讀取狀態)
    uploadButton_p6.disabled = true;
    uploadButton_p6.textContent = '填寫報表中...';
    loader_p6.style.display = 'block';
    statusMessage_p6.textContent = '';

    try {
        // 發送 fetch 請求
        const response = await fetch('/fill-report-data/', { // 對應 P6 路由
            method: 'POST',
            body: formData,
        });

        // 處理回應
        if (response.ok) {
            // --- 成功：觸發下載 CSV ---
            const blob = await response.blob();
            const downloadUrl = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = downloadUrl;
            
            a.download = 'Product_Test_Report_Filled.xlsx'; // 對應 main.py 的檔名
            
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(downloadUrl);
            statusMessage_p6.textContent = '✅ 報表填寫完成，已觸發 Excel 下載！';
            
        } else {
            // --- 失敗：顯示錯誤訊息 ---
            const errorData = await response.json();
            statusMessage_p6.textContent = `❌ 處理失敗：${errorData.detail || response.statusText}`;
        }

    } catch (error) {
        // --- 網路或其他錯誤 ---
        console.error('上傳失敗 (階段六):', error);
        statusMessage_p6.textContent = '❌ 網路錯誤，無法連線至伺服器。';
    } finally {
        // 恢復 UI
        uploadButton_p6.disabled = false;
        uploadButton_p6.textContent = '上傳並填寫報表';
        loader_p6.style.display = 'none';
        
        // 清空【所有】檔案選擇
        fileInput_excel_p6.value = ''; 
        fileInput_json_p6.value = '';
        fileInput_txt_p6.value = '';
        
        // 重設【所有】顯示文字
        fileNameDisplay_excel_p6.textContent = '尚未選擇 Excel 模板';
        fileNameDisplay_json_p6.textContent = '尚未選擇 JSON 映射';
        fileNameDisplay_txt_p6.textContent = '尚未選擇 TXT 報告';
        
        fileNameDisplay_excel_p6.classList.remove('selected');
        fileNameDisplay_json_p6.classList.remove('selected');
        fileNameDisplay_txt_p6.classList.remove('selected');
        
        checkPhaseSixFiles(); // 確保按鈕被禁用
    }
});