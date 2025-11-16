import pandas as pd
from collections import defaultdict
import json

def data_mapping_csv_to_dict(data_mapping_csv_path: str, output_json_file_path: str) -> bool:
    # =================================== 設定檔案路徑 ====================================
    #data_mapping_csv_path: 整理完的CSV檔案路徑
    #output_data_mapping_path: 輸出的JSON檔案路徑

    # ====================================== 設定區 =======================================
    data_mapping_table = pd.read_csv(fr"{data_mapping_csv_path}") # 開啟 csv 檔案

    parameter = data_mapping_table['parameter']            # 抓取 parameter 資料
    step = data_mapping_table['step']                      # 抓取 step 資料
    row = data_mapping_table['line']                       # 抓取 line 資料
    column = data_mapping_table['list']                    # 抓取 list 資料
    table_length = len(data_mapping_table['parameter'])    # 需要處理的資料數
    default_vaule = lambda: ["B"]                          # 設定內層字典
    data_mapping_dict = defaultdict(lambda: defaultdict(default_vaule)) # 設定巢狀結構字典

    # ====================================== 副程式 ======================================
    """將從 1 開始的數字索引轉換為 Excel 的欄位名稱 (A, B, ..., Z, AA, AB, ...)。"""
    def index_to_excel_column(index): 
        if index < 1: 
            return ""  # 處理非正數的情況
        
        result = ""    # 初始化
        index += 2     # 最少從 C 開始，偏移量

        while (index > 0):
            index -= 1                 # 將數字減 1，以便從 0 開始處理
            remainder = index % 26     # 計算餘數
            char = chr(65 + remainder) # 轉換為大寫字母 (A 的 ASCII 碼是 65)
            result = char + result     # 將字符加到結果的前面
            index //= 26               # 更新 index
        
        return result

    """遞迴地將 defaultdict 及其內層所有 defaultdict 轉換為標準 dict。"""
    def convert_defaultdict_to_dict(d):
        if isinstance(d, defaultdict):
            # 遍歷 d 的鍵值對，對值 (v) 進行遞迴轉換，最後將結果打包成一個標準 dict
            d = {k: convert_defaultdict_to_dict(v) for k, v in d.items()}
        return d # 如果 d 不是 defaultdict (例如是 list 或 string)，則直接返回 d

    # ====================================== 主程式 ======================================
    for i in range(table_length):
        key_parameter = parameter[i]        # 輸入第一層字典
        key_step = f'STEP.{step[i]}('       # 輸入第二層字典
        raw_value = f'{row[i]} {column[i]}' # 設定第三層字典前半

        all_values = data_mapping_dict[key_parameter][key_step] # 取得所有Value
        value_counts = len(all_values)                          # 取得Vaule數量

        excel_column_char = index_to_excel_column(value_counts) # 取得 Excel 欄位英文字母
        all_values.append(fr"{raw_value} {excel_column_char}")  # 第三層字典前半與英文字母輸入

    final_dict = convert_defaultdict_to_dict(data_mapping_dict) # 轉換成原本的 DICT 結構

    with open(output_json_file_path, 'w', encoding='utf-8') as f:
            json.dump(final_dict, f, ensure_ascii=False, indent=4)  # 產生 JSON 檔案
        
        # ==================================================================================
    return True 