import pandas as pd
import numpy as np
import re
import sys
from typing import Union # (確保型別提示)

# --- 步驟 1: 定義最終的「名稱清理」函式 ---
# 這是我們在 DataTransformer.py 中最終確定的版本
def sanitize_sheet_name(name: str) -> str:
    """
    清理名稱以符合 Excel 工作表規則，並套用自訂規則 ( / -> _, #1 -> _1 )
    """
    # 1. 將 '/' 替換為 '_' (例如 "MIN. LOAD ON/OFF TEST")
    sanitized = name.replace('/', '_')
    
    # 2. 將 '#1' (或 ' #1' 或 '# 1') 替換為 '_1'
    #    (r'_\1' 會把它替換為 '_' 和捕獲到的數字)
    sanitized = re.sub(r'\s*#\s*(\d+)', r'_\1', sanitized)
    
    # 3. (安全機制) 將「任何剩下」的 '#' 替換為 '_' (處理像 'SPECIAL#CASE' 這樣的邊界情況)
    sanitized = sanitized.replace('#', '_')
    
    # 4. 移除 Excel 不允許的「其他」字元
    invalid_chars_to_remove = r'[\[\]\*\\:\?]' 
    sanitized = re.sub(invalid_chars_to_remove, '', sanitized)
    
    # 5. 截斷到 Excel 允許的 31 個字元
    return sanitized[:31]

# --- 步驟 2: 主執行函式 (修改為可匯入的版本) ---
def normalize_csv(input_csv_path: str, output_csv_path: str) -> bool:
    """
    讀取 input_csv_path，套用 DataTransformer 的決策樹邏輯，
    將 'parameter' 欄位就地覆寫，並儲存到 output_csv_path。
    
    成功時回傳 True, 失敗時回傳 False。
    """
    print(f"--- [正規化模組] 1. 正在讀取檔案: {input_csv_path} ---")
    try:
        df = pd.read_csv(input_csv_path)
    except FileNotFoundError:
        print(f"❌ [正規化模組] 錯誤: 找不到檔案 '{input_csv_path}'")
        return False
    except Exception as e:
        print(f"❌ [正規化模組] 讀取 CSV 時發生錯誤: {e}")
        return False
        
    if 'parameter' not in df.columns:
        print(f"❌ [正規化模組] 錯誤: CSV 檔案中找不到 'parameter' 欄位。")
        return False

    print("--- [正規化模組] 2. 正在複製 DataTransformer.py 決策樹邏輯 ---")
    
    # (A) 核心邏輯: 依 (parameter, Vin, step) 分組
    config_df = df.groupby(['parameter', 'Vin', 'step'], dropna=False, sort=False).size().reset_index(name='limits_count')

    # (B) 建立決策輔助欄
    param_group = config_df.groupby('parameter', dropna=False, sort=False)
    config_df['param_size'] = param_group.transform('size')
    config_df['vin_nunique'] = param_group['Vin'].transform('nunique')
    config_df['param_rank'] = param_group.cumcount()
    
    sub_group = config_df.groupby(['parameter', 'Vin'], dropna=False, sort=False)
    config_df['sub_size'] = sub_group.transform('size')
    config_df['sub_rank'] = sub_group.cumcount()

    print("--- [正規化模組] 3. 正在產生新的 'parameter' 名稱 ---")
    
    # (C) 迭代並產生「未清理的」工作表名稱 (raw_sheet_name)
    output_list = []
    for index, row in config_df.iterrows():
        param_size = row['param_size']
        vin_nunique = row['vin_nunique']
        sub_size = row['sub_size']
        parameter = row['parameter'] # 這是 '...TEST#1'
        vin = row['Vin']
        
        vin_string = ""
        # 檢查 vin 是否為 NaN 或空字串
        if not pd.isna(vin) and str(vin).strip():
            # (使用 float() 再 int() 來處理可能的 ".0")
            vin_string = f" {int(float(vin))}V" 
            
        # 執行三層決策樹
        if param_size == 1:
            output_list.append(f"{parameter}")
        else:
            if vin_nunique == 1:
                rank = int(row['param_rank'] + 1)
                output_list.append(f"{parameter}_{rank}")
            else:
                if sub_size == 1:
                    output_list.append(f"{parameter}{vin_string}")
                else:
                    rank = int(row['sub_rank'] + 1)
                    output_list.append(f"{parameter}{vin_string}_{rank}")

    # (D) 清理名稱並將它們映射回 config_df
    config_df['raw_sheet_name'] = output_list
    config_df['new_parameter_name'] = config_df['raw_sheet_name'].apply(sanitize_sheet_name)

    print("--- [正規化模組] 4. 正在將新名稱合併並覆寫回 'parameter' 欄位 ---")
    
    # (E) 將「最終名稱」映射回原始的 df
    mapping_df = config_df[['parameter', 'Vin', 'step', 'new_parameter_name']]
    df = df.merge(mapping_df, on=['parameter', 'Vin', 'step'], how='left')

    # (F) 關鍵的覆寫步驟
    if df['new_parameter_name'].isnull().any():
        print("⚠️ [正規化模組] 警告: 合併時發生錯誤，某些 'parameter' 欄位無法被轉換。")
        # 即使有錯誤，我們還是繼續，但 'parameter' 欄位不會被覆寫
        df['new_parameter_name'] = df['new_parameter_name'].fillna(df['parameter']) # 用舊名稱填充
    
    df['parameter'] = df['new_parameter_name']
    print("✅ [正規化模組] 'parameter' 欄位已成功覆寫。")

    # (G) 清理輔助欄位
    df = df.drop(columns=['new_parameter_name'])
    
    # --- 步驟 3: 儲存到「輸出」檔案 ---
    try:
        df.to_csv(output_csv_path, index=False, encoding='utf-8')
        print(f"--- [正規化模組] 5. 成功！檔案已儲存回: {output_csv_path} ---")
        return True # 成功
    except PermissionError:
        print(f"❌ [正規化模組] 錯誤: 權限不足，無法儲存檔案 '{output_csv_path}'。")
        return False
    except Exception as e:
        print(f"❌ [正規化模組] 儲存 CSV 時發生錯誤: {e}")
        return False