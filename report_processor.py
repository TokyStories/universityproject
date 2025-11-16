import pandas as pd
import re
import difflib
from typing import Union, Tuple


def process_report_to_csv(input_excel_path: str, input_txt_path: str, output_csv_path: str):
    num1=num2=0
    """填入檔案位置"""
    """測試用檔案"""

    #________HPA65X-240
    #file_path = r"D:\someDATA\function_test_spec HPA65X-240-7.xlsx"
    #file_xtx  = r"D:\someDATA\ATE Report\90A0040412\2534010793.txt"
    #________HPA65X-120
    #file_path = r"D:\someDATA\function_test_spec HPA65X-120-3.xlsx"
    #file_xtx  = r"D:\someDATA\ATE Report\HPA65A.TXT\2529000705.txt"
    #________HPA90X-120
    #file_path = r"D:\someDATA\function_test_spec HPA90X-120-4.xlsx"
    #file_xtx  = r"D:\someDATA\ATE Report\96-42998-02-A0\2519010807.txt"
    #________HBU250-107
    #file_path = r"D:\someDATA\function_test_spec HBU250-107-4.xlsx"
    #file_xtx  = r"D:\someDATA\96-40318-02-A0-All.txt"
    file_path = input_excel_path
    file_xtx = input_txt_path

    exceldf = pd.read_excel(file_path)
    STEPdata = []
    TEST_PARAMETER=[]
    pdf_to_step =[]

    #函式名稱改成convert_load_value_int
    def convert_load_value_int(value_str: str) -> Union[int, Tuple[int, int], str]:
        s = value_str.strip().lower() 
        if s in ('max', 'min', 'forced', 'peak', 'no load', 'free'):
            return s 
        if '%' in s:
            content = s.replace('%', '')
            if '~' in content:
                try:
                    min_str, max_str = content.split('~')
                    return (int(min_str), int(max_str))
                except ValueError:
                    return s
            else:
                try:
                    return int(content)
                except ValueError:
                    return s
        elif '~' in s:
            try:
                min_str, max_str = s.split('~')
                return (int(min_str), int(max_str))
            except ValueError:
                return s
        return s
    #_____________________________________________________________找出step內容_____________________________________________________________
    find_line = []
    ATS=0

    with open(file_xtx, "r", encoding="utf-8") as archive:
        text_line = [txt.strip() for txt in archive.readlines()]
        keyword = "STEP"
        keyword2 = "Adapter/Charger ATS"
        
    for line in text_line:
        if keyword in line:
            find_line.append(line)
        if keyword2 in line:
            ATS=ATS+1
            if ATS>=2:            
                break
    #_____________________________________________________________step內容建立陣列_____________________________________________________________
    for line in find_line:
        io1 = ""
        io2 = ""
        test_name2 = ""
        voltage = ""
        step_match = re.search(r"(STEP\.\d+)", line)
        step = step_match.group(1) if step_match else ""
        step = step.split(".", 1)[-1].strip()
        after_colon = line.split(":", 1)[-1].strip()
        test_name_1 = after_colon.split("(", 1)[0].strip()
        all_parens = re.findall(r"\(([^)]*)\)", after_colon)

        for part in all_parens:
            part = part.strip()
            io1_match = re.search(r"io1\s*=\s*([^,()]+)", part)
            io2_match = re.search(r"io2\s*=\s*([^,()]+)", part)
            io_match = re.search(r"io\s*=\s*([^,()]+)", part)

            if io1_match:
                num1=num1+1
                content = io1_match.group(1).strip()
                key_match = re.search(r"(Forced|min|Peak|no load|\d+%|Free)", content, re.IGNORECASE)
                if key_match:
                    io1 = key_match.group(1).strip()

            if io2_match:
                num1=num1+1
                content = io2_match.group(1).strip()
                if "max" in content.lower():
                    io2 = "max"

            if "io1,2=" in part:
                io1 = "no load"
                io2 = "no load"
                
            if num1==0:
                if io_match:
                    content = io_match.group(1).strip()
                    key_match = re.search(
                        r"(\d+~\d+%)|(\d+%)|(Forced|min|max|Peak|no load|Free)", 
                        content, 
                        re.IGNORECASE
                    )
            
                    if key_match:
                        io1 = key_match.group(0).strip()
                    else:
                        io1 = content
                
            test_name2_match = re.search(r"(start|RIPPLE|ON/OFF)", part, re.IGNORECASE)
            if test_name2_match:
                test_name2 = test_name2_match.group(1).strip()
        
            voltage_match = re.search(r"(\d+\.?\d*\s*[AV])", part)
            if voltage_match and not voltage:
                voltage = voltage_match.group(1).strip()

        if not io1 and "Efficiency & Power Factor Test" in test_name_1:
            io1_full_match = re.search(r"io1\s*=\s*([^()\s]+)", after_colon)
            if io1_full_match:
                content = io1_full_match.group(1).strip()
                key_match = re.search(r"(Forced|min|Peak|no load|\d+%|Free)", content, re.IGNORECASE)
                if key_match:
                    io1 = key_match.group(1).strip()
                    
        if not voltage:
            match = re.search(r"(\d+\.?\d*\s*[AV])", after_colon)
            if match:
                voltage = match.group(1).strip()

        voltage = voltage.replace('V', '').replace('A', '')
        STEPdata.append([step, test_name_1, io1, io2, test_name2, voltage])

    STEPdf = pd.DataFrame(STEPdata, columns=["step","test name 1","io1","io2","test name 2","voltage"])
    print(STEPdf)
    print("=_"*50)
    #_____________________________________________________________確認PDF檢查項目_____________________________________________________________
    for i in range(len(exceldf)):
        item           = exceldf.at[i, 'item']
        Test_Parameter = exceldf.at[i, 'TEST PARAMETER']
        limits_i       = exceldf.at[i, 'LIMITS I']
        limits_ii      = exceldf.at[i, 'LIMITS II']
        Limits_iii     = exceldf.at[i, 'LIMITS III']
        limits_iv      = exceldf.at[i, 'LIMITS IV']
        limits_v       = exceldf.at[i, 'LIMITS V']
        limits_vi      = exceldf.at[i, 'LIMITS VI']
        limits_vii     = exceldf.at[i, 'LIMITS VII']
        limits_viii    = exceldf.at[i, 'LIMITS VIII']
        Pin            = exceldf.at[i, 'PIN']
        Condition_I    = exceldf.at[i, 'TEST CONDITION I']
        Load           = exceldf.at[i, 'TEST CONDITION III']
        Load           = Load.split(".", 1)[0].strip()
        Load_s         = Load.split("(", 1)[-1].strip()
        Load_s         = Load_s.split("s)", 1)[0].strip()
        Load           = Load.split(" ", 1)[0].strip()

    #_____________________________________________________________特定測試_____________________________________________________________    
        if difflib.SequenceMatcher(None,exceldf.at[i, 'TEST PARAMETER'],'SHORT CKT TEST').ratio() >= 0.7:
            TEST_PARAMETER.append([item,Test_Parameter,'Short Circuit Protection Test','','','',limits_i,limits_ii,Limits_iii,limits_iv,limits_v,limits_vi,limits_vii,limits_viii,Pin])
            
        elif exceldf.at[i, 'TEST PARAMETER'] == 'HOLD-UP TIME TEST':
            TEST_PARAMETER.append([item,Test_Parameter,'Hold Up & Sequence Test',Load,'',Condition_I,limits_i,limits_ii,Limits_iii,limits_iv,limits_v,limits_vi,limits_vii,limits_viii,Pin])

        elif difflib.SequenceMatcher(None,exceldf.at[i, 'TEST PARAMETER'],'START UP TIME').ratio() >= 0.7:
            if pd.notna(limits_iv):
                TEST_PARAMETER.append([item,Test_Parameter,'Efficiency & Power Factor Test',Load,'',Condition_I,limits_i,limits_ii,Limits_iii,limits_iv,limits_v,limits_vi,limits_vii,limits_viii,Pin])
            if pd.notna(limits_ii):
                TEST_PARAMETER.append([item,Test_Parameter,'Static Load Test',Load,'',Condition_I,limits_i,limits_ii,Limits_iii,limits_iv,limits_v,limits_vi,limits_vii,limits_viii,Pin])
            TEST_PARAMETER.append([item,Test_Parameter,'Turn On & Sequence Test',Load,'',Condition_I,limits_i,limits_ii,Limits_iii,limits_iv,limits_v,limits_vi,limits_vii,limits_viii,Pin])

        elif difflib.SequenceMatcher(None,exceldf.at[i, 'TEST PARAMETER'],'POWER CONSUMPTION TEST').ratio() >= 0.7:
            if  pd.notna(Limits_iii):
                TEST_PARAMETER.append([item,Test_Parameter,'Static Load Test',Load,'',Condition_I,limits_i,limits_ii,Limits_iii,limits_iv,limits_v,limits_vi,limits_vii,limits_viii,Pin])
            elif pd.notna(Pin):
                TEST_PARAMETER.append([item,Test_Parameter,'Input Power Integration Test',Load,'',Condition_I,limits_i,limits_ii,Limits_iii,limits_iv,limits_v,limits_vi,limits_vii,limits_viii,Pin])
                if pd.notna(limits_ii):
                    TEST_PARAMETER.append([item,Test_Parameter,'Static Load Test',Load,'',Condition_I,limits_i,limits_ii,Limits_iii,limits_iv,limits_v,limits_vi,limits_vii,limits_viii,Pin])

        elif difflib.SequenceMatcher(None,exceldf.at[i, 'TEST PARAMETER'],'O.T.P. TEST').ratio() >= 0.70:  
            if difflib.SequenceMatcher(None,exceldf.at[i, 'TEST PARAMETER'],'O.V.P. TEST').ratio() >= 0.95:  
                TEST_PARAMETER.append([item,Test_Parameter,'此項不測','','','',''])
            else: TEST_PARAMETER.append([item,Test_Parameter,'此項不測','','','',''])

        elif difflib.SequenceMatcher(None,exceldf.at[i, 'TEST PARAMETER'],'Dynamic Lood Test').ratio() >= 0.7:
            TEST_PARAMETER.append([item,Test_Parameter,'Peak Load Dynamic Test',Load,Load_s,Condition_I,limits_i,limits_ii,Limits_iii,limits_iv,limits_v,limits_vi,limits_vii,limits_viii,Pin])

        #____________________________________________________HPA65-240新增____________
        elif difflib.SequenceMatcher(None,exceldf.at[i, 'TEST PARAMETER'],'AVERAGE EFFICIENCY').ratio() >= 0.85:  
            TEST_PARAMETER.append([item,Test_Parameter,'Average Efficiency Test',Load,Load_s,Condition_I,limits_i,limits_ii,Limits_iii,limits_iv,limits_v,limits_vi,limits_vii,limits_viii,Pin])

        elif difflib.SequenceMatcher(None,exceldf.at[i, 'TEST PARAMETER'],'OVER CURRENT TEST').ratio() >= 0.7: 
            TEST_PARAMETER.append([item,Test_Parameter,'Static Load Test',Load,Load_s,Condition_I,limits_i,limits_ii,Limits_iii,limits_iv,limits_v,limits_vi,limits_vii,limits_viii,Pin])
            TEST_PARAMETER.append([item,Test_Parameter,'Over Load Protection Test',Load,Load_s,Condition_I,limits_i,limits_ii,Limits_iii,limits_iv,limits_v,limits_vi,limits_vii,limits_viii,Pin])
            
        elif difflib.SequenceMatcher(None,exceldf.at[i, 'TEST PARAMETER'],'MIN. LOAD ON/OFF').ratio() >= 0.7: 
            TEST_PARAMETER.append([item,Test_Parameter,'Static Load Test',Load,Load_s,Condition_I,limits_i,limits_ii,Limits_iii,limits_iv,limits_v,limits_vi,limits_vii,limits_viii,Pin])
            TEST_PARAMETER.append([item,Test_Parameter,'Turn On & Sequence Test',Load,Load_s,Condition_I,limits_i,limits_ii,Limits_iii,limits_iv,limits_v,limits_vi,limits_vii,limits_viii,Pin])
        elif difflib.SequenceMatcher(None,exceldf.at[i, 'TEST PARAMETER'],'AUX VOLTAGE TEST').ratio() >= 0.8:  
            TEST_PARAMETER.append([item,Test_Parameter,'此項不測','','','',''])
        #____________________________________________________HPA90-120新增____________
        elif difflib.SequenceMatcher(None,exceldf.at[i, 'TEST PARAMETER'],'PF TEST').ratio() >= 0.7:  
            TEST_PARAMETER.append([item,Test_Parameter,'Efficiency & Power Factor Test',Load,Load_s,Condition_I,limits_i,limits_ii,Limits_iii,limits_iv,limits_v,limits_vi,limits_vii,limits_viii,Pin])
        #____________________________________________________其他狀態__________________
        else:
            if pd.notna(limits_iv):
                TEST_PARAMETER.append([item,Test_Parameter,'Efficiency & Power Factor Test',Load,'',Condition_I,limits_i,limits_ii,Limits_iii,limits_iv,limits_v,limits_vi,limits_vii,limits_viii,Pin])
            elif pd.notna(limits_vi):
                TEST_PARAMETER.append([item,Test_Parameter,'Efficiency & Power Factor Test',Load,'',Condition_I,limits_i,limits_ii,Limits_iii,limits_iv,limits_v,limits_vi,limits_vii,limits_viii,Pin])
            else:
                TEST_PARAMETER.append([item,Test_Parameter,'Static Load Test',Load,'',Condition_I,limits_i,limits_ii,Limits_iii,limits_iv,limits_v,limits_vi,limits_vii,limits_viii,Pin])

    TEST_PARAMETER_df = pd.DataFrame(TEST_PARAMETER, columns=["item","TEST PARAMETER","test name","load","load_s","voltage","limits_i", "Vpp","Vo2","Iin","time","eff","PF","other","pin"])
    print(TEST_PARAMETER_df)
    print("==="*50)
    #_____________________________________________________________開始比對_____________________________________________________________
    #將內容改為str格式
    TEST_PARAMETER_df['test name'] = TEST_PARAMETER_df['test name'].astype(str)
    STEPdf['test name 1'] = STEPdf['test name 1'].astype(str)
    TEST_PARAMETER_df['voltage'] = TEST_PARAMETER_df['voltage'].astype(str)
    STEPdf['voltage'] = STEPdf['voltage'].astype(str)

    for i in range(len(TEST_PARAMETER_df)):
        for x in range(len(STEPdf)):
            num2=0
            OPTtext    = str(TEST_PARAMETER_df.at[i, "TEST PARAMETER"])
            Testname   = str(TEST_PARAMETER_df.at[i, "test name"])
            itemnum    = str(TEST_PARAMETER_df.at[i, "item"])
            vin        = TEST_PARAMETER_df.at[i, "voltage"]
            limits_I   = TEST_PARAMETER_df.at[i, 'limits_i']
            limits_II  = TEST_PARAMETER_df.at[i, 'Vpp']
            limits_III = TEST_PARAMETER_df.at[i, 'Vo2']
            limits_IV  = TEST_PARAMETER_df.at[i, 'Iin']
            limits_V   = TEST_PARAMETER_df.at[i, 'time']
            limits_VI  = TEST_PARAMETER_df.at[i, 'eff']
            limits_VII = TEST_PARAMETER_df.at[i, 'PF']
            limits_VIII= TEST_PARAMETER_df.at[i, 'other']
            pin        = TEST_PARAMETER_df.at[i, 'pin']
            load_s     = TEST_PARAMETER_df.at[i, 'load_s']
            Steptext   = str(STEPdf.at[x,"step"])
            test_wit   = str(STEPdf.at[x, "test name 2"]).strip().lower()
            io2O       = str(STEPdf.at[x, "io2"]).strip().lower()

            df_load_raw = str(TEST_PARAMETER_df.at[i, "load"])
            df_load = convert_load_value_int(df_load_raw) 
            step_load_raw = str(STEPdf.at[x, "io1"])
            step_load = convert_load_value_int(step_load_raw) 
        
            #將不同寫法意義相近的附載程度統整為統一格式
            if df_load  in ['max','peak','forced']:
                df_load = 'max'
            if step_load  in ['max','peak','forced']:
                step_load = 'max'
            if df_load  in ['no','no load']:
                df_load = 'no'
            if step_load  in ['no','no load']:
                step_load = 'no'    

            if isinstance(df_load, tuple) and isinstance(step_load, tuple):
                df_min, df_max = df_load
                step_min, step_max = step_load
                min_condition = (df_min >= step_min)
                max_condition = (df_min >= step_min)
        
                if min_condition and max_condition:
                    df_load = str ("match")
                    step_load = str ("match")
            elif isinstance(df_load, tuple) and isinstance(step_load, int):
                df_min, df_max = df_load
                if df_min <= step_load <= df_max:
                    df_load = str ("match")
                    step_load = str ("match")

            is_limits_III_valid = (
                pd.notna(limits_III) and 
                str(limits_III).strip() != ''
            )
            is_io2O_valid = (
                pd.notna(io2O) and 
                str(io2O).strip() != ''
            )

    #找對應的TEST PARAMETER
            if TEST_PARAMETER_df.at[i, "test name"] == STEPdf.at[x,"test name 1"]:  
    #定義特殊的測試
                if TEST_PARAMETER_df.at[i, "test name"] =='Turn On & Sequence Test':  
                    if df_load == step_load:  
                        if difflib.SequenceMatcher(None,exceldf.at[i, 'TEST PARAMETER'],'MIN. LOAD ON/OFF').ratio() >= 0.9: 
                            num2=1
                            if test_wit == 'NO\OFF':
                                pdf_to_step.append([itemnum,OPTtext,Testname,vin,Steptext,limits_I,limits_II,limits_III,limits_IV,limits_V,limits_VI,limits_VII,limits_VIII,pin])
                        else:
                            pdf_to_step.append([itemnum,OPTtext,Testname,vin,Steptext,limits_I,limits_II,limits_III,limits_IV,limits_V,limits_VI,limits_VII,limits_VIII,pin])
                if TEST_PARAMETER_df.at[i, "test name"] in ['Short Circuit Protection Test', 'Dynamic Load Setup', 'Hold Up & Sequence Test', 'Peak Load Dynamic Test']:
                    pdf_to_step.append([itemnum, OPTtext, Testname, vin, Steptext, limits_I, limits_II, limits_III, limits_IV, limits_V, limits_VI, limits_VII, limits_VIII, pin])

    #另外要求Vin相同的  
                if TEST_PARAMETER_df.at[i, "test name"] in ['Input Power Integration Test','Average Efficiency Test','Over Load Protection Test','Efficiency & Power Factor Test']:
                    if TEST_PARAMETER_df.at[i, "voltage"] == STEPdf.at[x,"voltage"]:
                        pdf_to_step.append([itemnum,OPTtext,Testname,vin,Steptext,limits_I,limits_II,limits_III,limits_IV,limits_V,limits_VI,limits_VII,limits_VIII,pin])
        
    #各種不同的Static Load Test
                if TEST_PARAMETER_df.at[i, "test name"] =='Static Load Test':           #針對Static Load Test的項目
                    if TEST_PARAMETER_df.at[i, "voltage"] == STEPdf.at[x,"voltage"]:    #限定相同電壓的項目
                        if is_limits_III_valid:                       
                            if is_io2O_valid:
                                io2O_clean = str(io2O).strip().lower() 
                                num2=num2+1
                                pdf_to_step.append([itemnum,OPTtext,Testname,vin,Steptext,limits_I,limits_II,limits_III,limits_IV,limits_V,limits_VI,limits_VII,limits_VIII,pin])
                                continue
                                
                    # elif pd.notna(load_s) and str(load_s).strip() != '':
                    #     if step_load == 'free':##===============================================================針對(3s)暫時廢棄
                    #         pdf_to_step.append([itemnum,OPTtext,Testname,vin,Steptext,limits_I,limits_II,limits_III,limits_IV,limits_V,limits_VI,limits_VII,limits_VIII,pin])                      

                        elif df_load == step_load: #限同負載
                            if  pd.notna(limits_II):
                                if test_wit =='ripple':
                                    pdf_to_step.append([itemnum,OPTtext,Testname,vin,Steptext,limits_I,limits_II,limits_III,limits_IV,limits_V,limits_VI,limits_VII,limits_VIII,pin])
                                    continue
                            else :
                                if test_wit !='ripple':
                                    if OPTtext =='PEAK LOAD TEST':
                                        pdf_to_step.append([itemnum,OPTtext,Testname,vin,Steptext,limits_I,limits_II,limits_III,limits_IV,limits_V,limits_VI,limits_VII,limits_VIII,pin]) 
                                        continue
                        if num2==0:
                            if df_load == step_load:
                                if difflib.SequenceMatcher(None,TEST_PARAMETER_df.at[i, 'TEST PARAMETER'],'OVER CURRENT TEST').ratio() >= 0.7:
                                    pdf_to_step.append([itemnum,OPTtext,Testname,vin,Steptext,limits_I,limits_II,limits_III,limits_IV,limits_V,limits_VI,limits_VII,limits_VIII,pin])
                                if difflib.SequenceMatcher(None,TEST_PARAMETER_df.at[i, 'TEST PARAMETER'],'AUX VOLTAGE TEST').ratio() >= 0.7:
                                    pdf_to_step.append([itemnum,OPTtext,Testname,vin,Steptext,limits_I,limits_II,limits_III,limits_IV,limits_V,limits_VI,limits_VII,limits_VIII,pin])

    PDF_to_Step = pd.DataFrame(pdf_to_step, columns=["item","PARAMETER","test name","Vin","step","limits_1","limits_2","limits_3","limits_4","limits_5","limits_6","limits_7","limits_8","pin"])
    PDF_to_Step2 = pd.DataFrame(pdf_to_step, columns=["item","PARAMETER","test name","Vin","step","Vdc","Vpp","Vdc2","Iin","time","eff","PF","other","pin"])
    #PDF_to_Step2.to_csv('test_data65240.csv', index=False, encoding='utf-8') 
    #print(PDF_to_Step)
    #print(PDF_to_Step2)
    print(PDF_to_Step[["item", "PARAMETER","test name","Vin", "step"]])





    
    df_input  = PDF_to_Step2
    #file_xtx  = r"D:\someDATA\2231008122.txt"#250107
    #file_xtx  = r"D:\someDATA\ATE Report\HPA65A.TXT\2529000705.txt"#65120
    #file_xtx  = r"D:\someDATA\ATE Report\90A0040412\2534010793.txt"#65240
    #file_xtx  = r"D:\someDATA\ATE Report\96-42998-02-A0\2519010807.txt"#90120
    keyword2  = "Adapter/Charger ATS"
    keyword_START = ("STEP.", "(Pre Test seq.1)")
    keyword_END = "==================================================="
    ATS=0
    num0=0
    all_find_lines = []
    is_collecting = False
    current_block = []
    find_where =[]
    vdc_where2=[]
    Vdc1F = ""
    Vdc2F = ""
    #新增Vdc1F = ""、Vdc2F = ""，避免未定義錯誤
    #print(df_input) 

    with open(file_xtx, "r", encoding="utf-8") as archive:
        text_line = [txt.strip() for txt in archive.readlines()]
        
    def convert_load_value(value_str: str) -> Union[float, Tuple[float, float], str]:###將 +XX.XX to +YY.YY拆開成(XX.XX , YY.YY)的副程式
        s = value_str.strip().lower() 
        if s in ('nan', 'NaN'):
            return 'NaN' 
        if 'to' in s:
            content = s.replace('v', '').replace(' ', '').replace('+', '')
            try:
                min_str, max_str = content.split('to')
                return (float(min_str), float(max_str))
            except ValueError:
                return s
        return s

    """讀取txt資料"""    
    for line in text_line:
        num0=num0+1
        if keyword2 in line:
            ATS=ATS+1
            if ATS>=2: 
                break             #確保只會抓一組TXT
        if keyword_END in line:
            if is_collecting:
                all_find_lines.append(current_block)   #資料加入"all_find_lines"
                is_collecting = False 
        if line.strip().startswith(keyword_START):
            current_block = []
            is_collecting = True 
        if is_collecting:
            current_block.append(line.rstrip())

    for i in range(len(df_input)):
        PDFv_raw = str(df_input.at[i, "Vdc"])
        PDFv = convert_load_value(PDFv_raw)  
        if isinstance(PDFv, tuple) and len(PDFv) == 2:
            min_val, max_val = PDFv
    PDFvA = int(round((min_val + max_val) / 2))
    PDFvA = str(PDFvA) #平均 Vdc

    for i in range(len(df_input)):
        PDFv_raw  = str(df_input.at[i, "Vdc"])
        PDFvA      = convert_load_value(PDFv_raw)
        if isinstance(PDFvA, tuple) and len(PDFvA) == 2:
            min_val, max_val = PDFvA
            for line in all_find_lines[0]:
                line3=line.split()
                if len(line3) > 1:
                    target_value = line3[0]
                    if re.match(r"^\d+V$", target_value, re.IGNORECASE):
                        content = target_value.replace('V', '').replace('v', '') 
                        target_value = int(content)
                        if min_val<=target_value<=max_val:
                            Vdc1F=str(target_value)+'V'
    #                        print(target_value)

    for i in range(len(df_input)):
        Vdc2_raw  = str(df_input.at[i, "Vdc2"])
        Vdc2      = convert_load_value(Vdc2_raw)
        if isinstance(Vdc2, tuple) and len(Vdc2) == 2:
            min_val, max_val = Vdc2
            for line in all_find_lines[0]:
                line3=line.split()
                #print(line3)
                if len(line3) > 1:
                    target_value = line3[0]
                    if re.match(r"^\d+V$", target_value, re.IGNORECASE):
                        content = target_value.replace('V', '').replace('v', '') 
                        target_value = int(content)
                        if min_val<=target_value<=max_val:
                            Vdc2F=str(target_value)+'V'
                            print(target_value)

    for i in range(len(df_input)):
        Witem        = df_input.at[i, "item"]
        PDFv_raw  = str(df_input.at[i, "Vdc"])
        PDFv      = convert_load_value(PDFv_raw)#Vdc 
        PDFvA      = Vdc1F
        Vpp       = df_input.at[i, "Vpp"]
        Vin       = df_input.at[i, "Vin"]
    ##    Vdc2      = df_input.at[i, "Vdc2"]
        Vdc2_raw  = str(df_input.at[i, "Vdc2"])
        Vdc2      = convert_load_value(Vdc2_raw)
        Iin       = df_input.at[i, "Iin"]
        time      = df_input.at[i, "time"]
        eff       = df_input.at[i, "eff"]
        PF        = df_input.at[i, "PF"]
        Pin       = df_input.at[i, "pin"]
        other     = df_input.at[i, "other"]
        whatstep  = int(df_input.at[i, "step"])#對照step幾
        parameter = df_input.at[i, "PARAMETER"]
        test_name = df_input.at[i, "test name"]
        print(parameter)

        line2=''
        line3=[]
        print("item:",df_input.at[i,"item"],df_input.at[i,"PARAMETER"],"STEP.",whatstep)    #顯示測試項目名稱與step數

        num0=-1
        num1=0  
        if PDFv!='NaN':
            print("要電壓")
            num2=0
            num0=-1
            num1=0
            if test_name=='Static Load Test':
                for line in all_find_lines[whatstep]:
                    num0=num0+1
                    vdc_where = []
                    V_where = []

                    line3=line.split()
                    for index, item in enumerate(line3): 
                        if item == '(V)':
                            V_where.append(index)
                    for index, item in enumerate(line3):
                        if item == 'Vout':
                            vdc_where.append(index)
                    common_column = set(vdc_where2).intersection(V_where)
                    if common_column:
                        list_A = list(common_column)[0]+1
                        print(f"成功提取的整數值: {list_A}")
                    vdc_where2=vdc_where
                    
                    if 'Vout' in line:
                        num1=1
                    if num1!=0:
                        if PDFvA in line:
                            print("找到Vout了",num0)
                            find_where.append([Witem,parameter,Vin,whatstep,num0,list_A, "Vdc"])
                            break
            elif test_name=='Peak Load Dynamic Test':
                for line in all_find_lines[whatstep]:
                    num0=num0+1
                    vdc_where = []
                    V_where = []
                    
                    line3=line.split()
                    for index, item in enumerate(line3): 
                        if item == 'Read(V)':
                            V_where.append(index)
                    for index, item in enumerate(line3):
                        if item == 'Vpk+':
                            vdc_where.append(index)
                    common_column = set(vdc_where2).intersection(V_where)
                    if common_column:
                        list_A = list(common_column)[0]+1
                        print(f"成功提取的整數值: {list_A}")
                    vdc_where2=vdc_where

                    if 'Vpk+' in line:
                        num1=1
                    if num1!=0:
                        if PDFvA in line:
                            print("找到Vout了",num0)
                            find_where.append([Witem,parameter,Vin,whatstep,num0,list_A,"Vdc"])
                            break

            else :#test_name=='Efficiency & Power Factor Test':
                for line in all_find_lines[whatstep]:
                    num0=num0+1
                    vdc_where = []
                    V_where = []
                    
                    line3=line.split()
                    for index, item in enumerate(line3): 
                        if item == '(V)':
                            V_where.append(index)
                    for index, item in enumerate(line3):
                        if item == 'Vdc':
                            vdc_where.append(index)
                    common_column = set(vdc_where2).intersection(V_where)
                    if common_column:
                        list_A = list(common_column)[0]+1
                        print(f"成功提取的整數值: {list_A}")
                    vdc_where2=vdc_where

                    if 'Reading' in line:
                        num1=1
                    if num1!=0:
                        if 'Vdc' in line:
                            num2=1
                    if num2!=0:
                        if PDFvA in line:
                            print("找到Vout了",num0)
                            find_where.append([Witem,parameter,Vin,whatstep,num0,list_A,"Vdc"])
                            break

        num0=-1
        num1=0
        if pd.notna(Vpp):
            print("要Vpp")
            for line in all_find_lines[whatstep]:
                num0=num0+1
                vdc_where = []
                V_where = []

                line3=line.split()
                for index, item in enumerate(line3): 
                    if item == '(mV)':
                        V_where.append(index)
                for index, item in enumerate(line3):
                    if item == 'Vpp':
                        vdc_where.append(index)
                common_column = set(vdc_where2).intersection(V_where)
                if common_column:
                    list_A = list(common_column)[0]+1
                    print(f"成功提取的整數值: {list_A}")
                vdc_where2=vdc_where
                if 'Vpp' in line:
                    num1=1
                if num1!=0:
                    if PDFvA in line:
                        print("找到Vpp了",num0)
                        find_where.append([Witem,parameter,Vin,whatstep,num0,list_A,"Vpp"])
                        break
                        
        num0=-1
        num1=0
        if Vdc2!='NaN':
            print(Vdc2)
            print("要Vdc2")
            for line in all_find_lines[whatstep]:
                    num0=num0+1
                    vdc_where = []
                    V_where = []

                    line3=line.split()
                    for index, item in enumerate(line3): 
                        if item == '(V)':
                            V_where.append(index)
                    for index, item in enumerate(line3):
                        if item == 'Vout':
                            vdc_where.append(index)
                    common_column = set(vdc_where2).intersection(V_where)
                    if common_column:
                        list_A = list(common_column)[0]+1
                        print(f"成功提取的整數值: {list_A}")
                    vdc_where2=vdc_where
                
                    if 'Vout' in line:
                        num1=1
                    if num1!=0:
                        if Vdc2F in line:
                            print("找到Vdc2了",num0)
                            find_where.append([Witem,parameter,Vin,whatstep,num0,list_A,"Vdc2"])
                            break
                            
        num0=-1
        num1=0
        if pd.notna(Iin):
            print("要Iin")
            for line in all_find_lines[whatstep]:
                num0=num0+1
                line3=line.split()
                target_element = 'Reading'
                try:
                    index_position = line3.index(target_element)
                    list_A = index_position + 1
                except ValueError:
                    pass
                if 'Iinrms' in line:
                    print("找到Iin了",num0)
                    find_where.append([Witem,parameter,Vin,whatstep,num0,list_A,"Iin"])
                    break

        num0=-1
        num1=0
        if pd.notna(Pin):
            print("要Pin")
            for line in all_find_lines[whatstep]:
                num0=num0+1
                line3=line.split()
                target_element = 'Reading'
                try:
                    index_position = line3.index(target_element)
                    list_A = index_position + 1
                    if test_name=='Short Circuit Protection Test':
                        list_A=list_A+2
                except ValueError:
                    pass
                if 'Reading' in line:
                        num1=1
                if num1!=0:
                    if 'Pin' in line:
                        print("找到pin了",num0)
                        find_where.append([Witem,parameter,Vin,whatstep,num0,list_A,"Pin"])
                        break
        
        num0=-1
        num1=0                      
        if pd.notna(time):
            print("要time")
            if test_name=='Hold Up & Sequence Test':
                for line in all_find_lines[whatstep]:
                    num0=num0+1
                    vdc_where = []
                    V_where = []

                    line3=line.split()
                    for index, item in enumerate(line3): 
                        if item == '(ms)':
                            V_where.append(index)
                    for index, item in enumerate(line3):
                        if item == 'Tholdup':
                            vdc_where.append(index)
                    common_column = set(vdc_where2).intersection(V_where)
                    if common_column:
                        list_A = list(common_column)[0]+1
                        print(f"成功提取的整數值: {list_A}")
                    vdc_where2=vdc_where
                    if 'Tholdup' in line:
                        num1=1
                    if num1!=0:
                        if PDFvA in line:
                            print("找到time了",num0)
                            find_where.append([Witem,parameter,Vin,whatstep,num0,list_A,"time"])
                            break
            if test_name=='Turn On & Sequence Test':
                for line in all_find_lines[whatstep]:
                    num0=num0+1
                    vdc_where = []
                    V_where = []
                    
                    line3=line.split()
                    for index, item in enumerate(line3): 
                        if item == '(ms)':
                            V_where.append(index)
                    for index, item in enumerate(line3):
                        if item == 'Ton':
                            vdc_where.append(index)
                    common_column = set(vdc_where2).intersection(V_where)
                    if common_column:
                        list_A = list(common_column)[0]+1
                        print(f"成功提取的整數值: {list_A}")
                    vdc_where2=vdc_where
                    if 'Ton' in line:
                        num1=1
                    if num1!=0:
                        if PDFvA in line:
                            print("找到time了",num0)
                            find_where.append([Witem,parameter,Vin,whatstep,num0,list_A,"time"])
                            break
        num0=-1
        num1=0
        if pd.notna(eff):
            num0=-1
            print("要eff")
            for line in all_find_lines[whatstep]:
                num0=num0+1
                line3=line.split()
                target_element = 'Reading'
                try:
                    index_position = line3.index(target_element)
                    list_A = index_position + 1
                except ValueError:
                    pass
                if 'Reading' in line:
                    num1=1
                if num1!=0:
                    if 'Eff' in line:
                        print("找到Eff了",num0)
                        find_where.append([Witem,parameter,Vin,whatstep,num0,list_A,"eff"])
                        num1=0
                        break
                        
        num0=-1
        num1=0
        if pd.notna(PF):
            print("要PF")
            for line in all_find_lines[whatstep]:
                num0=num0+1
                line3=line.split()
                target_element = 'Reading'
                try:
                    index_position = line3.index(target_element)
                    list_A = index_position + 1
                except ValueError:
                    pass
                if 'Reading' in line:
                    num1=1
                if num1!=0:
                    if 'PF' in line:
                        print("找到PF了",num0)
                        find_where.append([Witem,parameter,Vin,whatstep,num0,list_A,"PF"])
                        break

        #_____________________啊這_____________________________  
        if parameter=='EFFICIENCY':
            for line in all_find_lines[whatstep]:
                num0=num0+1
                line3=line.split()
                target_element = 'Reading'
                try:
                    index_position = line3.index(target_element)
                    list_A = index_position + 1
                except ValueError:
                    pass

                if 'Pin' in line:
                    print("找到Pin了",num0)
                    find_where.append([Witem,parameter,Vin,whatstep,num0,list_A,"Pout"])
                    num0=-1
                    break

        if parameter=='Dynamic Lood test':
            for line in all_find_lines[whatstep]:
                num0=num0+1
                vdc_where = []
                V_where = []

                line3=line.split()
                for index, item in enumerate(line3): 
                    if item == 'Read(V)':
                        V_where.append(index)
                for index, item in enumerate(line3):
                    if item == 'Vpk-':
                        vdc_where.append(index)
                common_column = set(vdc_where2).intersection(V_where)
                if common_column:
                    list_A = list(common_column)[0]+1
                    print(f"成功提取的整數值: {list_A}")
                vdc_where2=vdc_where

                if 'Vpk+' in line:
                    num1=1
                if num1!=0:
                    if PDFvA in line:
                        print("找到Vpk-了",num0)
                        find_where.append([Witem,parameter,Vin,whatstep,num0,list_A,"Vpk-"])
                        break

        if test_name=='Average Efficiency Test':
            for line in all_find_lines[whatstep]:
                num0=num0+1
                vdc_where = []
                V_where = []
                line3=line.split()
                    
                if 'Average Efficiency (%)' in line:
                    num1=1
                    list_A=len(line3)
                    print(list_A)
                    
                    print("AvEff",num0)
                    find_where.append([Witem,parameter,Vin,whatstep,num0,list_A,"AvEff"])
                    break
    #___________________________________________________抓PASS 
        if test_name=='Short Circuit Protection Test':
            if pd.isna(Pin): 
                for line in all_find_lines[whatstep]:
                    num0=num0+1
                    line3=line.split()
                    if 'PASS' in line:
                        list_A=len(line3)
                        print(list_A)
                        find_where.append([Witem,parameter,Vin,whatstep,num0,list_A,"PASS"])
        #if test_name=='Turn On & Sequence Test':
        #    if pd.isna(Pin): 
        #        for line in all_find_lines[whatstep]:
        #            num0=num0+1
        #            line3=line.split()
        #            if 'PASS' in line:
        #                list_A=len(line3)
        #                print(list_A)
        #                find_where.append([Witem,parameter,Vin,whatstep,num0,list_A,"PASS"])
        if parameter=='OVER CURRENT TEST':
            if pd.isna(Pin): 
                for line in all_find_lines[whatstep]:
                    num0=num0+1
                    line3=line.split()
                    if 'PASS' in line:
                        list_A=len(line3)
                        print(list_A)
                        find_where.append([Witem,parameter,Vin,whatstep,num0,list_A,"PASS"])
        if parameter=='MIN. LOAD ON/OFF TEST':
            if test_name=='Turn On & Sequence Test':
        #        if pd.isna(Pin): 
                    for line in all_find_lines[whatstep]:
                        num0=num0+1
                        line3=line.split()
                        if 'PASS' in line:
                            list_A=len(line3)
                            print(list_A)
                            find_where.append([Witem,parameter,Vin,whatstep,num0,list_A,"PASS"])
                    
    #___________________________________________________ 

        #print("-" * 50)
        #print('\n'.join(all_find_lines[whatstep]))
        #print("===" * 50)
        #print(all_find_lines[whatstep])
        print("===" * 50)
    find_where_pd = pd.DataFrame(find_where, columns=["item","parameter","Vin","step","line","list","test"])
    print(find_where_pd)
    find_where_pd.to_csv(output_csv_path, index=False, encoding='utf-8')
    return True