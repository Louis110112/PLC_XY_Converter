import pandas as pd # 導入 pandas 庫，用於處理 Excel 數據
import re # 導入 re 庫，用於正則表達式操作
import os # 導入 os 庫，用於操作文件系統，如切換目錄
import sys # 導入 sys 庫，用於訪問系統相關參數和函數
import argparse # 導入 argparse 庫，用於解析命令行參數
from datetime import datetime # 導入 datetime 模塊，用於獲取當前時間戳
from typing import Optional, Tuple # 導入 Optional 和 Tuple 類型提示


# === 0. 規則與預編譯正則 ===
# 定義用於匹配區域前綴的正則表達式 (例如 A16, B38 中的 A16)
AREA_RE = re.compile(r'^([A-Za-z]+\d+)')
# 定義用於匹配 CO 類型代碼的正則表達式 (例如 CO-01)
CO_RE = re.compile(r'^CO-\d+$')

# 設備中文名稱到英文代碼的映射字典
DEVICE_MAP = {
    "進風機": "IN_M", # 進風機對應 IN_M (Input Motor)
    "進氣機": "IN_M", # 進氣機作為進風機的同義詞，也對應 IN_M
    "排風機": "OUT_M", # 排風機對應 OUT_M (Output Motor)
    "排氣機": "OUT_M", # 排氣機作為排風機的同義詞，也對應 OUT_M
    "噴流風機": "M", # 噴流風機對應 M (Motor)
    "電動風門": "D", # 電動風門對應 D (Door)
}

# 訊號中文描述到英文後綴的映射字典
SUFFIX_MAP = {
    "運轉訊號": "STAT", # 運轉訊號對應 STAT (Status)
    "故障訊號": "ALM", # 故障訊號對應 ALM (Alarm)
}


# 輔助函數：判斷值是否為空 (None, NaN 或 pandas 的 NaT)
def is_na(value) -> bool:
    return value is None or (isinstance(value, float) and pd.isna(value)) or pd.isna(value)


# 提取文本中的區域前綴 (例如 "A16 進風機" 中提取 "A16")
def extract_area_prefix(text: str) -> str:
    match = AREA_RE.match(text) # 使用預編譯的正則表達式進行匹配
    return match.group(1) if match else "" # 如果匹配成功則返回第一組內容，否則返回空字符串


# 提取文本中的設備類型和編號 (例如 "進風機1" 中提取 "IN_M" 和 "1")
def extract_device_and_no(text: str) -> Tuple[Optional[str], str]:
    # 遍歷 DEVICE_MAP 中的中文設備名稱
    for zh, en in DEVICE_MAP.items():
        idx = text.find(zh) # 查找中文設備名稱在文本中的位置
        if idx != -1: # 如果找到
            after = text[idx + len(zh):] # 截取中文名稱之後的字符串
            num_match = re.search(r'(\d+)', after) # 在之後的字符串中查找數字 (設備編號)
            device_no = num_match.group(1) if num_match else "" # 如果找到數字則提取，否則為空
            return en, device_no # 返回英文設備代碼和設備編號
    return None, "" # 如果沒有匹配到任何設備，則返回 None 和空字符串


# 轉換 X 區段的名稱格式
def convert_x_name(name) -> str:
    if is_na(name): # 如果名稱為空或 NaN，則返回空字符串
        return ""
    s = str(name).strip() # 將名稱轉換為字符串並去除首尾空白

    area = extract_area_prefix(s) # 提取區域前綴
    dev_code, dev_no = extract_device_and_no(s) # 提取設備代碼和編號
    if not dev_code: # 如果沒有提取到設備代碼
        return re.sub(r'[^\w\-]', '', s) # 清理字符串，只保留字母、數字、下劃線和連字符

    suffix = "" # 初始化後綴為空
    for zh, suf in SUFFIX_MAP.items(): # 遍歷 SUFFIX_MAP 查找匹配的訊號後綴
        if zh in s: # 如果中文訊號描述在名稱中
            suffix = suf # 設定後綴
            break # 找到後即停止查找

    base = f"{dev_code}{dev_no}" # 構建基礎名稱 (設備代碼+編號)
    if area: # 如果有區域前綴，則添加到基礎名稱前
        base = f"{area}_{base}"
    if suffix: # 如果有後綴，則添加到基礎名稱後
        base = f"{base}_{suffix}"
    return base # 返回轉換後的名稱


# 轉換 Y 區段的名稱格式
def convert_y_name(name) -> str:
    if is_na(name): # 如果名稱為空或 NaN，則返回空字符串
        return ""
    s = str(name).strip() # 將名稱轉換為字符串並去除首尾空白

    if CO_RE.fullmatch(s): # 如果名稱完全匹配 CO-XX 格式，則直接返回
        return s

    area = extract_area_prefix(s) # 提取區域前綴
    dev_code, dev_no = extract_device_and_no(s) # 提取設備代碼和編號
    if dev_code is None: # 如果沒有提取到設備代碼，則直接返回原始字符串
        return s

    # 根據設備代碼構建描述字符串
    if dev_code in ("IN_M", "OUT_M", "M"):
        # 根據英文代碼獲取對應的英文標籤
        label = {"IN_M": "In Motor", "OUT_M": "Out Motor", "M": "Motor"}[dev_code]
        prefix = f"{area} " if area else "" # 如果有區域前綴，則加上空格
        return f"{prefix}{label}{dev_no}" # 返回格式化後的描述

    if dev_code == "D": # 如果是電動風門
        desc = f"DOOR{dev_no}" # 構建 DOORXX 格式
        return f"{area}_{desc}" if area else desc # 如果有區域前綴則添加，否則直接返回

    return s # 如果沒有匹配到上述規則，則返回原始字符串


# 解析命令行參數
def parse_args():
    parser = argparse.ArgumentParser(description="Convert A.xlsx (X/Y sections) to T.xlsx with REF/COMMENT/DESCRIPTION.")
    # 添加輸入文件參數
    parser.add_argument("-i", "--input", default="A.xlsx", help="Input Excel filename (default: A.xlsx)")
    # 添加輸出文件參數
    parser.add_argument("-o", "--output", default="T.xlsx", help="Output Excel filename (default: T.xlsx)")
    # 添加可選的工作表名稱或索引參數
    parser.add_argument("--sheet", default=None, help="Excel sheet name or index (default: first sheet)")
    return parser.parse_args() # 返回解析後的參數


# 主函數
def main():
    # 切換到腳本所在目錄，確保相對路徑正確，這對於讀取和寫入文件很重要
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)

    args = parse_args() # 解析命令行參數

    # === 1. 讀取輸入 Excel 文件 ===
    sheet_arg = args.sheet
    # 如果指定了工作表參數，並且是數字字符串，則轉換為整數索引
    if sheet_arg is not None and isinstance(sheet_arg, str) and sheet_arg.isdigit():
        sheet_arg = int(sheet_arg)
    # 讀取 Excel 文件，header=None 表示沒有標題行，dtype=str 確保所有數據都讀取為字符串
    # sheet_name 根據 sheet_arg 的值選擇，如果為 None 則讀取第一個工作表 (索引 0)
    df_a = pd.read_excel(args.input, header=None, dtype=str, sheet_name=(sheet_arg if sheet_arg is not None else 0))

    # 輔助函數：從 DataFrame 中獲取特定區段的數據
    def get_section(df, start_row: int, col_code: int, col_name: int, code_regex: str) -> pd.DataFrame:
        cols = df.columns.tolist() # 獲取所有列的索引列表
        # 檢查指定的代碼列和名稱列是否存在於 DataFrame 中
        if col_code not in cols or col_name not in cols:
            # 如果任一列不存在，則返回一個空的 DataFrame
            return pd.DataFrame(columns=["Code", "Name_zh"])
        # 從指定起始行開始，複製代碼列和名稱列的數據
        sec = df.loc[start_row:, [col_code, col_name]].copy()
        sec.columns = ["Code", "Name_zh"] # 重新命名列為 "Code" 和 "Name_zh"
        # 先過濾掉 Name_zh 為空 (None, NaN) 的行，確保只處理有數據的行
        sec = sec[~sec["Name_zh"].apply(is_na)]
        sec["Code"] = sec["Code"].astype(str).str.strip() # 將 Code 列轉換為字符串並去除空白
        sec["Name_zh"] = sec["Name_zh"].astype(str).str.strip() # 將 Name_zh 列轉換為字符串並去除空白
        # 進一步根據 code_regex 過濾 Code 列，na=False 確保 NaN 值不匹配
        sec = sec[sec["Code"].str.match(code_regex, na=False)]
        return sec # 返回處理後的區段數據

    # 提取 X 區段數據 (從 Excel 的 A16, B16 開始，對應索引 15, 0, 1)
    df_x = get_section(df_a, 15, 0, 1, r'^[Xx]\d+$')

    # 提取 Y 區段數據 (從 Excel 的 F16, G16 開始，對應索引 15, 5, 6)
    df_y = get_section(df_a, 15, 5, 6, r'^[Yy]\d+$')
    # 如果 Y 區段為空，且原始 DataFrame 不包含列 5 或 6，則打印警告信息
    if df_y.empty and (5 not in df_a.columns or 6 not in df_a.columns):
        print("[INFO] 未找到欄位 F/G（Y 區段），將只輸出 X 區段。")

    # === 2. 向量化處理並產出 T 文件所需數據 ===
    x_desc = df_x["Name_zh"].apply(convert_x_name) # 將 X 區段的中文名稱轉換為英文描述
    y_desc = df_y["Name_zh"].apply(convert_y_name) # 將 Y 區段的中文名稱轉換為英文描述

    # 創建 X 區段的輸出 DataFrame (REF, COMMENT, DESCRIPTION)
    df_tx = pd.DataFrame({
        "REF": df_x["Code"], # REF 列為 X 區段的 Code
        "COMMENT": x_desc, # COMMENT 列為 X 區段的英文描述
        "DESCRIPTION": "", # DESCRIPTION 列暫時為空
    })

    # 創建 Y 區段的輸出 DataFrame
    df_ty = pd.DataFrame({
        "REF": df_y["Code"], # REF 列為 Y 區段的 Code
        "COMMENT": y_desc, # COMMENT 列為 Y 區段的英文描述
        "DESCRIPTION": "", # DESCRIPTION 列暫時為空
    })

    # 將 X 區段和 Y 區段的結果合併，忽略原始索引
    df_t = pd.concat([df_tx, df_ty], ignore_index=True)
    output_path = args.output # 獲取輸出文件路徑
    try:
        df_t.to_excel(output_path, index=False) # 嘗試將結果輸出到 Excel 文件，不包含索引列
    except PermissionError: # 如果遇到權限錯誤 (例如文件正在被使用)
        base, ext = os.path.splitext(output_path) # 分離文件名和擴展名
        ts = datetime.now().strftime('%Y%m%d_%H%M%S') # 生成當前時間戳
        alt = f"{base}_{ts}{ext}" # 創建一個帶有時間戳的新文件名
        print(f"[WARN] {output_path} is in use. Writing to {alt} instead.") # 打印警告信息
        df_t.to_excel(alt, index=False) # 將結果輸出到新文件
        output_path = alt # 更新輸出路徑為新文件路徑
    print(f"已輸出 {output_path}（X 與 Y 區段，COMMENT 為英文結果，DESCRIPTION 僅表頭）") # 打印輸出成功信息


# 當腳本作為主程序執行時調用 main 函數
if __name__ == "__main__":
    main()


