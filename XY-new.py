import pandas as pd
import re
import os
import sys
import argparse
from datetime import datetime
from typing import Optional, Tuple


# === 0. 規則與預編譯正則 ===
AREA_RE = re.compile(r'^([A-Za-z]+\d+)')
CO_RE = re.compile(r'^CO-\d+$')

DEVICE_MAP = {
    "進風機": "IN_M",
    "進氣機": "IN_M",
    "排風機": "OUT_M",
    "排氣機": "OUT_M",
    "噴流風機": "M",
    "電動風門": "D",
}

SUFFIX_MAP = {
    "運轉訊號": "STAT",
    "故障訊號": "ALM",
}


def is_na(value) -> bool:
    return value is None or (isinstance(value, float) and pd.isna(value)) or pd.isna(value)


def extract_area_prefix(text: str) -> str:
    match = AREA_RE.match(text)
    return match.group(1) if match else ""


def extract_device_and_no(text: str) -> Tuple[Optional[str], str]:
    for zh, en in DEVICE_MAP.items():
        idx = text.find(zh)
        if idx != -1:
            after = text[idx + len(zh):]
            num_match = re.search(r'(\d+)', after)
            device_no = num_match.group(1) if num_match else ""
            return en, device_no
    return None, ""


def convert_x_name(name) -> str:
    if is_na(name):
        return ""
    s = str(name).strip()

    area = extract_area_prefix(s)
    dev_code, dev_no = extract_device_and_no(s)
    if not dev_code:
        return re.sub(r'[^\w\-]', '', s)

    suffix = ""
    for zh, suf in SUFFIX_MAP.items():
        if zh in s:
            suffix = suf
            break

    base = f"{dev_code}{dev_no}"
    if area:
        base = f"{area}_{base}"
    if suffix:
        base = f"{base}_{suffix}"
    return base


def convert_y_name(name) -> str:
    if is_na(name):
        return ""
    s = str(name).strip()

    if CO_RE.fullmatch(s):
        return s

    area = extract_area_prefix(s)
    dev_code, dev_no = extract_device_and_no(s)
    if dev_code is None:
        return s

    if dev_code in ("IN_M", "OUT_M", "M"):
        label = {"IN_M": "In Motor", "OUT_M": "Out Motor", "M": "Motor"}[dev_code]
        prefix = f"{area} " if area else ""
        return f"{prefix}{label}{dev_no}"

    if dev_code == "D":
        desc = f"DOOR{dev_no}"
        return f"{area}_{desc}" if area else desc

    return s


def parse_args():
    parser = argparse.ArgumentParser(description="Convert A.xlsx (X/Y sections) to T.xlsx with REF/COMMENT/DESCRIPTION.")
    parser.add_argument("-i", "--input", default="A.xlsx", help="Input Excel filename (default: A.xlsx)")
    parser.add_argument("-o", "--output", default="T.xlsx", help="Output Excel filename (default: T.xlsx)")
    parser.add_argument("--sheet", default=None, help="Excel sheet name or index (default: first sheet)")
    return parser.parse_args()


def main():
    # 切換到腳本所在目錄，確保相對路徑正確
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)

    args = parse_args()

    # === 1. 讀取表 A ===
    sheet_arg = args.sheet
    if sheet_arg is not None and isinstance(sheet_arg, str) and sheet_arg.isdigit():
        sheet_arg = int(sheet_arg)
    df_a = pd.read_excel(args.input, header=None, dtype=str, sheet_name=(sheet_arg if sheet_arg is not None else 0))

    def get_section(df, start_row: int, col_code: int, col_name: int, code_regex: str) -> pd.DataFrame:
        cols = df.columns.tolist()
        if col_code not in cols or col_name not in cols:
            return pd.DataFrame(columns=["Code", "Name_zh"])
        sec = df.loc[start_row:, [col_code, col_name]].copy()
        sec.columns = ["Code", "Name_zh"]
        # 先過濾掉 Name_zh 為空或 NaN 的行
        sec = sec[~sec["Name_zh"].apply(is_na)]
        sec["Code"] = sec["Code"].astype(str).str.strip()
        sec["Name_zh"] = sec["Name_zh"].astype(str).str.strip()
        sec = sec[sec["Code"].str.match(code_regex, na=False)]
        return sec

    # X 區段 (A16,B16 開始)
    df_x = get_section(df_a, 15, 0, 1, r'^[Xx]\d+$')

    # Y 區段 (F16,G16 開始)
    df_y = get_section(df_a, 15, 5, 6, r'^[Yy]\d+$')
    if df_y.empty and (5 not in df_a.columns or 6 not in df_a.columns):
        print("[INFO] 未找到欄位 F/G（Y 區段），將只輸出 X 區段。")

    # === 2. 向量化產出 T ===
    x_desc = df_x["Name_zh"].apply(convert_x_name)
    y_desc = df_y["Name_zh"].apply(convert_y_name)

    df_tx = pd.DataFrame({
        "REF": df_x["Code"],
        "COMMENT": x_desc,
        "DESCRIPTION": "",
    })

    df_ty = pd.DataFrame({
        "REF": df_y["Code"],
        "COMMENT": y_desc,
        "DESCRIPTION": "",
    })

    df_t = pd.concat([df_tx, df_ty], ignore_index=True)
    output_path = args.output
    try:
        df_t.to_excel(output_path, index=False)
    except PermissionError:
        base, ext = os.path.splitext(output_path)
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        alt = f"{base}_{ts}{ext}"
        print(f"[WARN] {output_path} is in use. Writing to {alt} instead.")
        df_t.to_excel(alt, index=False)
        output_path = alt
    print(f"已輸出 {output_path}（X 與 Y 區段，COMMENT 為英文結果，DESCRIPTION 僅表頭）")


if __name__ == "__main__":
    main()


