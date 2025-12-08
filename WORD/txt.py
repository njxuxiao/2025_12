import pandas as pd
import os
import re
from zipfile import BadZipFile

def sanitize_filename(filename):
    """
    清洗文件名，去掉Windows/Linux文件名中不允许的特殊字符
    """
    return re.sub(r'[\\/*?:"<>|]', "_", str(filename))

def convert_excel_to_txt(excel_filepath):
    """
    读取 Excel 文件，将每一个 有数据的 Sheet 转换为单独的 txt 文件。
    会自动跳过内容为空的 Sheet。
    """
    if not os.path.exists(excel_filepath):
        print(f"Error: 找不到文件 '{excel_filepath}'。跳过。")
        return

    if not excel_filepath.lower().endswith('.xlsx'):
        print(f"Error: 文件 '{excel_filepath}' 不是 .xlsx 文件。跳过。")
        return

    print(f"\n--- 正在处理: '{excel_filepath}' ---")
    
    try:
        xls = pd.ExcelFile(excel_filepath, engine='openpyxl')
        sheet_names = xls.sheet_names

        if not sheet_names:
            print("  在此文件中未找到任何 Sheet。")
            return

        base_filename = os.path.splitext(os.path.basename(excel_filepath))[0]
        output_dir = os.path.dirname(excel_filepath)

        for sheet_name in sheet_names:
            # 读取 Sheet
            df = pd.read_excel(xls, sheet_name=sheet_name)
            
            # --- 新增功能：检测是否为空 Sheet ---
            # 如果 df.empty 为 True，说明没有数据行（或者只有表头但没有内容）
            # 我们还可以更严格一点，去掉全是空的行再判断
            df_cleaned = df.dropna(how='all') 
            
            if df_cleaned.empty:
                # 只有当真的没有任何数据时，才跳过
                print(f"  -> [跳过] Sheet '{sheet_name}' 是空的 (无数据)。")
                continue
            # ----------------------------------

            print(f"  -> 正在转换 Sheet: '{sheet_name}'...")

            # 清理列名空格
            df.columns = df.columns.astype(str).str.strip()
            
            # 命名逻辑: Excel名_Sheet名.txt
            safe_sheet_name = sanitize_filename(sheet_name)
            new_txt_name = f"{base_filename}_{safe_sheet_name}.txt"
            output_path = os.path.join(output_dir, new_txt_name)
            
            # 保存
            df.to_csv(output_path, sep='\t', index=False, encoding='utf-8-sig')
            print(f"     成功! 已保存为: '{new_txt_name}'")
            
    except BadZipFile:
        print(f"  Error: 文件 '{excel_filepath}' 似乎已损坏。")
    except Exception as e:
        print(f"  发生意外错误: {e}")

if __name__ == '__main__':
    files_to_convert = [
        "1_21_07.xlsx", 
        "2_22_12.xlsx", 
        "3_22_07.xlsx", 
        "4_21_12.xlsx",
        "5_20_12.xlsx", 
        "6_19_12.xlsx",
        "7_19_07.xlsx",
        "blue_1.xlsx",
        "blue_2.xlsx"
    ]

    print("开始批量转换 (智能跳过空白 Sheet)...")

    for file in files_to_convert:
        convert_excel_to_txt(file)
    
    print("\n所有转换任务完成!")