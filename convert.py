import pandas as pd
import json
import os

# 設定檔名
EXCEL_FILE = 'osha_questions.xlsx.xlsx'
OUTPUT_FILE = 'questions.js'

def convert():
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ 找不到檔案: {EXCEL_FILE}")
        return

    all_data = []
    print("正在讀取 Excel 資料...")

    try:
        # 1. 讀取學科
        # 根據您的截圖 image_a80bfb.jpg，標題是 Option1, Option2...
        df_choice = pd.read_excel(EXCEL_FILE, sheet_name='Choice').fillna('')
        
        for i, row in df_choice.iterrows():
            # --- 處理答案 ---
            # 把數字答案轉成字串，去掉小數點 (例如 2.0 -> "2")
            ans_raw = str(row['Answer']).strip()
            if ans_raw.endswith('.0'):
                ans_raw = ans_raw[:-2]
            
            # --- 處理選項 ---
            # 對應 Excel 的 Option1 ~ Option4
            opts = [
                f"1. {row['Option1']}",
                f"2. {row['Option2']}",
                f"3. {row['Option3']}",
                f"4. {row['Option4']}"
            ]

            all_data.append({
                "year": row['Year'], 
                "batch": row['Batch'], 
                "id": row['ID'],
                "type": "choice", 
                "mode": row.get('Mode', '單選'), 
                "question": row['Question'],
                "options": opts,
                "answer": ans_raw,  # 這裡程式會輸出 "1", "2"... 對應上面的 1. 2.
                "note": row['Note']
            })
        
        # 2. 處理術科 (如果有 Essay 分頁才做)
        if 'Essay' in pd.ExcelFile(EXCEL_FILE).sheet_names:
            df_essay = pd.read_excel(EXCEL_FILE, sheet_name='Essay').fillna('')
            for _, row in df_essay.iterrows():
                all_data.append({
                    "year": row['Year'], "batch": row['Batch'], "id": row['ID'],
                    "type": "essay", "question": row['Question'],
                    "answer": row.get('RefAnswer', ''), 
                    "criteria": str(row.get('Criteria', '')).split('|')
                })

        # 3. 輸出檔案
        with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
            f.write(f"const questionBank = {json.dumps(all_data, ensure_ascii=False, indent=2)};")
        
        print(f"✅ 成功了！共處理了 {len(all_data)} 題。")
        print("請回到網頁按 F5 測試。")

    except Exception as e:
        print(f"❌ 轉換失敗: {e}")
        # 如果還是失敗，這行會告訴我們 Excel 到底讀到了什麼標題
        if 'df_choice' in locals():
             print("Excel 實際標題:", df_choice.columns.tolist())

if __name__ == "__main__":
    convert()