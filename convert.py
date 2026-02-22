import pandas as pd
import json
import os

# ---------------------------------------------------------------------------
# 職安衛學科題庫轉檔工具 (支援五層連動架構)
# ---------------------------------------------------------------------------

EXCEL_FILE = 'osha_questions.xlsx'
JS_FILE = 'questions.js'

def main():
    print(f"啟動題庫轉檔引擎：準備讀取 {EXCEL_FILE}...")
    
    # 初始化最終輸出的資料結構，分為「歷屆考試」與「分類法規」兩大塊
    question_bank = {
        "exam": [],
        "law": []
    }
    
    try:
        # 讀取 Excel 檔案
        xls = pd.ExcelFile(EXCEL_FILE)
        
        # ---------------------------------------------------------
        # 1. 處理「歷屆考試」工作表 (Exam)
        # ---------------------------------------------------------
        if 'Exam' in xls.sheet_names:
            print(">> 正在處理「歷屆考試(Exam)」工作表...")
            df_exam = pd.read_excel(xls, 'Exam')
            df_exam = df_exam.fillna("") # 處理所有的空值 (NaN)，替換為空字串
            
            for index, row in df_exam.iterrows():
                # 若題目為空，代表是空行，直接跳過
                if str(row.get('Question', '')).strip() == '':
                    continue
                    
                question_data = {
                    "level": str(row.get('Level', '')).strip(),     # 對應第二層：甲級安全/甲級衛生...
                    "batch": str(row.get('Batch', '')).strip(),     # 對應第三層：112-3...
                    "type": str(row.get('Type', '')).strip(),       # 對應第四層：單選/複選
                    "qNum": str(row.get('QNum', '')).strip(),       # 題號 (用於第五層切分範圍)
                    "question": str(row.get('Question', '')).strip(),
                    "options": [
                        str(row.get('A', '')).strip(),
                        str(row.get('B', '')).strip(),
                        str(row.get('C', '')).strip(),
                        str(row.get('D', '')).strip()
                    ],
                    "answer": str(row.get('Answer', '')).strip()
                }
                question_bank["exam"].append(question_data)
            print(f"   [成功] 已匯入 {len(question_bank['exam'])} 筆歷屆試題。")
        else:
            print("   [警告] 找不到名稱為「Exam」的工作表，將跳過歷屆考試題庫。")

        # ---------------------------------------------------------
        # 2. 處理「分類法規」工作表 (Law)
        # ---------------------------------------------------------
        if 'Law' in xls.sheet_names:
            print(">> 正在處理「分類法規(Law)」工作表...")
            df_law = pd.read_excel(xls, 'Law')
            df_law = df_law.fillna("") # 處理所有的空值 (NaN)
            
            for index, row in df_law.iterrows():
                if str(row.get('Question', '')).strip() == '':
                    continue
                    
                question_data = {
                    "category": str(row.get('Category', '')).strip(), # 對應法規第二層：職業安全衛生法...
                    "type": str(row.get('Type', '')).strip(),         # 題型：單選/複選
                    "qNum": str(row.get('QNum', '')).strip(),         # 題號
                    "question": str(row.get('Question', '')).strip(),
                    "options": [
                        str(row.get('A', '')).strip(),
                        str(row.get('B', '')).strip(),
                        str(row.get('C', '')).strip(),
                        str(row.get('D', '')).strip()
                    ],
                    "answer": str(row.get('Answer', '')).strip()
                }
                question_bank["law"].append(question_data)
            print(f"   [成功] 已匯入 {len(question_bank['law'])} 筆分類法規試題。")
        else:
            print("   [警告] 找不到名稱為「Law」的工作表，將跳過分類法規題庫。")

        # ---------------------------------------------------------
        # 3. 匯出成 JavaScript 可直接讀取的檔案
        # ---------------------------------------------------------
        print(f"\n準備將資料打包並寫入 {JS_FILE}...")
        
        # 將 Python 字典轉換為 JSON 格式字串
        json_str = json.dumps(question_bank, ensure_ascii=False, indent=4)
        
        # 將 JSON 包裝成全域常數 questionBank
        js_content = f"// 本檔案由 convert.py 自動生成，請勿手動修改\nconst questionBank = {json_str};\n"
        
        with open(JS_FILE, 'w', encoding='utf-8') as f:
            f.write(js_content)
            
        print(f"🎉 轉換作業完美結束！請確認目錄下已生成最新的 {JS_FILE}。")

    except FileNotFoundError:
        print(f"❌ [錯誤] 找不到檔案 '{EXCEL_FILE}'。請確認 Excel 檔案是否與本程式放在同一個資料夾下。")
    except Exception as e:
        print(f"❌ [錯誤] 轉換過程中發生系統例外：{e}")

if __name__ == "__main__":
    main()