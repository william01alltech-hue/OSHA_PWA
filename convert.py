import pandas as pd
import json

# ---------------------------------------------------------------------------
# 職安衛學科題庫轉檔工具 (內建資料清洗與 X 光掃描報告)
# ---------------------------------------------------------------------------

EXCEL_FILE = 'osha_questions.xlsx'
JS_FILE = 'questions.js'

def map_answer(ans):
    ans = str(ans).strip().split('.')[0]
    mapping = {'1': 'A', '2': 'B', '3': 'C', '4': 'D', 'A':'A', 'B':'B', 'C':'C', 'D':'D'}
    return mapping.get(ans.upper(), ans)

def determine_level(subject_str):
    subject_str = str(subject_str)
    if '甲級安全' in subject_str: return '甲級安全'
    if '甲級衛生' in subject_str: return '甲級衛生'
    if '乙級' in subject_str: return '乙級職安衛'
    return '其他'

def normalize_type(type_str):
    """資料清洗：把各種寫法的題型統一成標準格式"""
    type_str = str(type_str).strip()
    if '單' in type_str: return '單選'
    if '複' in type_str: return '複選'
    return type_str

def main():
    print(f"啟動題庫轉檔引擎 (X光掃描版)：準備讀取 {EXCEL_FILE}...\n")
    question_bank = {"exam": [], "law": []}
    
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
        
        # --- 處理 Exam ---
        if 'Exam' in xls.sheet_names:
            df_exam = pd.read_excel(xls, 'Exam').fillna("") 
            total_rows = len(df_exam)
            skipped_empty = 0
            
            print(f"🔍 [Exam] 掃描到 Excel 共有 {total_rows} 列資料...")
            
            for index, row in df_exam.iterrows():
                # 1. 抓取題目，支援多種可能欄位名稱
                q_text = str(row.get('題目內容', row.get('題目', ''))).strip()
                if not q_text:
                    skipped_empty += 1
                    continue
                
                # 2. 處理年度梯次
                year = str(row.get('年度', '')).split('.')[0].strip()
                batch = str(row.get('梯次', '')).split('.')[0].strip()
                combined_batch = f"{year}-{batch}" if year and batch else "未分類梯次"
                
                # 3. 抓取擴充資訊
                note_info = str(row.get('參考資訊', row.get('Note', ''))).strip()
                
                question_bank["exam"].append({
                    "level": determine_level(row.get('科目', '')),
                    "batch": combined_batch,
                    "type": normalize_type(row.get('模式', '')), # 自動清洗題型
                    "qNum": str(row.get('題目編號', index + 1)), 
                    "question": q_text,
                    "options": [str(row.get('選項1', '')).strip(), str(row.get('選項2', '')).strip(), str(row.get('選項3', '')).strip(), str(row.get('選項4', '')).strip()],
                    "answer": map_answer(row.get('正確答案', '')),
                    "law_name": str(row.get('法令名稱去條文', '')).strip(),
                    "law_article": str(row.get('法令條文', '')).strip(),
                    "note": note_info
                })
            print(f"✅ [Exam] 成功抓取 {len(question_bank['exam'])} 筆歷屆試題。 (略過了 {skipped_empty} 列沒有題目的空行)")
        
        # --- 處理 Law ---
        if 'Law' in xls.sheet_names:
            df_law = pd.read_excel(xls, 'Law').fillna("")
            total_rows = len(df_law)
            skipped_empty = 0
            
            print(f"🔍 [Law] 掃描到 Excel 共有 {total_rows} 列資料...")
            
            cols = df_law.columns.tolist()
            for index, row in df_law.iterrows():
                q_text = str(row.get('題目內容', row.get('題目', ''))).strip()
                if not q_text:
                    skipped_empty += 1
                    continue
                
                category_name = ""
                if '法令名稱去條文' in cols: category_name = str(row['法令名稱去條文']).strip()
                elif len(cols) > 14: category_name = str(row.iloc[14]).strip()
                    
                if not category_name or category_name.lower() == 'nan':
                    category_name = '其他'
                    
                note_info = str(row.get('參考資訊', row.get('Note', ''))).strip()
                    
                question_bank["law"].append({
                    "category": category_name,
                    "type": normalize_type(row.get('模式', '')), # 自動清洗題型
                    "qNum": str(row.get('題目編號', index + 1)),
                    "question": q_text,
                    "options": [str(row.get('選項1', '')).strip(), str(row.get('選項2', '')).strip(), str(row.get('選項3', '')).strip(), str(row.get('選項4', '')).strip()],
                    "answer": map_answer(row.get('正確答案', '')),
                    "law_name": str(row.get('法令名稱去條文', category_name)).strip(),
                    "law_article": str(row.get('法令條文', '')).strip(),
                    "note": note_info
                })
            print(f"✅ [Law] 成功抓取 {len(question_bank['law'])} 筆分類法規。 (略過了 {skipped_empty} 列沒有題目的空行)")

        json_str = json.dumps(question_bank, ensure_ascii=False, indent=4)
        with open(JS_FILE, 'w', encoding='utf-8') as f:
            f.write(f"// 自動生成題庫\nconst questionBank = {json_str};\n")
            
        print(f"\n🎉 轉檔完畢！")

    except Exception as e:
        print(f"❌ [錯誤] {e}")

if __name__ == "__main__":
    main()