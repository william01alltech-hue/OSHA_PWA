import pandas as pd
import json
import re

def convert_excel_to_js():
    questions = []

    # -------------------------------------
    # 1. 處理 Choice (學科)
    # -------------------------------------
    try:
        try:
            df_choice = pd.read_excel('osha_questions.xlsx', sheet_name='Choice')
        except:
            # 相容性：若無 Choice 分頁，嘗試讀取第一個
            df_choice = pd.read_excel('osha_questions.xlsx', sheet_name=0)
            
        print(f"📊 讀取學科題目：{len(df_choice)} 題")
        
        for _, row in df_choice.iterrows():
            # 支援中英文標題容錯
            ans = str(row.get('Answer', row.get('正確答案', ''))).replace('.0', '').strip()
            
            q_item = {
                "id": str(row.get('ID', row.get('題目編號', ''))),
                "year": int(row.get('Year', row.get('年度', 110))),
                "batch": int(row.get('Batch', row.get('梯次', 1))),
                "subject": str(row.get('Subject', row.get('科目', '不分'))).strip(), # ★ 新增科目欄位
                "mode": str(row.get('Mode', row.get('模式', ''))).strip(),
                "type": "choice",
                "question": str(row.get('Question', row.get('題目內容', ''))).strip(),
                "options": [
                    str(row.get('Opt1', row.get('選項1', ''))).strip(),
                    str(row.get('Opt2', row.get('選項2', ''))).strip(),
                    str(row.get('Opt3', row.get('選項3', ''))).strip(),
                    str(row.get('Opt4', row.get('選項4', ''))).strip()
                ],
                "answer": ans
            }
            questions.append(q_item)
    except Exception as e:
        print(f"⚠️ 學科讀取略過: {e}")

    # -------------------------------------
    # 2. 處理 Essay (術科)
    # -------------------------------------
    try:
        df_essay = pd.read_excel('osha_questions.xlsx', sheet_name='Essay')
        print(f"📝 讀取術科題目：{len(df_essay)} 題")

        for _, row in df_essay.iterrows():
            # 取得原始的評分標準文字
            raw_criteria = str(row.get('Criteria', row.get('關鍵字', ''))).strip()
            if raw_criteria == 'nan': raw_criteria = ""
            
            # 關鍵字提取
            stds = []
            match = re.search(r"關鍵字[：: ]*(.*)", raw_criteria)
            if match:
                kw_str = match.group(1).split('\n')[0]
                stds = re.split(r'[、,， ]+', kw_str)
                stds = [s.strip() for s in stds if s.strip()]
            else:
                # 若無「關鍵字：」前綴，則直接以逗號分割
                k_str = raw_criteria.replace('，', ',')
                stds = [k.strip() for k in k_str.split(',') if k.strip()]

            q_item = {
                "id": str(row.get('ID', row.get('題目編號', ''))),
                "year": int(row.get('Year', row.get('年度', 110))),
                "batch": int(row.get('Batch', row.get('梯次', 1))),
                "subject": str(row.get('Subject', row.get('考試類別', row.get('科目', '不分')))).strip(), # ★ 新增科目欄位
                "type": "essay",
                "question": str(row.get('Question', row.get('題目內容', ''))).strip(),
                "answer": str(row.get('RefAnswer', row.get('正確答案', ''))).strip(), 
                "criteria_display": raw_criteria,        
                "keywords": stds, # ★ 變更為 keywords 以配合 V38 系統
                "image": str(row.get('Image', '')) if 'Image' in row else ""
            }
            questions.append(q_item)
            
    except Exception as e:
        print(f"⚠️ 術科讀取略過: {e}")

    # -------------------------------------
    # 3. 輸出
    # -------------------------------------
    with open('questions.js', 'w', encoding='utf-8') as f:
        f.write(f"const questionBank = {json.dumps(questions, ensure_ascii=False, indent=2)};")
    
    print(f"✅ 轉檔完成！總計 {len(questions)} 題。")

if __name__ == "__main__":
    convert_excel_to_js()