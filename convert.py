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
            ans = str(row['Answer']).replace('.0', '').strip()
            q_item = {
                "id": str(row['ID']),
                "year": int(row['Year']),
                "batch": int(row['Batch']),
                "mode": str(row['Mode']).strip(),
                "type": "choice",
                "question": str(row['Question']).strip(),
                "options": [
                    str(row['Opt1']).strip(),
                    str(row['Opt2']).strip(),
                    str(row['Opt3']).strip(),
                    str(row['Opt4']).strip()
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
            # 取得原始的評分標準文字 (給 AI 讀懂語意用)
            raw_criteria = str(row['Criteria']).strip()
            if raw_criteria == 'nan': raw_criteria = ""
            
            # 關鍵字提取 (給電腦輔助標記用)
            stds = []
            match = re.search(r"關鍵字[：: ]*(.*)", raw_criteria)
            if match:
                kw_str = match.group(1).split('\n')[0]
                stds = re.split(r'[、,， ]+', kw_str)
                stds = [s.strip() for s in stds if s.strip()]

            q_item = {
                "id": str(row['ID']),
                "year": int(row['Year']),
                "batch": int(row['Batch']),
                "type": "essay",
                "question": str(row['Question']).strip(),
                "answer": str(row['RefAnswer']).strip(), # 標準參考解答
                "criteria_display": raw_criteria,        # 完整評分標準
                "standards": stds,                       # 關鍵字陣列
                "image": str(row['Image']) if 'Image' in row else ""
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