import json
import pandas as pd
import os
import re

def extract_name(text):
    if pd.isna(text): return ""
    return re.sub(r'[^a-z0-9]', '', str(text).lower())

def main():
    excel_file = "B.Tech 4th Semester Attendance Collection (Debarred List) (Responses).xlsx"
    df = pd.read_excel(excel_file)
    
    with open("ocr_cache.json", "r") as f:
        ocr_cache = json.load(f)
        
    standard_subjects = [
        "CSE12166 || Design and Analysis of Algorithms Lab",
        "CSE11111 || Formal Language and Automata",
        "CSE11110 || Design and Analysis of Algorithms",
        "PSG11021 || Human Values and Professional Ethics",
        "CSE11109 || Object Oriented Programming",
        "MTH11534 || Discrete Structures and Logic",
        "CSE12205 || Exploratory Data Analysis Lab",
        "CSE11112 || Introduction to Artificial Intelligence",
        "CSE11204 || Exploratory Data Analysis",
        "CSE12114 || Object Oriented Programming Lab",
        "MTH12531 || Numerical Techniques Lab",
        "CSE14170 || Mini Project-I"
    ]
    
    rows = []
    for index, row in df.iterrows():
        name = str(row["Student's Name (As per University Records)"]).strip()
        roll = str(row["Student's University Roll Number"]).strip()
        
        extracted_data = ocr_cache.get(name, {})
        subject_data = extracted_data.get("subject_wise", [])
        
        student_record = {
            "Student Name": name,
            "Roll Number": roll
        }
        
        for std_subj in standard_subjects:
            std_code = std_subj.split("||")[0].strip()
            
            # Find matching subject in extracted data
            matched = False
            for s in subject_data:
                extracted_subj = s.get("subject", "")
                if std_code in extracted_subj or std_code.replace(" ", "") in extracted_subj.replace(" ", ""):
                    tc = s.get("total_classes", 0)
                    ac = s.get("attended_classes", 0)
                    pct = round((ac / tc * 100), 2) if tc > 0 else 0
                    
                    student_record[f"{std_code} - Total"] = tc
                    student_record[f"{std_code} - Attended"] = ac
                    student_record[f"{std_code} - Percentage"] = pct
                    matched = True
                    break
                    
            if not matched:
                student_record[f"{std_code} - Total"] = 0
                student_record[f"{std_code} - Attended"] = 0
                student_record[f"{std_code} - Percentage"] = 0
                
        rows.append(student_record)
        
    # Create dataframe
    final_df = pd.DataFrame(rows)
    
    # Save to Excel
    out_file = "Subject_Wise_Attendance_Matrix.xlsx"
    final_df.to_excel(out_file, index=False)
    print(f"Successfully created {out_file}")

if __name__ == "__main__":
    main()
