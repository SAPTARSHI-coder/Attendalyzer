import os
import glob
import json
import re
import math
import time
import pandas as pd
from PIL import Image
import google.generativeai as genai
from dotenv import load_dotenv

# Setup Gemini API
load_dotenv()
API_KEY = os.getenv("GEMINI_API_KEY")
if not API_KEY:
    raise ValueError("GEMINI_API_KEY environment variable not set")
genai.configure(api_key=API_KEY)
model = genai.GenerativeModel('gemini-2.5-flash', generation_config={"response_mime_type": "application/json"})

def extract_name(text):
    if pd.isna(text): return ""
    return re.sub(r'[^a-z0-9]', '', str(text).lower())

def process_student_images(student_name, files):
    print(f"Processing {len(files)} files for student: {student_name}", flush=True)
    
    prompt = """
    You are an expert data extractor analyzing university attendance records from screenshots.
    Extract the following information:
    1. "subject_wise": A list of subjects showing the overall attendance summary. For each, provide "subject" (string), "total_classes" (integer), and "attended_classes" (integer).
    2. "date_wise": A comprehensive list of all date-specific attendance logs visible in ALL images. For each record, provide "date" (YYYY-MM-DD), "subject" (string), and "status" (strictly "Present" or "Absent").
    
    Format the output STRICTLY as a JSON object:
    {
      "subject_wise": [
        {"subject": "Math", "total_classes": 40, "attended_classes": 30}
      ],
      "date_wise": [
        {"date": "2024-04-10", "subject": "Math", "status": "Present"}
      ]
    }
    Make sure to combine data if there are multiple images.
    """
    
    try:
        inputs = [prompt]
        uploaded_files = []
        for f in files:
            if f.lower().endswith(('.png', '.jpg', '.jpeg')):
                try:
                    inputs.append(Image.open(f))
                except Exception as e:
                    print(f"  Error opening image {f}: {e}", flush=True)
            else:
                try:
                    print(f"  Uploading file {f}...", flush=True)
                    uploaded_file = genai.upload_file(f)
                    uploaded_files.append(uploaded_file)
                    inputs.append(uploaded_file)
                except Exception as e:
                    print(f"  Failed to upload {f}: {e}", flush=True)
                
        if len(inputs) == 1:
            print("  No valid images or files to process.", flush=True)
            return {"subject_wise": [], "date_wise": []}
            
        print("  Calling Gemini API...", flush=True)
        response = model.generate_content(inputs)
        print("  Success!", flush=True)
        
        for uf in uploaded_files:
            try:
                genai.delete_file(uf.name)
            except:
                pass
                
        return json.loads(response.text)
    except Exception as e:
        print(f"Error processing {student_name}: {e}", flush=True)
        return {"subject_wise": [], "date_wise": []}

def main():
    excel_file = "B.Tech 4th Semester Attendance Collection (Debarred List) (Responses).xlsx"
    df = pd.read_excel(excel_file)
    
    folder1 = "Upload UMS Attendance Screenshot-Report (Mandatory) (File responses)"
    folder2 = "Any attendance missed but you are present_ (provide ss that you are present that day in other classes) (File responses)"
    
    all_files = glob.glob(os.path.join(folder1, "*.*")) + glob.glob(os.path.join(folder2, "*.*"))
    
    results = []
    cache_file = "ocr_cache.json"
    if os.path.exists(cache_file):
        with open(cache_file, "r") as f:
            ocr_cache = json.load(f)
    else:
        ocr_cache = {}

    for index, row in df.iterrows():
        name = str(row["Student's Name (As per University Records)"]).strip()
        roll = str(row["Student's University Roll Number"]).strip()
        med_prov = str(row["Medical certificate provided?"]).strip()
        med_range = str(row["Medical certificate range written in certificate"]).strip()
        
        normalized_name = extract_name(name)
        
        student_files = []
        for f in all_files:
            basename = os.path.basename(f)
            if normalized_name and normalized_name in extract_name(basename):
                student_files.append(f)
                
        if not student_files and normalized_name:
            parts = name.lower().split()
            for f in all_files:
                basename = os.path.basename(f).lower()
                if any(len(p) > 3 and p in basename for p in parts):
                    student_files.append(f)
                    
        student_files = list(set(student_files))
        
        if name in ocr_cache:
            extracted_data = ocr_cache[name]
            print(f"Loaded {name} from cache. (Files matched: {len(student_files)})", flush=True)
        else:
            if student_files:
                extracted_data = process_student_images(name, student_files)
                ocr_cache[name] = extracted_data
                with open(cache_file, "w") as f:
                    json.dump(ocr_cache, f, indent=2)
                time.sleep(4)  # To avoid Gemini API rate limits (15 RPM)
            else:
                extracted_data = {"subject_wise": [], "date_wise": []}
                print(f"No files found for {name}", flush=True)
                
        # --- APPLY LOGIC ---
        subjects = extracted_data.get("subject_wise", [])
        dates = extracted_data.get("date_wise", [])
        
        date_map = {}
        for d in dates:
            d_date = d.get("date")
            if not d_date: continue
            if d_date not in date_map:
                date_map[d_date] = []
            date_map[d_date].append(d)
            
        corrected_dates = []
        for d_date, records in date_map.items():
            is_present_any = any(r.get("status", "").lower() == "present" for r in records)
            if is_present_any:
                for r in records:
                    if r.get("status", "").lower() == "absent":
                        r["status"] = "Present"
                        subj = r.get("subject", "Unknown Subject")
                        corrected_dates.append(f"{d_date} ({subj})")
                        for s in subjects:
                            # Fuzzy match
                            if s.get("subject", "")[:10].lower() in subj.lower() or subj[:10].lower() in s.get("subject", "").lower():
                                s["attended_classes"] = s.get("attended_classes", 0) + 1
                                break
                                
        total_overall_classes = 0
        total_overall_attended = 0
        subject_details = []
        
        for s in subjects:
            tc = s.get("total_classes", 0)
            ac = s.get("attended_classes", 0)
            if tc > 0:
                pct = (ac / tc) * 100
                subject_details.append(f"{s.get('subject', 'Unknown')}: {ac}/{tc} ({pct:.1f}%)")
                total_overall_classes += tc
                total_overall_attended += ac
                
        final_percentage = (total_overall_attended / total_overall_classes * 100) if total_overall_classes > 0 else 0
        
        is_medical = "yes" in med_prov.lower()
        threshold = 65.0 if is_medical else 75.0
        
        is_debarred = "DEBARRED" if final_percentage < threshold else "SAFE"
        
        classes_needed = 0
        if final_percentage < threshold and total_overall_classes > 0:
            req_classes = math.ceil((threshold / 100) * total_overall_classes)
            classes_needed = max(0, req_classes - total_overall_attended)
            
        results.append({
            "Name": name,
            "Roll Number": roll,
            "Subject-wise Details": " | ".join(subject_details),
            "Corrected Dates": ", ".join(corrected_dates) if corrected_dates else "None",
            "Total Classes": total_overall_classes,
            "Total Attended (Corrected)": total_overall_attended,
            "Final Percentage": f"{final_percentage:.2f}%",
            "Medical Given": "Yes" if is_medical else "No",
            "Medical Date Range": med_range if is_medical else "",
            "Debarment Logic": is_debarred,
            "Classes Needed to be Safe": classes_needed
        })
        
    final_df = pd.DataFrame(results)
    final_df.to_excel("Final_Corrected_Attendance.xlsx", index=False)
    print("Successfully generated Final_Corrected_Attendance.xlsx", flush=True)

if __name__ == "__main__":
    main()
