"""
Reprocesses students whose OCR cache is empty.
Uses exact screenshot filename from the Responses sheet.
"""
import json, os, re, time, sys, warnings
warnings.filterwarnings('ignore')

# Force unbuffered output
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', line_buffering=True)

print("=== Attendance OCR Reprocessor ===", flush=True)
print("Loading libraries...", flush=True)

import pandas as pd
from PIL import Image
import google.generativeai as genai
from dotenv import load_dotenv

load_dotenv()

API_KEY = os.getenv("GEMINI_API_KEY")
if not API_KEY:
    raise ValueError("GEMINI_API_KEY environment variable not set")
genai.configure(api_key=API_KEY)

# gemini-2.5-flash is the only working vision model on this key
MODEL_NAME = "gemini-2.5-flash"
model = genai.GenerativeModel(
    MODEL_NAME,
    generation_config={"response_mime_type": "application/json"}
)

SCREENSHOT_FOLDER = "Upload UMS Attendance Screenshot-Report (Mandatory) (File responses)"
CACHE_FILE        = "ocr_cache.json"
RESPONSES_EXCEL   = "B.Tech 4th Semester Attendance Collection (Debarred List) (Responses) (1).xlsx"

print(f"Using model: {MODEL_NAME}", flush=True)
print("Loading cache...", flush=True)

with open(CACHE_FILE, "r", encoding="utf-8") as f:
    ocr_cache = json.load(f)

print(f"Cache has {len(ocr_cache)} valid entries.", flush=True)
print("Loading responses Excel...", flush=True)

resp_df = pd.read_excel(RESPONSES_EXCEL)
resp_df.columns = resp_df.columns.str.strip()

NAME_COL       = "Student's Name (As per University Records)"
ROLL_COL       = "Student's University Roll Number"
SCREENSHOT_COL = "Upload UMS Attendance Screenshot/Report (Mandatory)"

PROMPT = """
You are an expert data extractor for university attendance records.
Analyze this UMS (University Management System) attendance screenshot carefully.

Extract:
1. "subject_wise": For EVERY subject row visible, extract:
   - "subject": full string shown (e.g. "CSE11111 || Formal Language and Automata")
   - "total_classes": integer total classes held
   - "attended_classes": integer classes attended by student

2. "date_wise": For every individual date row visible in detail tables, extract:
   - "date": in "YYYY-MM-DD" format
   - "subject": subject name
   - "status": "Present" or "Absent"

Return ONLY valid JSON, no markdown:
{
  "subject_wise": [
    {"subject": "CSE11111 || Formal Language and Automata", "total_classes": 40, "attended_classes": 30}
  ],
  "date_wise": [
    {"date": "2025-04-10", "subject": "Formal Language and Automata", "status": "Present"}
  ]
}
"""

def save_cache():
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(ocr_cache, f, indent=2, ensure_ascii=False)

def call_gemini(filepath):
    inputs = [PROMPT]
    uploaded = None
    try:
        ext = filepath.lower().rsplit('.', 1)[-1]
        if ext in ('jpg', 'jpeg', 'png', 'webp', 'gif'):
            img = Image.open(filepath)
            inputs.append(img)
        else:
            print(f"      Uploading {os.path.basename(filepath)}...", flush=True)
            uploaded = genai.upload_file(filepath)
            inputs.append(uploaded)

        # Retry loop for rate-limit (429) errors
        for wait_sec in [0, 65, 65, 120]:
            if wait_sec > 0:
                print(f"      Rate limited — waiting {wait_sec}s...", flush=True)
                time.sleep(wait_sec)
            try:
                resp   = model.generate_content(inputs)
                result = json.loads(resp.text)
                return result
            except Exception as e:
                err = str(e)
                if '429' in err:
                    continue   # retry after waiting
                elif 'JSONDecodeError' in type(e).__name__ or 'json' in err.lower():
                    print(f"      JSON parse error — raw: {getattr(resp,'text','?')[:200]}", flush=True)
                    return {"subject_wise": [], "date_wise": []}
                else:
                    print(f"      API Error: {err[:120]}", flush=True)
                    return None

        print("      Gave up after rate-limit retries.", flush=True)
        return None

    finally:
        if uploaded:
            try: genai.delete_file(uploaded.name)
            except: pass

# Build filename → path map
all_files = {f: os.path.join(SCREENSHOT_FOLDER, f) for f in os.listdir(SCREENSHOT_FOLDER)}

students = list(resp_df.iterrows())
total    = len(students)
done     = 0
skipped  = 0
failed   = 0

print(f"\nTotal students: {total}", flush=True)
print("─" * 60, flush=True)

for i, (_, row) in enumerate(students):
    name    = str(row[NAME_COL]).strip()
    roll    = str(row[ROLL_COL]).strip()
    ss_name = str(row[SCREENSHOT_COL]).strip()

    # Skip if already cached with real data
    cached = ocr_cache.get(name, {})
    if len(cached.get("subject_wise", [])) > 0:
        skipped += 1
        subj_count = len(cached["subject_wise"])
        print(f"  SKIP [{i+1}/{total}] {name} ({subj_count} subjects cached)", flush=True)
        continue

    # Find file
    filepath = all_files.get(ss_name)
    if not filepath:
        # Case-insensitive fallback
        ss_lower = ss_name.lower()
        for fname, fpath in all_files.items():
            if fname.lower() == ss_lower:
                filepath = fpath
                break

    if not filepath:
        print(f"  FILE NOT FOUND [{i+1}/{total}] {name} — '{ss_name}'", flush=True)
        failed += 1
        ocr_cache[name] = {"subject_wise": [], "date_wise": []}
        save_cache()
        continue

    print(f"  OCR [{i+1}/{total}] {name}", flush=True)
    print(f"    → {os.path.basename(filepath)}", flush=True)

    # Try up to 2 times
    result = None
    for attempt in range(2):
        result = call_gemini(filepath)
        if result is not None:
            break
        print(f"    Retrying in 30s (attempt {attempt+1})...", flush=True)
        time.sleep(30)

    if result is None:
        result = {"subject_wise": [], "date_wise": []}
        failed += 1
        print(f"    ✗ Failed after retries", flush=True)
    else:
        n_subj = len(result.get("subject_wise", []))
        print(f"    ✓ {n_subj} subjects extracted", flush=True)
        done += 1

    ocr_cache[name] = result
    save_cache()

    # 4-second gap → 15 req/min safely
    time.sleep(4)

print("\n" + "=" * 60, flush=True)
print(f"DONE.  Processed: {done}  |  Skipped: {skipped}  |  Failed: {failed}", flush=True)

# Auto-generate the final Excel
print("\nGenerating Final_Subject_Wise_Attendance.xlsx...", flush=True)
import subprocess
subprocess.run(["python", "fill_attendance_format.py"], check=True)
print("All complete!", flush=True)
