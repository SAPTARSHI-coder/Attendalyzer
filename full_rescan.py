"""
Full re-scan: wipes cache and re-processes every student's UMS screenshot
with a precise prompt that reads EXACTLY the numbers shown on screen.
Then generates the final Excel in the exact attendance.xlsx format.
"""
import json, os, re, time, sys, warnings, subprocess
warnings.filterwarnings('ignore')

import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', line_buffering=True)

print("=== FULL RESCAN - Attendance OCR ===", flush=True)

import pandas as pd
from PIL import Image
import google.generativeai as genai
from dotenv import load_dotenv

load_dotenv()

API_KEY = os.getenv("GEMINI_API_KEY")
if not API_KEY:
    raise ValueError("GEMINI_API_KEY environment variable not set")
genai.configure(api_key=API_KEY)
model = genai.GenerativeModel(
    "gemini-2.5-flash",
    generation_config={"response_mime_type": "application/json"}
)

SCREENSHOT_FOLDER = "Upload UMS Attendance Screenshot-Report (Mandatory) (File responses)"
CACHE_FILE        = "ocr_cache_v2.json"
RESPONSES_EXCEL   = "B.Tech 4th Semester Attendance Collection (Debarred List) (Responses) (1).xlsx"

# ── Strict prompt ─────────────────────────────────────────────────────────────
PROMPT = """
This is a UMS (University Management System) attendance report screenshot from an Indian university.

Your task: Read the attendance summary table EXACTLY as shown. Do NOT guess or estimate.

The table has columns like:
  Course Code | Course Name | Total Classes | Present | Absent | Percentage

For EACH subject row in the table, extract:
- "code": the course code exactly as shown (e.g. "CSE11111", "MTH11534", "PSG11021", "CSE12166")
- "name": the course name exactly as shown
- "total": integer - the total number of classes held (the column that shows the largest number)
- "present": integer - the number of classes the student attended/was present
- "percentage": the attendance percentage shown (as a number, e.g. 82.05 not "82.05%")

IMPORTANT RULES:
- Read ONLY what is printed. Do NOT add, guess, or calculate.
- If a cell is unclear, use 0.
- The code and name may be separated by || or be in separate columns — handle both.
- Do NOT confuse theory subjects with lab subjects. Labs have codes like CSE12xxx, MTH12xxx.
- List ALL rows visible in the table, even if percentage is 100%.

Return ONLY this JSON, no markdown, no explanation:
{
  "subjects": [
    {
      "code": "CSE11111",
      "name": "Formal Language and Automata",
      "total": 39,
      "present": 32,
      "percentage": 82.05
    }
  ]
}
"""

def call_gemini_image(filepath):
    inputs = [PROMPT]
    uploaded = None
    try:
        ext = filepath.lower().rsplit('.', 1)[-1]
        if ext in ('jpg', 'jpeg', 'png', 'webp', 'gif', 'bmp'):
            img = Image.open(filepath)
            inputs.append(img)
        else:
            uploaded = genai.upload_file(filepath)
            inputs.append(uploaded)

        for wait in [0, 70, 70, 120]:
            if wait > 0:
                print(f"      Rate limit hit, waiting {wait}s...", flush=True)
                time.sleep(wait)
            try:
                resp = model.generate_content(inputs)
                raw  = resp.text.strip()
                # Strip markdown code fences if present
                raw = re.sub(r'^```[a-z]*\n?', '', raw)
                raw = re.sub(r'\n?```$', '', raw)
                data = json.loads(raw)
                return data.get("subjects", [])
            except Exception as e:
                if '429' in str(e):
                    continue
                print(f"      Error: {str(e)[:120]}", flush=True)
                return []

        print("      Gave up after retries.", flush=True)
        return []
    finally:
        if uploaded:
            try: genai.delete_file(uploaded.name)
            except: pass

# ── Load responses sheet ───────────────────────────────────────────────────────
resp_df = pd.read_excel(RESPONSES_EXCEL)
resp_df.columns = resp_df.columns.str.strip()

NAME_COL  = "Student's Name (As per University Records)"
ROLL_COL  = "Student's University Roll Number"
SS_COL    = "Upload UMS Attendance Screenshot/Report (Mandatory)"

# ── Load or create fresh cache ─────────────────────────────────────────────────
if os.path.exists(CACHE_FILE):
    with open(CACHE_FILE, encoding='utf-8') as f:
        ocr_cache = json.load(f)
    print(f"Existing v2 cache: {len(ocr_cache)} entries", flush=True)
else:
    ocr_cache = {}
    print("Starting fresh v2 cache", flush=True)

def save():
    with open(CACHE_FILE, 'w', encoding='utf-8') as f:
        json.dump(ocr_cache, f, indent=2, ensure_ascii=False)

# Build file map
all_files = {fn: os.path.join(SCREENSHOT_FOLDER, fn)
             for fn in os.listdir(SCREENSHOT_FOLDER)}
all_lower  = {fn.lower(): fp for fn, fp in all_files.items()}

students = list(resp_df.iterrows())
print(f"Total students: {len(students)}\n{'─'*60}", flush=True)

done = skipped = failed = 0

for i, (_, row) in enumerate(students):
    name    = str(row[NAME_COL]).strip()
    roll    = str(row[ROLL_COL]).strip()
    ss_name = str(row[SS_COL]).strip()

    # Skip if already in v2 cache with data
    cached = ocr_cache.get(name, [])
    if len(cached) > 0:
        skipped += 1
        print(f"  SKIP [{i+1}/{len(students)}] {name} ({len(cached)} subjects)", flush=True)
        continue

    # Locate file
    filepath = all_files.get(ss_name) or all_lower.get(ss_name.lower())
    if not filepath:
        print(f"  FILE NOT FOUND [{i+1}/{len(students)}] {name} — '{ss_name}'", flush=True)
        ocr_cache[name] = []
        save()
        failed += 1
        continue

    print(f"  OCR [{i+1}/{len(students)}] {name}", flush=True)
    print(f"    → {os.path.basename(filepath)}", flush=True)

    subjects = call_gemini_image(filepath)
    n = len(subjects)
    print(f"    ✓ {n} subjects extracted", flush=True)

    # Sanity-check: cap attended at total
    for s in subjects:
        if s.get('present', 0) > s.get('total', 0):
            s['present'] = s['total']

    ocr_cache[name] = subjects
    save()
    done += 1
    time.sleep(4)  # stay under 15 RPM

print(f"\n{'='*60}", flush=True)
print(f"DONE. Processed:{done}  Skipped:{skipped}  Failed:{failed}", flush=True)

# Auto-generate Excel
print("\nGenerating Excel...", flush=True)
subprocess.run(["python", "-W", "ignore", "build_excel_v2.py"], check=True)
