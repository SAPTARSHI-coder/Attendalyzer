# 🎓 Attendalyzer: Automated Student Attendance Reconciliation System

> An AI-powered, audit-ready data pipeline that performs high-accuracy OCR on UMS screenshots, reconciles self-reported medical and event participation claims, applies institutional business logic, and produces a structured, color-coded Excel report for debarment analysis.

---

## 📋 Table of Contents

1. [📖 Overview](#overview)
2. [🛑 Problem Statement](#problem-statement)
3. [🏗️ System Architecture](#system-architecture)
4. [📁 Project Structure](#project-structure)
5. [🗄️ Data Sources](#data-sources)
6. [💻 Scripts — Detailed Reference](#scripts--detailed-reference)
7. [🧠 Business Logic](#business-logic)
8. [📊 Output Files](#output-files)
9. [🎨 Color Coding Reference](#color-coding-reference)
10. [🚀 Setup & Installation](#setup--installation)
11. [🏃 Running the Pipeline](#running-the-pipeline)
12. [⚙️ Configuration & Manual Overrides](#configuration--manual-overrides)
13. [📚 Subjects Reference](#subjects-reference)
14. [💾 Caching Architecture](#caching-architecture)
15. [🔧 Troubleshooting](#troubleshooting)
16. [📦 Dependencies](#dependencies)

---

## 📖 Overview

This project automates the end-of-semester attendance debarment process for a B.Tech 4th Semester cohort. Instead of manually reviewing hundreds of UMS (University Management System) screenshots and cross-checking student claims, this pipeline:

1. **OCR-scans** every student's uploaded UMS attendance screenshot using **Google Gemini 2.5 Flash** (multimodal vision model).
2. **Caches** extracted data in JSON to avoid redundant API calls across reruns.
3. **Parses** free-text missed-attendance claims to identify dates and subjects.
4. **Applies grant rules** — only one subject per date is granted correction.
5. **Builds a final Excel report** with subject-wise and corrected attendance for every student, including debarment flags.

---

## 🛑 Problem Statement

At semester end, students who fall below a certain attendance threshold are "debarred" from exams:

| Condition | Threshold |
|-----------|-----------|
| No medical certificate | **75%** overall attendance required |
| Valid medical certificate provided | **65%** overall attendance required |

Students submitted:
- UMS portal screenshots showing their recorded attendance.
- Claims for dates they believe were incorrectly marked absent.
- Medical certificates (where applicable).
- Event participation certificates.

Manually reconciling all of this for an entire batch was the bottleneck this system eliminates.

---

## 🏗️ System Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│  INPUT LAYER                                                         │
│  ┌──────────────────┐  ┌─────────────────────┐  ┌────────────────┐  │
│  │ UMS Screenshots  │  │ Google Form Excel   │  │ attendance.xlsx│  │
│  │ (images/PDFs)    │  │ (Responses Sheet)   │  │ (master list)  │  │
│  └────────┬─────────┘  └──────────┬──────────┘  └───────┬────────┘  │
│           │                       │                      │           │
└───────────┼───────────────────────┼──────────────────────┼───────────┘
            │                       │                      │
            ▼                       ▼                      │
┌─────────────────────────────────────────────────────────┐│
│  OCR ENGINE  (full_rescan.py / reprocess_empty.py)      ││
│  ┌────────────────────────────────────────────────────┐ ││
│  │  Google Gemini 2.5 Flash (Vision)                  │ ││
│  │  Prompt → JSON: { code, name, total, present, %  } │ ││
│  └──────────────────────────┬─────────────────────────┘ ││
│                             │                            ││
│                    ocr_cache_v2.json                     ││
│                    (persistent cache)                    ││
└──────────────────────────────────────────────────────────┘│
                              │                             │
                              ▼                             ▼
            ┌─────────────────────────────────────────────────┐
            │  RECONCILIATION ENGINE  (build_excel_v5.py)     │
            │                                                  │
            │  1. Load cache + master list + form responses    │
            │  2. Parse missed-attendance free text            │
            │     → extract dates & subject keywords          │
            │  3. Apply 1-subject-per-date grant rule          │
            │  4. Apply manual admin overrides                 │
            │  5. Compute corrected attendance per subject     │
            │  6. Compute corrected overall %                  │
            └──────────────────────┬───────────────────────────┘
                                   │
                                   ▼
                   ┌───────────────────────────────┐
                   │  OUTPUT: Attendance_Matrix_    │
                   │  FINAL_v6.xlsx                 │
                   │  (color-coded, audit-ready)    │
                   └───────────────────────────────┘
```

---

## 📁 Project Structure

```
attendance/
│
├── 📄 README.md                           ← This file
│
├── ── SCRIPTS ─────────────────────────────────────────────
│
├── 🔧 full_rescan.py                      ← STEP 1: OCR all UMS screenshots → ocr_cache_v2.json
├── 🔧 reprocess_empty.py                  ← STEP 1b: Re-OCR students with empty/missing cache entries
├── 🔧 attendance_engine.py                ← (Legacy v1) OCR + engine combined; date-wise extraction
├── 🔧 subject_extractor.py                ← (Legacy v1) Builds subject-wise matrix from ocr_cache.json
├── 🔧 build_excel_v4.py                   ← (Legacy v4) Earlier version of Excel builder
├── 🔧 build_excel_v5.py                   ← STEP 2: Final Excel builder (current, production-grade)
├── 🔧 fill_attendance_format.py           ← (Legacy v3) Fills T-/A-/P- columns into attendance.xlsx format
│
├── 🔍 check_cache.py                      ← Utility: inspect/audit the OCR cache
├── 🔍 inspect_events.py                   ← Utility: inspect event participation data in form responses
├── 🔍 inspect_form.py                     ← Utility: inspect raw form response columns
├── 🔍 test_api.py                         ← Utility: test Gemini API connectivity
│
├── ── CACHE ───────────────────────────────────────────────
│
├── 📦 ocr_cache.json                      ← (Legacy v1) Cache: subject_wise + date_wise format
├── 📦 ocr_cache_v2.json                   ← (Current) Cache: flat subjects list per student
│
├── ── DATA FILES ──────────────────────────────────────────
│
├── 📊 attendance.xlsx                     ← Master student list (Sl NO, Name, Roll Number)
├── 📊 B.Tech 4th Semester Attendance Collection (Debarred List) (Responses) (1).xlsx
│                                          ← Google Form responses (claims, medical, events)
│
├── ── OUTPUT FILES (versioned) ────────────────────────────
│
├── 📗 Attendance_Matrix_FINAL.xlsx        ← v1 output (T/A/P columns only)
├── 📗 Attendance_Matrix_FINAL_v2.xlsx     ← v2 output
├── 📗 Attendance_Matrix_FINAL_v3.xlsx     ← v3 output
├── 📗 Attendance_Matrix_FINAL_v4.xlsx     ← v4 output
├── 📗 Attendance_Matrix_FINAL_v5.xlsx     ← v5 output
├── 📗 Attendance_Matrix_FINAL_v6.xlsx     ← ✅ FINAL output (current production)
├── 📗 Final_Subject_Wise_Attendance.xlsx  ← Intermediate output from fill_attendance_format.py
├── 📗 Subject_Wise_Attendance_Matrix.xlsx ← Intermediate output from subject_extractor.py
│
└── ── GOOGLE FORM UPLOAD FOLDERS ──────────────────────────
    │
    ├── 📁 Upload UMS Attendance Screenshot-Report (Mandatory) (File responses)/
    │       Student-uploaded UMS screenshots (images/PDFs)
    │
    ├── 📁 Any attendance missed but you are present_ .../
    │       Supporting screenshots for missed attendance claims
    │
    └── 📁 Events participation details with date (certificates-proof) (File responses)/
            Event participation certificate uploads
```

---

## 🗄️ Data Sources

### 1. Google Form Responses Excel
**File:** `B.Tech 4th Semester Attendance Collection (Debarred List) (Responses) (1).xlsx`

The form collected the following columns (used by the pipeline):

| Column | Description |
|--------|-------------|
| `Student's Name (As per University Records)` | Official student name (key for matching) |
| `Student's University Roll Number` | Roll number (key for lookup) |
| `Upload UMS Attendance Screenshot/Report (Mandatory)` | Filename of the uploaded screenshot |
| `Medical certificate provided?` | Yes / No |
| `Medical certificate range written in certificate` | Date range on medical cert |
| `Any attendance missed but you are present?` | Free-text claim of wrongly-marked absences |
| `Events participation details with date` | Free-text event participation info |

---

### 2. Master Student List
**File:** `attendance.xlsx`

A clean roster with columns:
- `Sl NO` — serial number
- `Name` — student name
- `Roll Number` — university roll number

This is the authoritative source for generating rows in the output sheet.

---

### 3. UMS Screenshots
**Folder:** `Upload UMS Attendance Screenshot-Report (Mandatory) (File responses)/`

Each student uploads a screenshot from their UMS portal showing:
- Course Code | Course Name | Total Classes | Present | Absent | Percentage

Supported formats: `.jpg`, `.jpeg`, `.png`, `.webp`, `.gif`, `.bmp`, `.pdf`

---

## 💻 Scripts — Detailed Reference

---

### `full_rescan.py` — Primary OCR Script *(Current)*

**Purpose:** Scans every student's UMS screenshot using Gemini Vision and populates `ocr_cache_v2.json`.

**Key behaviors:**
- Loads existing cache; **skips** students already cached with ≥1 subject.
- Matches screenshots by the exact filename from the Google Form responses column.
- Handles rate limits with exponential back-off: waits 70s → 70s → 120s on 429 errors.
- Enforces a 4-second sleep between students to stay under the 15 RPM Gemini free tier limit.
- Applies a **sanity check**: if `present > total`, clamps `present = total`.
- After OCR, auto-invokes `build_excel_v2.py` to regenerate the Excel.

**Cache format (v2):**
```json
{
  "Student Name": [
    {
      "code": "CSE11111",
      "name": "Formal Language and Automata",
      "total": 39,
      "present": 32,
      "percentage": 82.05
    },
    ...
  ]
}
```

**Prompt strategy:** The Gemini prompt is deliberately strict — it instructs the model to read *only* what is printed on screen, not to calculate, and to return `0` for unclear cells. It differentiates theory subjects (`CSE11xxx`, `MTH11xxx`) from lab subjects (`CSE12xxx`, `MTH12xxx`).

---

### `reprocess_empty.py` — Re-OCR Fallback Script

**Purpose:** Targets only students whose cache entry is empty (`subject_wise: []`) and re-runs OCR on them. Useful after the initial scan if some files were missing or the API failed.

**Differences from `full_rescan.py`:**
- Uses the legacy v1 cache format (subject_wise + date_wise).
- Also extracts date-wise presence/absence logs if visible in the screenshot.
- After completion, auto-invokes `fill_attendance_format.py`.

---

### `build_excel_v5.py` — Final Excel Builder *(Current Production)*

**Purpose:** Reads `ocr_cache_v2.json`, applies all reconciliation logic, and writes `Attendance_Matrix_FINAL_v6.xlsx`.

**Processing pipeline inside this script:**

```
For each student in attendance.xlsx:
  1. Load subject data from ocr_cache_v2.json (fuzzy name match)
  2. Load medical info from form responses (keyed by roll number)
  3. Load missed-attendance claim text (keyed by roll number)
  4. Load event participation text (keyed by roll number)
  5. Parse claim text → extract (date, subject_code) pairs
  6. Apply 1-subject/date grant rule → granted_codes dict
  7. Apply manual admin overrides (MANUAL_OVERRIDES dict)
  8. For each of 12 subjects:
       corrected_attended = min(OCR_attended + granted, total)
       corrected_pct = corrected_attended / total × 100
  9. overall_corrected_pct = Σ corrected_attended / Σ total × 100
 10. Write row to Excel with all T-/A-/P-/CorrA/CorrP columns
```

**Excel layout:**

| Fixed Cols | Subject Blocks (×12) | Tail Cols |
|------------|----------------------|-----------|
| Sl NO, Name, Roll Number | T- / A- / P- / Corr A / Corr P | Overall Corr % / Medical? / Medical Dates / Claimed Missed / Event Details |

---

### `attendance_engine.py` — Legacy v1 Engine

**Purpose:** Earlier all-in-one script that performed OCR using a multi-image approach (sent all files for a student in one API call), then applied cross-date correction logic (if present in any subject on a date, grant absent siblings).

> ⚠️ Superseded by `full_rescan.py` + `build_excel_v5.py`. Retained for reference.

---

### `subject_extractor.py` — Legacy Subject Matrix Builder

**Purpose:** Reads the v1 OCR cache and builds a flat T-/A-/P- matrix. Predecessor to `fill_attendance_format.py`.

> ⚠️ Superseded. Retained for reference.

---

### `fill_attendance_format.py` — Legacy v3 Excel Formatter

**Purpose:** Takes the v1 cache, matches students to the master attendance.xlsx, and produces a color-coded multi-header Excel with T-/A-/P- columns per subject.

> ⚠️ Still called by `reprocess_empty.py` for the legacy cache path.

---

### Utility Scripts

| Script | Purpose |
|--------|---------|
| `check_cache.py` | Print cache summary (how many students, subjects per student) |
| `inspect_form.py` | Print all column names in the Google Form response Excel |
| `inspect_events.py` | Print raw event participation data per student |
| `test_api.py` | Sends a minimal test request to the Gemini API to verify key validity |

---

## 🧠 Business Logic

### 1. Missed Attendance Grant Rule

Students may claim: *"On date X, I was absent in Subject Y but was present in another class."*

The pipeline parses this free text and applies:

```
IF a date segment mentions exactly 1 subject → GRANT that date for that subject (+1 class)
IF a date segment mentions 2+ subjects     → REJECT (ambiguous, cannot verify)
```

**Date extraction supports:**
- Numeric formats: `9/4`, `09/04/25`, `9-4-2025`, `9.4`
- English ordinal formats: `9th April`, `27th March`, `16th January`

**Subject extraction uses keyword matching:**

| Subject Code | Keywords matched |
|-------------|-----------------|
| CSE11111 | flat, formal language, automata, fla |
| CSE11110 | daa, design and analysis of algo, algorithms |
| PSG11021 | ethics, human values, hvpe, professional ethics |
| CSE11109 | oop, object oriented, oops, java |
| MTH11534 | discrete, discrete math, dsl, discrete structures |
| CSE11112 | ai, intro to ai, artificial intelligence |
| CSE11204 | eda, exploratory data analysis, data analysis |
| CSE12205 | eda lab, exploratory data analysis lab |
| CSE12166 | daa lab, algorithms lab |
| CSE12114 | oop lab, oops lab, java lab |
| MTH12531 | nt lab, numerical techniques lab, numerical lab |
| CSE14170 | mini project, project |

---

### 2. Debarment Threshold

Computed on *corrected* overall percentage:

```
corrected_overall_pct = Σ corrected_attended / Σ total_classes × 100

IF medical cert provided:  threshold = 65%
ELSE:                      threshold = 75%

status = "DEBARRED" if corrected_overall_pct < threshold else "SAFE"
```

> Note: The final Excel (v6) shows **Overall Corr %** color-coded but does not include a separate DEBARRED column — the color makes status immediately visible.

---

### 3. Correction Cap

Granted corrections **never exceed** the total number of classes held:

```python
corrected_attended = min(OCR_attended + granted, total_classes)
```

---

### 4. Manual Admin Overrides

Hardcoded in `build_excel_v5.py`:

```python
MANUAL_OVERRIDES = {
    "Shubhajit Mandal": {
        "CSE12114": 6,   # OOP Lab — manually granted by admin
    },
}
```

Add entries here when an admin decision overrides the automated grant logic.

---

## 📊 Output Files

### `Attendance_Matrix_FINAL_v6.xlsx` *(Primary Output)*

**Sheet:** `Subject-wise Attendance`

| Column Group | Columns | Description |
|---|---|---|
| Identity | Sl NO, Name, Roll Number | Student identifiers |
| Per Subject (×12) | T-, A-, P-, Corr A, Corr P | Total / Attended (OCR) / Percentage / Corrected Attended / Corrected Percentage |
| Overall | Overall Corr % | Weighted corrected percentage across all subjects |
| Medical | Medical?, Medical Dates | Whether cert was submitted and the date range |
| Claims | Attendance Claimed Missed | Raw text from form (1-subject/date rule applied) |
| Events | Event Participation Details | Raw event data (informational; no attendance granted) |

**Freeze panes:** Row 1–2 (headers) and Column A–C (identity) are frozen.

---

## 🎨 Color Coding Reference

### Percentage Cells (P- and Corr P-)

| Color | Meaning | Threshold |
|-------|---------|-----------|
| 🟢 Green (`#C6EFCE`) | Safe — no action needed | ≥ 75% |
| 🟡 Yellow (`#FFEB9C`) | Warning — within medical grace zone | 65% – 74.9% |
| 🔴 Red (`#FFC7CE`) | Debarred zone | < 65% |

### Header Colors

| Color | Purpose |
|-------|---------|
| `#1F4E79` Dark Blue | Main identity columns header |
| `#2E75B6` Medium Blue | Subject name header row |
| `#375623` Dark Green | Overall Corrected % header |
| `#FFF2CC` Yellow | Medical columns |
| `#FCE4D6` Peach | Missed attendance claim column |
| `#DDEBF7` Light Blue | Event participation column |

---

## 🚀 Setup & Installation

### Prerequisites

- Python 3.9+
- A valid **Google Gemini API Key** (Gemini 2.5 Flash with vision capability)

### Install Dependencies

```bash
pip install pandas openpyxl pillow google-generativeai
```

Or install from a requirements file if available:

```bash
pip install -r requirements.txt
```

### Configure API Key

The scripts read the Gemini API key from an environment variable. Do not hardcode your key in the scripts.

1. Create a `.env` file in the root directory (this is ignored by Git).
2. Add your API key:
   ```env
   GEMINI_API_KEY=your_actual_api_key_here
   ```

*(Alternatively, you can set it directly in your terminal: `$env:GEMINI_API_KEY="your_api_key"` for PowerShell, or `export GEMINI_API_KEY="your_api_key"` for bash/zsh)*

---

## 🏃 Running the Pipeline

### Step 1 — Run OCR to populate the cache

```bash
python full_rescan.py
```

This will:
- Process each student whose entry is not yet in `ocr_cache_v2.json`
- Print progress to console
- Save cache incrementally after each student
- Auto-generate the Excel at the end

**Typical run time:** ~4–8 seconds per student (API call + rate-limit sleep).

---

### Step 1b — Re-OCR students with missing data (if needed)

```bash
python reprocess_empty.py
```

Run this if some students still show 0 subjects after Step 1.

---

### Step 2 — Regenerate Excel (standalone, uses existing cache)

```bash
python build_excel_v5.py
```

This skips OCR entirely and just rebuilds the Excel from the cached data. Use this when:
- You've made changes to `MANUAL_OVERRIDES`
- You've corrected the cache manually
- You want to regenerate after fixing subject keyword rules

---

### Inspecting the Cache

```bash
python check_cache.py
```

---

## ⚙️ Configuration & Manual Overrides

### Adding Manual Attendance Grants

In `build_excel_v5.py`, find and edit:

```python
MANUAL_OVERRIDES = {
    "Student Full Name (as in attendance.xlsx)": {
        "SUBJECT_CODE": number_of_extra_classes_to_grant,
    },
}
```

**Example:** Grant 3 extra classes in DAA Lab for a student:
```python
MANUAL_OVERRIDES = {
    "Shubhajit Mandal": {
        "CSE12114": 6,
    },
    "Another Student": {
        "CSE12166": 3,  # DAA Lab
    },
}
```

### Adding Subject Keywords

In `build_excel_v5.py`, find `SUBJ_KEYWORDS` and add entries:

```python
"CSE11111": ["flat", "formal language", "automata", "fla", "your-new-keyword"],
```

---

## 📚 Subjects Reference

| Subject Code | Short Name | Full Name |
|-------------|-----------|-----------|
| CSE11111 | FLAT | Formal Language and Automata |
| CSE11110 | DAA | Design and Analysis of Algorithms |
| PSG11021 | Ethics | Human Values and Professional Ethics |
| CSE11109 | OOP | Object Oriented Programming |
| MTH11534 | Discrete | Discrete Structures and Logic |
| CSE11112 | AI | Introduction to Artificial Intelligence |
| CSE11204 | EDA | Exploratory Data Analysis |
| CSE12205 | EDA Lab | Exploratory Data Analysis Lab |
| CSE12166 | DAA Lab | Design and Analysis of Algorithms Lab |
| CSE12114 | OOP Lab | Object Oriented Programming Lab |
| MTH12531 | NT Lab | Numerical Techniques Lab |
| CSE14170 | Mini Project | Mini Project-I |

---

## 💾 Caching Architecture

The pipeline uses a **two-version cache** system:

| File | Format | Used By |
|------|--------|---------|
| `ocr_cache.json` | `{name: {subject_wise: [...], date_wise: [...]}}` | Legacy scripts (v1) |
| `ocr_cache_v2.json` | `{name: [{code, name, total, present, percentage}]}` | Current scripts (v2) |

**Why incremental caching matters:**
- Gemini API has a free-tier limit of **15 requests per minute**.
- A full batch of ~50–60 students would take **~5 minutes** even under optimal conditions.
- The cache ensures interrupted runs resume without re-spending API quota.
- Cache is written after **every single student** to minimize loss on crashes.

---

## 🔧 Troubleshooting

### "FILE NOT FOUND" for a student

**Cause:** The filename stored in the Google Form responses column doesn't match any file in the screenshots folder.

**Fix:** Manually check the responses Excel for that student's filename and verify it exists in the folder. You can also manually add a cache entry:

```python
# In check_cache.py or a quick script:
import json
with open("ocr_cache_v2.json") as f:
    cache = json.load(f)

cache["Student Name"] = [
    {"code": "CSE11111", "name": "FLAT", "total": 40, "present": 35, "percentage": 87.5},
    # ... add all subjects
]

with open("ocr_cache_v2.json", "w") as f:
    json.dump(cache, f, indent=2)
```

---

### Rate limit errors (429)

The scripts have built-in retry logic. If you see repeated 429 errors:
- The free-tier quota (15 RPM) may be exhausted.
- Wait 1–2 minutes and restart — the cache ensures no data is lost.

---

### Subject shows 0/0 for a student

**Causes:**
1. OCR failed to extract that subject (screenshot was blurry/cropped).
2. Subject code wasn't visible in the screenshot.
3. The student's screenshot only shows a partial view.

**Fix:** Add a manual override in `MANUAL_OVERRIDES` after verifying the student's actual UMS data.

---

### Excel opens with encoding errors

The scripts use `sys.stdout.reconfigure(encoding='utf-8')`. If running on Windows PowerShell and seeing garbled output, run:
```powershell
$env:PYTHONIOENCODING = "utf-8"
python build_excel_v5.py
```

---

## 📦 Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| `pandas` | ≥ 1.5 | Reading Excel files, data manipulation |
| `openpyxl` | ≥ 3.1 | Writing formatted Excel output |
| `pillow` | ≥ 9.0 | Opening image files for Gemini Vision |
| `google-generativeai` | ≥ 0.5 | Gemini API client (OCR + JSON extraction) |
| `re` | stdlib | Regex for date/subject parsing |
| `json` | stdlib | Cache serialization |
| `collections` | stdlib | `defaultdict` for grouping |

---

## 🕒 Version History

| Version | Script | Key Change |
|---------|--------|-----------|
| v1 | `attendance_engine.py` | Initial: multi-image OCR + date-wise logic |
| v1b | `subject_extractor.py` | Subject matrix builder from v1 cache |
| v2 | `fill_attendance_format.py` | Proper Excel formatting with T-/A-/P- |
| v3 | `build_excel_v4.py` | Added corrected attendance columns |
| v4 | `full_rescan.py` + `build_excel_v5.py` | Switched to v2 cache, strict OCR prompt, grant rules, medical, events |
| **v5 (current)** | `build_excel_v5.py` | Added manual overrides, 5-column-per-subject layout, tail columns |

---

## ✍️ Author Notes

- This system was built iteratively across a single working session to meet an institutional deadline.
- The OCR accuracy depends entirely on screenshot quality — blurry, cropped, or rotated images will produce 0s.
- Event participation data is recorded in the output but **does not grant any attendance** — this is by institutional policy.
- The `MANUAL_OVERRIDES` mechanism allows an admin to apply decisions that cannot be derived from screenshots alone.

---

*Generated for B.Tech 4th Semester — Department of Computer Science & Engineering*
