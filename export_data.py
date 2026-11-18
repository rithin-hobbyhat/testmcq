"""
Export questions from Excel to optimized JSON files for the Practice Web App.
Extracts all 5,700+ questions, generates subjects catalog, and outputs
both consolidated and per-subject JSON files for ultra-fast mobile loading.
"""
import os
import json
import re
import openpyxl

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "data", "questions_bank.xlsx")
WEB_DATA_DIR = os.path.join(BASE_DIR, "web", "data")
BY_SUBJECT_DIR = os.path.join(WEB_DATA_DIR, "by_subject")

def sanitize_slug(text: str) -> str:
    """Generate safe filename slug from subject name or code."""
    slug = re.sub(r'[^a-zA-Z0-9_-]', '_', text.strip()).lower()
    return re.sub(r'_+', '_', slug).strip('_')

def export_all():
    os.makedirs(BY_SUBJECT_DIR, exist_ok=True)

    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(f"Excel file not found at: {EXCEL_PATH}")

    print(f"Loading questions from: {EXCEL_PATH}...")
    wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True, data_only=True)
    ws = wb['All Questions'] if 'All Questions' in wb.sheetnames else wb.active

    questions = []
    subjects_dict = {}

    for row_idx, row in enumerate(ws.iter_rows(values_only=True, min_row=2), start=2):
        if not row or not any(row):
            continue
        
        q_text = str(row[3] or '').strip()
        if not q_text:
            continue

        q_id = str(row[0] or f"Q{row_idx:04d}").strip()
        subj_name = str(row[1] or 'General').strip().upper()
        subj_code = str(row[2] or subj_name[:4]).strip().upper()
        
        opt_a = str(row[4] or '').strip()
        opt_b = str(row[5] or '').strip()
        opt_c = str(row[6] or '').strip()
        opt_d = str(row[7] or '').strip()
        
        ans = str(row[8] or 'A').strip().upper()
        if ans not in ['A', 'B', 'C', 'D']:
            ans = 'A'

        exp = str(row[9] or '').strip()
        diff = str(row[10] or 'Medium').strip()

        item = {
            "id": q_id,
            "subject": subj_name,
            "subject_code": subj_code,
            "question": q_text,
            "options": {
                "A": opt_a,
                "B": opt_b,
                "C": opt_c,
                "D": opt_d
            },
            "answer": ans,
            "explanation": exp,
            "difficulty": diff
        }

        questions.append(item)

        if subj_name not in subjects_dict:
            subjects_dict[subj_name] = {
                "name": subj_name,
                "code": subj_code,
                "slug": sanitize_slug(subj_code if subj_code else subj_name),
                "count": 0,
                "questions": []
            }
        
        subjects_dict[subj_name]["count"] += 1
        subjects_dict[subj_name]["questions"].append(item)

    wb.close()
    print(f"Processed {len(questions)} total questions across {len(subjects_dict)} subjects.")

    # 1. Write subjects catalog (metadata only)
    catalog = []
    for s_name, data in sorted(subjects_dict.items(), key=lambda x: x[0]):
        catalog.append({
            "name": data["name"],
            "code": data["code"],
            "slug": data["slug"],
            "count": data["count"]
        })

    subjects_file = os.path.join(WEB_DATA_DIR, "subjects.json")
    with open(subjects_file, "w", encoding="utf-8") as f:
        json.dump(catalog, f, ensure_ascii=False, indent=2)
    print(f"Saved subjects metadata: {subjects_file}")

    # 2. Write individual subject JSON files
    for s_name, data in subjects_dict.items():
        file_path = os.path.join(BY_SUBJECT_DIR, f"{data['slug']}.json")
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(data["questions"], f, ensure_ascii=False)
    print(f"Saved {len(subjects_dict)} individual subject JSON files in: {BY_SUBJECT_DIR}")

    # 3. Write consolidated questions.json (for instant offline PWA caching & fast full queries)
    all_file = os.path.join(WEB_DATA_DIR, "questions.json")
    with open(all_file, "w", encoding="utf-8") as f:
        json.dump(questions, f, ensure_ascii=False)
    print(f"Saved consolidated dataset: {all_file} ({os.path.getsize(all_file) / (1024*1024):.2f} MB)")

    return len(questions), len(catalog)

if __name__ == "__main__":
    export_all()
