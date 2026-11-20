"""
Export questions from Excel to optimized JSON files for the Practice Web App.
Extracts all questions, normalizes scientific and mathematical symbols,
generates subjects catalog, and outputs both consolidated and per-subject JSON files.
"""
import os
import json
import re
import openpyxl

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "data", "questions_bank.xlsx")
WEB_DATA_DIR = os.path.join(BASE_DIR, "web", "data")
DOCS_DATA_DIR = os.path.join(BASE_DIR, "docs", "data")

def sanitize_slug(text: str) -> str:
    """Generate safe filename slug from subject name or code."""
    slug = re.sub(r'[^a-zA-Z0-9_-]', '_', text.strip()).lower()
    return re.sub(r'_+', '_', slug).strip('_')

def clean_scientific_math(text: str) -> str:
    """Transform raw LaTeX commands and math artifacts into standard Unicode notation."""
    if not text or not isinstance(text, str):
        return text if text is not None else ''

    # 1. Degree & temperature
    text = re.sub(r'\^?\\circ\\text\{C\}|\^?\\circ\s*C|\\circ\s*C', '°C', text)
    text = re.sub(r'\^?\\circ\\text\{F\}|\^?\\circ\s*F|\\circ\s*F', '°F', text)
    text = re.sub(r'\^?\\circ|\\circ|\\degree', '°', text)

    # 2. Standard symbols
    replacements = [
        (r'\\times(?![a-zA-Z])', '×'),
        (r'\\pm(?![a-zA-Z])', '±'),
        (r'\\mu(?![a-zA-Z])', 'µ'),
        (r'\\Omega(?![a-zA-Z])', 'Ω'),
        (r'\\approx(?![a-zA-Z])', '≈'),
        (r'\\leq(?![a-zA-Z])', '≤'),
        (r'\\le(?![a-zA-Z])', '≤'),
        (r'\\geq(?![a-zA-Z])', '≥'),
        (r'\\ge(?![a-zA-Z])', '≥'),
        (r'\\neq(?![a-zA-Z])', '≠'),
        (r'\\ne(?![a-zA-Z])', '≠'),
        (r'\\Delta(?![a-zA-Z])', 'Δ'),
        (r'\\alpha(?![a-zA-Z])', 'α'),
        (r'\\beta(?![a-zA-Z])', 'β'),
        (r'\\gamma(?![a-zA-Z])', 'γ'),
        (r'\\theta(?![a-zA-Z])', 'θ'),
        (r'\\pi(?![a-zA-Z])', 'π'),
        (r'\\sim(?![a-zA-Z])', '~'),
    ]
    for pat, rep in replacements:
        text = re.sub(pat, rep, text)

    # 3. Subscripts
    sub_map = {'0':'₀','1':'₁','2':'₂','3':'₃','4':'₄','5':'₅','6':'₆','7':'₇','8':'₈','9':'₉','x':'ₓ','n':'ₙ','m':'ₘ'}
    text = re.sub(r'_([0-9xnm])\b', lambda m: sub_map.get(m.group(1), m.group(0)), text)
    text = re.sub(r'_\{([0-9xnm]+)\}', lambda m: ''.join(sub_map.get(c, c) for c in m.group(1)), text)

    # 4. Superscripts
    sup_map = {'0':'⁰','1':'¹','2':'²','3':'³','4':'⁴','5':'⁵','6':'⁶','7':'⁷','8':'⁸','9':'⁹','n':'ⁿ','x':'ˣ','st':'ˢᵗ','nd':'ⁿᵈ','rd':'ʳᵈ','th':'ᵗʰ'}
    text = re.sub(r'\^([0-9nx])\b', lambda m: sup_map.get(m.group(1), m.group(0)), text)
    text = re.sub(r'\^\\text\{([a-z]+)\}', lambda m: sup_map.get(m.group(1), m.group(1)), text)
    text = re.sub(r'\^\{([0-9]+)\}', lambda m: ''.join(sup_map.get(c, c) for c in m.group(1)), text)

    # 5. Chemical formulas with common LaTeX wrappers
    text = re.sub(r'\\text\{PbO\}_2', 'PbO₂', text)
    text = re.sub(r'\\text\{H\}_2\\text\{SO\}_4', 'H₂SO₄', text)
    text = re.sub(r'\\text\{PbSO\}_4', 'PbSO₄', text)
    text = re.sub(r'\\text\{CO\}_2', 'CO₂', text)
    text = re.sub(r'\\text\{O\}_2', 'O₂', text)
    text = re.sub(r'\\text\{NO\}_\\text\{x\}', 'NOₓ', text)
    text = re.sub(r'\\text\{NO\}_x', 'NOₓ', text)

    # 6. Unwrap \text{...}
    text = re.sub(r'\\text\{([^}]*)\}', r'\1', text)

    # 7. Clean up escaped backslashes and percentages
    text = text.replace(r'\_', '_').replace(r'\ ', ' ').replace(r'\%', '%')
    return text.strip()

def export_all():
    for target_dir in [WEB_DATA_DIR, DOCS_DATA_DIR]:
        os.makedirs(os.path.join(target_dir, "by_subject"), exist_ok=True)

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
        
        q_text = clean_scientific_math(str(row[3] or ''))
        if not q_text:
            continue

        q_id = str(row[0] or f"Q{row_idx:04d}").strip()
        subj_name = str(row[1] or 'General').strip().upper()
        subj_code = str(row[2] or subj_name[:4]).strip().upper()
        
        opt_a = clean_scientific_math(str(row[4] or ''))
        opt_b = clean_scientific_math(str(row[5] or ''))
        opt_c = clean_scientific_math(str(row[6] or ''))
        opt_d = clean_scientific_math(str(row[7] or ''))
        
        ans = str(row[8] or 'A').strip().upper()
        if ans not in ['A', 'B', 'C', 'D']:
            ans = 'A'

        exp = clean_scientific_math(str(row[9] or ''))
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

    for target_dir in [WEB_DATA_DIR, DOCS_DATA_DIR]:
        subjects_file = os.path.join(target_dir, "subjects.json")
        with open(subjects_file, "w", encoding="utf-8") as f:
            json.dump(catalog, f, ensure_ascii=False, indent=2)

        # 2. Write individual subject JSON files
        by_subj_dir = os.path.join(target_dir, "by_subject")
        for s_name, data in subjects_dict.items():
            file_path = os.path.join(by_subj_dir, f"{data['slug']}.json")
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(data["questions"], f, ensure_ascii=False)

        # 3. Write consolidated questions.json
        all_file = os.path.join(target_dir, "questions.json")
        with open(all_file, "w", encoding="utf-8") as f:
            json.dump(questions, f, ensure_ascii=False)
        print(f"Saved dataset in {target_dir}: {len(questions)} questions ({os.path.getsize(all_file) / (1024*1024):.2f} MB)")

    return len(questions), len(catalog)

if __name__ == "__main__":
    export_all()
