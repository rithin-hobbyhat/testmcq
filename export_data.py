"""
Export questions from Excel to optimized JSON files for the Practice Web App.
Extracts all questions, normalizes engineering, physics, scientific and mathematical
symbols into standard notation (Greek letters, delta, equations, radicals, etc.),
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
    """Transform raw LaTeX commands, engineering and physics symbols into standard typography."""
    if not text or not isinstance(text, str):
        return text if text is not None else ''

    # 1. Degree & temperature
    text = re.sub(r'\^?\\circ\\text\{C\}|\^?\\circ\s*C|\\circ\s*C', '°C', text)
    text = re.sub(r'\^?\\circ\\text\{F\}|\^?\\circ\s*F|\\circ\s*F', '°F', text)
    text = re.sub(r'\^?\\circ|\\circ|\\degree', '°', text)

    # 2. Standard LaTeX symbols
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

    # 3. Delta variants (Engineering pressure drops, temperature, elongation, boundary layer, electrical delta)
    text = re.sub(r'\bdelta_p_([a-zA-Z0-9]+)\b', r'Δp_\1', text)
    text = re.sub(r'\bdelta_p\b', 'Δp', text, flags=re.I)
    text = re.sub(r'\bdelta_T\b', 'ΔT', text)
    text = re.sub(r'\bdelta_t\b', 'Δt', text)
    text = re.sub(r'\bdelta_L([0-9]*)\b', r'ΔL\1', text)
    text = re.sub(r'\bdelta_V\b', 'ΔV', text)
    text = re.sub(r'\bdelta_h\b', 'Δh', text)
    text = re.sub(r'\bdelta_taper\b', 'δ_taper', text)
    text = re.sub(r'\bdelta_uniform\b', 'δ_uniform', text)
    text = re.sub(r'\bdelta_e\b', 'δₑ', text)
    text = re.sub(r'\bDelta\s+omega\b', 'Δω', text)
    text = re.sub(r'\bDelta-ferrite\b', 'δ-ferrite', text)
    text = re.sub(r'\bAlpha-ferrite\b', 'α-ferrite', text)
    text = re.sub(r'\bGamma-iron\b', 'γ-iron', text)
    
    # Delta connection in electrical stator/transformers
    text = re.sub(r'\bDelta\s*\(Delta\)', 'Δ (Delta)', text)
    text = re.sub(r'\bDelta-connected\b', 'Δ-connected', text)
    text = re.sub(r'\bDelta\s+stator\b', 'Δ stator', text)
    text = re.sub(r'\bDelta\s+connection\b', 'Δ connection', text)
    text = re.sub(r'\bDelta\s+winding\b', 'Δ winding', text)

    # Displacement thickness & boundary layer delta: delta*, delta
    text = re.sub(r'\bdelta\*', 'δ*', text)
    text = re.sub(r'\bdelta\b(?!\s*connec|\s*stator|\s*wind|\s*loop)', 'δ', text)

    # 4. Greek engineering & physics variables
    text = re.sub(r'\bsigma_([a-zA-Z0-9]+)\b', r'σ_\1', text)
    text = re.sub(r'\bsigma\b', 'σ', text)
    text = re.sub(r'\brho_([a-zA-Z0-9]+)\b', r'ρ_\1', text)
    text = re.sub(r'\brho\b', 'ρ', text)
    text = re.sub(r'\btheta\b', 'θ', text)
    text = re.sub(r'\btau_([a-zA-Z0-9]+)\b', r'τ_\1', text)
    text = re.sub(r'\btau\b', 'τ', text)
    text = re.sub(r'\bomega_([a-zA-Z0-9]+)\b', r'ω_\1', text)
    text = re.sub(r'\bomega\b', 'ω', text)
    text = re.sub(r'\bgamma_([a-zA-Z0-9]+)\b', r'γ_\1', text)
    text = re.sub(r'\bgamma\b', 'γ', text)
    text = re.sub(r'\blambda_([a-zA-Z0-9]+)\b', r'λ_\1', text)
    text = re.sub(r'\blambda\b', 'λ', text)
    text = re.sub(r'\bphi\b', 'ϕ', text)
    
    text = re.sub(r'\bmu_([a-zA-Z0-9]+)\b', r'µ_\1', text)
    text = re.sub(r'(?<=[0-9\s\*\+\-\/\(\[\=])mu(?=[0-9\s\*\+\-\/\)\]\,]|$)', 'µ', text)
    text = re.sub(r'\bmu\b', 'µ', text)

    text = re.sub(r'(?<=[0-9\s\*\+\-\/\(\[\=])pi(?=[0-9\s\*\+\-\/\)\]\,]|$)', 'π', text)
    text = re.sub(r'\bpi\b(?!\s*pipe|\s*pin|\s*pilot|\s*piston|\s*pitot)', 'π', text)

    text = re.sub(r'(?<=[0-9\s\*\+\-\/\(\[\=])nu(?=[0-9\s\*\+\-\/\)\]\,]|$)', 'ν', text)
    text = re.sub(r'\bnu\b(?!\s*Unit|\s*number)', 'ν', text)

    text = re.sub(r'\bepsilon_([a-zA-Z0-9]+)\b', r'ε_\1', text)
    text = re.sub(r'\bepsilon\b', 'ε', text)

    text = re.sub(r'\beta_([a-zA-Z0-9]+)\b', r'η_\1', text)
    text = re.sub(r'\beta\b(?!\s*carr|\s*hour)', 'η', text)

    text = re.sub(r'\balpha([0-9]*)\b', r'α\1', text)
    text = re.sub(r'\balpha_([a-zA-Z0-9]+)\b', r'α_\1', text)

    text = re.sub(r'\bbeta([0-9]*)\b', r'β\1', text)

    text = re.sub(r'\bpsi(?=\s*[\*\/=])', 'ψ', text)
    text = re.sub(r'=\s*psi\b', '= ψ', text)
    text = re.sub(r'\(psi\)', '(ψ)', text)

    # 5. Square roots: sqrt(...) -> √( ... )
    text = re.sub(r'\bsqrt\(([^()]+)\)', r'√(\1)', text)
    text = re.sub(r'\bsqrt\(([^\)]+)\)', r'√(\1)', text)

    # 6. Integrals
    text = re.sub(r'integral from 0 to ([a-zA-ZδΔ]+) of', r'∫₀^\1', text)
    text = re.sub(r'integral from ([a-zA-Z0-9]+) to ([a-zA-Z0-9]+) of', r'∫_{\1}^{\2}', text)

    # 7. Approximations
    text = re.sub(r'\bapprox\.?\b', '≈', text)

    # 8. Chemical formulas
    text = re.sub(r'\\text\{PbO\}_2', 'PbO₂', text)
    text = re.sub(r'\\text\{H\}_2\\text\{SO\}_4', 'H₂SO₄', text)
    text = re.sub(r'\\text\{PbSO\}_4', 'PbSO₄', text)
    text = re.sub(r'\\text\{CO\}_2', 'CO₂', text)
    text = re.sub(r'\\text\{O\}_2', 'O₂', text)
    text = re.sub(r'\\text\{NO\}_\\text\{x\}', 'NOₓ', text)
    text = re.sub(r'\\text\{NO\}_x', 'NOₓ', text)

    # 9. Unwrap \text{...}
    text = re.sub(r'\\text\{([^}]*)\}', r'\1', text)

    # 10. Clean up escaped backslashes and percentages
    text = text.replace(r'\_', '_').replace(r'\ ', ' ').replace(r'\%', '%')

    # 11. Engineering sub-indices: L1 -> L₁, A1 -> A₁, etc.
    sub_num = {'1':'₁','2':'₂','3':'₃','4':'₄'}
    text = re.sub(r'\b([yLAdEPVTrFrασ])[1-4]\b', lambda m: m.group(1) + sub_num[m.group(0)[-1]], text)
    text = text.replace('_max', 'ₘₐₓ').replace('_min', 'ₘᵢₙ').replace('_ph', 'ₚₕ')

    # 12. General subscripts
    sub_map = {'0':'₀','1':'₁','2':'₂','3':'₃','4':'₄','5':'₅','6':'₆','7':'₇','8':'₈','9':'₉','x':'ₓ','n':'ₙ','m':'ₘ','v':'ᵥ','p':'ₚ','t':'ₜ'}
    text = re.sub(r'_([0-9xnmvpt])\b', lambda m: sub_map.get(m.group(1), m.group(0)), text)
    text = re.sub(r'_\{([0-9xnmvpt]+)\}', lambda m: ''.join(sub_map.get(c, c) for c in m.group(1)), text)

    # 13. Superscripts
    sup_map = {'0':'⁰','1':'¹','2':'²','3':'³','4':'⁴','5':'⁵','6':'⁶','7':'⁷','8':'⁸','9':'⁹','n':'ⁿ','x':'ˣ','st':'ˢᵗ','nd':'ⁿᵈ','rd':'ʳᵈ','th':'ᵗʰ'}
    text = re.sub(r'\^([0-9nx])\b', lambda m: sup_map.get(m.group(1), m.group(0)), text)
    text = re.sub(r'\^\\text\{([a-z]+)\}', lambda m: sup_map.get(m.group(1), m.group(1)), text)
    text = re.sub(r'\^\{([0-9]+)\}', lambda m: ''.join(sup_map.get(c, c) for c in m.group(1)), text)

    # 14. Multiplication in equations: 2 * pi -> 2 · π
    text = re.sub(r'(\b[a-zA-Z0-9α-ωΑ-ΩΔδσμτρθηεγλπψων²³⁴⁰¹ⁿ]+)\s*\*\s*([a-zA-Z0-9α-ωΑ-ΩΔδσμτρθηεγλπψων²³⁴⁰¹ⁿ√])', r'\1 · \2', text)
    text = re.sub(r'(\))\s*\*\s*(\()', r'\1 · \2', text)

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
