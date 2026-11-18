"""
Seed initial sample questions with scientific & math symbols into multi-sheet Excel.
"""
import os
import shutil
import config
from excel_storage import ExcelStorage
from id_generator import SubjectManager, IdGenerator
from models import Question

def seed_sample_questions():
    # Remove old excel to start fresh with clean multi-sheet layout
    if os.path.exists(config.EXCEL_FILE_PATH):
        try:
            os.remove(config.EXCEL_FILE_PATH)
        except Exception:
            pass

    storage = ExcelStorage()
    subject_mgr = SubjectManager()

    sample_data = [
        {
            "subject": "Physics",
            "question": "According to Einstein's mass-energy equivalence formula E = mc², what does 'c' represent?",
            "option_a": "Speed of Sound (≈ 343 m/s)",
            "option_b": "Speed of Light in vacuum (≈ 3 × 10⁸ m/s)",
            "option_c": "Gravitational Constant (G)",
            "option_d": "Planck's Constant (ℏ)",
            "correct": "B",
            "explanation": "In E = mc², 'c' represents the speed of light in vacuum, approximately 299,792,458 m/s (3 × 10⁸ m/s).",
            "difficulty": "Easy"
        },
        {
            "subject": "Physics",
            "question": "What is the electric resistance formula based on Ohm's Law (V = I × R)?",
            "option_a": "R = V / I",
            "option_b": "R = V × I",
            "option_c": "R = I / V",
            "option_d": "R = V² × I",
            "correct": "A",
            "explanation": "From V = I × R, rearranging gives Resistance R = V / I (measured in Ohms, Ω).",
            "difficulty": "Easy"
        },
        {
            "subject": "Mathematics",
            "question": "What is the definite integral ∫₀¹ 3x² dx?",
            "option_a": "0",
            "option_b": "1",
            "option_c": "2",
            "option_d": "3",
            "correct": "B",
            "explanation": "∫ 3x² dx = [x³] from 0 to 1 = (1)³ - (0)³ = 1.",
            "difficulty": "Medium"
        },
        {
            "subject": "Mathematics",
            "question": "What is the trigonometric identity for sin²(θ) + cos²(θ)?",
            "option_a": "0",
            "option_b": "1",
            "option_c": "tan(θ)",
            "option_d": "sec²(θ)",
            "correct": "B",
            "explanation": "By the Pythagorean trigonometric identity, sin²(θ) + cos²(θ) = 1 for all real angles θ.",
            "difficulty": "Easy"
        },
        {
            "subject": "Chemistry",
            "question": "In the chemical reaction 2H₂ + O₂ → 2H₂O, what is the stoichiometric ratio of H₂ to O₂?",
            "option_a": "1 : 1",
            "option_b": "2 : 1",
            "option_c": "1 : 2",
            "option_d": "2 : 2",
            "correct": "B",
            "explanation": "Two moles of molecular Hydrogen (2H₂) react with one mole of molecular Oxygen (1O₂) to produce two moles of Water (2H₂O).",
            "difficulty": "Easy"
        },
        {
            "subject": "Biology",
            "question": "Which molecule stores and delivers energy for cellular activities (often called cellular energy currency)?",
            "option_a": "DNA",
            "option_b": "RNA",
            "option_c": "ATP (Adenosine Triphosphate)",
            "option_d": "Glucose",
            "correct": "C",
            "explanation": "ATP (Adenosine Triphosphate) captures chemical energy obtained from the breakdown of food molecules and releases it to fuel other cellular processes.",
            "difficulty": "Easy"
        }
    ]

    print("Populating multi-sheet Excel with scientific sample questions...")
    existing_ids = storage.get_existing_ids()
    for item in sample_data:
        subj = item["subject"]
        code = subject_mgr.get_code(subj)
        qid = IdGenerator.generate_next_id(code, existing_ids)
        q = Question(
            id=qid,
            subject=subj,
            subject_code=code,
            question=item["question"],
            option_a=item["option_a"],
            option_b=item["option_b"],
            option_c=item["option_c"],
            option_d=item["option_d"],
            correct_answer=item["correct"],
            explanation=item["explanation"],
            difficulty=item["difficulty"]
        )
        saved, msg = storage.add_question(q)
        if saved:
            existing_ids.append(qid)
            print(f"Added [{qid}] -> Master sheet & '{subj}' sheet.")

    print("\nExcel Multi-Sheet structure created:")
    sheets = storage.get_available_sheet_names()
    print("Sheets in workbook:", sheets)

if __name__ == "__main__":
    seed_sample_questions()
