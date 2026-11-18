"""
Core logic verification tests for Question model, ID Generator, Excel Multi-Sheet Storage, and Scientific Symbols.
"""
import os
import shutil
import unittest
import openpyxl
import config
from models import Question
from id_generator import SubjectManager, IdGenerator
from excel_storage import ExcelStorage

class TestCoreEngine(unittest.TestCase):
    def setUp(self):
        self.test_dir = os.path.join(config.BASE_DIR, "test_data")
        os.makedirs(self.test_dir, exist_ok=True)
        self.test_excel = os.path.join(self.test_dir, "test_questions.xlsx")
        self.test_json = os.path.join(self.test_dir, "test_subjects.json")
        self.storage = ExcelStorage(self.test_excel)
        self.subject_mgr = SubjectManager(self.test_json)

    def tearDown(self):
        if os.path.exists(self.test_dir):
            try:
                shutil.rmtree(self.test_dir)
            except Exception:
                pass

    def test_subject_and_id_generation(self):
        math_code = self.subject_mgr.get_code("Mathematics")
        self.assertEqual(math_code, "MATH")

        existing_ids = ["MATH-0001", "MATH-0002", "PHY-0001"]
        next_math_id = IdGenerator.generate_next_id("MATH", existing_ids)
        self.assertEqual(next_math_id, "MATH-0003")

        next_phy_id = IdGenerator.generate_next_id("PHY", existing_ids)
        self.assertEqual(next_phy_id, "PHY-0002")

        next_chem_id = IdGenerator.generate_next_id("CHEM", existing_ids)
        self.assertEqual(next_chem_id, "CHEM-0001")

        ai_code = self.subject_mgr.add_subject("Artificial Intelligence", "AI")
        self.assertEqual(ai_code, "AI")
        next_ai_id = IdGenerator.generate_next_id("AI", existing_ids)
        self.assertEqual(next_ai_id, "AI-0001")

    def test_multi_sheet_and_scientific_symbols(self):
        # 1. Add Physics question with symbols
        q_phy = Question(
            id="PHY-0001",
            subject="Physics",
            subject_code="PHY",
            question="What does 'c' represent in E = mc²?",
            option_a="Speed of light (≈ 3 × 10⁸ m/s)",
            option_b="Specific heat capacity",
            option_c="Planck's constant ℏ",
            option_d="Coulomb constant",
            correct_answer="A",
            explanation="c is the speed of light in vacuum (≈ 3 × 10⁸ m/s).",
            difficulty="Easy"
        )
        saved, msg = self.storage.add_question(q_phy)
        self.assertTrue(saved, f"Add failed: {msg}")

        # 2. Add Math question with calculus symbols
        q_math = Question(
            id="MATH-0001",
            subject="Mathematics",
            subject_code="MATH",
            question="Evaluate the definite integral ∫₀¹ 3x² dx.",
            option_a="0",
            option_b="1",
            option_c="2",
            option_d="3",
            correct_answer="B",
            explanation="∫ 3x² dx = [x³]₀¹ = 1 - 0 = 1",
            difficulty="Medium"
        )
        saved, msg = self.storage.add_question(q_math)
        self.assertTrue(saved, f"Add failed: {msg}")

        # 3. Add Chemistry question with chemical formulas
        q_chem = Question(
            id="CHEM-0001",
            subject="Chemistry",
            subject_code="CHEM",
            question="In 2H₂ + O₂ → 2H₂O, what is the reaction product?",
            option_a="Hydrogen peroxide (H₂O₂)",
            option_b="Water (H₂O)",
            option_c="Hydride ion (H⁻)",
            option_d="Hydroxide (OH⁻)",
            correct_answer="B",
            explanation="Water vapor/liquid is formed: 2H₂ + O₂ → 2H₂O.",
            difficulty="Easy"
        )
        saved, msg = self.storage.add_question(q_chem)
        self.assertTrue(saved, f"Add failed: {msg}")

        # 4. Verify sheets in Excel workbook
        wb = openpyxl.load_workbook(self.test_excel)
        sheet_names = wb.sheetnames
        self.assertIn(config.MASTER_SHEET_NAME, sheet_names)
        self.assertIn("Physics", sheet_names)
        self.assertIn("Mathematics", sheet_names)
        self.assertIn("Chemistry", sheet_names)

        # Verify Master Sheet has all 3 questions
        all_master = self.storage.load_all_questions()
        self.assertEqual(len(all_master), 3)

        # Verify Subject Sheets have only their respective questions
        phy_qs = self.storage.load_all_questions("Physics")
        self.assertEqual(len(phy_qs), 1)
        self.assertEqual(phy_qs[0].id, "PHY-0001")
        self.assertIn("E = mc²", phy_qs[0].question)

        math_qs = self.storage.load_all_questions("Mathematics")
        self.assertEqual(len(math_qs), 1)
        self.assertEqual(math_qs[0].id, "MATH-0001")
        self.assertIn("∫₀¹ 3x² dx", math_qs[0].question)

        chem_qs = self.storage.load_all_questions("Chemistry")
        self.assertEqual(len(chem_qs), 1)
        self.assertEqual(chem_qs[0].id, "CHEM-0001")
        self.assertIn("2H₂ + O₂ → 2H₂O", chem_qs[0].question)
        wb.close()

        # 5. Test updating with symbols
        q_phy.question = "In relativistic mechanics, E² = (pc)² + (m₀c²)² where m₀ is rest mass."
        updated, msg = self.storage.update_question(q_phy)
        self.assertTrue(updated, f"Update failed: {msg}")

        reloaded_phy = self.storage.load_all_questions("Physics")
        self.assertIn("E² = (pc)² + (m₀c²)²", reloaded_phy[0].question)

        # 6. Test deleting question
        deleted, msg = self.storage.delete_question("MATH-0001")
        self.assertTrue(deleted, f"Delete failed: {msg}")
        self.assertEqual(len(self.storage.load_all_questions()), 2)
        self.assertEqual(len(self.storage.load_all_questions("Mathematics")), 0)

if __name__ == "__main__":
    unittest.main()
