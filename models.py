"""
Data model for Question items in the Question Bank.
"""
from dataclasses import dataclass, field, asdict
from datetime import datetime
from typing import Optional, Dict, Any, List

@dataclass
class Question:
    id: str                               # Unique ID (e.g. MATH-0001)
    subject: str                          # Subject name (e.g. Mathematics)
    subject_code: str                     # Prefix code (e.g. MATH)
    question: str                         # Question statement
    option_a: str                         # Option A text
    option_b: str                         # Option B text
    option_c: str                         # Option C text
    option_d: str                         # Option D text
    correct_answer: str                   # 'A', 'B', 'C', or 'D'
    explanation: str = ""                 # Description / Explanation
    difficulty: str = "Medium"            # 'Easy', 'Medium', 'Hard'
    created_date: str = field(default_factory=lambda: datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    def validate(self) -> List[str]:
        """Validate question fields, returns list of validation error messages."""
        errors = []
        if not self.id or not self.id.strip():
            errors.append("Question ID is required.")
        if not self.subject or not self.subject.strip():
            errors.append("Subject is required.")
        if not self.question or not self.question.strip():
            errors.append("Question text cannot be empty.")
        if not self.option_a or not self.option_a.strip():
            errors.append("Option A cannot be empty.")
        if not self.option_b or not self.option_b.strip():
            errors.append("Option B cannot be empty.")
        if not self.option_c or not self.option_c.strip():
            errors.append("Option C cannot be empty.")
        if not self.option_d or not self.option_d.strip():
            errors.append("Option D cannot be empty.")
        if not self.correct_answer or self.correct_answer.upper() not in ["A", "B", "C", "D"]:
            errors.append("A valid Correct Answer (A, B, C, or D) must be selected.")
        return errors

    def to_row(self) -> List[Any]:
        """Convert Question to Excel row list matching EXCEL_COLUMNS order."""
        return [
            self.id.strip(),
            self.subject.strip(),
            self.subject_code.strip(),
            self.question.strip(),
            self.option_a.strip(),
            self.option_b.strip(),
            self.option_c.strip(),
            self.option_d.strip(),
            self.correct_answer.strip().upper(),
            self.explanation.strip(),
            self.difficulty.strip(),
            self.created_date
        ]

    @classmethod
    def from_row(cls, row: List[Any]) -> "Question":
        """Create Question instance from an Excel row."""
        def safe_str(val: Any) -> str:
            if val is None:
                return ""
            return str(val).strip()

        # Handle rows that might have missing trailing elements
        padded = list(row) + [""] * (12 - len(row))
        return cls(
            id=safe_str(padded[0]),
            subject=safe_str(padded[1]),
            subject_code=safe_str(padded[2]),
            question=safe_str(padded[3]),
            option_a=safe_str(padded[4]),
            option_b=safe_str(padded[5]),
            option_c=safe_str(padded[6]),
            option_d=safe_str(padded[7]),
            correct_answer=safe_str(padded[8]).upper() or "A",
            explanation=safe_str(padded[9]),
            difficulty=safe_str(padded[10]) or "Medium",
            created_date=safe_str(padded[11]) or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        )

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)
