"""
Subject registry and Unique ID generation engine.
Generates unique, sequential IDs formatted per subject (e.g. MATH-0001, PHY-0002).
"""
import json
import os
import re
from typing import Dict, List, Optional
import config

class SubjectManager:
    def __init__(self, config_path: str = config.SUBJECTS_CONFIG_PATH):
        self.config_path = config_path
        self.subjects: Dict[str, str] = {}
        self.load_subjects()

    def load_subjects(self) -> Dict[str, str]:
        """Load subjects and codes from file or initialize with defaults."""
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    self.subjects = json.load(f)
            except Exception:
                self.subjects = config.DEFAULT_SUBJECTS.copy()
                self.save_subjects()
        else:
            self.subjects = config.DEFAULT_SUBJECTS.copy()
            self.save_subjects()
        return self.subjects

    def save_subjects(self) -> None:
        """Save subjects map to JSON file."""
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(self.subjects, f, indent=4)
        except Exception as e:
            print(f"Error saving subjects config: {e}")

    def add_subject(self, name: str, code: Optional[str] = None) -> str:
        """
        Add a new subject. If code is not provided, auto-generate from name.
        Returns the assigned subject code.
        """
        name = name.strip()
        if not name:
            raise ValueError("Subject name cannot be empty.")

        if name in self.subjects:
            return self.subjects[name]

        if not code or not code.strip():
            # Derive 3-4 letter uppercase code
            words = re.findall(r'[A-Za-z0-9]+', name)
            if len(words) == 1:
                code = words[0][:4].upper()
            else:
                code = "".join(w[0] for w in words).upper()[:4]
        else:
            code = re.sub(r'[^A-Za-z0-9]', '', code).upper()

        if not code:
            code = "SUB"

        # Ensure uniqueness of code if another subject already uses it
        base_code = code
        counter = 1
        existing_codes = set(self.subjects.values())
        while code in existing_codes:
            code = f"{base_code}{counter}"
            counter += 1

        self.subjects[name] = code
        self.save_subjects()
        return code

    def get_code(self, subject_name: str) -> str:
        """Get the code prefix for a given subject."""
        if subject_name in self.subjects:
            return self.subjects[subject_name]
        # If subject is not yet in dictionary, register it
        return self.add_subject(subject_name)

    def get_all_subjects(self) -> List[str]:
        """Return sorted list of all subject names."""
        return sorted(list(self.subjects.keys()))


class IdGenerator:
    @staticmethod
    def extract_number(qid: str, prefix: str) -> int:
        """Extract the numeric suffix from an ID like MATH-0005 -> 5."""
        if not qid:
            return 0
        pattern = rf"^{re.escape(prefix)}-?(\d+)$"
        match = re.search(pattern, qid.strip(), re.IGNORECASE)
        if match:
            try:
                return int(match.group(1))
            except ValueError:
                return 0
        return 0

    @classmethod
    def generate_next_id(cls, subject_code: str, existing_ids: List[str], padding: int = 4) -> str:
        """
        Generate the next available sequential Unique ID for a given subject code.
        Example: if existing IDs are ['MATH-0001', 'MATH-0002'], returns 'MATH-0003'.
        """
        prefix = subject_code.strip().upper()
        max_num = 0

        for qid in existing_ids:
            if not qid:
                continue
            qid_clean = qid.strip()
            # Check if this ID belongs to the same subject prefix (with or without hyphen)
            if qid_clean.upper().startswith(f"{prefix}-") or qid_clean.upper().startswith(prefix):
                num = cls.extract_number(qid_clean, prefix)
                if num > max_num:
                    max_num = num

        next_num = max_num + 1
        formatted_id = f"{prefix}-{str(next_num).zfill(padding)}"
        return formatted_id
