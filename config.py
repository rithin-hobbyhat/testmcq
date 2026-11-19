"""
Configuration settings for the Question Bank Authoring Studio.
"""
import os
import sys

# Base directories
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)

# File paths
EXCEL_FILE_PATH = os.path.join(DATA_DIR, "questions_bank.xlsx")
SUBJECTS_CONFIG_PATH = os.path.join(DATA_DIR, "subjects.json")

# Default Subjects and their Unique ID prefixes
DEFAULT_SUBJECTS = {
    "FLUID MECHANICS AND HYDRAULIC MACHINES": "FMHM"
}

# Excel Column Headers (Order is strictly maintained)
EXCEL_COLUMNS = [
    "Question ID",
    "Subject",
    "Subject Code",
    "Question",
    "Option A",
    "Option B",
    "Option C",
    "Option D",
    "Correct Answer",
    "Explanation",
    "Difficulty",
    "Created Date"
]

# Master sheet name
MASTER_SHEET_NAME = "All Questions"

# Scientific, Math, Chemistry & Physics Symbols for Quick Insertion
SCIENTIFIC_SYMBOLS = {
    "Math & Calculus": [
        "±", "×", "÷", "√", "∛", "∫", "∬", "∭", "∮", "∂", "∑", "∏",
        "≠", "≈", "≡", "≤", "≥", "<", ">", "∞", "∝", "∠", "⊥", "∥",
        "log", "ln", "lim", "sin", "cos", "tan", "cot", "sec", "csc"
    ],
    "Greek & Physics": [
        "α", "β", "γ", "δ", "ε", "ζ", "η", "θ", "ι", "κ", "λ", "μ",
        "ν", "ξ", "π", "ρ", "σ", "τ", "υ", "φ", "χ", "ψ", "ω",
        "Δ", "Γ", "Λ", "Σ", "Φ", "Ψ", "Ω", "ℏ", "ħ", "Å"
    ],
    "Superscript": [
        "⁰", "¹", "²", "³", "⁴", "⁵", "⁶", "⁷", "⁸", "⁹", "⁺", "⁻", "⁼", "⁽", "⁾", "ⁿ", "ⁱ", "ˣ", "ʸ"
    ],
    "Subscript": [
        "₀", "₁", "₂", "₃", "₄", "₅", "₆", "₇", "₈", "₉", "₊", "₋", "₌", "₍", "₎", "ₐ", "ₑ", "ₒ", "ₓ"
    ],
    "Chemistry & Arrows": [
        "→", "⇌", "⇄", "⇒", "⇔", "↑", "↓", "↔",
        "°", "℃", "℉", "K", "mol", "pH", "pKa", "ΔH", "ΔG", "ΔS",
        "α-ray", "β-ray", "γ-ray", "e⁻", "p⁺", "n⁰"
    ]
}

# UI Appearance Settings
APP_TITLE = "MOQ Question Bank Authoring Studio"
APP_VERSION = "v3.0"
WINDOW_MIN_WIDTH = 1250
WINDOW_MIN_HEIGHT = 800

# Color Theme Palette - High-Contrast Black & White Monochrome
THEME = {
    "primary": "#27272A",       # Dark zinc monochrome
    "primary_hover": "#3F3F46",
    "success": "#18181B",       # Monochrome dark
    "success_hover": "#27272A",
    "danger": "#7F1D1D",        # Dark red
    "danger_hover": "#991B1B",
    "warning": "#27272A",       
    "warning_hover": "#3F3F46",
    "card_dark": "#0D0D0F",     # Pure dark
    "card_light": "#18181B",
    "accent": "#FFFFFF"         # Pure white
}
