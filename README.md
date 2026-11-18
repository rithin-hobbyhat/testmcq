# 📝 MOQ Question Bank Authoring Studio (Version 2.2)

A modern Windows desktop application designed to author, manage, and store Multiple Choice Questions (MCQs / Mock Questions) directly into a structured, multi-sheet Excel spreadsheet (`data/questions_bank.xlsx`).

---

## 🌟 1-Click Launch with Custom 'R' Icon

You now have a **dedicated desktop shortcut with the custom 'R' icon**!

### 🖱️ How to Open:
1. **On your Windows Desktop**: Look for the **`MOQ Question Bank`** icon with the **"R"** logo and double-click it.
2. **Inside the `MOQ` folder**:
   - Double-click **`MOQ Question Bank.lnk`**
   - Or double-click **`Launch_App.bat`**
   - Or double-click **`Launch_Silent.vbs`** (opens silently with no terminal window)

---

## 🌟 Key Features

### 1. 🔣 Collapsible Scientific & Mathematical Symbol Palette
- Quick-access symbol toolbar with category tabs:
  - **Math & Calculus**: `±`, `×`, `÷`, `√`, `∛`, `∫`, `∬`, `∮`, `∂`, `∑`, `∏`, `≠`, `≈`, `≤`, `≥`, `∞`, etc.
  - **Greek & Physics**: `α`, `β`, `γ`, `δ`, `ε`, `θ`, `λ`, `μ`, `π`, `ρ`, `σ`, `φ`, `ψ`, `ω`, `Δ`, `Ω`, `ℏ`, `Å`
  - **Superscripts**: `⁰`, `¹`, `²`, `³`, `⁴`, `⁵`, `⁶`, `⁷`, `⁸`, `⁹`, `⁺`, `⁻`, `ⁿ`, `ˣ`, `ʸ`
  - **Subscripts**: `₀`, `₁`, `₂`, `₃`, `₄`, `₅`, `₆`, `₇`, `₈`, `₉`, `₊`, `₋`, `ₐ`, `ₑ`, `ₒ`, `ₓ`
  - **Chemistry & Reactions**: `→`, `⇌`, `⇄`, `⇒`, `↑`, `↓`, `°`, `℃`, `℉`, `mol`, `pH`, `ΔH`, `e⁻`, `p⁺`, `n⁰`
- **Collapsible / Expandable**: Clean 1-click button to open or close the palette to save screen space.

### 2. ⚡ High-Speed Data Entry (Faster than manual Excel entry)
- **`⚡ Save & Next (Ctrl + Enter)`**: Saves question to Excel and immediately clears the form for the next question without moving your hands from the keyboard.
- **`Ctrl + S`**: Save question to Excel.
- **`Ctrl + N`**: Clear form / New question.
- **`Alt + A` / `Alt + B` / `Alt + C` / `Alt + D`**: Instantly select right answer.

### 3. 📑 Subject-Wise Multi-Sheet Excel Storage
- Automatically saves questions into separate Excel sheets named after each subject (`Physics`, `Mathematics`, `Chemistry`, `Biology`, etc.) and a master `All Questions` sheet.
- Styled with dark navy headers, bold text, auto-adjusted column widths, word wrap, and zebra row striping.
- **"📊 Open in Excel" Button**: 1-click button inside the app to view and open the spreadsheet directly in Microsoft Excel.

### 4. 🆔 Automatic Subject-Based Unique IDs
- Sequential Unique IDs are automatically generated per subject prefix (e.g. `PHY-0001`, `MATH-0001`, `CHEM-0001`).

---

## 📁 Project Files

```
MOQ/
├── MOQ Question Bank.lnk   # Windows Desktop Shortcut with 'R' icon
├── app_icon.ico            # High-resolution Windows 'R' icon
├── app_icon.png            # PNG logo
├── app.py                  # Main desktop GUI application
├── config.py               # Settings & symbols configuration
├── models.py               # Question data model & validation
├── id_generator.py         # Subject manager & unique ID engine
├── excel_storage.py        # Multi-sheet Excel engine (openpyxl)
├── seed_data.py            # Sample question generator
├── test_core.py            # Unit tests
├── Launch_App.bat          # Batch launcher
├── Launch_Silent.vbs       # Silent background launcher
└── data/
    ├── questions_bank.xlsx # Excel database with subject tabs
    ├── subjects.json       # Configured subjects and prefix codes
    └── backups/            # Auto backups
```
