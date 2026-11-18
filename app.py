"""
MOQ Question Bank Authoring Studio (Version 2.1)
High-Productivity Desktop Application for Fast MCQ Authoring & Excel Storage.
Features:
- Collapsible/Expandable Scientific & Math Symbol Palette
- Rapid Data Entry: Keyboard Shortcuts (Ctrl+S, Ctrl+Enter for Save & Next, Alt+A/B/C/D)
- Duplicate Question Template
- Subject-wise Multi-Sheet Excel Storage with Auto-Generated Unique IDs
"""
import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import customtkinter as ctk

import config
from models import Question
from id_generator import SubjectManager, IdGenerator
from excel_storage import ExcelStorage

# Configure CustomTkinter Appearance
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class AddSubjectDialog(ctk.CTkToplevel):
    """Modal dialog to add a new subject with customizable prefix code."""
    def __init__(self, parent, subject_manager: SubjectManager, on_success_callback):
        super().__init__(parent)
        self.parent = parent
        self.subject_manager = subject_manager
        self.on_success = on_success_callback

        self.title("Add New Subject")
        self.geometry("450x330")
        self.resizable(False, False)
        self.grab_set()

        # Center dialog
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 450) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 330) // 2
        self.geometry(f"+{x}+{y}")

        self._build_ui()

    def _build_ui(self):
        container = ctk.CTkFrame(self, corner_radius=12)
        container.pack(fill="both", expand=True, padx=20, pady=20)

        title_lbl = ctk.CTkLabel(
            container,
            text="Add New Subject",
            font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold")
        )
        title_lbl.pack(pady=(15, 10))

        desc_lbl = ctk.CTkLabel(
            container,
            text="Enter subject name and its unique prefix code\n(e.g., Physics -> PHY, Biology -> BIO)",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="gray"
        )
        desc_lbl.pack(pady=(0, 15))

        self.name_entry = ctk.CTkEntry(
            container,
            placeholder_text="Subject Name (e.g. Environmental Science)",
            height=38,
            font=ctk.CTkFont(family="Segoe UI", size=13)
        )
        self.name_entry.pack(fill="x", padx=20, pady=(0, 10))
        self.name_entry.bind("<KeyRelease>", self._auto_suggest_code)

        self.code_entry = ctk.CTkEntry(
            container,
            placeholder_text="Unique Prefix Code (e.g. ENV)",
            height=38,
            font=ctk.CTkFont(family="Segoe UI", size=13)
        )
        self.code_entry.pack(fill="x", padx=20, pady=(0, 15))

        btn_frame = ctk.CTkFrame(container, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=(5, 10))

        cancel_btn = ctk.CTkButton(
            btn_frame,
            text="Cancel",
            fg_color="#475569",
            hover_color="#334155",
            command=self.destroy,
            width=100,
            height=36
        )
        cancel_btn.pack(side="left", padx=(0, 10))

        save_btn = ctk.CTkButton(
            btn_frame,
            text="Save Subject",
            fg_color=config.THEME["primary"],
            hover_color=config.THEME["primary_hover"],
            command=self._save_subject,
            height=36
        )
        save_btn.pack(side="right", fill="x", expand=True)

    def _auto_suggest_code(self, event=None):
        name = self.name_entry.get().strip()
        if name and not self.code_entry.get().strip():
            words = name.split()
            if len(words) == 1:
                suggested = words[0][:4].upper()
            else:
                suggested = "".join(w[0] for w in words).upper()[:4]
            self.code_entry.delete(0, "end")
            self.code_entry.insert(0, suggested)

    def _save_subject(self):
        name = self.name_entry.get().strip()
        code = self.code_entry.get().strip().upper()

        if not name:
            messagebox.showwarning("Validation Error", "Please enter a subject name.", parent=self)
            return

        try:
            assigned_code = self.subject_manager.add_subject(name, code if code else None)
            self.on_success(name, assigned_code)
            self.destroy()
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)


class QuestionBankApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # App Window Settings
        self.title(f"{config.APP_TITLE} ({config.APP_VERSION}) - High-Speed Authoring")
        self.geometry(f"{config.WINDOW_MIN_WIDTH}x{config.WINDOW_MIN_HEIGHT}")
        self.minsize(1100, 750)

        # Set Custom 'R' Window & Taskbar Icon
        ico_path = os.path.join(config.BASE_DIR, "app_icon.ico")
        if os.path.exists(ico_path):
            try:
                self.iconbitmap(ico_path)
            except Exception:
                pass

        # Backend Services
        self.storage = ExcelStorage()
        self.subject_manager = SubjectManager()

        # State Variables
        self.editing_question_id: Optional[str] = None
        self.all_questions: list[Question] = []
        self.filtered_questions: list[Question] = []
        self.last_focused_widget = None
        self.is_symbols_palette_open = False  # Palette starts collapsed by default for fast clean typing

        # Build UI
        self._setup_styles()
        self._build_header()
        self._build_main_layout()
        self._build_statusbar()
        self._setup_keyboard_shortcuts()

        # Load Initial Data
        self._refresh_all_data()

    def _setup_styles(self):
        """Configure ttk Treeview styles."""
        self.style = ttk.Style()
        self.style.theme_use("clam")
        
        self.style.configure(
            "Custom.Treeview",
            background="#1E293B",
            foreground="#F8FAFC",
            fieldbackground="#1E293B",
            rowheight=32,
            font=("Segoe UI", 10),
            borderwidth=0
        )
        self.style.configure(
            "Custom.Treeview.Heading",
            background="#0F172A",
            foreground="#38BDF8",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            padding=6
        )
        self.style.map(
            "Custom.Treeview",
            background=[("selected", "#2563EB")],
            foreground=[("selected", "#FFFFFF")]
        )

    def _setup_keyboard_shortcuts(self):
        """Global productivity hotkeys for rapid data entry."""
        self.bind("<Control-s>", lambda e: self._on_save_button_clicked())
        self.bind("<Control-Return>", lambda e: self._save_and_next())
        self.bind("<Control-n>", lambda e: self._reset_form())
        self.bind("<Alt-a>", lambda e: self.correct_ans_var.set("A"))
        self.bind("<Alt-b>", lambda e: self.correct_ans_var.set("B"))
        self.bind("<Alt-c>", lambda e: self.correct_ans_var.set("C"))
        self.bind("<Alt-d>", lambda e: self.correct_ans_var.set("D"))

    def _build_header(self):
        """Top banner with stats, open excel button, and theme switcher."""
        self.header_frame = ctk.CTkFrame(self, height=65, corner_radius=0, fg_color=("#1E293B", "#0F172A"))
        self.header_frame.pack(fill="x", side="top")

        title_frame = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        title_frame.pack(side="left", padx=20, pady=12)

        logo_badge = ctk.CTkLabel(
            title_frame,
            text="MOQ",
            font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
            text_color="#FFFFFF"
        )
        logo_badge.pack(side="left", padx=(0, 10))

        title_lbl = ctk.CTkLabel(
            title_frame,
            text="Question Bank Studio",
            font=ctk.CTkFont(family="Segoe UI", size=17, weight="bold"),
            text_color="#F8FAFC"
        )
        title_lbl.pack(side="left")

        # Stats Badges
        self.total_badge = ctk.CTkLabel(
            self.header_frame,
            text="Total Questions: 0",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#27272A",
            text_color="#FFFFFF",
            corner_radius=8,
            padx=12,
            pady=4
        )
        self.total_badge.pack(side="left", padx=10)

        self.sheets_badge = ctk.CTkLabel(
            self.header_frame,
            text="Subject Sheets: 0",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#18181B",
            text_color="#D4D4D8",
            corner_radius=8,
            padx=12,
            pady=4
        )
        self.sheets_badge.pack(side="left", padx=5)

        # Quick Shortcuts Hint
        shortcuts_lbl = ctk.CTkLabel(
            self.header_frame,
            text="Hotkeys: [Ctrl+Enter: Save & Next] [Ctrl+S: Save] [Ctrl+N: Clear] [Alt+A/B/C/D: Ans]",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color="#94A3B8"
        )
        shortcuts_lbl.pack(side="left", padx=15)

        # Right-side Header Buttons
        right_btn_frame = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        right_btn_frame.pack(side="right", padx=20)

        # Theme Switch
        self.theme_switch = ctk.CTkSwitch(
            right_btn_frame,
            text="Dark Mode",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            command=self._toggle_theme
        )
        self.theme_switch.select()
        self.theme_switch.pack(side="right", padx=10)

        # Open in Excel Button
        open_excel_btn = ctk.CTkButton(
            right_btn_frame,
            text="Open in Excel",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#27272A",
            hover_color="#3F3F46",
            command=self._open_excel_file,
            width=130,
            height=34
        )
        open_excel_btn.pack(side="right", padx=4)

        # Practice Window Button (Web / Mobile)
        practice_btn = ctk.CTkButton(
            right_btn_frame,
            text="Practice MCQs (Web / Phone)",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#FFFFFF",
            text_color="#000000",
            hover_color="#E4E4E7",
            command=self._open_practice_window,
            width=210,
            height=34
        )
        practice_btn.pack(side="right", padx=4)

    def _build_main_layout(self):
        """Build Split layout: Left Panel (Entry Form + Symbols) & Right Panel (Table Explorer)."""
        self.main_paned = ctk.CTkFrame(self, fg_color="transparent")
        self.main_paned.pack(fill="both", expand=True, padx=15, pady=10)

        # ----------------- LEFT PANEL: FORM & SYMBOLS -----------------
        self.left_panel = ctk.CTkScrollableFrame(
            self.main_paned,
            width=540,
            corner_radius=12,
            fg_color=("#F1F5F9", "#1E293B")
        )
        self.left_panel.pack(side="left", fill="both", expand=False, padx=(0, 10))

        self._build_form(self.left_panel)

        # ----------------- RIGHT PANEL: TABLE & PREVIEW -----------------
        self.right_panel = ctk.CTkFrame(
            self.main_paned,
            corner_radius=12,
            fg_color=("#F1F5F9", "#1E293B")
        )
        self.right_panel.pack(side="right", fill="both", expand=True)

        self._build_table_view(self.right_panel)

    def _build_form(self, container):
        """Build Question Entry Form with Collapsible Symbols Toolbar."""
        # Form Header
        form_hdr_frame = ctk.CTkFrame(container, fg_color="transparent")
        form_hdr_frame.pack(fill="x", pady=(5, 8))

        self.form_mode_lbl = ctk.CTkLabel(
            form_hdr_frame,
            text="New Question Entry",
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            text_color="#FFFFFF"
        )
        self.form_mode_lbl.pack(side="left")

        # Subject & Unique ID Section
        subj_card = ctk.CTkFrame(container, corner_radius=8, fg_color=("#E2E8F0", "#0F172A"))
        subj_card.pack(fill="x", pady=(0, 8), padx=5)

        subj_row = ctk.CTkFrame(subj_card, fg_color="transparent")
        subj_row.pack(fill="x", padx=10, pady=8)

        subj_lbl = ctk.CTkLabel(subj_row, text="Subject:", font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"))
        subj_lbl.grid(row=0, column=0, sticky="w", padx=(0, 5))

        subjects = self.subject_manager.get_all_subjects()
        self.subject_var = ctk.StringVar(value=subjects[0] if subjects else "General Knowledge")
        self.subject_dropdown = ctk.CTkOptionMenu(
            subj_row,
            variable=self.subject_var,
            values=subjects,
            command=self._on_subject_changed,
            width=180,
            height=32
        )
        self.subject_dropdown.grid(row=0, column=1, padx=5, sticky="w")

        add_subj_btn = ctk.CTkButton(
            subj_row,
            text="+ Add Subject",
            width=95,
            height=32,
            fg_color="#475569",
            hover_color="#334155",
            command=self._show_add_subject_dialog
        )
        add_subj_btn.grid(row=0, column=2, padx=5, sticky="w")

        # Unique ID Live Preview Badge
        id_row = ctk.CTkFrame(subj_card, fg_color="transparent")
        id_row.pack(fill="x", padx=10, pady=(0, 8))

        id_title = ctk.CTkLabel(id_row, text="Auto Unique ID (Subject-wise):", font=ctk.CTkFont(family="Segoe UI", size=12))
        id_title.pack(side="left")

        self.id_badge = ctk.CTkLabel(
            id_row,
            text="GK-0001",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color="#047857",
            text_color="#FFFFFF",
            corner_radius=6,
            padx=10,
            pady=2
        )
        self.id_badge.pack(side="left", padx=10)

        # ----------------- COLLAPSIBLE SCIENTIFIC SYMBOLS PALETTE -----------------
        self.symbols_card = ctk.CTkFrame(container, corner_radius=8, fg_color=("#E2E8F0", "#0F172A"))
        self.symbols_card.pack(fill="x", pady=(0, 8), padx=5)

        # Toggle Button Header
        sym_hdr = ctk.CTkFrame(self.symbols_card, fg_color="transparent")
        sym_hdr.pack(fill="x", padx=10, pady=6)

        self.toggle_palette_btn = ctk.CTkButton(
            sym_hdr,
            text="Scientific & Math Symbols Palette ▼",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#27272A",
            hover_color="#3F3F46",
            height=30,
            command=self._toggle_symbols_palette
        )
        self.toggle_palette_btn.pack(fill="x", expand=True)

        # Body of Symbols Palette (Hidden by default)
        self.symbols_body_frame = ctk.CTkFrame(self.symbols_card, fg_color="transparent")

        # Category Selector for Symbols
        self.symbol_cat_var = ctk.StringVar(value="Math & Calculus")
        self.sym_cat_menu = ctk.CTkSegmentedButton(
            self.symbols_body_frame,
            values=list(config.SCIENTIFIC_SYMBOLS.keys()),
            variable=self.symbol_cat_var,
            command=self._render_symbol_buttons,
            height=26,
            font=ctk.CTkFont(family="Segoe UI", size=11)
        )
        self.sym_cat_menu.pack(fill="x", padx=10, pady=(4, 6))

        # Buttons container for symbols
        self.symbols_btn_frame = ctk.CTkFrame(self.symbols_body_frame, fg_color="transparent")
        self.symbols_btn_frame.pack(fill="x", padx=6, pady=(0, 8))

        self._render_symbol_buttons("Math & Calculus")

        # ----------------- QUESTION & OPTIONS INPUTS -----------------
        q_lbl = ctk.CTkLabel(container, text="Question Statement: * (Supports symbols)", font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"))
        q_lbl.pack(anchor="w", padx=5, pady=(4, 2))

        self.question_text = ctk.CTkTextbox(
            container,
            height=85,
            font=ctk.CTkFont(family="Segoe UI", size=13),
            corner_radius=8
        )
        self.question_text.pack(fill="x", padx=5, pady=(0, 8))
        self._bind_focus(self.question_text)

        # Options A, B, C, D
        opts_lbl = ctk.CTkLabel(container, text="Options: * (Use Alt+A/B/C/D to pick right answer)", font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"))
        opts_lbl.pack(anchor="w", padx=5, pady=(0, 4))

        # Option A
        opt_a_frame = ctk.CTkFrame(container, fg_color="transparent")
        opt_a_frame.pack(fill="x", padx=5, pady=2)
        ctk.CTkLabel(opt_a_frame, text="A:", font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"), width=25).pack(side="left")
        self.opt_a_entry = ctk.CTkEntry(opt_a_frame, placeholder_text="Option A text", height=32, font=ctk.CTkFont(family="Segoe UI", size=12))
        self.opt_a_entry.pack(side="left", fill="x", expand=True)
        self._bind_focus(self.opt_a_entry)

        # Option B
        opt_b_frame = ctk.CTkFrame(container, fg_color="transparent")
        opt_b_frame.pack(fill="x", padx=5, pady=2)
        ctk.CTkLabel(opt_b_frame, text="B:", font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"), width=25).pack(side="left")
        self.opt_b_entry = ctk.CTkEntry(opt_b_frame, placeholder_text="Option B text", height=32, font=ctk.CTkFont(family="Segoe UI", size=12))
        self.opt_b_entry.pack(side="left", fill="x", expand=True)
        self._bind_focus(self.opt_b_entry)

        # Option C
        opt_c_frame = ctk.CTkFrame(container, fg_color="transparent")
        opt_c_frame.pack(fill="x", padx=5, pady=2)
        ctk.CTkLabel(opt_c_frame, text="C:", font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"), width=25).pack(side="left")
        self.opt_c_entry = ctk.CTkEntry(opt_c_frame, placeholder_text="Option C text", height=32, font=ctk.CTkFont(family="Segoe UI", size=12))
        self.opt_c_entry.pack(side="left", fill="x", expand=True)
        self._bind_focus(self.opt_c_entry)

        # Option D
        opt_d_frame = ctk.CTkFrame(container, fg_color="transparent")
        opt_d_frame.pack(fill="x", padx=5, pady=2)
        ctk.CTkLabel(opt_d_frame, text="D:", font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"), width=25).pack(side="left")
        self.opt_d_entry = ctk.CTkEntry(opt_d_frame, placeholder_text="Option D text", height=32, font=ctk.CTkFont(family="Segoe UI", size=12))
        self.opt_d_entry.pack(side="left", fill="x", expand=True)
        self._bind_focus(self.opt_d_entry)

        # Correct Answer Selection & Difficulty Row
        ans_diff_frame = ctk.CTkFrame(container, corner_radius=8, fg_color=("#E2E8F0", "#0F172A"))
        ans_diff_frame.pack(fill="x", padx=5, pady=8)

        # Correct Answer Segmented
        ans_sub = ctk.CTkFrame(ans_diff_frame, fg_color="transparent")
        ans_sub.pack(fill="x", padx=10, pady=6)

        ctk.CTkLabel(ans_sub, text="Right Answer: *", font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"), text_color="#10B981").pack(side="left", padx=(0, 10))
        self.correct_ans_var = ctk.StringVar(value="A")
        self.ans_segmented = ctk.CTkSegmentedButton(
            ans_sub,
            values=["A", "B", "C", "D"],
            variable=self.correct_ans_var,
            selected_color="#10B981",
            selected_hover_color="#059669",
            height=30,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold")
        )
        self.ans_segmented.pack(side="left", fill="x", expand=True)

        # Difficulty
        diff_sub = ctk.CTkFrame(ans_diff_frame, fg_color="transparent")
        diff_sub.pack(fill="x", padx=10, pady=(0, 6))

        ctk.CTkLabel(diff_sub, text="Difficulty Level:", font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold")).pack(side="left", padx=(0, 10))
        self.diff_var = ctk.StringVar(value="Medium")
        self.diff_segmented = ctk.CTkSegmentedButton(
            diff_sub,
            values=["Easy", "Medium", "Hard"],
            variable=self.diff_var,
            selected_color="#3B82F6",
            height=28
        )
        self.diff_segmented.pack(side="left", fill="x", expand=True)

        # Explanation / Description Text
        exp_lbl = ctk.CTkLabel(container, text="Description / Explanation of Right Answer:", font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"))
        exp_lbl.pack(anchor="w", padx=5, pady=(4, 2))

        self.explanation_text = ctk.CTkTextbox(
            container,
            height=75,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            corner_radius=8
        )
        self.explanation_text.pack(fill="x", padx=5, pady=(0, 10))
        self._bind_focus(self.explanation_text)

        # High-Speed Action Buttons Row
        self.btn_action_frame = ctk.CTkFrame(container, fg_color="transparent")
        self.btn_action_frame.pack(fill="x", padx=5, pady=(0, 10))

        # Save & Next Button (Super Fast)
        self.save_next_btn = ctk.CTkButton(
            self.btn_action_frame,
            text="Save & Next (Ctrl+Enter)",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color="#27272A",
            hover_color="#3F3F46",
            height=40,
            command=self._save_and_next
        )
        self.save_next_btn.pack(side="left", fill="x", expand=True, padx=(0, 6))

        self.save_btn = ctk.CTkButton(
            self.btn_action_frame,
            text="Save (Ctrl+S)",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color=config.THEME["primary"],
            hover_color=config.THEME["primary_hover"],
            height=40,
            width=110,
            command=self._on_save_button_clicked
        )
        self.save_btn.pack(side="left", padx=(0, 6))

        self.clear_btn = ctk.CTkButton(
            self.btn_action_frame,
            text="Clear",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color="#18181B",
            hover_color="#27272A",
            width=80,
            height=40,
            command=self._reset_form
        )
        self.clear_btn.pack(side="right")

    def _toggle_symbols_palette(self):
        """Expand or collapse the scientific symbols palette with one click."""
        if self.is_symbols_palette_open:
            self.symbols_body_frame.pack_forget()
            self.toggle_palette_btn.configure(
                text="Scientific & Math Symbols Palette ▼",
                fg_color="#27272A"
            )
            self.is_symbols_palette_open = False
        else:
            self.symbols_body_frame.pack(fill="x", padx=4, pady=(0, 4))
            self.toggle_palette_btn.configure(
                text="Scientific & Math Symbols Palette ▲",
                fg_color="#3F3F46"
            )
            self.is_symbols_palette_open = True

    def _bind_focus(self, widget):
        """Track last focused text widget to insert symbols into."""
        def on_focus(event):
            self.last_focused_widget = widget
        widget.bind("<FocusIn>", on_focus)

    def _render_symbol_buttons(self, category_name: str):
        """Render click-to-insert buttons for the selected symbol category."""
        for child in self.symbols_btn_frame.winfo_children():
            child.destroy()

        symbols = config.SCIENTIFIC_SYMBOLS.get(category_name, [])
        cols = 10
        for idx, sym in enumerate(symbols):
            r = idx // cols
            c = idx % cols
            btn = ctk.CTkButton(
                self.symbols_btn_frame,
                text=sym,
                width=38,
                height=28,
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                fg_color=("#CBD5E1", "#334155"),
                hover_color=("#94A3B8", "#475569"),
                text_color=("#0F172A", "#F8FAFC"),
                command=lambda s=sym: self._insert_symbol(s)
            )
            btn.grid(row=r, column=c, padx=2, pady=2, sticky="ew")

    def _insert_symbol(self, symbol: str):
        """Insert the clicked symbol into the active input widget."""
        target = self.last_focused_widget or self.question_text
        try:
            if isinstance(target, ctk.CTkTextbox):
                target.insert("insert", symbol)
                target.focus_set()
            elif isinstance(target, ctk.CTkEntry):
                target.insert("insert", symbol)
                target.focus_set()
        except Exception:
            self.question_text.insert("insert", symbol)
            self.question_text.focus_set()

    def _build_table_view(self, container):
        """Build Table Explorer and Multi-Sheet View on the Right Panel."""
        toolbar = ctk.CTkFrame(container, fg_color="transparent")
        toolbar.pack(fill="x", padx=10, pady=(10, 8))

        # Search Box
        self.search_entry = ctk.CTkEntry(
            toolbar,
            placeholder_text="Search question statement, options, ID...",
            height=36,
            font=ctk.CTkFont(family="Segoe UI", size=12)
        )
        self.search_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.search_entry.bind("<KeyRelease>", self._on_search_query_changed)

        # Sheet / Subject Filter Dropdown
        self.filter_sheet_var = ctk.StringVar(value="All Questions (Master Sheet)")
        self.filter_dropdown = ctk.CTkOptionMenu(
            toolbar,
            variable=self.filter_sheet_var,
            values=["All Questions (Master Sheet)"],
            command=self._on_filter_changed,
            width=200,
            height=36
        )
        self.filter_dropdown.pack(side="left", padx=(0, 10))

        # Refresh Button
        refresh_btn = ctk.CTkButton(
            toolbar,
            text="Refresh",
            width=80,
            height=36,
            fg_color="#27272A",
            hover_color="#3F3F46",
            command=self._refresh_all_data
        )
        refresh_btn.pack(side="right")

        # Table Frame with Scrollbars
        table_container = ctk.CTkFrame(container, fg_color="transparent")
        table_container.pack(fill="both", expand=True, padx=10, pady=(0, 8))

        columns = ("id", "subject", "question", "correct", "difficulty", "date")
        self.tree = ttk.Treeview(
            table_container,
            columns=columns,
            show="headings",
            style="Custom.Treeview",
            selectmode="browse"
        )

        self.tree.heading("id", text="Unique ID")
        self.tree.heading("subject", text="Subject Sheet")
        self.tree.heading("question", text="Question Statement")
        self.tree.heading("correct", text="Ans")
        self.tree.heading("difficulty", text="Difficulty")
        self.tree.heading("date", text="Created Date")

        self.tree.column("id", width=95, anchor="center")
        self.tree.column("subject", width=120, anchor="w")
        self.tree.column("question", width=260, anchor="w")
        self.tree.column("correct", width=50, anchor="center")
        self.tree.column("difficulty", width=80, anchor="center")
        self.tree.column("date", width=130, anchor="center")

        v_scroll = ttk.Scrollbar(table_container, orient="vertical", command=self.tree.yview)
        h_scroll = ttk.Scrollbar(table_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        v_scroll.pack(side="right", fill="y")
        h_scroll.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)

        self.tree.bind("<<TreeviewSelect>>", self._on_row_selected)
        self.tree.bind("<Double-1>", self._on_row_double_clicked)

        # Bottom Quick Action Bar
        action_bar = ctk.CTkFrame(container, fg_color="transparent")
        action_bar.pack(fill="x", padx=10, pady=(0, 8))

        self.edit_btn = ctk.CTkButton(
            action_bar,
            text="Edit Selected",
            fg_color="#27272A",
            hover_color="#3F3F46",
            command=self._load_selected_for_edit,
            height=34,
            width=110
        )
        self.edit_btn.pack(side="left", padx=(0, 6))

        self.duplicate_btn = ctk.CTkButton(
            action_bar,
            text="Duplicate as New",
            fg_color="#27272A",
            hover_color="#3F3F46",
            command=self._duplicate_selected_question,
            height=34,
            width=135
        )
        self.duplicate_btn.pack(side="left", padx=(0, 6))

        self.delete_btn = ctk.CTkButton(
            action_bar,
            text="Delete",
            fg_color="#7F1D1D",
            hover_color="#991B1B",
            command=self._delete_selected_question,
            height=34,
            width=90
        )
        self.delete_btn.pack(side="left", padx=(0, 6))

        self.count_info_lbl = ctk.CTkLabel(
            action_bar,
            text="Showing 0 questions",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="gray"
        )
        self.count_info_lbl.pack(side="right", padx=10)

        # Preview Details Card at bottom right
        self.preview_card = ctk.CTkFrame(container, corner_radius=8, fg_color=("#E2E8F0", "#0F172A"), height=130)
        self.preview_card.pack(fill="x", padx=10, pady=(0, 8))

        self.preview_title = ctk.CTkLabel(
            self.preview_card,
            text="Question Details Preview",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color="#FFFFFF"
        )
        self.preview_title.pack(anchor="w", padx=10, pady=(6, 2))

        self.preview_content = ctk.CTkLabel(
            self.preview_card,
            text="Select any question in the table above to preview full options, right answer, and description.",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            justify="left",
            wraplength=550
        )
        self.preview_content.pack(anchor="w", padx=10, pady=(0, 6))

    def _build_statusbar(self):
        """Bottom status bar with information and file path."""
        self.statusbar = ctk.CTkFrame(self, height=26, corner_radius=0, fg_color=("#CBD5E1", "#0F172A"))
        self.statusbar.pack(fill="x", side="bottom")

        self.status_lbl = ctk.CTkLabel(
            self.statusbar,
            text="Ready. Excel file: " + config.EXCEL_FILE_PATH,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=("gray20", "gray70")
        )
        self.status_lbl.pack(side="left", padx=15)

    # ----------------- LOGIC & HANDLERS -----------------

    def _refresh_all_data(self):
        """Reload all questions and update UI components."""
        self.all_questions = self.storage.load_all_questions()
        
        # Update Subject Dropdown values
        subjects = self.subject_manager.get_all_subjects()
        self.subject_dropdown.configure(values=subjects)

        # Update Filter Sheet Dropdown
        available_sheets = self.storage.get_available_sheet_names()
        filter_options = ["All Questions (Master Sheet)"] + [s for s in available_sheets if s != config.MASTER_SHEET_NAME]
        self.filter_dropdown.configure(values=filter_options)

        # Update Live ID preview
        self._update_next_id_preview()

        # Update Table
        self._apply_filters()

        # Update Header Counters
        self.total_badge.configure(text=f"Total Questions: {len(self.all_questions)}")
        subject_sheets_count = len([s for s in available_sheets if s != config.MASTER_SHEET_NAME])
        self.sheets_badge.configure(text=f"Subject Sheets: {subject_sheets_count}")

    def _update_next_id_preview(self):
        """Update the ID badge based on current subject."""
        if self.editing_question_id:
            self.id_badge.configure(text=f"{self.editing_question_id} (Editing)", fg_color="#D97706")
            return

        subject_name = self.subject_var.get()
        subject_code = self.subject_manager.get_code(subject_name)
        existing_ids = self.storage.get_existing_ids()
        next_id = IdGenerator.generate_next_id(subject_code, existing_ids)
        self.id_badge.configure(text=next_id, fg_color="#047857")

    def _on_subject_changed(self, choice):
        self._update_next_id_preview()

    def _show_add_subject_dialog(self):
        def on_added(name, code):
            self.subject_var.set(name)
            self._refresh_all_data()
            self._set_status(f"Added subject '{name}' with prefix '{code}'.")

        AddSubjectDialog(self, self.subject_manager, on_added)

    def _apply_filters(self):
        """Filter table based on search query and selected sheet."""
        query = self.search_entry.get().strip().lower()
        selected_filter = self.filter_sheet_var.get().strip()

        if selected_filter == "All Questions (Master Sheet)":
            base_list = self.all_questions
        else:
            filt_clean = selected_filter.lower()
            base_list = [
                q for q in self.all_questions
                if self.storage._sanitize_sheet_title(q.subject).strip().lower() == filt_clean
                or q.subject.strip().lower() == filt_clean
                or (q.subject_code and q.subject_code.strip().lower() == filt_clean)
                or (q.id and q.id.strip().lower().startswith(filt_clean))
            ]

        self.filtered_questions = []
        for q in base_list:
            if query:
                searchable_text = f"{q.id} {q.subject} {q.question} {q.option_a} {q.option_b} {q.option_c} {q.option_d} {q.explanation}".lower()
                if query not in searchable_text:
                    continue
            self.filtered_questions.append(q)

        # Repopulate treeview
        for item in self.tree.get_children():
            self.tree.delete(item)

        for q in self.filtered_questions:
            q_summary = (q.question[:45] + "...") if len(q.question) > 45 else q.question
            self.tree.insert("", "end", iid=q.id, values=(
                q.id,
                q.subject,
                q_summary,
                q.correct_answer,
                q.difficulty,
                q.created_date
            ))

        self.count_info_lbl.configure(text=f"Showing {len(self.filtered_questions)} of {len(self.all_questions)} questions")

    def _on_search_query_changed(self, event=None):
        self._apply_filters()

    def _on_filter_changed(self, choice):
        self._apply_filters()

    def _on_row_selected(self, event=None):
        selected_items = self.tree.selection()
        if not selected_items:
            return
        qid = selected_items[0]
        q = next((item for item in self.all_questions if item.id == qid), None)
        if q:
            self.preview_title.configure(text=f"[{q.id}] {q.subject} ({q.difficulty})")
            exp_text = f"\nExplanation: {q.explanation}" if q.explanation else ""
            preview_str = (
                f"Question: {q.question}\n"
                f"A: {q.option_a} | B: {q.option_b} | C: {q.option_c} | D: {q.option_d}\n"
                f"Correct Answer: Option {q.correct_answer}{exp_text}"
            )
            self.preview_content.configure(text=preview_str)

    def _on_row_double_clicked(self, event=None):
        self._load_selected_for_edit()

    def _load_selected_for_edit(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showinfo("Select Question", "Please select a question from the table to edit.")
            return

        qid = selected_items[0]
        q = next((item for item in self.all_questions if item.id == qid), None)
        if not q:
            return

        self.editing_question_id = q.id
        self.form_mode_lbl.configure(text=f"Editing Question: {q.id}", text_color="#FFFFFF")
        self.save_btn.configure(text="Update in Excel", fg_color=config.THEME["warning"], hover_color=config.THEME["warning_hover"])

        self.subject_var.set(q.subject)
        self._update_next_id_preview()

        self.question_text.delete("1.0", "end")
        self.question_text.insert("1.0", q.question)

        self.opt_a_entry.delete(0, "end")
        self.opt_a_entry.insert(0, q.option_a)

        self.opt_b_entry.delete(0, "end")
        self.opt_b_entry.insert(0, q.option_b)

        self.opt_c_entry.delete(0, "end")
        self.opt_c_entry.insert(0, q.option_c)

        self.opt_d_entry.delete(0, "end")
        self.opt_d_entry.insert(0, q.option_d)

        self.correct_ans_var.set(q.correct_answer.upper())
        self.diff_var.set(q.difficulty)

        self.explanation_text.delete("1.0", "end")
        self.explanation_text.insert("1.0", q.explanation)

        self._set_status(f"Loaded question '{q.id}' for editing.")
        self.question_text.focus_set()

    def _duplicate_selected_question(self):
        """Duplicate selected question as a new question draft with new unique ID."""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showinfo("Select Question", "Please select a question from the table to duplicate.")
            return

        qid = selected_items[0]
        q = next((item for item in self.all_questions if item.id == qid), None)
        if not q:
            return

        self.editing_question_id = None
        self.form_mode_lbl.configure(text="New Question (Duplicated Template)", text_color="#FFFFFF")
        self.save_btn.configure(text="Save (Ctrl+S)", fg_color=config.THEME["primary"], hover_color=config.THEME["primary_hover"])

        self.subject_var.set(q.subject)
        self._update_next_id_preview()

        self.question_text.delete("1.0", "end")
        self.question_text.insert("1.0", q.question)

        self.opt_a_entry.delete(0, "end")
        self.opt_a_entry.insert(0, q.option_a)

        self.opt_b_entry.delete(0, "end")
        self.opt_b_entry.insert(0, q.option_b)

        self.opt_c_entry.delete(0, "end")
        self.opt_c_entry.insert(0, q.option_c)

        self.opt_d_entry.delete(0, "end")
        self.opt_d_entry.insert(0, q.option_d)

        self.correct_ans_var.set(q.correct_answer.upper())
        self.diff_var.set(q.difficulty)

        self.explanation_text.delete("1.0", "end")
        self.explanation_text.insert("1.0", q.explanation)

        self._set_status(f"Duplicated question from '{q.id}'. Edit and click Save.")
        self.question_text.focus_set()

    def _save_and_next(self):
        """Save question and automatically prepare form for the next question without interruption."""
        saved = self._save_question_internal(show_alert=False)
        if saved:
            self._reset_form()
            self._set_status("Saved! Ready for next question.")
            self.question_text.focus_set()

    def _on_save_button_clicked(self):
        """Standard save button click."""
        self._save_question_internal(show_alert=True)

    def _save_question_internal(self, show_alert: bool = True) -> bool:
        """Internal logic for saving questions."""
        subject = self.subject_var.get().strip()
        subject_code = self.subject_manager.get_code(subject)
        question_text = self.question_text.get("1.0", "end").strip()
        opt_a = self.opt_a_entry.get().strip()
        opt_b = self.opt_b_entry.get().strip()
        opt_c = self.opt_c_entry.get().strip()
        opt_d = self.opt_d_entry.get().strip()
        correct_ans = self.correct_ans_var.get().strip().upper()
        difficulty = self.diff_var.get().strip()
        explanation = self.explanation_text.get("1.0", "end").strip()

        if self.editing_question_id:
            qid = self.editing_question_id
        else:
            existing_ids = self.storage.get_existing_ids()
            qid = IdGenerator.generate_next_id(subject_code, existing_ids)

        new_q = Question(
            id=qid,
            subject=subject,
            subject_code=subject_code,
            question=question_text,
            option_a=opt_a,
            option_b=opt_b,
            option_c=opt_c,
            option_d=opt_d,
            correct_answer=correct_ans,
            explanation=explanation,
            difficulty=difficulty
        )

        errors = new_q.validate()
        if errors:
            error_msg = "\n• " + "\n• ".join(errors)
            messagebox.showwarning("Incomplete Information", f"Please resolve the following before saving:\n{error_msg}")
            return False

        if self.editing_question_id:
            success, msg = self.storage.update_question(new_q)
            action_desc = "updated"
        else:
            success, msg = self.storage.add_question(new_q)
            action_desc = "saved"

        if success:
            self._set_status(f"Successfully {action_desc} question {new_q.id} in Master sheet and '{subject}' sheet!")
            self._refresh_all_data()
            if show_alert:
                messagebox.showinfo("Success", f"Question [{new_q.id}] has been {action_desc} into Excel:\n- 'All Questions' sheet\n- '{subject}' subject sheet")
                self._reset_form()
            return True
        else:
            messagebox.showerror("Save Error", f"Failed to save question:\n{msg}")
            return False

    def _delete_selected_question(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showinfo("Select Question", "Please select a question from the table to delete.")
            return

        qid = selected_items[0]
        confirm = messagebox.askyesno(
            "Confirm Delete",
            f"Are you sure you want to permanently delete question '{qid}' from the Excel workbook?",
            icon="warning"
        )
        if confirm:
            success, msg = self.storage.delete_question(qid)
            if success:
                self._set_status(f"Deleted question {qid}.")
                if self.editing_question_id == qid:
                    self._reset_form()
                self._refresh_all_data()
                messagebox.showinfo("Deleted", f"Question '{qid}' was removed from all sheets.")
            else:
                messagebox.showerror("Delete Error", f"Could not delete question:\n{msg}")

    def _reset_form(self):
        self.editing_question_id = None
        self.form_mode_lbl.configure(text="New Question Entry", text_color="#FFFFFF")
        self.save_btn.configure(text="Save (Ctrl+S)", fg_color=config.THEME["primary"], hover_color=config.THEME["primary_hover"])

        self.question_text.delete("1.0", "end")
        self.opt_a_entry.delete(0, "end")
        self.opt_b_entry.delete(0, "end")
        self.opt_c_entry.delete(0, "end")
        self.opt_d_entry.delete(0, "end")
        self.correct_ans_var.set("A")
        self.diff_var.set("Medium")
        self.explanation_text.delete("1.0", "end")

        self._update_next_id_preview()
        self._set_status("Form ready for new question entry.")

    def _open_practice_window(self):
        """Export latest questions and launch the interactive MCQ Practice Arena in default browser."""
        try:
            self._set_status("Exporting latest questions for Practice Arena...")
            from export_data import export_all
            export_all()

            server_script = os.path.join(config.BASE_DIR, "server.py")
            subprocess.Popen([sys.executable, server_script])
            self._set_status("Practice Arena launched in your default browser!")
        except Exception as e:
            messagebox.showerror("Practice Arena Error", f"Could not launch practice mode: {e}")

    def _open_excel_file(self):
        file_path = config.EXCEL_FILE_PATH
        if not os.path.exists(file_path):
            self.storage._ensure_workbook_exists()

        try:
            os.startfile(file_path)
            self._set_status("Opened Excel workbook.")
        except Exception as e:
            messagebox.showerror("Error Opening File", f"Could not open file: {e}")

    def _toggle_theme(self):
        if self.theme_switch.get() == 1:
            ctk.set_appearance_mode("Dark")
            self.style.configure(
                "Custom.Treeview",
                background="#1E293B",
                foreground="#F8FAFC",
                fieldbackground="#1E293B"
            )
            self.style.configure("Custom.Treeview.Heading", background="#0F172A", foreground="#38BDF8")
        else:
            ctk.set_appearance_mode("Light")
            self.style.configure(
                "Custom.Treeview",
                background="#FFFFFF",
                foreground="#0F172A",
                fieldbackground="#FFFFFF"
            )
            self.style.configure("Custom.Treeview.Heading", background="#E2E8F0", foreground="#1E3A8A")

    def _set_status(self, msg: str):
        self.status_lbl.configure(text=msg)

if __name__ == "__main__":
    app = QuestionBankApp()
    app.mainloop()
