"""
Excel Storage Engine for Question Bank.
Handles all read, write, update, and delete operations on .xlsx files with modern openpyxl styling.
Maintains both a consolidated 'All Questions' master sheet and individual subject-specific sheets.
"""
import os
import re
import shutil
import datetime
from typing import List, Optional, Tuple, Dict, Any
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import config
from models import Question

class ExcelStorage:
    def __init__(self, file_path: str = config.EXCEL_FILE_PATH):
        self.file_path = file_path
        self._ensure_workbook_exists()

    def _sanitize_sheet_title(self, name: str) -> str:
        """Sanitize sheet title for Excel (max 31 chars, remove invalid characters)."""
        clean = re.sub(r'[\\/*?:\[\]]', '', name.strip())
        return clean[:31] if clean else "Subject"

    def _ensure_workbook_exists(self) -> None:
        """Create the workbook with formatted master sheet if it doesn't exist."""
        if not os.path.exists(self.file_path):
            wb = openpyxl.Workbook()
            ws_master = wb.active
            ws_master.title = config.MASTER_SHEET_NAME
            self._write_header_row(ws_master)
            self._apply_sheet_styling(ws_master)
            self._save_workbook_safely(wb)
        else:
            # Upgrade existing sheet title if needed
            try:
                wb = openpyxl.load_workbook(self.file_path)
                if config.MASTER_SHEET_NAME not in wb.sheetnames:
                    # Rename first sheet to master sheet name
                    ws_first = wb.active
                    ws_first.title = config.MASTER_SHEET_NAME
                    self._save_workbook_safely(wb)
                wb.close()
            except Exception:
                pass

    def _write_header_row(self, ws) -> None:
        """Write column headers to a worksheet."""
        for col_idx, header in enumerate(config.EXCEL_COLUMNS, 1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid") # Navy Blue
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def _apply_sheet_styling(self, ws) -> None:
        """Apply freeze panes, gridlines, and sensible column widths."""
        ws.freeze_panes = "A2"
        if ws.views and ws.views.sheetView:
            ws.views.sheetView[0].showGridLines = True

        custom_widths = {
            1: 16,  # Question ID
            2: 20,  # Subject
            3: 14,  # Subject Code
            4: 45,  # Question
            5: 25,  # Option A
            6: 25,  # Option B
            7: 25,  # Option C
            8: 25,  # Option D
            9: 16,  # Correct Answer
            10: 45, # Explanation
            11: 14, # Difficulty
            12: 22  # Created Date
        }

        for col_idx, width in custom_widths.items():
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width

        ws.row_dimensions[1].height = 28

    def _format_data_row(self, ws, row_idx: int) -> None:
        """Style data rows with clean borders and typography."""
        thin_border = Border(
            left=Side(style='thin', color='CBD5E1'),
            right=Side(style='thin', color='CBD5E1'),
            top=Side(style='thin', color='CBD5E1'),
            bottom=Side(style='thin', color='CBD5E1')
        )
        
        fill_color = "F8FAFC" if row_idx % 2 == 0 else "FFFFFF"
        row_fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

        ws.row_dimensions[row_idx].height = 24

        for col_idx in range(1, len(config.EXCEL_COLUMNS) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = Font(name="Segoe UI", size=10)
            cell.fill = row_fill
            cell.border = thin_border

            if col_idx in [1, 3, 9, 11, 12]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

            if col_idx == 9 and cell.value:
                cell.font = Font(name="Segoe UI", size=10, bold=True, color="047857")

    def _save_workbook_safely(self, wb) -> Tuple[bool, str]:
        """Save workbook with lock detection and auto-backup."""
        try:
            if os.path.exists(self.file_path):
                backup_dir = os.path.join(os.path.dirname(self.file_path), "backups")
                os.makedirs(backup_dir, exist_ok=True)
                backup_file = os.path.join(backup_dir, f"backup_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx")
                try:
                    shutil.copy2(self.file_path, backup_file)
                except Exception:
                    pass

            wb.save(self.file_path)
            return True, "Success"
        except PermissionError:
            msg = (
                f"Cannot save to '{os.path.basename(self.file_path)}' because it is currently OPEN in Excel or another program.\n\n"
                "Please CLOSE the Excel file and try saving again."
            )
            return False, msg
        except Exception as e:
            return False, f"Unexpected error saving Excel file: {str(e)}"

    def _get_master_sheet(self, wb):
        """Retrieve master sheet from workbook."""
        if config.MASTER_SHEET_NAME in wb.sheetnames:
            return wb[config.MASTER_SHEET_NAME]
        elif "Master Question Bank" in wb.sheetnames:
            ws = wb["Master Question Bank"]
            ws.title = config.MASTER_SHEET_NAME
            return ws
        else:
            ws = wb.create_sheet(title=config.MASTER_SHEET_NAME, index=0)
            self._write_header_row(ws)
            self._apply_sheet_styling(ws)
            return ws

    def _get_or_create_subject_sheet(self, wb, subject_name: str):
        """Retrieve existing sheet for subject or create and style a new one."""
        sanitized = self._sanitize_sheet_title(subject_name)
        # Check case-insensitively first
        for name in wb.sheetnames:
            if name.strip().lower() == sanitized.strip().lower():
                return wb[name]
        ws = wb.create_sheet(title=sanitized)
        self._write_header_row(ws)
        self._apply_sheet_styling(ws)
        return ws

    def load_all_questions(self, sheet_name: Optional[str] = None) -> List[Question]:
        """Load questions from Excel workbook with high performance read_only mode."""
        self._ensure_workbook_exists()
        questions = []
        try:
            wb = openpyxl.load_workbook(self.file_path, read_only=True, data_only=True)
            ws = None
            if sheet_name:
                # Find matching sheet case-insensitively
                for s in wb.sheetnames:
                    if s.strip().lower() == sheet_name.strip().lower():
                        ws = wb[s]
                        break
            if ws is None:
                ws = wb[config.MASTER_SHEET_NAME] if config.MASTER_SHEET_NAME in wb.sheetnames else wb.active

            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or not any(row):
                    continue
                if not row[0] and not row[3]:
                    continue
                q = Question.from_row(list(row))
                questions.append(q)
            wb.close()
        except Exception as e:
            print(f"Error loading questions: {e}")
        return questions

    def get_existing_ids(self, subject: Optional[str] = None) -> List[str]:
        """Return a list of all existing Question IDs in the workbook."""
        questions = self.load_all_questions()
        if subject:
            return [q.id for q in questions if q.subject == subject and q.id]
        return [q.id for q in questions if q.id]

    def get_available_sheet_names(self) -> List[str]:
        """Return all sheet names in the workbook."""
        self._ensure_workbook_exists()
        try:
            wb = openpyxl.load_workbook(self.file_path, read_only=True)
            names = wb.sheetnames
            wb.close()
            return names
        except Exception:
            return [config.MASTER_SHEET_NAME]

    def add_question(self, question: Question) -> Tuple[bool, str]:
        """
        Add question to both the 'All Questions' master sheet and the subject's dedicated sheet.
        """
        self._ensure_workbook_exists()
        try:
            wb = openpyxl.load_workbook(self.file_path)
            ws_master = self._get_master_sheet(wb)

            # Check for duplicate ID in master sheet
            for row in ws_master.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True):
                if row and row[0] and str(row[0]).strip().upper() == question.id.strip().upper():
                    wb.close()
                    return False, f"Question ID '{question.id}' already exists in the database!"

            row_data = question.to_row()

            # 1. Append to Master Sheet
            next_master_row = ws_master.max_row + 1
            for col_idx, val in enumerate(row_data, 1):
                ws_master.cell(row=next_master_row, column=col_idx, value=val)
            self._format_data_row(ws_master, next_master_row)
            self._apply_sheet_styling(ws_master)

            # 2. Append to Subject-Specific Sheet
            ws_subject = self._get_or_create_subject_sheet(wb, question.subject)
            next_subj_row = ws_subject.max_row + 1
            for col_idx, val in enumerate(row_data, 1):
                ws_subject.cell(row=next_subj_row, column=col_idx, value=val)
            self._format_data_row(ws_subject, next_subj_row)
            self._apply_sheet_styling(ws_subject)

            saved, msg = self._save_workbook_safely(wb)
            wb.close()
            return saved, msg
        except Exception as e:
            return False, str(e)

    def update_question(self, question: Question) -> Tuple[bool, str]:
        """
        Update an existing question across both the master sheet and subject sheets.
        """
        self._ensure_workbook_exists()
        try:
            wb = openpyxl.load_workbook(self.file_path)
            row_data = question.to_row()
            qid_target = question.id.strip().upper()

            # 1. Update in Master Sheet
            ws_master = self._get_master_sheet(wb)
            for r in range(2, ws_master.max_row + 1):
                val = ws_master.cell(row=r, column=1).value
                if val and str(val).strip().upper() == qid_target:
                    for col_idx, cell_val in enumerate(row_data, 1):
                        ws_master.cell(row=r, column=col_idx, value=cell_val)
                    self._format_data_row(ws_master, r)
                    break

            # 2. Update in Subject Sheets
            subj_sheet_target = self._sanitize_sheet_title(question.subject)
            for sheet_name in list(wb.sheetnames):
                if sheet_name == config.MASTER_SHEET_NAME:
                    continue
                ws = wb[sheet_name]
                for r in range(2, ws.max_row + 1):
                    val = ws.cell(row=r, column=1).value
                    if val and str(val).strip().upper() == qid_target:
                        if sheet_name == subj_sheet_target:
                            # Update in place
                            for col_idx, cell_val in enumerate(row_data, 1):
                                ws.cell(row=r, column=col_idx, value=cell_val)
                            self._format_data_row(ws, r)
                        else:
                            # Subject changed, remove from old sheet
                            ws.delete_rows(r)
                        break

            # Ensure present in new subject sheet
            ws_target = self._get_or_create_subject_sheet(wb, question.subject)
            found_in_target = False
            for r in range(2, ws_target.max_row + 1):
                val = ws_target.cell(row=r, column=1).value
                if val and str(val).strip().upper() == qid_target:
                    found_in_target = True
                    break

            if not found_in_target:
                next_r = ws_target.max_row + 1
                for col_idx, cell_val in enumerate(row_data, 1):
                    ws_target.cell(row=next_r, column=col_idx, value=cell_val)
                self._format_data_row(ws_target, next_r)

            self._apply_sheet_styling(ws_target)

            saved, msg = self._save_workbook_safely(wb)
            wb.close()
            return saved, msg
        except Exception as e:
            return False, str(e)

    def delete_question(self, question_id: str) -> Tuple[bool, str]:
        """
        Delete a question from both master sheet and its subject sheet.
        """
        self._ensure_workbook_exists()
        try:
            wb = openpyxl.load_workbook(self.file_path)
            qid_target = question_id.strip().upper()

            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                for r in range(2, ws.max_row + 1):
                    val = ws.cell(row=r, column=1).value
                    if val and str(val).strip().upper() == qid_target:
                        ws.delete_rows(r)
                        for row_idx in range(2, ws.max_row + 1):
                            self._format_data_row(ws, row_idx)
                        self._apply_sheet_styling(ws)
                        break

            saved, msg = self._save_workbook_safely(wb)
            wb.close()
            return saved, msg
        except Exception as e:
            return False, str(e)
