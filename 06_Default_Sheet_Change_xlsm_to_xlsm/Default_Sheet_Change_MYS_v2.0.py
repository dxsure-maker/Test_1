import os
import shutil
import threading
import traceback
import time
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
import unicodedata

import pywintypes
import win32com.client as win32


class ExcelCopier:
    """Handle the Excel work so the UI stays clean."""

    def __init__(self, logger):
        self.logger = logger
        self.constants = self._build_constants()

    def run(
        self,
        main_file: Path,
        sheet_names,
        exclude_sheet_names,
        precondition_enabled,
        target_folder: Path,
        output_subfolder: str,
        progress_cb=None,
        cancel_cb=None,
    ):
        excel_files = self._collect_excel_files(target_folder)
        if not excel_files:
            raise RuntimeError("No Excel files (*.xlsx, *.xlsm) found in the target folder.")

        self.logger(f"Found {len(excel_files)} file(s) to process.")

        # Prepare output folder under the target folder
        output_folder = target_folder / output_subfolder
        output_folder.mkdir(parents=True, exist_ok=True)
        self.logger(f"Output folder: {output_folder}")

        total = len(excel_files)
        if progress_cb:
            progress_cb(0, total)

        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        # [OPT-5] Disable automatic recalculation during the entire run.
        # Formulas won't recalculate on every paste/write, saving significant
        # time for complex workbooks.  We restore the mode in finally.
        _calc_disabled = False
        try:
            excel.Calculation = self.constants["xlCalculationManual"]
            _calc_disabled = True
        except Exception:
            pass

        post_delete_targets = []
        try:
            base_name = main_file.stem
            ext = main_file.suffix

            for idx, file_name in enumerate(excel_files, start=1):
                if cancel_cb and cancel_cb():
                    self.logger("Cancellation requested before starting next file; stopping.")
                    break
                source_path = target_folder / file_name
                temp_target = output_folder / f"{base_name}({idx}){ext}"
                final_target = output_folder / file_name

                self.logger(f"[{idx}/{len(excel_files)}] Preparing {file_name}")
                shutil.copy2(main_file, temp_target)

                completed = self._copy_multiple_sheets(
                    excel,
                    source_path,
                    temp_target,
                    sheet_names,
                    exclude_sheet_names,
                    precondition_enabled,
                    cancel_cb=cancel_cb,
                )
                if not completed:
                    if temp_target.exists():
                        try:
                            temp_target.unlink()
                        except Exception:
                            pass
                    self.logger(f"[{idx}/{len(excel_files)}] Cancelled during {file_name}; partial file removed.")
                    break

                if final_target.exists():
                    final_target.unlink()
                temp_target.rename(final_target)

                if exclude_sheet_names:
                    post_delete_targets.append(final_target)

                self.logger(f"[{idx}/{len(excel_files)}] Done -> {file_name}")
                if progress_cb:
                    progress_cb(idx, total)
                # cancellation check happens at top of next iteration
        finally:
            if _calc_disabled:
                try:
                    excel.Calculation = self.constants["xlCalculationAutomatic"]
                except Exception:
                    pass
            excel.DisplayAlerts = True
            excel.EnableEvents = True
            excel.Quit()

        # Post-delete: 파일마다 Excel 인스턴스를 새로 생성합니다.
        # 단일 인스턴스로 여러 파일을 처리하면 COM 내부 상태가 누적되어
        # 시트 삭제가 적용되지 않는 이전 현상이 재발할 수 있습니다.
        # 안정성을 위해 v1.0 방식(파일당 새 인스턴스)을 유지합니다.
        if exclude_sheet_names and post_delete_targets:
            for target in post_delete_targets:
                if cancel_cb and cancel_cb():
                    break
                excel_cleanup = win32.DispatchEx("Excel.Application")
                excel_cleanup.Visible = False
                excel_cleanup.DisplayAlerts = False
                excel_cleanup.EnableEvents = False
                try:
                    self._delete_sheets_in_file(excel_cleanup, target, exclude_sheet_names, cancel_cb=cancel_cb)
                finally:
                    excel_cleanup.DisplayAlerts = True
                    excel_cleanup.EnableEvents = True
                    excel_cleanup.Quit()

    def _collect_excel_files(self, folder: Path):
        files = []
        for entry in folder.iterdir():
            if entry.is_file() and entry.suffix.lower() in {".xlsx", ".xlsm"}:
                files.append(entry.name)
        return sorted(files, key=str.lower)

    def _copy_multiple_sheets(
        self,
        excel,
        source_path: Path,
        target_path: Path,
        sheet_names,
        exclude_sheet_names,
        precondition_enabled,
        cancel_cb=None,
    ):
        source_wb = excel.Workbooks.Open(str(source_path), ReadOnly=True)
        target_wb = excel.Workbooks.Open(str(target_path))

        try:
            for sheet_name in sheet_names:
                if cancel_cb and cancel_cb():
                    return False
                sheet_name = sheet_name.strip()
                if not sheet_name:
                    continue

                source_sheet = self._get_sheet(source_wb, sheet_name)
                if source_sheet is None:
                    self.logger(f"  - Warning: sheet '{sheet_name}' not found in {source_path.name}, skipped.")
                    continue

                target_sheet = self._get_sheet(target_wb, sheet_name)
                if target_sheet is None:
                    target_sheet = target_wb.Worksheets.Add(After=target_wb.Worksheets(target_wb.Worksheets.Count))
                    target_sheet.Name = sheet_name

                if precondition_enabled and self._is_test_case_sheet(sheet_name):
                    completed = self._copy_test_case_partial(
                        excel, source_sheet, target_sheet, cancel_cb=cancel_cb
                    )
                    if not completed:
                        return False
                else:
                    self._copy_sheet_data(excel, source_sheet, target_sheet)
                if cancel_cb and cancel_cb():
                    return False

            # Copy missing defined names to avoid #NAME? errors when formulas rely on them
            self._copy_missing_names(source_wb, target_wb)
            # [OPT-5] Use workbook-scoped Calculate instead of the much more
            # expensive CalculateFullRebuild (which rebuilds ALL open workbooks).
            try:
                target_wb.Calculate()
            except Exception:
                pass

            target_wb.Save()
        finally:
            target_wb.Close(SaveChanges=False)
            source_wb.Close(SaveChanges=False)
        return True

    def _copy_sheet_data(self, excel, source_sheet, target_sheet):
        target_sheet.Cells.Clear()

        source_range = source_sheet.UsedRange
        if source_range is None:
            return

        address = source_range.Address
        start_row, start_col, end_row, end_col = self._get_used_range_bounds(source_sheet)

        # Copy values and formulas
        target_sheet.Range(address).Formula = source_range.Formula

        # Copy formats
        max_cells = 200000
        cell_count = (end_row - start_row + 1) * (end_col - start_col + 1)
        if cell_count > max_cells:
            self._paste_formats_in_chunks(excel, source_sheet, target_sheet, start_row, end_row, start_col, end_col)
        else:
            try:
                source_range.Copy()
                target_sheet.Range(address).PasteSpecial(self.constants["xlPasteFormats"])
            except pywintypes.com_error as exc:
                if self._is_paste_too_long_error(exc):
                    self._paste_formats_in_chunks(excel, source_sheet, target_sheet, start_row, end_row, start_col, end_col)
                else:
                    raise
            finally:
                excel.CutCopyMode = False

        # [OPT-4a] Column widths: one PasteSpecial call instead of a per-column loop.
        # xlPasteColumnWidths (8) sets all column widths in a single COM roundtrip.
        try:
            source_range.Copy()
            target_sheet.Range(address).PasteSpecial(self.constants["xlPasteColumnWidths"])
        except Exception:
            # Fallback: per-column loop
            for col_idx in range(source_range.Columns.Count):
                src_col = source_range.Column + col_idx
                try:
                    target_sheet.Columns(src_col).ColumnWidth = source_sheet.Columns(src_col).ColumnWidth
                except Exception:
                    pass
        finally:
            excel.CutCopyMode = False

        # [OPT-4b] Row heights: try one bulk COM call for uniform-height sheets
        # (most sheets have one consistent row height).  Falls back to per-row only
        # when rows actually differ.
        self._apply_row_heights(
            source_sheet, target_sheet,
            source_start=start_row, source_end=end_row,
            target_start=start_row,
        )

        # Preserve outline/group settings within used area (and last used cell)
        self._copy_outline_region(
            source_sheet,
            target_sheet,
            source_start_row=start_row,
            source_end_row=end_row,
            target_start_row=start_row,
            source_start_col=start_col,
            source_end_col=end_col,
            target_start_col=start_col,
        )

    def _get_sheet(self, workbook, name):
        try:
            return workbook.Worksheets(name)
        except Exception:
            return None

    def _get_used_range_bounds(self, sheet):
        used_range = sheet.UsedRange
        if used_range is None:
            return 1, 1, 1, 1
        start_row = used_range.Row
        start_col = used_range.Column
        end_row = start_row + used_range.Rows.Count - 1
        end_col = start_col + used_range.Columns.Count - 1
        try:
            last_cell = sheet.Cells.SpecialCells(self.constants["xlCellTypeLastCell"])
            end_row = max(end_row, last_cell.Row)
            end_col = max(end_col, last_cell.Column)
        except Exception:
            pass
        return start_row, start_col, end_row, end_col

    def _normalize_sheet_name(self, name):
        return unicodedata.normalize("NFKC", str(name)).strip().casefold()

    def _normalize_cell_value(self, value):
        if value is None:
            return ""
        return unicodedata.normalize("NFKC", str(value)).strip().casefold()

    def _is_test_case_sheet(self, name):
        return self._normalize_sheet_name(name) == self._normalize_sheet_name("Test Case")

    def _find_row_in_column(self, sheet, col_index, target_value, start_row=None, end_row_hint=None):
        # [OPT-1] Bulk column read: read the entire column slice in ONE COM call
        # (sheet.Range(...).Value) instead of calling sheet.Cells(row, col).Value
        # once per row.  For a 500-row sheet this cuts ~500 COM roundtrips to 1.
        if end_row_hint is not None:
            end_row = end_row_hint
        else:
            _, _, end_row, _ = self._get_used_range_bounds(sheet)

        lookup = self._normalize_cell_value(target_value)
        row_start = start_row if start_row is not None else 1
        if row_start > end_row:
            return None

        try:
            col_range = sheet.Range(
                sheet.Cells(row_start, col_index),
                sheet.Cells(end_row, col_index),
            )
            values = col_range.Value
        except Exception:
            values = None

        if values is None:
            return None

        # Single-cell returns a scalar; multi-cell returns a tuple of 1-tuples.
        if not isinstance(values, (list, tuple)):
            values = ((values,),)

        for i, row_val in enumerate(values):
            cell_val = row_val[0] if isinstance(row_val, (list, tuple)) else row_val
            if self._normalize_cell_value(cell_val) == lookup:
                return row_start + i

        return None

    def _copy_test_case_partial(self, excel, source_sheet, target_sheet, cancel_cb=None):
        if cancel_cb and cancel_cb():
            return False
        marker_value = "T1.1"
        end_marker = "END"

        # [OPT-6] Cache bounds once per sheet to avoid 5 redundant COM calls
        # (_get_used_range_bounds makes 2 COM calls each time it runs).
        _, _, source_end_row_bound, source_last_col = self._get_used_range_bounds(source_sheet)
        _, _, target_last_row, _ = self._get_used_range_bounds(target_sheet)

        source_start_row = self._find_row_in_column(
            source_sheet, 3, marker_value, end_row_hint=source_end_row_bound
        )
        if source_start_row is None:
            self.logger("  - PreCondition: 'T1.1' not found in source column C; full copy used.")
            self._copy_sheet_data(excel, source_sheet, target_sheet)
            return True

        source_end_row = self._find_row_in_column(
            source_sheet, 1, end_marker,
            start_row=source_start_row,
            end_row_hint=source_end_row_bound,
        )
        if source_end_row is None:
            self.logger("  - PreCondition: 'END' not found in source column A; full copy used.")
            self._copy_sheet_data(excel, source_sheet, target_sheet)
            return True

        target_start_row = self._find_row_in_column(
            target_sheet, 3, marker_value, end_row_hint=target_last_row
        )
        if target_start_row is None:
            self.logger("  - PreCondition: 'T1.1' not found in target column C; full copy used.")
            self._copy_sheet_data(excel, source_sheet, target_sheet)
            return True

        if cancel_cb and cancel_cb():
            return False

        end_col = source_last_col  # Already retrieved above — no extra COM call.

        if target_last_row >= target_start_row:
            target_sheet.Rows(f"{target_start_row}:{target_last_row}").Delete()

        row_count = source_end_row - source_start_row + 1
        source_range = source_sheet.Range(
            source_sheet.Cells(source_start_row, 1), source_sheet.Cells(source_end_row, end_col)
        )
        target_range = target_sheet.Range(
            target_sheet.Cells(target_start_row, 1),
            target_sheet.Cells(target_start_row + row_count - 1, end_col),
        )

        # Copy values and formulas
        target_range.Formula = source_range.Formula

        # Copy formats
        max_cells = 200000
        cell_count = row_count * end_col
        if cell_count > max_cells:
            self._paste_formats_in_chunks_offset(
                excel,
                source_sheet,
                target_sheet,
                source_start_row,
                target_start_row,
                row_count,
                1,
                end_col,
            )
        else:
            try:
                source_range.Copy()
                target_range.PasteSpecial(self.constants["xlPasteFormats"])
            except pywintypes.com_error as exc:
                if self._is_paste_too_long_error(exc):
                    self._paste_formats_in_chunks_offset(
                        excel,
                        source_sheet,
                        target_sheet,
                        source_start_row,
                        target_start_row,
                        row_count,
                        1,
                        end_col,
                    )
                else:
                    raise
            finally:
                excel.CutCopyMode = False

        # [OPT-4a] Column widths via PasteSpecial instead of per-column loop.
        try:
            source_range.Copy()
            target_range.PasteSpecial(self.constants["xlPasteColumnWidths"])
        except Exception:
            for col_idx in range(1, end_col + 1):
                try:
                    target_sheet.Columns(col_idx).ColumnWidth = source_sheet.Columns(col_idx).ColumnWidth
                except Exception:
                    pass
        finally:
            excel.CutCopyMode = False

        # [OPT-4b] Row heights with bulk-set optimisation.
        self._apply_row_heights(
            source_sheet, target_sheet,
            source_start=source_start_row, source_end=source_end_row,
            target_start=target_start_row,
        )

        self._copy_outline_region(
            source_sheet,
            target_sheet,
            source_start_row=source_start_row,
            source_end_row=source_end_row,
            target_start_row=target_start_row,
            source_start_col=1,
            source_end_col=end_col,
            target_start_col=1,
        )

        return True

    # ------------------------------------------------------------------
    # [OPT-4b] Helper: apply row heights in bulk when all source rows share
    # the same height (most common case), falling back to per-row only when
    # heights actually differ.
    # ------------------------------------------------------------------
    def _apply_row_heights(self, source_sheet, target_sheet,
                           source_start, source_end, target_start):
        row_count = source_end - source_start + 1
        if row_count <= 0:
            return
        try:
            src_rng = source_sheet.Range(
                source_sheet.Rows(source_start),
                source_sheet.Rows(source_end),
            )
            uniform_height = src_rng.RowHeight  # None if mixed, float if all equal
            if uniform_height is not None:
                # All rows share one height — set the whole target range at once.
                target_sheet.Range(
                    target_sheet.Rows(target_start),
                    target_sheet.Rows(target_start + row_count - 1),
                ).RowHeight = uniform_height
                return
        except Exception:
            pass
        # Fallback: per-row loop for sheets with varying row heights.
        for offset in range(row_count):
            try:
                target_sheet.Rows(target_start + offset).RowHeight = \
                    source_sheet.Rows(source_start + offset).RowHeight
            except Exception:
                pass

    def _delete_sheets(self, workbook, sheet_names):
        if not sheet_names:
            return
        sheet_map = {self._normalize_sheet_name(ws.Name): ws for ws in workbook.Worksheets}
        for sheet_name in sheet_names:
            if not sheet_name:
                continue
            key = self._normalize_sheet_name(sheet_name)
            if not key:
                continue
            sheet = sheet_map.get(key)
            if sheet is None:
                self.logger(f"  - Exclude: sheet '{sheet_name}' not found in default file, skipped.")
                continue
            try:
                sheet_display_name = sheet.Name
                if workbook.Worksheets.Count <= 1:
                    self.logger(
                        f"  - Exclude: cannot delete '{sheet_display_name}' because it is the last remaining sheet."
                    )
                    continue
                sheet.Delete()
                self.logger(f"  - Exclude: deleted sheet '{sheet_display_name}' from output.")
            except Exception:
                self.logger(f"  - Exclude: failed to delete sheet '{sheet_display_name}', skipped.")

    def _delete_sheets_in_file(self, excel, file_path: Path, sheet_names, cancel_cb=None):
        if cancel_cb and cancel_cb():
            return
        wb = excel.Workbooks.Open(str(file_path))
        try:
            self._delete_sheets(wb, sheet_names)
            try:
                wb.Calculate()
            except Exception:
                pass
            wb.Save()
        finally:
            wb.Close(SaveChanges=False)

    def _is_paste_too_long_error(self, exc):
        try:
            if len(exc.args) >= 3 and exc.args[2]:
                desc = str(exc.args[2][2])
                if "too long" in desc.lower() or "데이터가 너무 길어서" in desc:
                    return True
                code = exc.args[2][5]
                if code == -2146827284:
                    return True
        except Exception:
            pass
        return False

    def _paste_formats_in_chunks(self, excel, source_sheet, target_sheet, start_row, end_row, start_col, end_col):
        col_count = end_col - start_col + 1
        if col_count <= 0:
            return
        max_cells = 200000
        row_chunk = max(1, min(end_row - start_row + 1, max_cells // col_count))
        for row in range(start_row, end_row + 1, row_chunk):
            chunk_end = min(end_row, row + row_chunk - 1)
            try:
                src = source_sheet.Range(source_sheet.Cells(row, start_col), source_sheet.Cells(chunk_end, end_col))
                dst = target_sheet.Range(target_sheet.Cells(row, start_col), target_sheet.Cells(chunk_end, end_col))
                src.Copy()
                dst.PasteSpecial(self.constants["xlPasteFormats"])
            finally:
                try:
                    excel.CutCopyMode = False
                except Exception:
                    pass

    def _paste_formats_in_chunks_offset(
        self,
        excel,
        source_sheet,
        target_sheet,
        source_start_row,
        target_start_row,
        row_count,
        start_col,
        end_col,
    ):
        col_count = end_col - start_col + 1
        if col_count <= 0 or row_count <= 0:
            return
        max_cells = 200000
        row_chunk = max(1, min(row_count, max_cells // col_count))
        for offset in range(0, row_count, row_chunk):
            chunk_rows = min(row_chunk, row_count - offset)
            src_start = source_start_row + offset
            src_end = src_start + chunk_rows - 1
            tgt_start = target_start_row + offset
            tgt_end = tgt_start + chunk_rows - 1
            try:
                src = source_sheet.Range(source_sheet.Cells(src_start, start_col), source_sheet.Cells(src_end, end_col))
                dst = target_sheet.Range(target_sheet.Cells(tgt_start, start_col), target_sheet.Cells(tgt_end, end_col))
                src.Copy()
                dst.PasteSpecial(self.constants["xlPasteFormats"])
            finally:
                try:
                    excel.CutCopyMode = False
                except Exception:
                    pass

    def _copy_outline_region(
        self,
        source_sheet,
        target_sheet,
        source_start_row,
        source_end_row,
        target_start_row,
        source_start_col,
        source_end_col,
        target_start_col,
    ):
        try:
            target_sheet.Outline.SummaryRow = source_sheet.Outline.SummaryRow
            target_sheet.Outline.SummaryColumn = source_sheet.Outline.SummaryColumn
        except Exception:
            pass

        # Mixed ranges report Range.Hidden=False in Excel COM, even when some
        # rows are actually collapsed. For outline reliability, copy row/column
        # state directly cell-band by cell-band instead of trying to infer it
        # from a mixed Range object.
        row_count = source_end_row - source_start_row + 1
        for offset in range(row_count):
            src_row_idx = source_start_row + offset
            tgt_row_idx = target_start_row + offset
            try:
                src_row = source_sheet.Rows(src_row_idx)
                tgt_row = target_sheet.Rows(tgt_row_idx)
                tgt_row.OutlineLevel = src_row.OutlineLevel
                tgt_row.Hidden = src_row.Hidden
            except Exception:
                pass

        col_count = source_end_col - source_start_col + 1
        for offset in range(col_count):
            src_col_idx = source_start_col + offset
            tgt_col_idx = target_start_col + offset
            try:
                src_col = source_sheet.Columns(src_col_idx)
                tgt_col = target_sheet.Columns(tgt_col_idx)
                tgt_col.OutlineLevel = src_col.OutlineLevel
                tgt_col.Hidden = src_col.Hidden
            except Exception:
                pass

    def _copy_missing_names(self, source_wb, target_wb):
        normalize = self._normalize_sheet_name
        existing = {normalize(n.Name) for n in target_wb.Names}
        for name_obj in source_wb.Names:
            key = normalize(name_obj.Name)
            if key in existing:
                continue
            # Prefer workbook scope; fall back to sheet scope if needed
            try:
                target_wb.Names.Add(Name=name_obj.Name, RefersTo=name_obj.RefersTo)
                existing.add(key)
                continue
            except Exception:
                pass
            try:
                parent = getattr(name_obj, "Parent", None)
                if parent is not None and getattr(parent, "Name", None):
                    ws = self._get_sheet(target_wb, parent.Name)
                    if ws is not None:
                        ws.Names.Add(Name=name_obj.Name, RefersTo=name_obj.RefersTo)
                        existing.add(key)
                        continue
            except Exception:
                pass
            self.logger(f"  - Warning: failed to copy defined name '{name_obj.Name}'")

    def _build_constants(self):
        consts = {
            "xlPasteFormats": -4122,
            "xlCellTypeLastCell": 11,
            "xlPasteColumnWidths": 8,       # [OPT-4a] batch column-width paste
            "xlCalculationManual": -4135,   # [OPT-5] suppress auto-recalc
            "xlCalculationAutomatic": -4105,
        }
        try:
            from win32com.client import constants as excel_consts
            for key, default in list(consts.items()):
                consts[key] = getattr(excel_consts, key, default)
        except Exception:
            pass
        return consts



THEME = {
    "title": "Default Sheet Change MYS v2.0",
    "bg": "#0F172A",
    "panel": "#111827",
    "surface": "#1F2937",
    "surface_alt": "#0B1220",
    "text": "#E5E7EB",
    "muted": "#94A3B8",
    "accent": "#14B8A6",
    "accent_hover": "#0D9488",
    "accent_text": "#ECFEFF",
    "secondary": "#334155",
    "secondary_hover": "#475569",
    "border": "#233044",
    "progress_trough": "#0B1220",
    "list_bg": "#0B1220",
    "log_bg": "#08111F",
}


class App:
    def __init__(self, root):
        self.root = root
        self.root.title(THEME["title"])
        self.root.geometry("1180x840")
        self.root.minsize(1100, 800)

        self.main_file_var = tk.StringVar()
        self.sheet_names_var = tk.StringVar(value="Test Case,Test Setup,History")
        self.exclude_sheet_names_var = tk.StringVar(value="InOut정의방법")
        self.precondition_var = tk.BooleanVar(value=True)
        self.folder_var = tk.StringVar()
        self.output_subfolder_var = tk.StringVar(value="_output")
        self.running = False
        self.cancel_requested = False
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_text = tk.StringVar(value="0/0")
        self.eta_text = tk.StringVar(value="--")
        self._start_time = None

        self.style = ttk.Style()
        self._setup_theme()
        self._build_ui()

    def _setup_theme(self):
        self.style.theme_use("clam")
        self.root.configure(bg=THEME["bg"])

        self.style.configure(".", background=THEME["bg"], foreground=THEME["text"])
        self.style.configure("App.TFrame", background=THEME["bg"])
        self.style.configure("Card.TFrame", background=THEME["panel"])
        self.style.configure(
            "HeaderTitle.TLabel",
            background=THEME["bg"],
            foreground=THEME["text"],
            font=("Malgun Gothic", 18, "bold"),
        )
        self.style.configure(
            "HeaderSub.TLabel",
            background=THEME["bg"],
            foreground=THEME["muted"],
            font=("Malgun Gothic", 10),
        )
        self.style.configure(
            "CardTitle.TLabel",
            background=THEME["panel"],
            foreground=THEME["text"],
            font=("Malgun Gothic", 11, "bold"),
        )
        self.style.configure(
            "Body.TLabel",
            background=THEME["panel"],
            foreground=THEME["text"],
            font=("Malgun Gothic", 10),
        )
        self.style.configure(
            "Hint.TLabel",
            background=THEME["panel"],
            foreground=THEME["muted"],
            font=("Malgun Gothic", 9),
        )
        self.style.configure(
            "Value.TLabel",
            background=THEME["panel"],
            foreground=THEME["text"],
            font=("Consolas", 10),
        )
        self.style.configure(
            "App.TEntry",
            fieldbackground=THEME["surface"],
            foreground=THEME["text"],
            padding=(10, 8),
        )
        self.style.map(
            "App.TEntry",
            fieldbackground=[("disabled", THEME["surface_alt"]), ("readonly", THEME["surface"])],
            foreground=[("disabled", THEME["muted"])],
        )
        self.style.configure(
            "Accent.TButton",
            background=THEME["accent"],
            foreground=THEME["accent_text"],
            padding=(18, 10),
            font=("Malgun Gothic", 10, "bold"),
        )
        self.style.map(
            "Accent.TButton",
            background=[("active", THEME["accent_hover"]), ("disabled", THEME["secondary"])],
            foreground=[("disabled", THEME["muted"])],
        )
        self.style.configure(
            "Secondary.TButton",
            background=THEME["secondary"],
            foreground=THEME["text"],
            padding=(16, 10),
            font=("Malgun Gothic", 10),
        )
        self.style.map(
            "Secondary.TButton",
            background=[("active", THEME["secondary_hover"]), ("disabled", THEME["surface"])],
            foreground=[("disabled", THEME["muted"])],
        )
        self.style.configure(
            "App.TCheckbutton",
            background=THEME["panel"],
            foreground=THEME["text"],
            font=("Malgun Gothic", 10),
        )
        self.style.map("App.TCheckbutton", background=[("active", THEME["panel"])])
        self.style.configure(
            "Panel.Horizontal.TProgressbar",
            background=THEME["accent"],
            troughcolor=THEME["progress_trough"],
            thickness=14,
        )

    def _build_labeled_entry(self, parent, row, label, variable, button_text=None, button_command=None, hint=None):
        ttk.Label(parent, text=label, style="CardTitle.TLabel").grid(row=row, column=0, sticky="w")
        entry = ttk.Entry(parent, textvariable=variable, style="App.TEntry")
        entry.grid(row=row + 1, column=0, sticky="ew", pady=(6, 0))
        if button_text and button_command:
            ttk.Button(parent, text=button_text, command=button_command, style="Secondary.TButton").grid(
                row=row + 1, column=1, sticky="e", padx=(12, 0)
            )
        if hint:
            ttk.Label(parent, text=hint, style="Hint.TLabel", wraplength=520, justify="left").grid(
                row=row + 2, column=0, columnspan=2, sticky="w", pady=(6, 14)
            )
        return entry

    def _build_ui(self):
        outer = ttk.Frame(self.root, style="App.TFrame", padding=(20, 18, 20, 20))
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)

        header = ttk.Frame(outer, style="App.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 16))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="Default Sheet Change MYS v2.0", style="HeaderTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )

        body = ttk.Frame(outer, style="App.TFrame")
        body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(0, weight=7, uniform="panel")
        body.columnconfigure(1, weight=8, uniform="panel")
        body.rowconfigure(0, weight=1)

        left = ttk.Frame(body, style="Card.TFrame", padding=18)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.columnconfigure(0, weight=1)
        left.columnconfigure(1, weight=0)
        row = 0
        self._build_labeled_entry(
            left,
            row,
            "기준 파일",
            self.main_file_var,
            button_text="파일 선택",
            button_command=self.pick_main_file,
            hint="기준이 되는 Main Excel 파일을 선택합니다.",
        )
        row += 3
        self._build_labeled_entry(
            left,
            row,
            "복제 시트",
            self.sheet_names_var,
            hint="복제할 시트 이름을 입력합니다. 구분자: ,",
        )
        row += 3

        ttk.Label(left, text="PreCondition", style="CardTitle.TLabel").grid(row=row, column=0, sticky="w", columnspan=2)
        precondition_box = ttk.Frame(left, style="Card.TFrame")
        precondition_box.grid(row=row + 1, column=0, columnspan=2, sticky="ew", pady=(6, 14))
        ttk.Checkbutton(
            precondition_box,
            text="체크 해제 시 기존 복제 동작 유지",
            variable=self.precondition_var,
            style="App.TCheckbutton",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))
        row += 3

        self._build_labeled_entry(
            left,
            row,
            "제외 시트",
            self.exclude_sheet_names_var,
            hint="복제 대상에서 제외할 시트 이름을 입력합니다. 구분자: ,",
        )
        row += 3
        self._build_labeled_entry(
            left,
            row,
            "대상 폴더",
            self.folder_var,
            button_text="폴더 선택",
            button_command=self.pick_folder,
            hint="복제 대상 Excel 파일들이 있는 폴더를 선택합니다.",
        )
        row += 3
        self._build_labeled_entry(
            left,
            row,
            "결과 폴더명",
            self.output_subfolder_var,
            hint="대상 폴더 하위에 생성될 결과 파일 폴더의 이름을 입력합니다. 경로 구분자는 사용할 수 없습니다.",
        )

        right = ttk.Frame(body, style="Card.TFrame", padding=18)
        right.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        right.columnconfigure(0, weight=1)
        right.rowconfigure(2, weight=1)
        right.rowconfigure(5, weight=2)

        ttk.Label(right, text="대상 파일", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            right,
            text="선택한 폴더 안의 xlsx/xlsm 파일 목록입니다.",
            style="Hint.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(6, 10))

        list_frame = tk.Frame(right, bg=THEME["surface_alt"], highlightbackground=THEME["border"], highlightthickness=1)
        list_frame.grid(row=2, column=0, sticky="nsew")
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
        self.file_list = tk.Listbox(
            list_frame,
            height=10,
            bg=THEME["list_bg"],
            fg=THEME["text"],
            selectbackground=THEME["accent"],
            selectforeground=THEME["accent_text"],
            highlightthickness=0,
            relief="flat",
            borderwidth=0,
            activestyle="none",
        )
        self.file_list.grid(row=0, column=0, sticky="nsew")

        self.refresh_btn = ttk.Button(right, text="목록 새로고침", command=self.refresh_file_list, style="Secondary.TButton")
        self.refresh_btn.grid(row=3, column=0, sticky="e", pady=(12, 18))

        ttk.Label(right, text="실행 로그", style="CardTitle.TLabel").grid(row=4, column=0, sticky="w")
        log_frame = tk.Frame(right, bg=THEME["surface_alt"], highlightbackground=THEME["border"], highlightthickness=1)
        log_frame.grid(row=5, column=0, sticky="nsew", pady=(10, 0))
        log_frame.grid_rowconfigure(0, weight=1)
        log_frame.grid_columnconfigure(0, weight=1)
        self.log_box = ScrolledText(
            log_frame,
            height=14,
            state="disabled",
            bg=THEME["log_bg"],
            fg=THEME["text"],
            insertbackground=THEME["text"],
            selectbackground=THEME["accent"],
            selectforeground=THEME["accent_text"],
            relief="flat",
            borderwidth=0,
            padx=10,
            pady=10,
            font=("Consolas", 10),
        )
        self.log_box.grid(row=0, column=0, sticky="nsew")

        footer = ttk.Frame(outer, style="Card.TFrame", padding=(18, 14))
        footer.grid(row=2, column=0, sticky="ew", pady=(16, 0))
        footer.columnconfigure(0, weight=1)

        progress_area = ttk.Frame(footer, style="Card.TFrame")
        progress_area.grid(row=0, column=0, sticky="ew")
        progress_area.columnconfigure(0, weight=1)
        ttk.Label(progress_area, text="진행 상태", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(
            progress_area,
            variable=self.progress_var,
            maximum=1.0,
            style="Panel.Horizontal.TProgressbar",
        )
        self.progress.grid(row=1, column=0, sticky="ew", pady=(10, 8))

        progress_meta = ttk.Frame(progress_area, style="Card.TFrame")
        progress_meta.grid(row=2, column=0, sticky="ew")
        progress_meta.columnconfigure(1, weight=1)
        ttk.Label(progress_meta, text="완료:", style="Hint.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(progress_meta, textvariable=self.progress_text, style="Value.TLabel").grid(
            row=0, column=1, sticky="w", padx=(8, 0)
        )
        ttk.Label(progress_meta, text="예상 남은 시간:", style="Hint.TLabel").grid(row=0, column=2, sticky="w", padx=(24, 0))
        ttk.Label(progress_meta, textvariable=self.eta_text, style="Value.TLabel").grid(
            row=0, column=3, sticky="w", padx=(8, 0)
        )

        button_area = ttk.Frame(footer, style="Card.TFrame")
        button_area.grid(row=0, column=1, sticky="e", padx=(16, 0))
        self.cancel_btn = ttk.Button(button_area, text="중단", command=self.cancel_run, style="Secondary.TButton")
        self.cancel_btn.grid(row=0, column=0, padx=(0, 10))
        self.run_btn = ttk.Button(button_area, text="수행", command=self.run, style="Accent.TButton")
        self.run_btn.grid(row=0, column=1)

    def pick_main_file(self):
        path = filedialog.askopenfilename(
            title="Select main Excel file",
            filetypes=[("Excel Macro-Enabled", "*.xlsm"), ("Excel Workbook", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.main_file_var.set(path)

    def pick_folder(self):
        path = filedialog.askdirectory(title="Select target folder")
        if path:
            self.folder_var.set(path)
            self.refresh_file_list()

    def refresh_file_list(self):
        folder = Path(self.folder_var.get())
        self.file_list.delete(0, tk.END)
        if not folder.exists():
            return
        for name in sorted(os.listdir(folder), key=str.lower):
            if name.lower().endswith((".xlsx", ".xlsm")):
                self.file_list.insert(tk.END, name)

    def run(self):
        if self.running:
            return

        main_file = Path(self.main_file_var.get()).expanduser()
        sheet_names = [s.strip() for s in self.sheet_names_var.get().split(",") if s.strip()]
        exclude_sheet_names = [
            s.strip() for s in self.exclude_sheet_names_var.get().split(",") if s.strip()
        ]
        precondition_enabled = bool(self.precondition_var.get())
        folder = Path(self.folder_var.get()).expanduser()
        output_subfolder = self.output_subfolder_var.get().strip() or "_output"

        normalize = lambda value: unicodedata.normalize("NFKC", str(value)).strip().casefold()
        overlap = []
        seen_overlap = set()
        exclude_keys = {normalize(name) for name in exclude_sheet_names}
        for name in sheet_names:
            key = normalize(name)
            if key in exclude_keys and key not in seen_overlap:
                overlap.append(name)
                seen_overlap.add(key)

        if not main_file.is_file():
            messagebox.showerror("Missing main file", "Please select a valid main Excel file.")
            return
        if not folder.is_dir():
            messagebox.showerror("Missing folder", "Please select a valid target folder.")
            return
        if not sheet_names:
            messagebox.showerror("Sheet names", "Please enter at least one sheet name.")
            return
        if any(sep in output_subfolder for sep in ("\\", "/")):
            messagebox.showerror("Output folder", "Output subfolder name should not contain path separators.")
            return
        if overlap:
            messagebox.showwarning("경고", f"복제 시트와 제거 시트가 겹칩니다: {', '.join(overlap)}")

        self._set_running(True)
        self.cancel_requested = False
        self._log("Starting...")
        self._update_progress(0, 0)
        self._start_time = None
        self.eta_text.set("--")

        thread = threading.Thread(
            target=self._run_in_thread,
            args=(main_file, sheet_names, exclude_sheet_names, precondition_enabled, folder, output_subfolder),
            daemon=True,
        )
        thread.start()

    def cancel_run(self):
        if not self.running:
            return
        if messagebox.askyesno("Cancel", "중지하시겠습니까?"):
            self.cancel_requested = True
            self._log("Cancellation requested. Stopping immediately; current file may be discarded.")

    def _run_in_thread(self, main_file, sheet_names, exclude_sheet_names, precondition_enabled, folder, output_subfolder):
        copier = ExcelCopier(self._log)
        try:
            copier.run(
                main_file,
                sheet_names,
                exclude_sheet_names,
                precondition_enabled,
                folder,
                output_subfolder,
                progress_cb=self._update_progress,
                cancel_cb=lambda: self.cancel_requested,
            )
            if self.cancel_requested:
                self._log("Processing stopped by user. Current file aborted if mid-copy.")
                self._show_message_async("중단", "사용자 요청으로 즉시 중지했습니다.")
            else:
                self._log("Completed successfully.")
                self._show_message_async("완료", "모든 파일 처리가 끝났습니다.")
            self.root.after(0, self.refresh_file_list)
        except Exception as exc:
            msg = f"Error: {exc}"
            self._log(msg)
            self._log(traceback.format_exc())
            self._show_message_async("오류", msg)
        finally:
            self._set_running(False)

    def _set_running(self, flag: bool):
        def toggle():
            self.running = flag
            run_state = "disabled" if flag else "normal"
            cancel_state = "normal" if flag else "disabled"
            for widget in (self.run_btn, self.refresh_btn):
                widget.configure(state=run_state)
            self.cancel_btn.configure(state=cancel_state)
        self.root.after(0, toggle)

    def _log(self, message: str):
        def append():
            self.log_box.configure(state="normal")
            self.log_box.insert(tk.END, message + "\n")
            self.log_box.see(tk.END)
            self.log_box.configure(state="disabled")
        self.root.after(0, append)

    def _show_message_async(self, title, message):
        self.root.after(0, lambda: messagebox.showinfo(title, message))

    def _update_progress(self, current: int, total: int):
        def apply():
            if total <= 0:
                self.progress_var.set(0.0)
                self.progress_text.set("0/0")
                self.eta_text.set("--")
            else:
                frac = min(max(current / total, 0.0), 1.0)
                self.progress_var.set(frac)
                self.progress_text.set(f"{current}/{total}")
                now = time.perf_counter()
                if self._start_time is None:
                    self._start_time = now
                elapsed = now - self._start_time
                if current <= 0 or elapsed <= 0:
                    self.eta_text.set("--")
                else:
                    remaining = total - current
                    per_item = elapsed / current
                    eta_sec = remaining * per_item
                    if eta_sec >= 3600:
                        self.eta_text.set(f"{eta_sec/3600:.1f}h 남음")
                    elif eta_sec >= 120:
                        self.eta_text.set(f"{eta_sec/60:.0f}m 남음")
                    elif eta_sec >= 60:
                        self.eta_text.set(f"{eta_sec/60:.1f}m 남음")
                    else:
                        self.eta_text.set(f"{eta_sec:.0f}s 남음")
        self.root.after(0, apply)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
