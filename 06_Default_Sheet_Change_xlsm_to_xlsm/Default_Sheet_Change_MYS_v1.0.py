import os
import shutil
import threading
import traceback
import time
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import pywintypes
import win32com.client as win32


class ExcelCopier:
    """Handle the Excel work so the UI stays clean."""

    def __init__(self, logger):
        self.logger = logger
        self.constants = self._build_constants()

    def run(self, main_file: Path, sheet_names, target_folder: Path, output_subfolder: str, progress_cb=None, cancel_cb=None):
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

                completed = self._copy_multiple_sheets(excel, source_path, temp_target, sheet_names, cancel_cb=cancel_cb)
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

                self.logger(f"[{idx}/{len(excel_files)}] Done -> {file_name}")
                if progress_cb:
                    progress_cb(idx, total)
                # cancellation check happens at top of next iteration
        finally:
            excel.DisplayAlerts = True
            excel.EnableEvents = True
            excel.Quit()

    def _collect_excel_files(self, folder: Path):
        files = []
        for entry in folder.iterdir():
            if entry.is_file() and entry.suffix.lower() in {".xlsx", ".xlsm"}:
                files.append(entry.name)
        return sorted(files, key=str.lower)

    def _copy_multiple_sheets(self, excel, source_path: Path, target_path: Path, sheet_names, cancel_cb=None):
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

                self._copy_sheet_data(excel, source_sheet, target_sheet)
                if cancel_cb and cancel_cb():
                    return False

            # Copy missing defined names to avoid #NAME? errors when formulas rely on them
            self._copy_missing_names(source_wb, target_wb)
            try:
                excel.CalculateFullRebuild()
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
        start_row = source_range.Row
        start_col = source_range.Column
        end_row = start_row + source_range.Rows.Count - 1
        end_col = start_col + source_range.Columns.Count - 1
        try:
            last_cell = source_sheet.Cells.SpecialCells(self.constants["xlCellTypeLastCell"])
            end_row = max(end_row, last_cell.Row)
            end_col = max(end_col, last_cell.Column)
        except Exception:
            pass

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

        # Column widths
        for col_idx in range(source_range.Columns.Count):
            src_col = source_range.Column + col_idx
            target_sheet.Columns(src_col).ColumnWidth = source_sheet.Columns(src_col).ColumnWidth

        # Row heights
        for row_idx in range(source_range.Rows.Count):
            src_row = source_range.Row + row_idx
            target_sheet.Rows(src_row).RowHeight = source_sheet.Rows(src_row).RowHeight

        # Preserve outline/group settings within used area (and last used cell)
        self._copy_outline(source_sheet, target_sheet, start_row, end_row, start_col, end_col)

    def _get_sheet(self, workbook, name):
        try:
            return workbook.Worksheets(name)
        except Exception:
            return None

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

    def _copy_outline(self, source_sheet, target_sheet, start_row, end_row, start_col, end_col):
        try:
            target_sheet.Outline.SummaryRow = source_sheet.Outline.SummaryRow
            target_sheet.Outline.SummaryColumn = source_sheet.Outline.SummaryColumn
        except Exception:
            pass

        for row_idx in range(start_row, end_row + 1):
            try:
                src_row = source_sheet.Rows(row_idx)
                tgt_row = target_sheet.Rows(row_idx)
                tgt_row.OutlineLevel = src_row.OutlineLevel
                tgt_row.Hidden = src_row.Hidden
            except Exception:
                pass

        for col_idx in range(start_col, end_col + 1):
            try:
                src_col = source_sheet.Columns(col_idx)
                tgt_col = target_sheet.Columns(col_idx)
                tgt_col.OutlineLevel = src_col.OutlineLevel
                tgt_col.Hidden = src_col.Hidden
            except Exception:
                pass

    def _copy_missing_names(self, source_wb, target_wb):
        existing = {n.Name.lower() for n in target_wb.Names}
        for name_obj in source_wb.Names:
            key = name_obj.Name.lower()
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
        consts = {"xlPasteFormats": -4122, "xlCellTypeLastCell": 11}
        try:
            from win32com.client import constants as excel_consts
            consts["xlPasteFormats"] = getattr(excel_consts, "xlPasteFormats", consts["xlPasteFormats"])
            consts["xlCellTypeLastCell"] = getattr(excel_consts, "xlCellTypeLastCell", consts["xlCellTypeLastCell"])
        except Exception:
            pass
        return consts


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Default Sheet Change MYS v1.0")
        self.root.geometry("700x520")

        self.main_file_var = tk.StringVar()
        self.sheet_names_var = tk.StringVar(value="Test Case,Test Setup,History")
        self.folder_var = tk.StringVar()
        self.output_subfolder_var = tk.StringVar(value="_output")
        self.running = False
        self.cancel_requested = False
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_text = tk.StringVar(value="0/0")
        self.eta_text = tk.StringVar(value="--")
        self._start_time = None

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}

        frm = ttk.Frame(self.root)
        frm.pack(fill="both", expand=True, padx=12, pady=12)

        # Main file selector
        ttk.Label(frm, text="Main Excel (기준 파일):").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.main_file_var, width=70).grid(row=0, column=1, sticky="we", **pad)
        ttk.Button(frm, text="Browse", command=self.pick_main_file).grid(row=0, column=2, **pad)

        # Sheet names
        ttk.Label(frm, text="Sheet name(s) (comma separated):").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.sheet_names_var, width=70).grid(row=1, column=1, sticky="we", **pad)

        # Target folder
        ttk.Label(frm, text="Target folder (대상 폴더):").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.folder_var, width=70).grid(row=2, column=1, sticky="we", **pad)
        ttk.Button(frm, text="Browse", command=self.pick_folder).grid(row=2, column=2, **pad)

        # Output subfolder name
        ttk.Label(frm, text="Output subfolder (대상 폴더 하위):").grid(row=3, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.output_subfolder_var, width=70).grid(row=3, column=1, sticky="we", **pad)

        # File list
        ttk.Label(frm, text="Excel files found in folder:").grid(row=4, column=0, sticky="w", **pad)
        self.file_list = tk.Listbox(frm, height=6)
        self.file_list.grid(row=4, column=1, sticky="we", **pad)
        frm.grid_columnconfigure(1, weight=1)

        self.refresh_btn = ttk.Button(frm, text="Refresh list", command=self.refresh_file_list)
        self.refresh_btn.grid(row=4, column=2, sticky="n", **pad)

        # Progress bar
        ttk.Label(frm, text="Progress:").grid(row=5, column=0, sticky="w", **pad)
        prog_frame = ttk.Frame(frm)
        prog_frame.grid(row=5, column=1, sticky="we", **pad)
        prog_frame.columnconfigure(0, weight=1)
        self.progress = ttk.Progressbar(prog_frame, variable=self.progress_var, maximum=1.0)
        self.progress.grid(row=0, column=0, sticky="we")
        ttk.Label(prog_frame, textvariable=self.progress_text, width=8).grid(row=0, column=1, sticky="e", padx=(8, 0))

        # ETA row
        ttk.Label(frm, text="ETA:").grid(row=6, column=0, sticky="w", **pad)
        ttk.Label(frm, textvariable=self.eta_text).grid(row=6, column=1, sticky="w", **pad)

        # Cancel / Run buttons (stacked)
        self.cancel_btn = ttk.Button(frm, text="Cancel", command=self.cancel_run)
        self.cancel_btn.grid(row=5, column=2, sticky="new", padx=8, pady=(4, 2))
        self.run_btn = ttk.Button(frm, text="Run", command=self.run)
        self.run_btn.grid(row=6, column=2, sticky="new", padx=8, pady=(2, 6), ipady=6)

        # Log area
        ttk.Label(frm, text="Log:").grid(row=7, column=0, sticky="nw", **pad)
        self.log_box = ScrolledText(frm, height=12, state="disabled")
        self.log_box.grid(row=7, column=1, columnspan=2, sticky="nsew", **pad)
        frm.grid_rowconfigure(7, weight=1)

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
        folder = Path(self.folder_var.get()).expanduser()
        output_subfolder = self.output_subfolder_var.get().strip() or "_output"

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

        self._set_running(True)
        self.cancel_requested = False
        self._log("Starting...")
        self._update_progress(0, 0)
        self._start_time = None
        self.eta_text.set("--")

        thread = threading.Thread(
            target=self._run_in_thread,
            args=(main_file, sheet_names, folder, output_subfolder),
            daemon=True,
        )
        thread.start()

    def cancel_run(self):
        if not self.running:
            return
        if messagebox.askyesno("Cancel", "중지하시겠습니까?"):
            self.cancel_requested = True
            self._log("Cancellation requested. Stopping immediately; current file may be discarded.")

    def _run_in_thread(self, main_file, sheet_names, folder, output_subfolder):
        copier = ExcelCopier(self._log)
        try:
            copier.run(
                main_file,
                sheet_names,
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
        self.root.after(0, apply)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
