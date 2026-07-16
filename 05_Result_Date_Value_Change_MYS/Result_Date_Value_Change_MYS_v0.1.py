import os
import queue
import re
import signal
import sys
import threading
import time
import traceback
from datetime import datetime
from pathlib import Path
from tkinter import (
    BooleanVar,
    Listbox,
    StringVar,
    Tk,
    filedialog,
    messagebox,
)
from tkinter import ttk

try:
    import pythoncom
    import win32com.client
except ImportError:
    pythoncom = None
    win32com = None


APP_TITLE = "Result_Date_Value_Change_MYS_v0.1"
APP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
RUN_LOG_PATH = APP_DIR / "Result_Date_Value_Change_MYS_v0.1_run.log"
ERROR_LOG_PATH = APP_DIR / "Result_Date_Value_Change_MYS_v0.1_error.log"
EXCEL_EXTENSIONS = {".xlsx", ".xlsm"}
TIMESTAMP_SUFFIX_PATTERN = re.compile(r"^(?P<prefix>.+_)(?P<timestamp>\d+)$")
EXCEL_SERIAL_EPOCH = datetime(1899, 12, 30)


def write_run_log(message: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with RUN_LOG_PATH.open("a", encoding="utf-8") as log_file:
        log_file.write(f"[{timestamp}] {message}\n")


class ResultDateValueChangeApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1120x830")
        self.root.minsize(1040, 800)
        self.root.configure(bg="#eef6ff")
        self.root.option_add("*Font", ("Malgun Gothic", 10))

        self.folder_path = StringVar()
        self.include_subfolders = BooleanVar(value=False)
        self.change_summary = BooleanVar(value=True)
        self.change_filename = BooleanVar(value=True)
        self.status_text = StringVar(value="폴더를 선택하세요.")
        self.remaining_text = StringVar(value="남은 시간: -")

        self.summary_vars = [StringVar() for _ in range(6)]
        self.filename_vars = [StringVar() for _ in range(6)]
        self.files: list[Path] = []
        self.progress_queue: queue.Queue[tuple] = queue.Queue()
        self.is_running = False

        self._configure_styles()
        self._build_ui()
        self.fill_current_datetime()

    def _configure_styles(self) -> None:
        self.style = ttk.Style(self.root)
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        self.style.configure(".", font=("Malgun Gothic", 10))
        self.style.configure("App.TFrame", background="#eef6ff")
        self.style.configure("Card.TFrame", background="#ffffff", relief="solid", borderwidth=1)
        self.style.configure("Row.TFrame", background="#ffffff")
        self.style.configure("Header.TLabel", background="#eef6ff", foreground="#123a63", font=("Malgun Gothic", 18, "bold"))
        self.style.configure("SubHeader.TLabel", background="#eef6ff", foreground="#45627f", font=("Malgun Gothic", 10))
        self.style.configure("Section.TLabel", background="#ffffff", foreground="#123a63", font=("Malgun Gothic", 11, "bold"))
        self.style.configure("Text.TLabel", background="#ffffff", foreground="#1f2937")
        self.style.configure("Status.TLabel", background="#eef6ff", foreground="#123a63", font=("Malgun Gothic", 10, "bold"))
        self.style.configure("Muted.TLabel", background="#eef6ff", foreground="#45627f")
        self.style.configure("Card.TCheckbutton", background="#ffffff", foreground="#123a63", padding=(2, 4))
        self.style.configure("Primary.TButton", background="#1d64b8", foreground="#ffffff", padding=(18, 9), font=("Malgun Gothic", 10, "bold"))
        self.style.configure("Secondary.TButton", background="#dbeafe", foreground="#123a63", padding=(18, 9), font=("Malgun Gothic", 10, "bold"))
        self.style.configure("Date.TEntry", padding=(5, 4), fieldbackground="#f8fbff", foreground="#111827")
        self.style.configure("Path.TEntry", padding=(6, 5), fieldbackground="#f8fbff", foreground="#111827")
        self.style.configure("Blue.Horizontal.TProgressbar", troughcolor="#dbeafe", background="#2563eb", lightcolor="#2563eb", darkcolor="#2563eb")
        self.style.map("Primary.TButton", background=[("active", "#174f91"), ("disabled", "#93b4d8")])
        self.style.map("Secondary.TButton", background=[("active", "#bfdbfe"), ("disabled", "#e5e7eb")])

    def _build_ui(self) -> None:
        root_frame = ttk.Frame(self.root, padding=(22, 18, 22, 18), style="App.TFrame")
        root_frame.pack(fill="both", expand=True)

        ttk.Label(root_frame, text="결과 리포트 날짜/파일명 변경", style="Header.TLabel").pack(anchor="w")
        ttk.Label(
            root_frame,
            text=".xlsx / .xlsm 파일의 summary B7 값과 파일명 타임스탬프를 일괄 변경합니다.",
            style="SubHeader.TLabel",
        ).pack(anchor="w", pady=(2, 14))

        folder_card = ttk.Frame(root_frame, padding=(16, 14, 16, 14), style="Card.TFrame")
        folder_card.pack(fill="x", pady=(0, 12))

        ttk.Label(folder_card, text="대상 폴더", style="Section.TLabel").pack(anchor="w", pady=(0, 8))
        folder_frame = ttk.Frame(folder_card, style="Row.TFrame")
        folder_frame.pack(fill="x")

        ttk.Button(folder_frame, text="폴더 선택", width=16, command=self.select_folder, style="Primary.TButton").pack(side="left")
        ttk.Entry(folder_frame, textvariable=self.folder_path, state="readonly", style="Path.TEntry").pack(
            side="left", fill="x", expand=True, padx=(10, 0), ipady=3
        )

        ttk.Checkbutton(
            folder_card,
            text="하위 폴더 포함",
            variable=self.include_subfolders,
            command=self.refresh_file_list,
            style="Card.TCheckbutton",
        ).pack(anchor="w", pady=(10, 0))

        list_frame = ttk.Frame(root_frame, padding=(16, 14, 16, 14), style="Card.TFrame")
        list_frame.pack(fill="both", expand=True, pady=(0, 12))

        ttk.Label(list_frame, text="선정된 파일 리스트", style="Section.TLabel").pack(anchor="w")
        list_inner = ttk.Frame(list_frame, style="Row.TFrame")
        list_inner.pack(fill="both", expand=True, pady=(8, 0))
        list_inner.grid_rowconfigure(0, weight=1)
        list_inner.grid_columnconfigure(0, weight=1)

        self.file_listbox = Listbox(
            list_inner,
            height=13,
            bg="#f8fbff",
            fg="#111827",
            selectbackground="#2563eb",
            selectforeground="#ffffff",
            activestyle="none",
            relief="flat",
            highlightthickness=1,
            highlightcolor="#93c5fd",
            highlightbackground="#c7dbf2",
            font=("Consolas", 10),
        )
        self.file_listbox.grid(row=0, column=0, sticky="nsew")

        y_scrollbar = ttk.Scrollbar(list_inner, orient="vertical", command=self.file_listbox.yview)
        y_scrollbar.grid(row=0, column=1, sticky="ns")
        x_scrollbar = ttk.Scrollbar(list_inner, orient="horizontal", command=self.file_listbox.xview)
        x_scrollbar.grid(row=1, column=0, sticky="ew")
        self.file_listbox.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)

        date_card = ttk.Frame(root_frame, padding=(16, 14, 16, 14), style="Card.TFrame")
        date_card.pack(fill="x", pady=(0, 12))
        ttk.Label(date_card, text="변경 값 입력", style="Section.TLabel").pack(anchor="w", pady=(0, 8))

        self._build_datetime_row(
            date_card,
            check_var=self.change_summary,
            label="Summary 변경",
            values=self.summary_vars,
        )
        self._build_datetime_row(
            date_card,
            check_var=self.change_filename,
            label="파일명 변경",
            values=self.filename_vars,
        )

        action_frame = ttk.Frame(root_frame, style="App.TFrame")
        action_frame.pack(fill="x", pady=(2, 10))

        ttk.Button(action_frame, text="갱신", width=14, command=self.fill_current_datetime, style="Secondary.TButton").pack(side="left")
        ttk.Button(action_frame, text="실행", width=14, command=self.run, style="Primary.TButton").pack(side="left", padx=(10, 0))

        self.progress_bar = ttk.Progressbar(root_frame, orient="horizontal", mode="determinate", style="Blue.Horizontal.TProgressbar")
        self.progress_bar.pack(fill="x", pady=(8, 4))

        progress_text_frame = ttk.Frame(root_frame, style="App.TFrame")
        progress_text_frame.pack(fill="x")
        ttk.Label(progress_text_frame, textvariable=self.status_text, style="Status.TLabel").pack(side="left")
        ttk.Label(progress_text_frame, textvariable=self.remaining_text, style="Muted.TLabel").pack(side="right")

    def _build_datetime_row(
        self,
        parent: ttk.Frame,
        check_var: BooleanVar,
        label: str,
        values: list[StringVar],
    ) -> None:
        row = ttk.Frame(parent, style="Row.TFrame")
        row.pack(fill="x", pady=5)

        ttk.Checkbutton(row, variable=check_var, style="Card.TCheckbutton").grid(row=0, column=0, sticky="w")
        ttk.Label(row, text=label, width=14, anchor="w", style="Text.TLabel").grid(row=0, column=1, sticky="w", padx=(8, 16))

        labels = ["년", "월", "일", "시 (0~23)", "분", "초"]
        widths = [8, 6, 6, 8, 6, 6]
        for idx, (caption, width) in enumerate(zip(labels, widths)):
            field = ttk.Frame(row, style="Row.TFrame")
            field.grid(row=0, column=idx + 2, sticky="w", padx=(0, 12))
            ttk.Entry(field, textvariable=values[idx], width=width, justify="center", style="Date.TEntry").pack(side="left", ipady=2)
            ttk.Label(field, text=caption, style="Text.TLabel").pack(side="left", padx=(4, 0))

        row.grid_columnconfigure(8, weight=1)

    def fill_current_datetime(self) -> None:
        now = datetime.now()
        values = [
            f"{now.year:04d}",
            f"{now.month:02d}",
            f"{now.day:02d}",
            f"{now.hour:02d}",
            f"{now.minute:02d}",
            f"{now.second:02d}",
        ]
        for target in (self.summary_vars, self.filename_vars):
            for var, value in zip(target, values):
                var.set(value)

    def select_folder(self) -> None:
        self.clear_selection()
        selected = filedialog.askdirectory(title="변경할 폴더 선택")
        if not selected:
            self.status_text.set("폴더 선택이 취소되었습니다.")
            return
        self.folder_path.set(selected)
        self.refresh_file_list()

    def clear_selection(self) -> None:
        self.folder_path.set("")
        self.files = []
        self.file_listbox.delete(0, "end")
        self.progress_bar["value"] = 0
        self.remaining_text.set("남은 시간: -")
        self.status_text.set("폴더를 선택하세요.")

    def refresh_file_list(self) -> None:
        self.files = []
        self.file_listbox.delete(0, "end")

        folder = self.folder_path.get()
        if not folder:
            return

        base = Path(folder)
        if not base.exists():
            self.status_text.set("선택한 폴더가 존재하지 않습니다.")
            return

        iterator = base.rglob("*") if self.include_subfolders.get() else base.glob("*")
        self.files = sorted(
            path
            for path in iterator
            if path.is_file()
            and path.suffix.lower() in EXCEL_EXTENSIONS
            and not path.name.startswith("~$")
        )

        for path in self.files:
            display_path = path.relative_to(base) if path.is_relative_to(base) else path
            self.file_listbox.insert("end", str(display_path))

        self.status_text.set(f"선정된 파일: {len(self.files)}개")

    def run(self) -> None:
        if self.is_running:
            return
        if not self.change_summary.get() and not self.change_filename.get():
            messagebox.showerror(APP_TITLE, "하나 이상의 기능을 체크해야 수행 가능합니다.")
            return
        if not self.files:
            messagebox.showerror(APP_TITLE, "대상 엑셀 파일이 없습니다. 폴더를 선택하세요.")
            return

        try:
            summary_dt = self.parse_datetime(self.summary_vars) if self.change_summary.get() else None
            filename_dt = self.parse_datetime(self.filename_vars) if self.change_filename.get() else None
        except ValueError as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return

        self.is_running = True
        self.progress_bar["value"] = 0
        self.status_text.set("실행 중...")
        self.remaining_text.set("남은 시간: 계산 중")

        worker = threading.Thread(
            target=self.process_files,
            args=(summary_dt, filename_dt),
            daemon=True,
        )
        worker.start()
        self.root.after(100, self.consume_progress_queue)

    def parse_datetime(self, vars_: list[StringVar]) -> datetime:
        raw = [var.get().strip() for var in vars_]
        try:
            year, month, day, hour, minute, second = [int(value) for value in raw[:6]]
        except ValueError as exc:
            raise ValueError("날짜/시간 입력 칸에는 숫자를 입력하세요.") from exc

        if not 0 <= hour <= 23:
            raise ValueError("시간은 24시간 기준 0~23 범위로 입력하세요.")

        try:
            return datetime(year, month, day, hour, minute, second)
        except ValueError as exc:
            raise ValueError(f"유효하지 않은 날짜/시간입니다: {exc}") from exc

    def process_files(self, summary_dt: datetime | None, filename_dt: datetime | None) -> None:
        total_steps = len(self.files) * int(summary_dt is not None) + len(self.files) * int(filename_dt is not None)
        completed = 0
        start_time = time.time()
        failures: list[str] = []
        current_files = list(self.files)

        if summary_dt is not None:
            completed = self.change_summary_cells(
                current_files,
                summary_dt,
                failures,
                completed,
                total_steps,
                start_time,
            )

        if filename_dt is not None:
            renamed_files: list[Path] = []
            timestamp = filename_dt.strftime("%Y%m%d%H%M%S")
            for file_path in current_files:
                try:
                    renamed_path = self.change_file_name(file_path, timestamp)
                    if renamed_path is not None:
                        renamed_files.append(renamed_path)
                except Exception as exc:
                    failures.append(f"{file_path.name}: 파일명 변경 실패 - {exc}")
                completed += 1
                self.progress_queue.put(("progress", completed, total_steps, start_time))

            if renamed_files:
                self.files = renamed_files

        self.progress_queue.put(("done", failures))

    def change_summary_cells(
        self,
        file_paths: list[Path],
        dt_value: datetime,
        failures: list[str],
        completed: int,
        total_steps: int,
        start_time: float,
    ) -> int:
        if pythoncom is None or win32com is None:
            for file_path in file_paths:
                failures.append(
                    f"{file_path.name}: Summary 변경 실패 - pywin32가 설치되어 있지 않아 Excel 저장을 사용할 수 없습니다."
                )
                completed += 1
                self.progress_queue.put(("progress", completed, total_steps, start_time))
            return completed

        pythoncom.CoInitialize()
        excel = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.EnableEvents = False
            try:
                excel.AskToUpdateLinks = False
            except Exception:
                pass

            for file_path in file_paths:
                workbook = None
                try:
                    workbook = excel.Workbooks.Open(
                        str(file_path.resolve()),
                        UpdateLinks=0,
                        ReadOnly=False,
                        IgnoreReadOnlyRecommended=True,
                        AddToMru=False,
                    )
                    worksheet = workbook.Worksheets("summary")
                    cell = worksheet.Range("B7")
                    cell.Value2 = self.datetime_to_excel_serial(dt_value)
                    cell.NumberFormat = "yyyy-mm-dd  h:mm:ss AM/PM"
                    workbook.Save()
                except Exception as exc:
                    failures.append(f"{file_path.name}: Summary 변경 실패 - {exc}")
                finally:
                    if workbook is not None:
                        try:
                            workbook.Close(SaveChanges=False)
                        except Exception:
                            pass
                    completed += 1
                    self.progress_queue.put(("progress", completed, total_steps, start_time))
        finally:
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()

        return completed

    @staticmethod
    def datetime_to_excel_serial(dt_value: datetime) -> float:
        delta = dt_value - EXCEL_SERIAL_EPOCH
        return delta.days + (delta.seconds + delta.microseconds / 1_000_000) / 86400

    def change_file_name(self, file_path: Path, timestamp: str) -> Path | None:
        stem = file_path.stem
        match = TIMESTAMP_SUFFIX_PATTERN.match(stem)
        if not match:
            return None

        new_path = file_path.with_name(f"{match.group('prefix')}{timestamp}{file_path.suffix}")
        if new_path == file_path:
            return file_path
        if new_path.exists():
            raise FileExistsError(f"동일한 파일명이 이미 존재합니다: {new_path.name}")

        os.rename(file_path, new_path)
        return new_path

    def consume_progress_queue(self) -> None:
        try:
            while True:
                message = self.progress_queue.get_nowait()
                if message[0] == "progress":
                    _, completed, total, start_time = message
                    percent = int(completed / total * 100) if total else 100
                    self.progress_bar["value"] = percent
                    self.status_text.set(f"진행률: {percent}% ({completed}/{total})")
                    self.remaining_text.set(f"남은 시간: {self.format_remaining(completed, total, start_time)}")
                elif message[0] == "done":
                    _, failures = message
                    self.is_running = False
                    self.progress_bar["value"] = 100
                    self.remaining_text.set("남은 시간: 0초")
                    self.refresh_file_list()
                    if failures:
                        self.status_text.set(f"완료 - 오류 {len(failures)}건")
                        messagebox.showwarning(APP_TITLE, "일부 파일 처리에 실패했습니다.\n\n" + "\n".join(failures[:20]))
                    else:
                        self.status_text.set("완료")
                        messagebox.showinfo(APP_TITLE, "작업이 완료되었습니다.")
                    return
        except queue.Empty:
            pass

        if self.is_running:
            self.root.after(100, self.consume_progress_queue)

    def format_remaining(self, completed: int, total: int, start_time: float) -> str:
        if completed <= 0 or total <= completed:
            return "계산 중" if completed <= 0 else "0초"

        elapsed = time.time() - start_time
        seconds = max(0, int((elapsed / completed) * (total - completed)))
        minutes, sec = divmod(seconds, 60)
        hours, minutes = divmod(minutes, 60)

        if hours:
            return f"{hours}시간 {minutes}분 {sec}초"
        if minutes:
            return f"{minutes}분 {sec}초"
        return f"{sec}초"


if __name__ == "__main__":
    window = None

    def ignore_sigint(signum, frame) -> None:
        write_run_log("SIGINT ignored. Close the UI window to exit.")

    try:
        signal.signal(signal.SIGINT, ignore_sigint)
        write_run_log("Application starting.")
        window = Tk()
        window.protocol("WM_DELETE_WINDOW", window.destroy)
        app = ResultDateValueChangeApp(window)
        write_run_log("Tk mainloop entered.")
        window.mainloop()
        write_run_log("Tk mainloop exited normally.")
    except Exception:
        ERROR_LOG_PATH.write_text(traceback.format_exc(), encoding="utf-8")
        write_run_log("Application crashed. See error log.")
        raise
