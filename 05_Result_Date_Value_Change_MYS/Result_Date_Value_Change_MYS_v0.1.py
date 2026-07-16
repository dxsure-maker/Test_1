import os
import queue
import re
import signal
import sys
import threading
import time
import traceback
from datetime import datetime
from html import unescape
from pathlib import Path, PurePosixPath
from tkinter import (
    BooleanVar,
    Listbox,
    StringVar,
    Tk,
    filedialog,
    messagebox,
)
from tkinter import ttk
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile


APP_TITLE = "Result_Date_Value_Change_MYS_v0.1"
APP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
RUN_LOG_PATH = APP_DIR / "Result_Date_Value_Change_MYS_v0.1_run.log"
ERROR_LOG_PATH = APP_DIR / "Result_Date_Value_Change_MYS_v0.1_error.log"
EXCEL_EXTENSIONS = {".xlsx", ".xlsm"}
TIMESTAMP_SUFFIX_PATTERN = re.compile(r"^(?P<prefix>.+_)(?P<timestamp>\d+)$")
EXCEL_SERIAL_EPOCH = datetime(1899, 12, 30)
SUMMARY_DATE_FORMAT = "yyyy-mm-dd  h:mm:ss AM/PM"
SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
REL_ID_ATTR = f"{{{OFFICE_REL_NS}}}id"


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
        for file_path in file_paths:
            try:
                self.change_summary_cell(file_path, dt_value)
            except Exception as exc:
                failures.append(f"{file_path.name}: Summary 변경 실패 - {exc}")
            completed += 1
            self.progress_queue.put(("progress", completed, total_steps, start_time))

        return completed

    def change_summary_cell(self, file_path: Path, dt_value: datetime) -> None:
        serial_text = self.format_excel_serial(self.datetime_to_excel_serial(dt_value))
        temp_path = file_path.with_name(f"{file_path.name}.tmp")
        summary_sheet_path = None
        patched_styles_xml = None
        date_style_id = None

        try:
            with ZipFile(file_path, "r") as source_zip:
                summary_sheet_path = self.find_summary_sheet_path(source_zip)
                summary_sheet_xml = source_zip.read(summary_sheet_path)
                base_style_id = self.get_b7_style_id(summary_sheet_xml)
                if "xl/styles.xml" in source_zip.namelist():
                    patched_styles_xml, date_style_id = self.ensure_summary_date_style(
                        source_zip.read("xl/styles.xml"),
                        base_style_id,
                    )
                with ZipFile(temp_path, "w", compression=ZIP_DEFLATED, allowZip64=True) as target_zip:
                    for item in source_zip.infolist():
                        data = source_zip.read(item.filename)
                        if item.filename == summary_sheet_path:
                            data = self.patch_b7_value(data, serial_text, date_style_id)
                        elif item.filename == "xl/styles.xml" and patched_styles_xml is not None:
                            data = patched_styles_xml
                        target_zip.writestr(item, data)
                    target_zip.comment = source_zip.comment
            os.replace(temp_path, file_path)
        except Exception:
            if temp_path.exists():
                try:
                    temp_path.unlink()
                except Exception:
                    pass
            raise

        if summary_sheet_path is None:
            raise ValueError("summary 시트 XML을 찾지 못했습니다.")

    def find_summary_sheet_path(self, workbook_zip: ZipFile) -> str:
        workbook_root = ET.fromstring(workbook_zip.read("xl/workbook.xml"))
        summary_rel_id = None
        for sheet in workbook_root.findall(f".//{{{SPREADSHEET_NS}}}sheet"):
            if sheet.attrib.get("name", "").casefold() == "summary":
                summary_rel_id = sheet.attrib.get(REL_ID_ATTR)
                break
        if not summary_rel_id:
            raise ValueError("summary 시트가 없습니다.")

        rels_root = ET.fromstring(workbook_zip.read("xl/_rels/workbook.xml.rels"))
        for rel in rels_root.findall(f"{{{PACKAGE_REL_NS}}}Relationship"):
            if rel.attrib.get("Id") == summary_rel_id:
                target = rel.attrib.get("Target", "")
                if target.startswith("/"):
                    return target.lstrip("/")
                return PurePosixPath("xl", target).as_posix()

        raise ValueError("summary 시트 관계 정보를 찾지 못했습니다.")

    @staticmethod
    def get_b7_style_id(sheet_xml: bytes) -> int:
        encoding = ResultDateValueChangeApp.detect_xml_encoding(sheet_xml)
        xml_text = sheet_xml.decode(encoding)
        match = re.search(r"<c\b(?=[^>]*\br=(['\"])B7\1)[^>]*\bs=(['\"])(\d+)\2", xml_text)
        return int(match.group(3)) if match else 0

    def ensure_summary_date_style(self, styles_xml: bytes, base_style_id: int) -> tuple[bytes, int]:
        encoding = self.detect_xml_encoding(styles_xml)
        styles_text = styles_xml.decode(encoding)
        target_num_fmt_id = self.find_num_fmt_id(styles_text, SUMMARY_DATE_FORMAT)
        if target_num_fmt_id is None:
            target_num_fmt_id = self.next_custom_num_fmt_id(styles_text)
            styles_text = self.insert_num_fmt(styles_text, target_num_fmt_id, SUMMARY_DATE_FORMAT)

        cell_xfs_match = re.search(r"<cellXfs\b[^>]*>(.*?)</cellXfs>", styles_text, re.DOTALL)
        if not cell_xfs_match:
            raise ValueError("styles.xml에서 cellXfs를 찾지 못했습니다.")

        xf_nodes = re.findall(r"<xf\b[^>]*(?:/>|>.*?</xf>)", cell_xfs_match.group(1), re.DOTALL)
        if not xf_nodes:
            raise ValueError("styles.xml에 셀 스타일 정보가 없습니다.")

        base_xf = xf_nodes[base_style_id] if 0 <= base_style_id < len(xf_nodes) else xf_nodes[0]
        expected_xf = self.with_date_num_fmt(base_xf, target_num_fmt_id)
        for index, xf_node in enumerate(xf_nodes):
            if xf_node == expected_xf:
                return styles_text.encode(encoding), index

        new_style_id = len(xf_nodes)
        styles_text = styles_text[: cell_xfs_match.end(1)] + expected_xf + styles_text[cell_xfs_match.end(1) :]
        updated_start_tag = self.set_xml_attr(cell_xfs_match.group(0).split(">", 1)[0] + ">", "count", str(new_style_id + 1))
        styles_text = styles_text[: cell_xfs_match.start()] + updated_start_tag + styles_text[cell_xfs_match.start(1) :]
        return styles_text.encode(encoding), new_style_id

    @staticmethod
    def find_num_fmt_id(styles_text: str, format_code: str) -> int | None:
        for match in re.finditer(r"<numFmt\b[^>]*\bnumFmtId=(['\"])(\d+)\1[^>]*\bformatCode=(['\"])(.*?)\3[^>]*/?>", styles_text, re.DOTALL):
            if unescape(match.group(4)) == format_code:
                return int(match.group(2))
        return None

    @staticmethod
    def next_custom_num_fmt_id(styles_text: str) -> int:
        ids = [int(match.group(2)) for match in re.finditer(r"\bnumFmtId=(['\"])(\d+)\1", styles_text)]
        return max([163, *ids]) + 1

    def insert_num_fmt(self, styles_text: str, num_fmt_id: int, format_code: str) -> str:
        escaped_format_code = escape(format_code, {'"': "&quot;"})
        num_fmt_xml = f'<numFmt numFmtId="{num_fmt_id}" formatCode="{escaped_format_code}"/>'
        num_fmts_match = re.search(r"<numFmts\b[^>]*>", styles_text)
        if num_fmts_match:
            close_index = styles_text.find("</numFmts>", num_fmts_match.end())
            if close_index < 0:
                raise ValueError("styles.xml의 numFmts 종료 태그를 찾지 못했습니다.")
            current_count = len(re.findall(r"<numFmt\b", styles_text[num_fmts_match.end() : close_index]))
            styles_text = styles_text[:close_index] + num_fmt_xml + styles_text[close_index:]
            updated_start_tag = self.set_xml_attr(num_fmts_match.group(0), "count", str(current_count + 1))
            return styles_text[: num_fmts_match.start()] + updated_start_tag + styles_text[num_fmts_match.end() :]

        style_sheet_match = re.search(r"<styleSheet\b[^>]*>", styles_text)
        if not style_sheet_match:
            raise ValueError("styles.xml의 styleSheet 시작 태그를 찾지 못했습니다.")
        num_fmts_xml = f'<numFmts count="1">{num_fmt_xml}</numFmts>'
        return styles_text[: style_sheet_match.end()] + num_fmts_xml + styles_text[style_sheet_match.end() :]

    @staticmethod
    def with_date_num_fmt(xf_xml: str, num_fmt_id: int) -> str:
        start_end = xf_xml.find(">")
        start_tag = xf_xml[: start_end + 1]
        tail = xf_xml[start_end + 1 :]
        start_tag = ResultDateValueChangeApp.set_xml_attr(start_tag, "numFmtId", str(num_fmt_id))
        start_tag = ResultDateValueChangeApp.set_xml_attr(start_tag, "applyNumberFormat", "1")
        return start_tag + tail

    @staticmethod
    def set_xml_attr(start_tag: str, name: str, value: str) -> str:
        attr_pattern = rf"\s+{re.escape(name)}=(['\"]).*?\1"
        replacement = f' {name}="{value}"'
        if re.search(attr_pattern, start_tag, re.DOTALL):
            return re.sub(attr_pattern, replacement, start_tag, count=1, flags=re.DOTALL)
        insert_at = start_tag.rfind("/>") if start_tag.rstrip().endswith("/>") else start_tag.rfind(">")
        if insert_at < 0:
            raise ValueError("XML 시작 태그 형식이 올바르지 않습니다.")
        return start_tag[:insert_at] + replacement + start_tag[insert_at:]

    def patch_b7_value(self, sheet_xml: bytes, serial_text: str, style_id: int | None) -> bytes:
        encoding = self.detect_xml_encoding(sheet_xml)
        xml_text = sheet_xml.decode(encoding)
        cell_pattern = re.compile(
            r"(<c\b(?=[^>]*\br=(['\"])B7\2)[^>]*)(?:>(.*?)</c>|/>)",
            re.DOTALL,
        )

        def replace_cell(match: re.Match) -> str:
            start_tag = re.sub(r"\s+t=(['\"]).*?\1", "", match.group(1), flags=re.DOTALL)
            if style_id is not None:
                if re.search(r"\s+s=(['\"])\d+\1", start_tag):
                    start_tag = re.sub(r"\s+s=(['\"])\d+\1", f' s="{style_id}"', start_tag, count=1)
                else:
                    start_tag = f'{start_tag} s="{style_id}"'
            body = match.group(3)
            if body is None:
                return f"{start_tag}><v>{serial_text}</v></c>"

            body = re.sub(r"<is\b.*?</is>", "", body, flags=re.DOTALL)
            if re.search(r"<v>.*?</v>", body, flags=re.DOTALL):
                body = re.sub(r"<v>.*?</v>", f"<v>{serial_text}</v>", body, count=1, flags=re.DOTALL)
            else:
                body = f"<v>{serial_text}</v>{body}"
            return f"{start_tag}>{body}</c>"

        patched_text, replace_count = cell_pattern.subn(replace_cell, xml_text, count=1)
        if replace_count:
            return patched_text.encode(encoding)

        return self.patch_b7_value_with_xml_parser(sheet_xml, serial_text, style_id)

    @staticmethod
    def detect_xml_encoding(xml_bytes: bytes) -> str:
        match = re.match(br"\s*<\?xml[^>]*encoding=[\"']([^\"']+)[\"']", xml_bytes)
        if match:
            return match.group(1).decode("ascii")
        return "utf-8"

    @staticmethod
    def patch_b7_value_with_xml_parser(sheet_xml: bytes, serial_text: str, style_id: int | None) -> bytes:
        ET.register_namespace("", SPREADSHEET_NS)
        root = ET.fromstring(sheet_xml)
        sheet_data = root.find(f"{{{SPREADSHEET_NS}}}sheetData")
        if sheet_data is None:
            raise ValueError("summary 시트의 sheetData를 찾지 못했습니다.")

        row = None
        for candidate in sheet_data.findall(f"{{{SPREADSHEET_NS}}}row"):
            if candidate.attrib.get("r") == "7":
                row = candidate
                break
        if row is None:
            row = ET.SubElement(sheet_data, f"{{{SPREADSHEET_NS}}}row", {"r": "7"})

        cell = None
        for candidate in row.findall(f"{{{SPREADSHEET_NS}}}c"):
            if candidate.attrib.get("r") == "B7":
                cell = candidate
                break
        if cell is None:
            cell = ET.SubElement(row, f"{{{SPREADSHEET_NS}}}c", {"r": "B7"})

        cell.attrib.pop("t", None)
        if style_id is not None:
            cell.set("s", str(style_id))
        for child in list(cell):
            if child.tag in {f"{{{SPREADSHEET_NS}}}v", f"{{{SPREADSHEET_NS}}}is"}:
                cell.remove(child)
        value_node = ET.Element(f"{{{SPREADSHEET_NS}}}v")
        value_node.text = serial_text
        cell.insert(0, value_node)
        return ET.tostring(root, encoding="utf-8", xml_declaration=sheet_xml.lstrip().startswith(b"<?xml"))

    @staticmethod
    def datetime_to_excel_serial(dt_value: datetime) -> float:
        delta = dt_value - EXCEL_SERIAL_EPOCH
        return delta.days + (delta.seconds + delta.microseconds / 1_000_000) / 86400

    @staticmethod
    def format_excel_serial(serial_value: float) -> str:
        return f"{serial_value:.10f}".rstrip("0").rstrip(".")

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
