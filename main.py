import csv
import os
import re
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


STUDENT_COLUMNS = [
    "EventName",
    "Location",
    "Description",
    "Weekday",
    "StartTime",
    "EndTime",
    "IsCourse",
]

PROFESSOR_COLUMNS = [
    "ProfessorName",
    "Email",
    "EventName",
    "Location",
    "Description",
    "Weekday",
    "StartTime",
    "EndTime",
]

DATETIME_FORMAT = "%Y-%m-%d %H:%M"
DATE_FORMAT = "%Y-%m-%d"
DEFAULT_STEP_MINUTES = 30
CSV_FILE_TYPES = [("CSV 文件", "*.csv"), ("所有文件", "*.*")]


def ensure_csv_path(path_text: str, default_name: str) -> str:
    """
    Normalize the path. If a directory is provided, append the default csv name.
    Create parent directories when needed.
    """
    path_text = path_text.strip()
    if not path_text:
        raise ValueError("请选择 CSV 输出路径")

    if os.path.isdir(path_text):
        path_text = os.path.join(path_text, default_name)

    parent = os.path.dirname(path_text)
    if parent and not os.path.exists(parent):
        os.makedirs(parent, exist_ok=True)
    return path_text


def validate_weekday(value: str) -> int:
    weekday = int(value)
    if weekday < 1 or weekday > 7:
        raise ValueError("Weekday 必须在 1-7 之间")
    return weekday


def parse_time(value: str) -> datetime:
    try:
        return datetime.strptime(value.strip(), DATETIME_FORMAT)
    except ValueError as exc:
        raise ValueError(f"时间格式应为 {DATETIME_FORMAT}") from exc


def parse_date(value: str) -> datetime:
    try:
        return datetime.strptime(value.strip(), DATE_FORMAT)
    except ValueError as exc:
        raise ValueError(f"日期格式应为 {DATE_FORMAT}") from exc


def validate_email(value: str) -> str:
    if not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", value):
        raise ValueError("请输入合法的邮箱地址")
    return value


class CSVApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("CSV 生成工具")
        self.geometry("1200x720")

        self.student_data = []
        self.professor_data = []
        self.selected_student_index: int | None = None
        self.selected_professor_index: int | None = None

        self._build_ui()

    def _build_ui(self) -> None:
        self._build_menu()
        notebook = ttk.Notebook(self)
        student_frame = ttk.Frame(notebook, padding=10)
        professor_frame = ttk.Frame(notebook, padding=10)
        notebook.add(student_frame, text="学生日程")
        notebook.add(professor_frame, text="教师办公时间")
        notebook.pack(fill="both", expand=True)

        self._build_student_tab(student_frame)
        self._build_professor_tab(professor_frame)

    def _build_menu(self) -> None:
        menubar = tk.Menu(self)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="打开学生 CSV", command=self.import_student_csv)
        file_menu.add_command(label="导出学生 CSV", command=self.export_student_csv)
        file_menu.add_separator()
        file_menu.add_command(label="打开教师 CSV", command=self.import_prof_csv)
        file_menu.add_command(label="导出教师 CSV", command=self.export_prof_csv)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.destroy)
        menubar.add_cascade(label="文件", menu=file_menu)

        self.config(menu=menubar)

    def _build_student_tab(self, frame: ttk.Frame) -> None:
        form = ttk.Frame(frame)
        form.pack(side="left", fill="y", padx=(0, 15))

        self.student_vars = {
            "EventName": tk.StringVar(),
            "Location": tk.StringVar(),
            "Description": tk.StringVar(),
            "Weekday": tk.StringVar(),
            "StartTime": tk.StringVar(),
            "EndTime": tk.StringVar(),
            "IsCourse": tk.IntVar(value=1),
        }

        self._add_labeled_entry(form, "EventName", "课程/事件名称")
        self._add_labeled_entry(form, "Location", "地点")
        self._add_labeled_entry(form, "Description", "描述（可选）")
        self._add_labeled_entry(form, "Weekday", "星期（1-7）")
        self._add_labeled_entry(form, "StartTime", "开始时间 (YYYY-MM-DD HH:MM)")
        self._add_labeled_entry(form, "EndTime", "结束时间 (YYYY-MM-DD HH:MM)")
        self._add_time_picker_button(
            form, self.student_vars["StartTime"], self.student_vars["EndTime"]
        )

        is_course_frame = ttk.Frame(form)
        is_course_frame.pack(fill="x", pady=5)
        ttk.Label(is_course_frame, text="是否课程 (1=是/0=否)").pack(anchor="w")
        ttk.Checkbutton(
            is_course_frame,
            text="课程",
            variable=self.student_vars["IsCourse"],
            onvalue=1,
            offvalue=0,
        ).pack(anchor="w")

        button_row = ttk.Frame(form)
        button_row.pack(fill="x", pady=10)
        ttk.Button(button_row, text="保存/更新", command=self.save_student_entry).pack(
            side="left", padx=5
        )
        ttk.Button(button_row, text="删除选中", command=self.delete_student_entry).pack(
            side="left", padx=5
        )
        ttk.Button(button_row, text="清空表单", command=self.clear_student_form).pack(
            side="left", padx=5
        )

        right_panel = ttk.Frame(frame)
        right_panel.pack(side="left", fill="both", expand=True)

        self.student_tree = ttk.Treeview(
            right_panel,
            columns=STUDENT_COLUMNS,
            show="headings",
            selectmode="browse",
            height=18,
        )
        for col in STUDENT_COLUMNS:
            self.student_tree.heading(col, text=col)
            self.student_tree.column(col, width=140, anchor="w")
        self.student_tree.pack(fill="both", expand=True)
        self.student_tree.bind("<<TreeviewSelect>>", self.on_student_select)

        export_frame = ttk.Frame(frame)
        export_frame.pack(fill="x", pady=10)
        self.student_path_var = tk.StringVar(
            value=os.path.join(os.getcwd(), "student_schedule.csv")
        )
        self._build_path_selector(
            export_frame,
            "学生日程 CSV 输出路径",
            self.student_path_var,
            "student_schedule.csv",
        )
        ttk.Button(
            export_frame,
            text="打开 CSV",
            command=self.import_student_csv,
        ).pack(side="right", padx=5)
        ttk.Button(
            export_frame,
            text="导出 student_schedule.csv",
            command=self.export_student_csv,
        ).pack(side="right", padx=5)

    def _build_professor_tab(self, frame: ttk.Frame) -> None:
        form = ttk.Frame(frame)
        form.pack(side="left", fill="y", padx=(0, 15))

        self.prof_vars = {
            "ProfessorName": tk.StringVar(),
            "Email": tk.StringVar(),
            "EventName": tk.StringVar(value="Office Hour"),
            "Location": tk.StringVar(),
            "Description": tk.StringVar(),
            "Weekday": tk.StringVar(),
            "StartTime": tk.StringVar(),
            "EndTime": tk.StringVar(),
        }

        self._add_labeled_entry(form, "ProfessorName", "教师姓名")
        self._add_labeled_entry(form, "Email", "邮箱")
        self._add_labeled_entry(form, "EventName", "事件名称")
        self._add_labeled_entry(form, "Location", "地点")
        self._add_labeled_entry(form, "Description", "描述（可选）")
        self._add_labeled_entry(form, "Weekday", "星期（1-7）")
        self._add_labeled_entry(form, "StartTime", "开始时间 (YYYY-MM-DD HH:MM)")
        self._add_labeled_entry(form, "EndTime", "结束时间 (YYYY-MM-DD HH:MM)")
        self._add_time_picker_button(
            form, self.prof_vars["StartTime"], self.prof_vars["EndTime"]
        )

        button_row = ttk.Frame(form)
        button_row.pack(fill="x", pady=10)
        ttk.Button(button_row, text="保存/更新", command=self.save_prof_entry).pack(
            side="left", padx=5
        )
        ttk.Button(button_row, text="删除选中", command=self.delete_prof_entry).pack(
            side="left", padx=5
        )
        ttk.Button(button_row, text="清空表单", command=self.clear_prof_form).pack(
            side="left", padx=5
        )

        right_panel = ttk.Frame(frame)
        right_panel.pack(side="left", fill="both", expand=True)

        self.prof_tree = ttk.Treeview(
            right_panel,
            columns=PROFESSOR_COLUMNS,
            show="headings",
            selectmode="browse",
            height=18,
        )
        for col in PROFESSOR_COLUMNS:
            self.prof_tree.heading(col, text=col)
            self.prof_tree.column(col, width=140, anchor="w")
        self.prof_tree.pack(fill="both", expand=True)
        self.prof_tree.bind("<<TreeviewSelect>>", self.on_prof_select)

        export_frame = ttk.Frame(frame)
        export_frame.pack(fill="x", pady=10)
        self.prof_path_var = tk.StringVar(
            value=os.path.join(os.getcwd(), "professors.csv")
        )
        self._build_path_selector(
            export_frame,
            "教师办公时间 CSV 输出路径",
            self.prof_path_var,
            "professors.csv",
        )
        ttk.Button(
            export_frame,
            text="打开 CSV",
            command=self.import_prof_csv,
        ).pack(side="right", padx=5)
        ttk.Button(
            export_frame,
            text="导出 professors.csv",
            command=self.export_prof_csv,
        ).pack(side="right", padx=5)

    def _add_labeled_entry(
        self, parent: ttk.Frame, key: str, label_text: str
    ) -> ttk.Entry:
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=4)
        ttk.Label(row, text=label_text).pack(anchor="w")
        entry = ttk.Entry(
            row, textvariable=self.student_vars.get(key) or self.prof_vars.get(key)
        )
        entry.pack(fill="x")
        return entry

    def _add_time_picker_button(
        self, parent: ttk.Frame, start_var: tk.StringVar, end_var: tk.StringVar
    ) -> None:
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=4)

        def open_picker() -> None:
            base_date = self._detect_base_date(start_var.get(), end_var.get())
            start_clock = self._extract_clock(start_var.get())
            end_clock = self._extract_clock(end_var.get())

            def apply_range(date_str: str, start_time: str, end_time: str) -> None:
                start_var.set(f"{date_str} {start_time}")
                end_var.set(f"{date_str} {end_time}")

            TimeRangeDialog(
                self,
                base_date=base_date,
                start_clock=start_clock,
                end_clock=end_clock,
                on_apply=apply_range,
            )

        ttk.Button(row, text="拖拽选择时间段", command=open_picker).pack(anchor="w")

    @staticmethod
    def _detect_base_date(*candidates: str) -> str:
        for val in candidates:
            try:
                dt = parse_time(val)
                return dt.strftime(DATE_FORMAT)
            except Exception:
                try:
                    dt = parse_date(val)
                    return dt.strftime(DATE_FORMAT)
                except Exception:
                    continue
        return datetime.now().strftime(DATE_FORMAT)

    @staticmethod
    def _extract_clock(value: str) -> str | None:
        try:
            dt = parse_time(value)
            return dt.strftime("%H:%M")
        except Exception:
            return None

    def _build_path_selector(
        self,
        parent: ttk.Frame,
        label: str,
        path_var: tk.StringVar,
        default_name: str,
    ) -> None:
        ttk.Label(parent, text=label).pack(side="left", padx=(0, 5))
        entry = ttk.Entry(parent, textvariable=path_var, width=60)
        entry.pack(side="left", padx=5, fill="x", expand=True)

        def browse_dir() -> None:
            chosen = filedialog.askdirectory()
            if chosen:
                path_var.set(os.path.join(chosen, default_name))

        def browse_file() -> None:
            current = path_var.get().strip()
            initial_dir = os.path.dirname(current) if current else os.getcwd()
            initial_file = os.path.basename(current) if current else default_name
            chosen = filedialog.asksaveasfilename(
                title="选择保存位置",
                defaultextension=".csv",
                filetypes=CSV_FILE_TYPES,
                initialdir=initial_dir,
                initialfile=initial_file,
            )
            if chosen:
                path_var.set(chosen)

        ttk.Button(parent, text="选择目录", command=browse_dir).pack(side="left", padx=5)
        ttk.Button(parent, text="选择文件", command=browse_file).pack(side="left", padx=5)

    def save_student_entry(self) -> None:
        try:
            data = self._collect_student_input()
        except ValueError as exc:
            messagebox.showerror("输入有误", str(exc))
            return

        if self.selected_student_index is not None:
            self.student_data[self.selected_student_index] = data
            self.selected_student_index = None
        else:
            self.student_data.append(data)

        self.refresh_tree(self.student_tree, self.student_data, STUDENT_COLUMNS)
        self.clear_student_form()

    def delete_student_entry(self) -> None:
        if self.selected_student_index is None:
            messagebox.showinfo("提示", "请先在列表中选择要删除的日程")
            return
        self.student_data.pop(self.selected_student_index)
        self.selected_student_index = None
        self.refresh_tree(self.student_tree, self.student_data, STUDENT_COLUMNS)
        self.clear_student_form()

    def on_student_select(self, _event=None) -> None:
        selection = self.student_tree.selection()
        if not selection:
            return
        index = int(selection[0])
        self.selected_student_index = index
        record = self.student_data[index]
        for key, var in self.student_vars.items():
            var.set(record[key])

    def clear_student_form(self) -> None:
        for key, var in self.student_vars.items():
            if key == "IsCourse":
                var.set(1)
            else:
                var.set("")
        self.selected_student_index = None
        self.student_tree.selection_remove(self.student_tree.selection())

    def export_student_csv(self) -> None:
        save_path = self._ask_save_path(
            self.student_path_var, "student_schedule.csv", "保存学生日程 CSV"
        )
        if not save_path:
            return
        try:
            path = ensure_csv_path(save_path, "student_schedule.csv")
            self._write_csv(path, STUDENT_COLUMNS, self.student_data)
            self.student_path_var.set(path)
            messagebox.showinfo("已导出", f"学生日程已保存到:\n{path}")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("导出失败", str(exc))

    def import_student_csv(self) -> None:
        path = filedialog.askopenfilename(
            title="打开学生日程 CSV", filetypes=CSV_FILE_TYPES
        )
        if not path:
            return
        try:
            data = self._load_csv_data(path, STUDENT_COLUMNS)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("打开失败", str(exc))
            return

        self.student_data = data
        self.refresh_tree(self.student_tree, self.student_data, STUDENT_COLUMNS)
        self.clear_student_form()
        self.student_path_var.set(path)
        messagebox.showinfo("已加载", f"已从下列文件载入 {len(data)} 条学生日程:\n{path}")

    def _collect_student_input(self) -> dict:
        values = {
            k: v.get().strip() if k != "IsCourse" else v.get()
            for k, v in self.student_vars.items()
        }
        required = ["EventName", "Location", "Weekday", "StartTime", "EndTime"]
        for field in required:
            if not str(values[field]).strip():
                raise ValueError(f"{field} 为必填项")

        weekday = validate_weekday(values["Weekday"])
        start_dt = parse_time(values["StartTime"])
        end_dt = parse_time(values["EndTime"])
        if start_dt.date() != end_dt.date():
            raise ValueError("开始时间和结束时间必须在同一天")
        if end_dt <= start_dt:
            raise ValueError("结束时间必须大于开始时间")

        is_course_val = int(values["IsCourse"])
        if is_course_val not in (0, 1):
            raise ValueError("IsCourse 只能为 0 或 1")

        return {
            "EventName": values["EventName"],
            "Location": values["Location"],
            "Description": values["Description"],
            "Weekday": str(weekday),
            "StartTime": start_dt.strftime(DATETIME_FORMAT),
            "EndTime": end_dt.strftime(DATETIME_FORMAT),
            "IsCourse": str(is_course_val),
        }

    def save_prof_entry(self) -> None:
        try:
            data = self._collect_prof_input()
        except ValueError as exc:
            messagebox.showerror("输入有误", str(exc))
            return

        if self.selected_professor_index is not None:
            self.professor_data[self.selected_professor_index] = data
            self.selected_professor_index = None
        else:
            self.professor_data.append(data)

        self.refresh_tree(self.prof_tree, self.professor_data, PROFESSOR_COLUMNS)
        self.clear_prof_form()

    def delete_prof_entry(self) -> None:
        if self.selected_professor_index is None:
            messagebox.showinfo("提示", "请先在列表中选择要删除的教师记录")
            return
        self.professor_data.pop(self.selected_professor_index)
        self.selected_professor_index = None
        self.refresh_tree(self.prof_tree, self.professor_data, PROFESSOR_COLUMNS)
        self.clear_prof_form()

    def on_prof_select(self, _event=None) -> None:
        selection = self.prof_tree.selection()
        if not selection:
            return
        index = int(selection[0])
        self.selected_professor_index = index
        record = self.professor_data[index]
        for key, var in self.prof_vars.items():
            var.set(record[key])

    def clear_prof_form(self) -> None:
        for _, var in self.prof_vars.items():
            var.set("")
        self.prof_vars["EventName"].set("Office Hour")
        self.selected_professor_index = None
        self.prof_tree.selection_remove(self.prof_tree.selection())

    def export_prof_csv(self) -> None:
        save_path = self._ask_save_path(
            self.prof_path_var, "professors.csv", "保存教师办公时间 CSV"
        )
        if not save_path:
            return
        try:
            path = ensure_csv_path(save_path, "professors.csv")
            self._write_csv(path, PROFESSOR_COLUMNS, self.professor_data)
            self.prof_path_var.set(path)
            messagebox.showinfo("已导出", f"教师办公时间已保存到:\n{path}")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("导出失败", str(exc))

    def import_prof_csv(self) -> None:
        path = filedialog.askopenfilename(
            title="打开教师办公时间 CSV", filetypes=CSV_FILE_TYPES
        )
        if not path:
            return
        try:
            data = self._load_csv_data(path, PROFESSOR_COLUMNS)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("打开失败", str(exc))
            return

        self.professor_data = data
        self.refresh_tree(self.prof_tree, self.professor_data, PROFESSOR_COLUMNS)
        self.clear_prof_form()
        self.prof_path_var.set(path)
        messagebox.showinfo("已加载", f"已从下列文件载入 {len(data)} 条教师记录:\n{path}")

    def _collect_prof_input(self) -> dict:
        values = {k: v.get().strip() for k, v in self.prof_vars.items()}
        required = ["ProfessorName", "Email", "EventName", "Location", "Weekday", "StartTime", "EndTime"]
        for field in required:
            if not values[field]:
                raise ValueError(f"{field} 为必填项")

        validate_email(values["Email"])
        weekday = validate_weekday(values["Weekday"])
        start_dt = parse_time(values["StartTime"])
        end_dt = parse_time(values["EndTime"])
        if start_dt.date() != end_dt.date():
            raise ValueError("开始时间和结束时间必须在同一天")
        if end_dt <= start_dt:
            raise ValueError("结束时间必须大于开始时间")

        return {
            "ProfessorName": values["ProfessorName"],
            "Email": values["Email"],
            "EventName": values["EventName"],
            "Location": values["Location"],
            "Description": values["Description"],
            "Weekday": str(weekday),
            "StartTime": start_dt.strftime(DATETIME_FORMAT),
            "EndTime": end_dt.strftime(DATETIME_FORMAT),
        }

    def _write_csv(self, path: str, columns: list[str], data: list[dict]) -> None:
        if not data:
            raise ValueError("列表为空，至少添加一条记录后再导出")
        with open(path, "w", newline="", encoding="utf-8") as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=columns)
            writer.writeheader()
            for row in data:
                writer.writerow(row)

    def _ask_save_path(
        self, path_var: tk.StringVar, default_name: str, title: str
    ) -> str | None:
        """
        Prompt for a save location, prefilled with the current or default file name.
        """
        current = path_var.get().strip() or os.path.join(os.getcwd(), default_name)
        initial_dir = os.path.dirname(current) or os.getcwd()
        initial_file = os.path.basename(current) or default_name
        chosen = filedialog.asksaveasfilename(
            title=title,
            defaultextension=".csv",
            filetypes=CSV_FILE_TYPES,
            initialdir=initial_dir,
            initialfile=initial_file,
        )
        if not chosen:
            return None
        path_var.set(chosen)
        return chosen

    def _load_csv_data(self, path: str, columns: list[str]) -> list[dict]:
        with open(path, newline="", encoding="utf-8") as csvfile:
            reader = csv.DictReader(csvfile)
            if not reader.fieldnames:
                raise ValueError("CSV 文件缺少表头")
            missing = [col for col in columns if col not in reader.fieldnames]
            if missing:
                raise ValueError(f"缺少必要列: {', '.join(missing)}")
            rows: list[dict] = []
            for row in reader:
                rows.append({col: str(row.get(col, "") or "").strip() for col in columns})
            if not rows:
                raise ValueError("CSV 文件为空")
            return rows

    @staticmethod
    def refresh_tree(tree: ttk.Treeview, data: list[dict], columns: list[str]) -> None:
        tree.delete(*tree.get_children())
        for idx, row in enumerate(data):
            tree.insert("", "end", iid=str(idx), values=[row[col] for col in columns])


class TimeRangeDialog(tk.Toplevel):
    def __init__(
        self,
        master: tk.Tk,
        base_date: str,
        start_clock: str | None,
        end_clock: str | None,
        on_apply,
        step_minutes: int = DEFAULT_STEP_MINUTES,
    ):
        super().__init__(master)
        self.title(f"选择时间段（{base_date}）")
        self.resizable(False, False)
        self.step_minutes = step_minutes
        self.on_apply = on_apply
        self.base_date = base_date
        self.date_var = tk.StringVar(value=base_date)

        self.geometry("+200+120")

        self.slot_height = 14
        self.times = self._build_times()
        self.canvas_height = len(self.times) * self.slot_height
        self.canvas_width = 520

        self.start_idx = 0
        self.end_idx = 1
        if start_clock in self.times:
            self.start_idx = self.times.index(start_clock)
        if end_clock in self.times:
            self.end_idx = max(self.start_idx + 1, self.times.index(end_clock))

        self._build_ui()
        self._draw_scale()
        self._draw_selection()

    def _build_times(self) -> list[str]:
        times = []
        minutes = 0
        while minutes < 24 * 60:
            times.append(f"{minutes // 60:02d}:{minutes % 60:02d}")
            minutes += self.step_minutes
        return times

    def _build_ui(self) -> None:
        container = ttk.Frame(self, padding=10)
        container.pack(fill="both", expand=True)

        date_row = ttk.Frame(container)
        date_row.pack(fill="x", pady=(0, 6))
        ttk.Label(date_row, text="日期 (YYYY-MM-DD)").pack(side="left")
        ttk.Entry(date_row, textvariable=self.date_var, width=12).pack(
            side="left", padx=4
        )
        ttk.Button(date_row, text="前一天", command=lambda: self._shift_date(-1)).pack(
            side="left", padx=2
        )
        ttk.Button(date_row, text="今天", command=lambda: self._set_date_string(datetime.now().strftime(DATE_FORMAT))).pack(
            side="left", padx=2
        )
        ttk.Button(date_row, text="后一天", command=lambda: self._shift_date(1)).pack(
            side="left", padx=2
        )
        ttk.Button(date_row, text="应用日期", command=self._apply_date_entry).pack(
            side="left", padx=6
        )

        self.canvas = tk.Canvas(
            container,
            width=self.canvas_width,
            height=self.canvas_height,
            bg="#f7f7f7",
            highlightthickness=1,
            highlightbackground="#ccc",
        )
        self.canvas.pack(fill="both", expand=True)

        info_frame = ttk.Frame(container)
        info_frame.pack(fill="x", pady=8)
        self.range_label = ttk.Label(info_frame, text="")
        self.range_label.pack(side="left")

        btn_frame = ttk.Frame(container)
        btn_frame.pack(fill="x")
        ttk.Button(btn_frame, text="确定", command=self._apply).pack(
            side="right", padx=5
        )
        ttk.Button(btn_frame, text="取消", command=self.destroy).pack(
            side="right", padx=5
        )

        self.canvas.bind("<Button-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)

    def _draw_scale(self) -> None:
        self.canvas.delete("scale")
        for idx, t in enumerate(self.times):
            y = idx * self.slot_height
            if t.endswith("00"):
                self.canvas.create_line(
                    60, y, self.canvas_width - 10, y, fill="#ccc", tags="scale"
                )
                self.canvas.create_text(40, y, text=t, anchor="e", tags="scale")

    def _draw_selection(self) -> None:
        self.canvas.delete("selection")
        start_y = self.start_idx * self.slot_height
        end_y = self.end_idx * self.slot_height
        self.canvas.create_rectangle(
            60,
            start_y,
            self.canvas_width - 10,
            end_y,
            fill="#cde9ff",
            outline="#4a90e2",
            width=2,
            tags="selection",
        )
        start_time = self.times[self.start_idx]
        end_time = self.times[self.end_idx] if self.end_idx < len(self.times) else "24:00"
        self.range_label.config(
            text=f"{self.base_date} {start_time} - {self.base_date} {end_time}"
        )

    def _on_press(self, event) -> None:
        idx = self._y_to_idx(event.y)
        if idx >= len(self.times) - 1:
            idx = len(self.times) - 2
        self.start_idx = max(0, idx)
        self.end_idx = self.start_idx + 1
        self._draw_selection()

    def _on_drag(self, event) -> None:
        idx = self._y_to_idx(event.y)
        idx = min(idx, len(self.times) - 1)
        if idx <= self.start_idx:
            self.start_idx = max(0, min(idx, len(self.times) - 2))
            self.end_idx = self.start_idx + 1
        else:
            self.end_idx = min(len(self.times) - 1, idx)
            if self.end_idx <= self.start_idx:
                self.end_idx = min(len(self.times) - 1, self.start_idx + 1)
        self._draw_selection()

    def _on_release(self, _event) -> None:
        self._draw_selection()

    def _apply(self) -> None:
        start_time = self.times[self.start_idx]
        end_time = self.times[self.end_idx] if self.end_idx < len(self.times) else "24:00"
        self.on_apply(self.base_date, start_time, end_time)
        self.destroy()

    def _apply_date_entry(self) -> None:
        self._set_date_string(self.date_var.get())

    def _set_date_string(self, date_str: str) -> None:
        try:
            dt = parse_date(date_str)
            self.base_date = dt.strftime(DATE_FORMAT)
            self.date_var.set(self.base_date)
            self._draw_selection()
        except ValueError as exc:
            messagebox.showerror("日期格式错误", str(exc))

    def _shift_date(self, delta_days: int) -> None:
        try:
            dt = parse_date(self.date_var.get())
        except Exception:
            dt = datetime.now()
        dt = dt + timedelta(days=delta_days)
        self._set_date_string(dt.strftime(DATE_FORMAT))

    def _y_to_idx(self, y: int) -> int:
        idx = int(y // self.slot_height)
        return max(0, min(len(self.times) - 1, idx))


if __name__ == "__main__":
    app = CSVApp()
    app.mainloop()
