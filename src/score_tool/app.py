from __future__ import annotations

import os
import sys
import traceback
from pathlib import Path
from tkinter import BooleanVar, StringVar, Tk, filedialog, messagebox, simpledialog
from tkinter import ttk

try:
    from .excel_processor import build_configs, infer_score_name, merge_sources, preview_sources
except ImportError:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    from score_tool.excel_processor import build_configs, infer_score_name, merge_sources, preview_sources

APP_TITLE = "成绩汇总工具"
EXCEL_FILETYPES = [("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")]


class ScoreToolApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1080x720")
        self.root.minsize(940, 620)

        self.files: list[Path] = []
        self.score_names: dict[str, str] = {}
        self.output_path = StringVar(value=str(Path.cwd() / "成绩汇总.xlsx"))
        self.include_log_sheet = BooleanVar(value=True)
        self.status_text = StringVar(value="请选择 Excel 文件。")

        self._build_layout()

    def _build_layout(self) -> None:
        root = self.root
        root.columnconfigure(0, weight=1)
        root.rowconfigure(1, weight=1)

        top = ttk.Frame(root, padding=12)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(0, weight=1)

        title = ttk.Label(top, text="Excel 成绩汇总工具", font=("Microsoft YaHei UI", 16, "bold"))
        title.grid(row=0, column=0, sticky="w")
        subtitle = ttk.Label(top, text="汇总成绩、重修置顶，并自动标记满分、低分、异常低分和疑似前后同学成绩对调。")
        subtitle.grid(row=1, column=0, sticky="w", pady=(4, 0))

        main = ttk.Frame(root, padding=(12, 0, 12, 8))
        main.grid(row=1, column=0, sticky="nsew")
        main.columnconfigure(0, weight=3)
        main.columnconfigure(1, weight=2)
        main.rowconfigure(0, weight=1)

        left = ttk.LabelFrame(main, text="源文件", padding=10)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(0, weight=1)

        columns = ("path", "score_name")
        self.file_tree = ttk.Treeview(left, columns=columns, show="headings", selectmode="extended")
        self.file_tree.heading("path", text="Excel 文件")
        self.file_tree.heading("score_name", text="成绩列名")
        self.file_tree.column("path", width=460, anchor="w")
        self.file_tree.column("score_name", width=160, anchor="w")
        self.file_tree.grid(row=0, column=0, sticky="nsew")
        self.file_tree.bind("<Double-1>", self._edit_score_name)

        file_scroll = ttk.Scrollbar(left, orient="vertical", command=self.file_tree.yview)
        file_scroll.grid(row=0, column=1, sticky="ns")
        self.file_tree.configure(yscrollcommand=file_scroll.set)

        file_buttons = ttk.Frame(left)
        file_buttons.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Button(file_buttons, text="添加 Excel", command=self.add_files).pack(side="left")
        ttk.Button(file_buttons, text="移除选中", command=self.remove_selected).pack(side="left", padx=6)
        ttk.Button(file_buttons, text="清空", command=self.clear_files).pack(side="left")
        ttk.Button(file_buttons, text="修改成绩列名", command=self.edit_selected_score_name).pack(side="left", padx=6)

        right = ttk.Frame(main)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)

        output_box = ttk.LabelFrame(right, text="输出设置", padding=10)
        output_box.grid(row=0, column=0, sticky="ew")
        output_box.columnconfigure(0, weight=1)
        ttk.Entry(output_box, textvariable=self.output_path).grid(row=0, column=0, sticky="ew")
        ttk.Button(output_box, text="选择保存位置", command=self.choose_output).grid(row=0, column=1, padx=(6, 0))
        ttk.Checkbutton(output_box, text="导出处理日志工作表", variable=self.include_log_sheet).grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))
        review_note = ttk.Label(
            output_box,
            text="核查标记会写入主表最后一列；问题成绩格会高亮，详细原因会写入“核查明细”工作表。",
            foreground="#6B4E00",
            wraplength=420,
        )
        review_note.grid(row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))

        log_box = ttk.LabelFrame(right, text="预检与日志", padding=10)
        log_box.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
        log_box.columnconfigure(0, weight=1)
        log_box.rowconfigure(0, weight=1)
        self.log_text = ttk.Treeview(log_box, columns=("type", "message"), show="headings")
        self.log_text.heading("type", text="类型")
        self.log_text.heading("message", text="内容")
        self.log_text.column("type", width=70, anchor="center")
        self.log_text.column("message", width=360, anchor="w")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        log_scroll = ttk.Scrollbar(log_box, orient="vertical", command=self.log_text.yview)
        log_scroll.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.tag_configure("重点核查", background="#F4CCCC")
        self.log_text.tag_configure("核查", background="#FFF2CC")
        self.log_text.tag_configure("警告", background="#FCE5CD")
        self.log_text.tag_configure("错误", background="#F4CCCC")

        bottom = ttk.Frame(root, padding=(12, 0, 12, 12))
        bottom.grid(row=2, column=0, sticky="ew")
        bottom.columnconfigure(0, weight=1)
        ttk.Label(bottom, textvariable=self.status_text).grid(row=0, column=0, sticky="w")
        ttk.Button(bottom, text="预检", command=self.precheck).grid(row=0, column=1, padx=6)
        ttk.Button(bottom, text="生成汇总表", command=self.generate).grid(row=0, column=2, padx=6)
        ttk.Button(bottom, text="打开输出目录", command=self.open_output_dir).grid(row=0, column=3)

    def add_files(self) -> None:
        paths = filedialog.askopenfilenames(title="选择 Excel 文件", filetypes=EXCEL_FILETYPES)
        if not paths:
            return
        for raw_path in paths:
            path = Path(raw_path)
            if path.suffix.lower() not in {".xlsx", ".xlsm"}:
                self._add_log("警告", f"暂不支持 {path.name}，请另存为 .xlsx 后再导入。")
                continue
            if path not in self.files:
                self.files.append(path)
                self.score_names[str(path)] = infer_score_name(path)
        if self.files and self.output_path.get() == str(Path.cwd() / "成绩汇总.xlsx"):
            self.output_path.set(str(self.files[0].parent / "成绩汇总.xlsx"))
        self._refresh_file_tree()
        self.status_text.set(f"已选择 {len(self.files)} 个 Excel 文件。")

    def remove_selected(self) -> None:
        selected = self.file_tree.selection()
        if not selected:
            return
        selected_paths = {self.file_tree.item(item, "values")[0] for item in selected}
        self.files = [path for path in self.files if str(path) not in selected_paths]
        for path in selected_paths:
            self.score_names.pop(path, None)
        self._refresh_file_tree()
        self.status_text.set(f"已选择 {len(self.files)} 个 Excel 文件。")

    def clear_files(self) -> None:
        self.files.clear()
        self.score_names.clear()
        self._refresh_file_tree()
        self._clear_logs()
        self.status_text.set("已清空文件列表。")

    def choose_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="保存汇总表",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialfile="成绩汇总.xlsx",
        )
        if path:
            self.output_path.set(path)

    def edit_selected_score_name(self) -> None:
        selected = self.file_tree.selection()
        if not selected:
            messagebox.showinfo(APP_TITLE, "请先选中一个文件。")
            return
        self._prompt_edit_score_name(selected[0])

    def _edit_score_name(self, _event: object) -> None:
        selected = self.file_tree.selection()
        if selected:
            self._prompt_edit_score_name(selected[0])

    def _prompt_edit_score_name(self, item_id: str) -> None:
        values = self.file_tree.item(item_id, "values")
        if not values:
            return
        path_text, old_name = values[0], values[1]
        new_name = simpledialog.askstring(APP_TITLE, "请输入这个文件对应的成绩列名：", initialvalue=old_name)
        if new_name is None:
            return
        new_name = new_name.strip()
        if not new_name:
            messagebox.showwarning(APP_TITLE, "成绩列名不能为空。")
            return
        self.score_names[path_text] = new_name
        self._refresh_file_tree()

    def precheck(self) -> None:
        if not self._ensure_ready(check_output=False):
            return
        self._clear_logs()
        try:
            previews = preview_sources(self._configs())
            for preview in previews:
                self.score_names[str(preview.path)] = preview.score_name
                self._add_log("文件", f"{preview.path.name} | 工作表：{preview.sheet_name} | 表头行：{preview.header_row} | 人数：{preview.student_count} | 成绩列：{preview.score_name}")
                for warning in preview.warnings:
                    self._add_log("警告", warning)
            self._refresh_file_tree()
            self.status_text.set("预检完成。确认无误后可生成汇总表。")
        except Exception as exc:
            self._show_error("预检失败", exc)

    def generate(self) -> None:
        if not self._ensure_ready(check_output=True):
            return
        self._clear_logs()
        try:
            result = merge_sources(
                self._configs(),
                Path(self.output_path.get()),
                include_log_sheet=self.include_log_sheet.get(),
            )
            for preview in result.source_previews:
                self._add_log("文件", f"{preview.path.name} → {preview.score_name}，{preview.student_count} 人")
            for warning in result.warnings:
                self._add_log("警告", warning)
            if result.review_count:
                self._add_log("核查", f"已生成 {result.review_count} 条核查标记，其中重点核查 {result.focus_review_count} 条。")
                if result.focus_review_count:
                    self._add_log("重点核查", "存在疑似异常低分或前后同学成绩对调，请优先查看“核查明细”工作表。")
            else:
                self._add_log("核查", "未发现需要标记的满分、低分或异常低分。")
            self.status_text.set(f"已生成：{result.output_path}，共 {result.row_count} 人，核查 {result.review_count} 条。")
            messagebox.showinfo(
                APP_TITLE,
                f"汇总完成。\n\n输出文件：{result.output_path}\n核查标记：{result.review_count} 条\n重点核查：{result.focus_review_count} 条",
            )
        except Exception as exc:
            self._show_error("生成失败", exc)

    def open_output_dir(self) -> None:
        path = Path(self.output_path.get()).expanduser()
        folder = path.parent if path.suffix else path
        if not folder.exists():
            folder = Path.cwd()
        os.startfile(folder)

    def _configs(self):
        return build_configs(self.files, self.score_names)

    def _ensure_ready(self, check_output: bool) -> bool:
        if not self.files:
            messagebox.showwarning(APP_TITLE, "请先添加 Excel 文件。")
            return False
        if check_output and not self.output_path.get().strip():
            messagebox.showwarning(APP_TITLE, "请选择输出文件路径。")
            return False
        return True

    def _refresh_file_tree(self) -> None:
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        for path in self.files:
            self.file_tree.insert("", "end", values=(str(path), self.score_names.get(str(path), infer_score_name(path))))

    def _clear_logs(self) -> None:
        for item in self.log_text.get_children():
            self.log_text.delete(item)

    def _add_log(self, item_type: str, message: str) -> None:
        tags = (item_type,) if item_type in {"重点核查", "核查", "警告", "错误"} else ()
        self.log_text.insert("", "end", values=(item_type, message), tags=tags)

    def _show_error(self, title: str, exc: Exception) -> None:
        self._add_log("错误", str(exc))
        self._add_log("详情", traceback.format_exc(limit=3))
        messagebox.showerror(APP_TITLE, f"{title}：\n{exc}")
        self.status_text.set(title)


def main() -> None:
    root = Tk()
    try:
        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    ScoreToolApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()