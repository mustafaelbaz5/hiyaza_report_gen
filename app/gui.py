from __future__ import annotations

import os
import re
import sys
import subprocess
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .excel_analyzer import analyze_excel
from .models import AnalysisResult, ReportInput
from .report_generator import generate_report
from .text_utils import safe_filename


APP_TITLE = "مولد تقارير تسليم الحيازات الزراعية"
NAVY = "#173F67"
TEAL = "#168B8B"
BG = "#F5F7FA"
CARD = "#FFFFFF"
TEXT = "#1F2933"
MUTED = "#64717C"
BORDER = "#D9E2E8"

ARABIC_MONTHS = [
    "يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
    "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر",
]


def today_arabic() -> str:
    value = datetime.now()
    return f"{value.day:02d} {ARABIC_MONTHS[value.month - 1]} {value.year}"


def resource_path(relative: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base / relative


class ReportGeneratorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1040x760")
        self.minsize(920, 680)
        self.configure(bg=BG)
        self.option_add("*Font", ("Segoe UI", 10))
        self.analysis: AnalysisResult | None = None
        self._configure_style()
        self._build_ui()

    def _configure_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TFrame", background=BG)
        style.configure("Card.TFrame", background=CARD, relief="solid", borderwidth=1)
        style.configure("Title.TLabel", background=BG, foreground=NAVY, font=("Segoe UI", 20, "bold"))
        style.configure("Subtitle.TLabel", background=BG, foreground=MUTED, font=("Segoe UI", 10))
        style.configure("Section.TLabel", background=CARD, foreground=NAVY, font=("Segoe UI", 12, "bold"))
        style.configure("Field.TLabel", background=CARD, foreground=TEXT, font=("Segoe UI", 10, "bold"))
        style.configure("Hint.TLabel", background=CARD, foreground=MUTED, font=("Segoe UI", 9))
        style.configure("Primary.TButton", background=NAVY, foreground="white", padding=(16, 9), font=("Segoe UI", 10, "bold"))
        style.map("Primary.TButton", background=[("active", "#215583")])
        style.configure("Accent.TButton", background=TEAL, foreground="white", padding=(18, 10), font=("Segoe UI", 10, "bold"))
        style.map("Accent.TButton", background=[("active", "#1B9D9D")])
        style.configure("Secondary.TButton", background="#E8EEF3", foreground=NAVY, padding=(12, 8))
        style.configure("Treeview", rowheight=28, background="white", fieldbackground="white", foreground=TEXT)
        style.configure("Treeview.Heading", background=NAVY, foreground="white", font=("Segoe UI", 10, "bold"))

    def _build_ui(self):
        container = ttk.Frame(self, padding=(28, 22))
        container.pack(fill="both", expand=True)

        ttk.Label(container, text=APP_TITLE, style="Title.TLabel", anchor="e").pack(fill="x")
        ttk.Label(
            container,
            text="اختر ملف Excel، راجع البيانات المحسوبة، ثم أنشئ تقرير Word من القالب المعتمد",
            style="Subtitle.TLabel",
            anchor="e",
        ).pack(fill="x", pady=(2, 16))

        body = ttk.Panedwindow(container, orient="horizontal")
        body.pack(fill="both", expand=True)
        left = ttk.Frame(body, style="Card.TFrame", padding=18)
        right = ttk.Frame(body, style="Card.TFrame", padding=18)
        body.add(left, weight=5)
        body.add(right, weight=4)
        self._build_inputs(left)
        self._build_preview(right)

        footer = ttk.Frame(container)
        footer.pack(fill="x", pady=(14, 0))
        self.status_var = tk.StringVar(value="جاهز لاختيار ملف Excel")
        ttk.Label(footer, textvariable=self.status_var, style="Subtitle.TLabel", anchor="w").pack(side="left", fill="x", expand=True)
        ttk.Button(footer, text="إنشاء تقرير Word", style="Accent.TButton", command=self.generate).pack(side="right")

    def _entry_row(self, parent, row, label, variable, browse=None, readonly=False):
        ttk.Label(parent, text=label, style="Field.TLabel", anchor="e").grid(row=row, column=2, sticky="e", padx=(10, 0), pady=6)
        entry = ttk.Entry(parent, textvariable=variable, justify="right", state="readonly" if readonly else "normal")
        entry.grid(row=row, column=0 if browse else 1, columnspan=2 if browse else 1, sticky="ew", pady=6)
        if browse:
            ttk.Button(parent, text="استعراض", style="Secondary.TButton", command=browse).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=6)
            entry.grid(row=row, column=1, sticky="ew", pady=6)
        return entry

    def _build_inputs(self, parent):
        parent.columnconfigure(1, weight=1)
        ttk.Label(parent, text="المدخلات", style="Section.TLabel", anchor="e").grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))

        self.excel_var = tk.StringVar()
        self.template_var = tk.StringVar(value=str(resource_path("assets/DST_Agricultural_Association_Report_Template.docx")))
        self.code_var = tk.StringVar()
        self.date_var = tk.StringVar(value=today_arabic())
        self.association_var = tk.StringVar()
        self.type_var = tk.StringVar(value="ائتمان زراعي")
        self.village_var = tk.StringVar()

        self._entry_row(parent, 1, "ملف Excel", self.excel_var, self.choose_excel)
        self._entry_row(parent, 2, "قالب Word", self.template_var, self.choose_template)
        ttk.Button(parent, text="تحليل ملف Excel", style="Primary.TButton", command=self.analyze).grid(row=3, column=0, columnspan=3, sticky="ew", pady=(8, 14))

        ttk.Separator(parent).grid(row=4, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        ttk.Label(parent, text="بيانات التقرير", style="Section.TLabel", anchor="e").grid(row=5, column=0, columnspan=3, sticky="ew")
        self._entry_row(parent, 6, "اسم الجمعية", self.association_var)
        ttk.Label(parent, text="نوع الجمعية", style="Field.TLabel", anchor="e").grid(row=7, column=2, sticky="e", padx=(10, 0), pady=6)
        ttk.Combobox(parent, textvariable=self.type_var, values=["ائتمان زراعي", "إصلاح زراعي"], state="readonly", justify="right").grid(row=7, column=1, sticky="ew", pady=6)
        self._code_entry_row(parent, 8)
        self._entry_row(parent, 9, "اسم البلد / الحزمة", self.village_var)
        self._entry_row(parent, 10, "تاريخ التسليم", self.date_var)

        ttk.Label(parent, text="ملاحظات خاصة بالتسليمة", style="Field.TLabel", anchor="e").grid(row=11, column=0, columnspan=3, sticky="e", pady=(8, 4))
        self.notes_text = tk.Text(parent, height=5, wrap="word", font=("Segoe UI", 10), relief="solid", borderwidth=1, padx=8, pady=8)
        self.notes_text.tag_configure("rtl", justify="right")
        self.notes_text.grid(row=12, column=0, columnspan=3, sticky="nsew")
        parent.rowconfigure(12, weight=1)
        ttk.Label(parent, text="يمكنك تعديل أي قيمة بعد التحليل وقبل إنشاء التقرير", style="Hint.TLabel", anchor="e").grid(row=13, column=0, columnspan=3, sticky="ew", pady=(5, 0))

    def _code_entry_row(self, parent, row):
        ttk.Label(parent, text="كود الجمعية", style="Field.TLabel", anchor="e").grid(
            row=row, column=2, sticky="e", padx=(10, 0), pady=6
        )
        entry = ttk.Entry(parent, textvariable=self.code_var, justify="right")
        entry.grid(row=row, column=1, sticky="ew", pady=6)
        ttk.Button(parent, text="لصق", style="Secondary.TButton", command=self.paste_association_code).grid(
            row=row, column=0, sticky="w", padx=(0, 8), pady=6
        )
        entry.bind("<Control-v>", self._paste_code_event)
        entry.bind("<Control-V>", self._paste_code_event)
        self.code_entry = entry

    def _paste_code_event(self, _event=None):
        self.paste_association_code()
        return "break"

    def paste_association_code(self):
        try:
            clipboard = self.clipboard_get().strip()
        except tk.TclError:
            messagebox.showinfo("لا يوجد نص", "لا يوجد نص منسوخ في الحافظة.", parent=self)
            return
        labeled = re.search(r"(?:كود|code)\s*(?:الجمعية|association)?\s*[:：-]\s*([^\r\n]+)", clipboard, re.IGNORECASE)
        value = labeled.group(1).strip() if labeled else clipboard
        self.code_var.set(value.splitlines()[0].strip())
        self.code_entry.focus_set()
        self.code_entry.selection_range(0, "end")

    def _build_preview(self, parent):
        ttk.Label(parent, text="معاينة التحليل", style="Section.TLabel", anchor="e").pack(fill="x", pady=(0, 10))
        self.tree = ttk.Treeview(parent, columns=("value", "metric"), show="headings", height=9)
        self.tree.heading("metric", text="المؤشر")
        self.tree.heading("value", text="القيمة")
        self.tree.column("metric", width=210, anchor="e")
        self.tree.column("value", width=120, anchor="center")
        self.tree.pack(fill="x")

        ttk.Label(parent, text="الحالات المكتشفة", style="Section.TLabel", anchor="e").pack(fill="x", pady=(18, 6))
        self.detected_notes = tk.Text(parent, height=8, wrap="word", state="disabled", font=("Segoe UI", 9), relief="solid", borderwidth=1, padx=8, pady=8)
        self.detected_notes.tag_configure("rtl", justify="right")
        self.detected_notes.pack(fill="both", expand=True)

        ttk.Button(parent, text="نسخ الحالات إلى الملاحظات", style="Secondary.TButton", command=self.copy_detected_notes).pack(fill="x", pady=(10, 0))

    def choose_excel(self):
        path = filedialog.askopenfilename(title="اختر ملف Excel", filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if path:
            self.excel_var.set(path)
            self.analysis = None
            self.status_var.set("تم اختيار ملف Excel. اضغط تحليل ملف Excel.")

    def choose_template(self):
        path = filedialog.askopenfilename(title="اختر قالب Word", filetypes=[("Word documents", "*.docx")])
        if path:
            self.template_var.set(path)

    def analyze(self):
        try:
            self.status_var.set("جارٍ تحليل ملف Excel...")
            self.update_idletasks()
            result = analyze_excel(self.excel_var.get())
            self.analysis = result
            self.association_var.set(result.association_name)
            self.type_var.set(result.association_type)
            self.village_var.set(result.village_name)
            self._show_analysis(result)
            self.status_var.set("اكتمل التحليل. راجع البيانات ثم أنشئ التقرير.")
        except Exception as exc:
            self.status_var.set("تعذر تحليل الملف")
            messagebox.showerror("خطأ في التحليل", str(exc), parent=self)

    def _show_analysis(self, result: AnalysisResult):
        for item in self.tree.get_children():
            self.tree.delete(item)
        rows = [
            ("إجمالي عدد القطع", result.total_parcels),
            ("عدد الأحواض", result.basin_count),
            ("القطع المسجلة", f"{result.registered_count}  ({result.registered_percentage:.1f}%)"),
            ("القطع غير المسجلة", f"{result.unregistered_count}  ({result.unregistered_percentage:.1f}%)"),
            ("منها غير محيزة", f"{result.non_held_count}  ({result.non_held_percentage:.1f}%)"),
        ]
        for metric, value in rows:
            self.tree.insert("", "end", values=(value, metric))
        notes = list(result.detected_notes)
        if result.warnings:
            notes.extend(["تنبيه: " + x for x in result.warnings])
        self.detected_notes.configure(state="normal")
        self.detected_notes.delete("1.0", "end")
        self.detected_notes.insert("1.0", "\n".join("• " + x for x in notes) if notes else "لا توجد حالات خاصة مكتشفة.", "rtl")
        self.detected_notes.configure(state="disabled")

    def copy_detected_notes(self):
        if not self.analysis:
            messagebox.showinfo("تنبيه", "حلل ملف Excel أولًا.", parent=self)
            return
        text = "\n".join(self.analysis.detected_notes)
        self.notes_text.delete("1.0", "end")
        self.notes_text.insert("1.0", text or "لا يوجد", "rtl")

    def generate(self):
        if not self.analysis:
            messagebox.showwarning("بيانات غير مكتملة", "يجب تحليل ملف Excel أولًا.", parent=self)
            return
        if not self.code_var.get().strip():
            messagebox.showwarning("كود الجمعية", "أدخل كود الجمعية قبل إنشاء التقرير.", parent=self)
            return
        if not self.association_var.get().strip():
            messagebox.showwarning("اسم الجمعية", "اسم الجمعية غير موجود.", parent=self)
            return
        template = self.template_var.get().strip()
        if not template:
            messagebox.showwarning("قالب Word", "اختر قالب Word.", parent=self)
            return

        suggested = "تقرير_تسليم_" + safe_filename(self.association_var.get()) + ".docx"
        destination = filedialog.asksaveasfilename(
            title="حفظ تقرير Word",
            defaultextension=".docx",
            initialfile=suggested,
            filetypes=[("Word document", "*.docx")],
        )
        if not destination:
            return
        try:
            report_data = ReportInput(
                analysis=self.analysis,
                association_name=self.association_var.get().strip(),
                association_type=self.type_var.get().strip(),
                association_code=self.code_var.get().strip(),
                village_name=self.village_var.get().strip() or self.analysis.village_name,
                delivery_date=self.date_var.get().strip(),
                special_notes=self.notes_text.get("1.0", "end").strip() or "لا يوجد",
            )
            output = generate_report(template, destination, report_data)
            self.status_var.set(f"تم إنشاء التقرير: {output.name}")
            if messagebox.askyesno("تم إنشاء التقرير", "تم إنشاء ملف Word بنجاح. هل تريد فتح مكان الحفظ؟", parent=self):
                self._open_folder(output.parent)
        except Exception as exc:
            self.status_var.set("تعذر إنشاء التقرير")
            messagebox.showerror("خطأ في إنشاء التقرير", str(exc), parent=self)

    @staticmethod
    def _open_folder(folder: Path):
        if os.name == "nt":
            os.startfile(folder)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(folder)])
        else:
            subprocess.Popen(["xdg-open", str(folder)])


def run_app():
    app = ReportGeneratorApp()
    app.mainloop()
