"""
PythonOffice - a small Microsoft Office-inspired desktop suite made with Tkinter.

PyCharm-friendly:
- No external runtime dependencies for the app itself.
- Run this file directly from PyCharm.
- Build a Windows .exe with build_exe.bat on Windows.
"""

from __future__ import annotations

import csv
import json
import os
import re
import subprocess
import sys
import tkinter as tk
from dataclasses import dataclass, asdict
from pathlib import Path
from tkinter import filedialog, messagebox, ttk, colorchooser, simpledialog


APP_NAME = "PythonOffice"
APP_VERSION = "1.0.0"
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)


# -----------------------------
# Shared helpers
# -----------------------------
def open_with_default_app(path: Path) -> None:
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
    except Exception as exc:
        messagebox.showerror(APP_NAME, f"Не удалось открыть файл:\n{exc}")


def show_about(parent: tk.Misc) -> None:
    messagebox.showinfo(
        "О PythonOffice",
        f"{APP_NAME} {APP_VERSION}\n\n"
        "Небольшой офисный пакет на Python/Tkinter.\n"
        "Writer • Calc • Impress\n\n"
        "Проект создан так, чтобы его было удобно запускать и развивать в PyCharm.",
        parent=parent,
    )


# -----------------------------
# Writer
# -----------------------------
class WriterTab(ttk.Frame):
    def __init__(self, master, status_cb):
        super().__init__(master)
        self.status_cb = status_cb
        self.current_path: Path | None = None
        self._build_ui()

    def _build_ui(self):
        toolbar = ttk.Frame(self)
        toolbar.pack(fill="x", padx=8, pady=(8, 4))

        ttk.Button(toolbar, text="Новый", command=self.new_doc).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Открыть", command=self.open_doc).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Сохранить", command=self.save_doc).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Сохранить как", command=self.save_as).pack(side="left", padx=2)
        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Button(toolbar, text="Жирный", command=lambda: self._toggle_tag("bold")).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Курсив", command=lambda: self._toggle_tag("italic")).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Подчёркнутый", command=lambda: self._toggle_tag("underline")).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Цвет", command=self._choose_color).pack(side="left", padx=2)

        size_var = tk.StringVar(value="12")
        ttk.Label(toolbar, text="Размер:").pack(side="left", padx=(10, 2))
        size_box = ttk.Combobox(toolbar, width=5, textvariable=size_var, values=[str(n) for n in range(8, 31)], state="readonly")
        size_box.pack(side="left")
        size_box.bind("<<ComboboxSelected>>", lambda _e: self._set_font_size(int(size_var.get())))

        body = ttk.Frame(self)
        body.pack(fill="both", expand=True, padx=8, pady=4)

        yscroll = ttk.Scrollbar(body, orient="vertical")
        yscroll.pack(side="right", fill="y")
        self.text = tk.Text(
            body,
            wrap="word",
            undo=True,
            font=("Segoe UI", 12),
            padx=28,
            pady=24,
            bg="white",
            fg="#202020",
            insertwidth=2,
            yscrollcommand=yscroll.set,
        )
        self.text.pack(fill="both", expand=True)
        yscroll.config(command=self.text.yview)

        self.text.tag_configure("bold", font=("Segoe UI", 12, "bold"))
        self.text.tag_configure("italic", font=("Segoe UI", 12, "italic"))
        self.text.tag_configure("underline", font=("Segoe UI", 12, "underline"))

        bottom = ttk.Frame(self)
        bottom.pack(fill="x", padx=8, pady=(2, 8))
        self.info = ttk.Label(bottom, text="Новый документ")
        self.info.pack(side="left")

        self.text.bind("<<Modified>>", self._on_modified)
        self._on_modified()

    def _on_modified(self, _event=None):
        content = self.text.get("1.0", "end-1c")
        words = len(re.findall(r"\S+", content))
        chars = len(content)
        self.status_cb(f"Writer • {words} слов • {chars} символов")
        self.text.edit_modified(False)

    def new_doc(self):
        self.text.delete("1.0", "end")
        self.current_path = None
        self.info.config(text="Новый документ")
        self.status_cb("Writer • новый документ")

    def open_doc(self):
        path = filedialog.askopenfilename(
            title="Открыть документ",
            filetypes=[("Текст", "*.txt"), ("Все файлы", "*.*")],
        )
        if not path:
            return
        try:
            data = Path(path).read_text(encoding="utf-8")
            self.text.delete("1.0", "end")
            self.text.insert("1.0", data)
            self.current_path = Path(path)
            self.info.config(text=self.current_path.name)
            self.status_cb(f"Открыт: {self.current_path.name}")
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Ошибка открытия:\n{exc}")

    def save_doc(self):
        if self.current_path is None:
            return self.save_as()
        try:
            self.current_path.write_text(self.text.get("1.0", "end-1c"), encoding="utf-8")
            self.info.config(text=self.current_path.name)
            self.status_cb(f"Сохранено: {self.current_path.name}")
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Ошибка сохранения:\n{exc}")

    def save_as(self):
        path = filedialog.asksaveasfilename(
            title="Сохранить документ",
            defaultextension=".txt",
            filetypes=[("Текстовый файл", "*.txt"), ("Все файлы", "*.*")],
        )
        if not path:
            return
        self.current_path = Path(path)
        self.save_doc()

    def _selection(self):
        try:
            return self.text.index("sel.first"), self.text.index("sel.last")
        except tk.TclError:
            return None

    def _toggle_tag(self, tag):
        sel = self._selection()
        if not sel:
            return
        start, end = sel
        if tag in self.text.tag_names("sel.first"):
            self.text.tag_remove(tag, start, end)
        else:
            self.text.tag_add(tag, start, end)

    def _choose_color(self):
        sel = self._selection()
        if not sel:
            return
        color = colorchooser.askcolor(title="Цвет текста")[1]
        if not color:
            return
        tag = f"color_{color.replace('#', '')}"
        self.text.tag_configure(tag, foreground=color)
        self.text.tag_add(tag, *sel)

    def _set_font_size(self, size: int):
        sel = self._selection()
        if not sel:
            return
        tag = f"size_{size}"
        self.text.tag_configure(tag, font=("Segoe UI", size))
        self.text.tag_add(tag, *sel)


# -----------------------------
# Calc
# -----------------------------
class CalcTab(ttk.Frame):
    def __init__(self, master, status_cb):
        super().__init__(master)
        self.status_cb = status_cb
        self.path: Path | None = None
        self._build()

    def _build(self):
        top = ttk.Frame(self)
        top.pack(fill="x", padx=8, pady=8)

        ttk.Button(top, text="Новая таблица", command=self.new_sheet).pack(side="left", padx=2)
        ttk.Button(top, text="Импорт CSV", command=self.import_csv).pack(side="left", padx=2)
        ttk.Button(top, text="Экспорт CSV", command=self.export_csv).pack(side="left", padx=2)
        ttk.Button(top, text="Сохранить JSON", command=self.save_json).pack(side="left", padx=2)
        ttk.Button(top, text="Открыть JSON", command=self.open_json).pack(side="left", padx=2)
        ttk.Button(top, text="Формула", command=self.formula_help).pack(side="left", padx=2)

        holder = ttk.Frame(self)
        holder.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        self.tree = ttk.Treeview(holder, columns=[f"C{i}" for i in range(1, 11)], show="headings")
        for i in range(1, 11):
            self.tree.heading(f"C{i}", text=chr(64 + i))
            self.tree.column(f"C{i}", width=110, anchor="center")
        self.tree.pack(side="left", fill="both", expand=True)

        sy = ttk.Scrollbar(holder, orient="vertical", command=self.tree.yview)
        sy.pack(side="right", fill="y")
        sx = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        sx.pack(fill="x", padx=8)
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)

        self.tree.bind("<Double-1>", self.edit_cell)
        self._populate(20, 10)

    def _populate(self, rows, cols):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for r in range(rows):
            values = [""] * cols
            self.tree.insert("", "end", iid=str(r), values=values)

    def new_sheet(self):
        self.path = None
        self._populate(20, 10)
        self.status_cb("Calc • новая таблица")

    def edit_cell(self, event):
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id:
            return
        x, y, w, h = self.tree.bbox(row_id, col_id)
        if not w:
            return
        current = self.tree.set(row_id, col_id)
        entry = ttk.Entry(self.tree)
        entry.insert(0, current)
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()

        def finish(_event=None):
            value = entry.get()
            entry.destroy()
            self.tree.set(row_id, col_id, value)
            self.status_cb(f"Calc • ячейка {col_id[1:]} изменена")

        entry.bind("<Return>", finish)
        entry.bind("<Escape>", lambda _e: entry.destroy())
        entry.bind("<FocusOut>", finish)

    def _matrix(self):
        return [list(self.tree.item(i, "values")) for i in self.tree.get_children()]

    def import_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV", "*.csv"), ("Все файлы", "*.*")])
        if not path:
            return
        try:
            with open(path, newline="", encoding="utf-8-sig") as f:
                rows = list(csv.reader(f))
            rows = [row[:10] + [""] * max(0, 10 - len(row)) for row in rows]
            self._populate(max(20, len(rows)), 10)
            for r, values in enumerate(rows):
                self.tree.item(str(r), values=values)
            self.path = Path(path)
            self.status_cb(f"Импортировано: {self.path.name}")
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Ошибка CSV:\n{exc}")

    def export_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                csv.writer(f).writerows(self._matrix())
            self.status_cb(f"CSV сохранён: {Path(path).name}")
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Ошибка экспорта:\n{exc}")

    def save_json(self):
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if not path:
            return
        try:
            Path(path).write_text(json.dumps(self._matrix(), ensure_ascii=False, indent=2), encoding="utf-8")
            self.status_cb(f"JSON сохранён: {Path(path).name}")
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Ошибка сохранения:\n{exc}")

    def open_json(self):
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not path:
            return
        try:
            matrix = json.loads(Path(path).read_text(encoding="utf-8"))
            matrix = [list(map(str, r))[:10] for r in matrix]
            self._populate(max(20, len(matrix)), 10)
            for r, values in enumerate(matrix):
                values += [""] * (10 - len(values))
                self.tree.item(str(r), values=values[:10])
            self.path = Path(path)
            self.status_cb(f"Открыт: {self.path.name}")
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Ошибка JSON:\n{exc}")

    def formula_help(self):
        messagebox.showinfo(
            "Формулы",
            "Доступные примеры в MVP:\n\n"
            "=SUM(A1:A5)\n"
            "=AVERAGE(B1:B5)\n"
            "=A1+B1\n\n"
            "Формулы пока отображаются как обычный текст, "
            "но структура таблицы готова для расширения вычислителем.",
        )


# -----------------------------
# Impress
# -----------------------------
@dataclass
class Slide:
    title: str = "Новый слайд"
    body: str = "Добавьте текст…"
    bg: str = "#f3f6fb"


class ImpressTab(ttk.Frame):
    def __init__(self, master, status_cb):
        super().__init__(master)
        self.status_cb = status_cb
        self.slides: list[Slide] = [Slide()]
        self.index = 0
        self.path: Path | None = None
        self._build()

    def _build(self):
        left = ttk.Frame(self, width=240)
        left.pack(side="left", fill="y", padx=(8, 4), pady=8)

        ttk.Label(left, text="Слайды", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        self.listbox = tk.Listbox(left, width=28, activestyle="none")
        self.listbox.pack(fill="both", expand=True, pady=6)
        self.listbox.bind("<<ListboxSelect>>", self.select_slide)

        btns = ttk.Frame(left)
        btns.pack(fill="x")
        ttk.Button(btns, text="+", width=4, command=self.add_slide).pack(side="left")
        ttk.Button(btns, text="−", width=4, command=self.delete_slide).pack(side="left", padx=3)
        ttk.Button(btns, text="↑", width=4, command=lambda: self.move_slide(-1)).pack(side="left")
        ttk.Button(btns, text="↓", width=4, command=lambda: self.move_slide(1)).pack(side="left", padx=3)

        right = ttk.Frame(self)
        right.pack(side="left", fill="both", expand=True, padx=(4, 8), pady=8)

        bar = ttk.Frame(right)
        bar.pack(fill="x", pady=(0, 8))
        ttk.Button(bar, text="Новая", command=self.new_deck).pack(side="left", padx=2)
        ttk.Button(bar, text="Сохранить", command=self.save).pack(side="left", padx=2)
        ttk.Button(bar, text="Открыть", command=self.open).pack(side="left", padx=2)
        ttk.Button(bar, text="Экспорт TXT", command=self.export_txt).pack(side="left", padx=2)
        ttk.Button(bar, text="Цвет", command=self.change_bg).pack(side="left", padx=2)
        ttk.Button(bar, text="Показ", command=self.show_slideshow).pack(side="left", padx=2)

        self.editor = ttk.Frame(right)
        self.editor.pack(fill="both", expand=True)

        ttk.Label(self.editor, text="Заголовок").pack(anchor="w")
        self.title_var = tk.StringVar()
        ttk.Entry(self.editor, textvariable=self.title_var, font=("Segoe UI", 15, "bold")).pack(fill="x", pady=(2, 12))

        ttk.Label(self.editor, text="Текст слайда").pack(anchor="w")
        self.body_text = tk.Text(self.editor, height=12, wrap="word", font=("Segoe UI", 12))
        self.body_text.pack(fill="both", expand=True, pady=2)

        ttk.Button(self.editor, text="Применить изменения", command=self.apply_current).pack(anchor="e", pady=8)

        self.refresh_list()
        self.load_current()

    def refresh_list(self):
        self.listbox.delete(0, "end")
        for i, slide in enumerate(self.slides, start=1):
            self.listbox.insert("end", f"{i}. {slide.title}")
        if self.slides:
            self.listbox.selection_clear(0, "end")
            self.listbox.selection_set(self.index)
            self.listbox.activate(self.index)

    def load_current(self):
        if not self.slides:
            return
        slide = self.slides[self.index]
        self.title_var.set(slide.title)
        self.body_text.delete("1.0", "end")
        self.body_text.insert("1.0", slide.body)

    def apply_current(self):
        if not self.slides:
            return
        self.slides[self.index].title = self.title_var.get() or "Без заголовка"
        self.slides[self.index].body = self.body_text.get("1.0", "end-1c")
        self.refresh_list()
        self.status_cb(f"Impress • слайд {self.index + 1} изменён")

    def select_slide(self, _event=None):
        sel = self.listbox.curselection()
        if not sel:
            return
        self.apply_current()
        self.index = sel[0]
        self.load_current()

    def add_slide(self):
        self.apply_current()
        self.slides.insert(self.index + 1, Slide(title=f"Слайд {len(self.slides) + 1}"))
        self.index += 1
        self.refresh_list()
        self.load_current()

    def delete_slide(self):
        if len(self.slides) <= 1:
            return
        self.slides.pop(self.index)
        self.index = min(self.index, len(self.slides) - 1)
        self.refresh_list()
        self.load_current()

    def move_slide(self, delta):
        self.apply_current()
        new_i = self.index + delta
        if not (0 <= new_i < len(self.slides)):
            return
        self.slides[self.index], self.slides[new_i] = self.slides[new_i], self.slides[self.index]
        self.index = new_i
        self.refresh_list()
        self.load_current()

    def new_deck(self):
        self.slides = [Slide()]
        self.index = 0
        self.path = None
        self.refresh_list()
        self.load_current()

    def change_bg(self):
        color = colorchooser.askcolor(title="Фон слайда")[1]
        if color and self.slides:
            self.apply_current()
            self.slides[self.index].bg = color

    def save(self):
        self.apply_current()
        if not self.path:
            path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("Presentation JSON", "*.json")])
            if not path:
                return
            self.path = Path(path)
        self.path.write_text(
            json.dumps([asdict(s) for s in self.slides], ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        self.status_cb(f"Impress • сохранено: {self.path.name}")

    def open(self):
        path = filedialog.askopenfilename(filetypes=[("Presentation JSON", "*.json")])
        if not path:
            return
        try:
            raw = json.loads(Path(path).read_text(encoding="utf-8"))
            self.slides = [Slide(**item) for item in raw] or [Slide()]
            self.index = 0
            self.path = Path(path)
            self.refresh_list()
            self.load_current()
            self.status_cb(f"Impress • открыт: {self.path.name}")
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Ошибка открытия презентации:\n{exc}")

    def export_txt(self):
        self.apply_current()
        path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text", "*.txt")])
        if not path:
            return
        parts = []
        for i, slide in enumerate(self.slides, start=1):
            parts.append(f"=== СЛАЙД {i} ===\n{slide.title}\n\n{slide.body}\n")
        Path(path).write_text("\n".join(parts), encoding="utf-8")
        self.status_cb(f"Impress • экспорт: {Path(path).name}")

    def show_slideshow(self):
        self.apply_current()
        win = tk.Toplevel(self)
        win.title("PythonOffice • Показ слайдов")
        win.geometry("1000x650")
        win.configure(bg="#111")

        idx = {"value": 0}
        frame = tk.Frame(win, bg=self.slides[0].bg)
        frame.pack(fill="both", expand=True, padx=28, pady=28)

        title = tk.Label(frame, text="", font=("Segoe UI", 30, "bold"), bg=self.slides[0].bg, fg="#111", wraplength=850)
        title.pack(anchor="nw", padx=35, pady=(40, 20))
        body = tk.Label(frame, text="", font=("Segoe UI", 18), bg=self.slides[0].bg, fg="#222", justify="left", anchor="nw", wraplength=850)
        body.pack(fill="both", expand=True, padx=35, pady=10)

        def render():
            s = self.slides[idx["value"]]
            frame.configure(bg=s.bg)
            title.configure(text=s.title, bg=s.bg)
            body.configure(text=s.body, bg=s.bg)
            win.title(f"Показ слайдов • {idx['value'] + 1}/{len(self.slides)}")

        def next_slide(_e=None):
            idx["value"] = min(len(self.slides) - 1, idx["value"] + 1)
            render()

        def prev_slide(_e=None):
            idx["value"] = max(0, idx["value"] - 1)
            render()

        win.bind("<Right>", next_slide)
        win.bind("<space>", next_slide)
        win.bind("<Left>", prev_slide)
        win.bind("<Escape>", lambda _e: win.destroy())
        render()
        win.focus_set()


# -----------------------------
# App shell
# -----------------------------
class PythonOfficeApp(tk.Tk):

    def __init__(self):
        super().__init__()

        self.title(f"{APP_NAME} {APP_VERSION}")
        self.geometry("1280x820")
        self.minsize(1000, 650)
        self.configure(bg="#eef2f7")

        # Сначала создаём стиль
        self._setup_style()

        # Потом меню и верхнюю часть
        self._build_menu()
        self._build_header()

        # ВАЖНО:
        # status_var должен существовать ДО создания вкладок,
        # потому что WriterTab вызывает status_cb() во время запуска.
        self._build_status()

        # И только после этого создаём вкладки
        self._build_tabs()
    def _setup_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TNotebook", tabposition="n", background="#eef2f7")
        style.configure("TNotebook.Tab", padding=(18, 10), font=("Segoe UI", 10, "bold"))
        style.configure("TButton", font=("Segoe UI", 10))
        style.configure("TLabel", font=("Segoe UI", 10))
        style.configure("Header.TLabel", font=("Segoe UI", 20, "bold"))
        style.configure("Subtitle.TLabel", font=("Segoe UI", 10))

    def _build_menu(self):
        menu = tk.Menu(self)
        file_menu = tk.Menu(menu, tearoff=False)
        file_menu.add_command(label="О Writer", command=lambda: self.tabs.select(self.writer))
        file_menu.add_command(label="О Calc", command=lambda: self.tabs.select(self.calc))
        file_menu.add_command(label="О Impress", command=lambda: self.tabs.select(self.impress))
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=self.destroy)
        help_menu = tk.Menu(menu, tearoff=False)
        help_menu.add_command(label="О программе", command=lambda: show_about(self))
        menu.add_cascade(label="Файл", menu=file_menu)
        menu.add_cascade(label="Помощь", menu=help_menu)
        self.config(menu=menu)

    def _build_header(self):
        header = ttk.Frame(self)
        header.pack(fill="x", padx=18, pady=(16, 8))
        ttk.Label(header, text="PythonOffice", style="Header.TLabel").pack(side="left")
        ttk.Label(
            header,
            text="   Writer  •  Calc  •  Impress",
            style="Subtitle.TLabel",
        ).pack(side="left", pady=(7, 0))

    def _build_tabs(self):
        self.tabs = ttk.Notebook(self)
        self.tabs.pack(fill="both", expand=True, padx=14, pady=8)

        self.writer = WriterTab(self.tabs, self.set_status)
        self.calc = CalcTab(self.tabs, self.set_status)
        self.impress = ImpressTab(self.tabs, self.set_status)

        self.tabs.add(self.writer, text="Writer")
        self.tabs.add(self.calc, text="Calc")
        self.tabs.add(self.impress, text="Impress")

    def _build_status(self):
        self.status_var = tk.StringVar(value="Готово")
        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=14, pady=(0, 10))
        ttk.Label(bar, textvariable=self.status_var).pack(side="left")

    def set_status(self, message: str):
        self.status_var.set(message)


def main():
    app = PythonOfficeApp()
    app.mainloop()


if __name__ == "__main__":
    main()
