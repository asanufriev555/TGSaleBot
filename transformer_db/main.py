# -*- coding: utf-8 -*-
"""Каталог проектов поставки трансформаторов (ГОСТ Р 52719).
Python 3.8 + tkinter + sqlite3 — работает на Windows 7 без дополнительных библиотек."""
import csv
import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

from db import DB
from scanner import Scanner

APP_DIR = os.path.dirname(os.path.abspath(sys.argv[0] if getattr(sys, "frozen", False) else __file__))
CONFIG = os.path.join(APP_DIR, "config.json")
ALL = "(все)"

PROJECT_FIELDS = [  # key, подпись, тип
    ("name", "Название проекта", "s"), ("year", "Год", "i"), ("category", "Тип трансформатора", "cat"),
    ("designation", "Обозначение (напр. ТДН-16000/110)", "s"), ("power_kva", "Мощность, кВА", "f"),
    ("voltage_hv", "Напряжение ВН, кВ", "f"), ("voltage_lv", "Напряжение НН, кВ", "f"),
    ("cooling", "Охлаждение", "s"), ("quantity", "Количество, шт", "i"), ("path", "Папка проекта", "s"),
    ("notes", "Примечания", "s"),
]


def load_config():
    try:
        with open(CONFIG, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(cfg):
    try:
        with open(CONFIG, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=1)
    except Exception:
        pass


def conv(kind, s):
    s = (s or "").strip()
    if s == "":
        return None
    try:
        if kind == "i":
            return int(float(s.replace(",", ".")))
        if kind == "f":
            return float(s.replace(",", "."))
    except ValueError:
        raise ValueError("Некорректное число: %s" % s)
    return s


class FormDialog(tk.Toplevel):
    """Универсальная форма редактирования: fields = [(key, label, value, choices|None)]."""

    def __init__(self, parent, title, fields, width=48):
        tk.Toplevel.__init__(self, parent)
        self.title(title)
        self.transient(parent)
        self.result = None
        self.vars = {}
        frm = ttk.Frame(self, padding=10)
        frm.pack(fill="both", expand=True)
        for i, (key, label, value, choices) in enumerate(fields):
            ttk.Label(frm, text=label).grid(row=i, column=0, sticky="w", pady=2, padx=(0, 8))
            v = tk.StringVar(value="" if value is None else str(value))
            self.vars[key] = v
            w = ttk.Combobox(frm, textvariable=v, values=choices, width=width - 3) if choices \
                else ttk.Entry(frm, textvariable=v, width=width)
            w.grid(row=i, column=1, sticky="ew", pady=2)
        bt = ttk.Frame(frm)
        bt.grid(row=len(fields), column=0, columnspan=2, pady=(10, 0), sticky="e")
        ttk.Button(bt, text="Сохранить", command=self.ok).pack(side="left", padx=4)
        ttk.Button(bt, text="Отмена", command=self.destroy).pack(side="left")
        self.bind("<Return>", lambda e: self.ok())
        self.bind("<Escape>", lambda e: self.destroy())
        self.grab_set()
        self.wait_window(self)

    def ok(self):
        self.result = {k: v.get() for k, v in self.vars.items()}
        self.destroy()


class App(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.title("Каталог проектов: трансформаторы (ГОСТ Р 52719)")
        self.geometry("1150x680")
        self.minsize(900, 520)
        self.cfg = load_config()
        self.db = None
        self.q = queue.Queue()
        self.scanner = None
        self.thread = None
        self._build()
        self.open_db(self.cfg.get("db", os.path.join(APP_DIR, "transformers.db")))
        self.after(150, self._poll)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------------- каркас ----------------
    def _build(self):
        top = ttk.Frame(self, padding=(8, 6))
        top.pack(fill="x")
        ttk.Label(top, text="Файл базы:").pack(side="left")
        self.db_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.db_var, state="readonly").pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(top, text="Открыть / создать…", command=self.choose_db).pack(side="left")
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.tab_scan, self.tab_proj, self.tab_sum, self.tab_par = (ttk.Frame(self.nb) for _ in range(4))
        for t, n in ((self.tab_scan, "Сканирование"), (self.tab_proj, "Проекты"),
                     (self.tab_sum, "По типам и годам"), (self.tab_par, "Параметры")):
            self.nb.add(t, text=n)
        self._build_scan()
        self._build_projects()
        self._build_summary()
        self._build_params()
        self.nb.bind("<<NotebookTabChanged>>", lambda e: self.refresh_all())

    # ---------------- сканирование ----------------
    def _build_scan(self):
        f = ttk.Frame(self.tab_scan, padding=10)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="Адрес папки (локальный путь или сетевой, напр. \\\\SERVER\\Проекты\\Трансформаторы):").pack(anchor="w")
        row = ttk.Frame(f)
        row.pack(fill="x", pady=4)
        self.path_var = tk.StringVar()
        ttk.Entry(row, textvariable=self.path_var).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="Обзор…", command=self.browse).pack(side="left", padx=6)
        row2 = ttk.Frame(f)
        row2.pack(fill="x", pady=4)
        self.btn_scan = ttk.Button(row2, text="Сканировать", command=self.start_scan)
        self.btn_scan.pack(side="left")
        self.btn_stop = ttk.Button(row2, text="Остановить", command=self.stop_scan, state="disabled")
        self.btn_stop.pack(side="left", padx=6)
        self.status = tk.StringVar(value="Готово")
        ttk.Label(row2, textvariable=self.status).pack(side="left", padx=12)
        ttk.Label(f, text="Правки, сделанные вручную, при повторном сканировании не перезаписываются.",
                  foreground="#555").pack(anchor="w")
        self.log = tk.Text(f, height=20, state="disabled", wrap="none")
        sb = ttk.Scrollbar(f, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log.pack(fill="both", expand=True, pady=(6, 0))

    def browse(self):
        d = filedialog.askdirectory(initialdir=self.path_var.get() or None, title="Выберите папку")
        if d:
            self.path_var.set(os.path.normpath(d))

    def log_add(self, s):
        self.log.configure(state="normal")
        self.log.insert("end", s + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def start_scan(self):
        path = self.path_var.get().strip().strip('"')
        if not path:
            messagebox.showwarning("Сканирование", "Введите адрес папки.")
            return
        if not os.path.isdir(path):
            messagebox.showerror("Сканирование", "Папка недоступна:\n%s\nПроверьте адрес и подключение к сети." % path)
            return
        self.db.set("root_path", path)
        self.btn_scan.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.status.set("Идёт сканирование…")
        self.log_add("=== Сканирование: %s ===" % path)
        # отдельное подключение к SQLite для рабочего потока
        db_path = self.db.path

        def work():
            wdb = DB(db_path)
            self.scanner = Scanner(wdb, log=lambda s: self.q.put(("log", s)),
                                   progress=lambda n: self.q.put(("prog", n)))
            try:
                n = self.scanner.run(path)
                self.q.put(("done", "Готово. Проектов: %d" % n))
            except Exception as e:
                self.q.put(("done", "Ошибка: %s" % e))
            finally:
                wdb.close()

        self.thread = threading.Thread(target=work, daemon=True)
        self.thread.start()

    def stop_scan(self):
        if self.scanner:
            self.scanner.cancel = True

    def _poll(self):
        try:
            while True:
                kind, val = self.q.get_nowait()
                if kind == "log":
                    self.log_add(val)
                elif kind == "prog":
                    self.status.set("Найдено файлов: %s" % val)
                else:
                    self.status.set(val)
                    self.log_add(val)
                    self.btn_scan.configure(state="normal")
                    self.btn_stop.configure(state="disabled")
                    self.refresh_all()
        except queue.Empty:
            pass
        self.after(150, self._poll)

    # ---------------- проекты ----------------
    def _build_projects(self):
        f = ttk.Frame(self.tab_proj, padding=8)
        f.pack(fill="both", expand=True)
        flt = ttk.Frame(f)
        flt.pack(fill="x")
        self.f_year, self.f_cat, self.f_text = tk.StringVar(value=ALL), tk.StringVar(value=ALL), tk.StringVar()
        ttk.Label(flt, text="Год:").pack(side="left")
        self.cb_year = ttk.Combobox(flt, textvariable=self.f_year, width=8, state="readonly")
        self.cb_year.pack(side="left", padx=(4, 12))
        ttk.Label(flt, text="Тип:").pack(side="left")
        self.cb_cat = ttk.Combobox(flt, textvariable=self.f_cat, width=36, state="readonly")
        self.cb_cat.pack(side="left", padx=(4, 12))
        ttk.Label(flt, text="Поиск:").pack(side="left")
        e = ttk.Entry(flt, textvariable=self.f_text, width=24)
        e.pack(side="left", padx=4)
        for w in (self.cb_year, self.cb_cat):
            w.bind("<<ComboboxSelected>>", lambda ev: self.refresh_projects())
        e.bind("<KeyRelease>", lambda ev: self.refresh_projects())
        cols = [("name", "Проект", 200), ("year", "Год", 50), ("category", "Тип", 210), ("designation", "Обозначение", 140),
                ("power", "Мощн., кВА", 80), ("volt", "Напряжение, кВ", 100), ("qty", "Шт", 40),
                ("gost", "ГОСТ Р 52719", 90), ("files", "Файлов", 55), ("manual", "Ручн.", 45)]
        self.tree = ttk.Treeview(f, columns=[c[0] for c in cols], show="headings", selectmode="browse")
        for k, t, w in cols:
            self.tree.heading(k, text=t, command=lambda k=k: self.sort_by(k))
            self.tree.column(k, width=w, anchor="w")
        sb = ttk.Scrollbar(f, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        bt = ttk.Frame(f)
        bt.pack(side="bottom", fill="x", pady=(6, 0))
        for text, cmd in (("Добавить", self.add_project), ("Изменить", self.edit_project), ("Удалить", self.del_project),
                          ("Открыть папку", self.open_folder), ("Подробнее (файлы, ГОСТ)", self.show_details),
                          ("Экспорт в CSV…", self.export_csv)):
            ttk.Button(bt, text=text, command=cmd).pack(side="left", padx=(0, 6))
        self.count_var = tk.StringVar()
        ttk.Label(bt, textvariable=self.count_var).pack(side="right")
        sb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True, pady=6)
        self.tree.bind("<Double-1>", lambda e: self.edit_project())
        self._sort = ("year", True)

    def sort_by(self, col):
        items = [(self.tree.set(i, col), i) for i in self.tree.get_children("")]
        rev = self._sort == (col, False)
        def key(x):
            try:
                return (0, float(x[0].replace(",", ".").split("/")[0]))
            except ValueError:
                return (1, x[0].lower())
        items.sort(key=key, reverse=rev)
        for idx, (_, i) in enumerate(items):
            self.tree.move(i, "", idx)
        self._sort = (col, rev)

    def refresh_projects(self):
        if not self.db:
            return
        years, cats = [ALL] + [str(y) for y in self.db.years()], [ALL] + self.db.categories()
        self.cb_year["values"], self.cb_cat["values"] = years, cats
        if self.f_year.get() not in years:
            self.f_year.set(ALL)
        if self.f_cat.get() not in cats:
            self.f_cat.set(ALL)
        rows = self.db.projects(None if self.f_year.get() == ALL else self.f_year.get(),
                                None if self.f_cat.get() == ALL else self.f_cat.get(),
                                self.f_text.get().strip() or None)
        self.tree.delete(*self.tree.get_children())
        fmt = lambda v: "" if v is None else ("%g" % v)
        for r in rows:
            volt = "/".join(fmt(v) for v in (r["voltage_hv"], r["voltage_lv"]) if v is not None)
            gost = "%d/%d" % (r["gost_found"], r["gost_total"]) if r["gost_total"] else ""
            self.tree.insert("", "end", iid=str(r["id"]), values=(
                r["name"], r["year"] or "", r["category"], r["designation"], fmt(r["power_kva"]), volt,
                r["quantity"] or "", gost, r["files_count"], "да" if r["manual"] else ""))
        self.count_var.set("Записей: %d" % len(rows))

    def selected_id(self):
        s = self.tree.selection()
        if not s:
            messagebox.showinfo("Проекты", "Выберите проект в списке.")
            return None
        return int(s[0])

    def categories_list(self):
        names = [r["category"] for r in self.db.type_rules()] + ["Не определён"]
        return list(dict.fromkeys(names + self.db.categories()))

    def project_form(self, title, rec):
        fields = []
        for key, label, kind in PROJECT_FIELDS:
            val = rec[key] if rec is not None else ""
            fields.append((key, label, val, self.categories_list() if kind == "cat" else None))
        dlg = FormDialog(self, title, fields)
        if dlg.result is None:
            return None
        try:
            return {key: conv(kind if kind != "cat" else "s", dlg.result[key]) for key, _, kind in PROJECT_FIELDS}
        except ValueError as e:
            messagebox.showerror("Ошибка ввода", str(e))
            return None

    def add_project(self):
        data = self.project_form("Новый проект", None)
        if data and data["name"]:
            data = {k: v for k, v in data.items() if v is not None}
            if not data.get("path"):
                data.pop("path", None)
            self.db.add_project(**data)
            self.refresh_all()

    def edit_project(self):
        pid = self.selected_id()
        if pid is None:
            return
        data = self.project_form("Изменить проект", self.db.project(pid))
        if data and data["name"]:
            if not data.get("path"):
                data.pop("path")
            self.db.update_project(pid, **data)
            self.refresh_all()

    def del_project(self):
        pid = self.selected_id()
        if pid and messagebox.askyesno("Удаление", "Удалить проект из базы? Файлы на диске не затрагиваются.\n"
                                       "При следующем сканировании он появится снова."):
            self.db.delete_project(pid)
            self.refresh_all()

    def open_folder(self):
        pid = self.selected_id()
        if pid is None:
            return
        p = self.db.project(pid)["path"]
        if not os.path.isdir(p):
            messagebox.showwarning("Папка", "Папка недоступна:\n%s" % p)
            return
        try:
            if os.name == "nt":
                os.startfile(p)
            else:
                subprocess.Popen(["xdg-open", p])
        except Exception as e:
            messagebox.showerror("Папка", str(e))

    def show_details(self):
        pid = self.selected_id()
        if pid is None:
            return
        p = self.db.project(pid)
        w = tk.Toplevel(self)
        w.title(p["name"])
        w.geometry("760x520")
        t = tk.Text(w, wrap="word")
        t.pack(fill="both", expand=True)
        t.insert("end", "Проект: %s\nПапка: %s\n\n" % (p["name"], p["path"]))
        t.insert("end", "Соответствие чек-листу ГОСТ Р 52719: %d из %d параметров найдено в ТЗ\n" % (p["gost_found"], p["gost_total"]))
        if p["gost_missing"]:
            t.insert("end", "Не найдено в ТЗ:\n  • " + p["gost_missing"].replace("; ", "\n  • ") + "\n")
        t.insert("end", "\nФайлы:\n")
        for f in self.db.project_files(pid):
            t.insert("end", "  %s %s  (%s, %d КБ)%s\n" % ("[ТЗ]" if f["is_spec"] else "    ", f["path"], f["mtime"],
                                                          f["size"] // 1024, "  → " + f["designation"] if f["designation"] else ""))
        t.configure(state="disabled")

    def export_csv(self):
        fn = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV (Excel)", "*.csv")],
                                          initialfile="transformers.csv")
        if not fn:
            return
        rows = self.db.projects(None if self.f_year.get() == ALL else self.f_year.get(),
                                None if self.f_cat.get() == ALL else self.f_cat.get(),
                                self.f_text.get().strip() or None)
        with open(fn, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow([l for _, l, _ in PROJECT_FIELDS] + ["ГОСТ найдено", "ГОСТ всего", "Не найдено в ТЗ"])
            for r in rows:
                w.writerow([r[k] if r[k] is not None else "" for k, _, _ in PROJECT_FIELDS] +
                           [r["gost_found"], r["gost_total"], r["gost_missing"]])
        messagebox.showinfo("Экспорт", "Сохранено: %s" % fn)

    # ---------------- сводка ----------------
    def _build_summary(self):
        f = ttk.Frame(self.tab_sum, padding=8)
        f.pack(fill="both", expand=True)
        self.mode = tk.StringVar(value="type")
        row = ttk.Frame(f)
        row.pack(fill="x")
        ttk.Radiobutton(row, text="Тип → Год → Проекты", variable=self.mode, value="type", command=self.refresh_summary).pack(side="left")
        ttk.Radiobutton(row, text="Год → Тип → Проекты", variable=self.mode, value="year", command=self.refresh_summary).pack(side="left", padx=12)
        self.stree = ttk.Treeview(f, columns=("n", "p"), show="tree headings")
        self.stree.heading("#0", text="Группа / проект")
        self.stree.heading("n", text="Кол-во")
        self.stree.heading("p", text="Мощность, кВА")
        self.stree.column("#0", width=520)
        self.stree.column("n", width=80, anchor="e")
        self.stree.column("p", width=120, anchor="e")
        sb = ttk.Scrollbar(f, command=self.stree.yview)
        self.stree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.stree.pack(fill="both", expand=True, pady=6)

    def refresh_summary(self):
        self.stree.delete(*self.stree.get_children())
        rows = self.db.projects()
        k1, k2 = (("category", "year") if self.mode.get() == "type" else ("year", "category"))
        groups = {}
        for r in rows:
            groups.setdefault(r[k1] if r[k1] is not None else "Без года", {}).setdefault(
                r[k2] if r[k2] is not None else "Без года", []).append(r)
        for a in sorted(groups, key=lambda x: str(x), reverse=(k1 == "year")):
            cnt = sum(len(v) for v in groups[a].values())
            na = self.stree.insert("", "end", text=str(a), values=(cnt, ""), open=True)
            for b in sorted(groups[a], key=lambda x: str(x), reverse=(k2 == "year")):
                lst = groups[a][b]
                nb = self.stree.insert(na, "end", text=str(b), values=(len(lst), ""))
                for r in lst:
                    self.stree.insert(nb, "end", text=r["name"] + ("  [%s]" % r["designation"] if r["designation"] else ""),
                                      values=("", "%g" % r["power_kva"] if r["power_kva"] else ""))

    # ---------------- параметры ----------------
    def _build_params(self):
        f = ttk.Frame(self.tab_par, padding=8)
        f.pack(fill="both", expand=True)
        f.columnconfigure(0, weight=1)
        f.columnconfigure(1, weight=2)
        f.rowconfigure(1, weight=1)
        f.rowconfigure(3, weight=1)
        # настройки
        ttk.Label(f, text="Настройки сканирования (двойной щелчок — изменить)").grid(row=0, column=0, sticky="w")
        self.set_tree = ttk.Treeview(f, columns=("v",), show="tree headings", height=8)
        self.set_tree.heading("#0", text="Параметр")
        self.set_tree.heading("v", text="Значение")
        self.set_tree.column("#0", width=140)
        self.set_tree.grid(row=1, column=0, sticky="nsew", padx=(0, 8))
        self.set_tree.bind("<Double-1>", self.edit_setting)
        ttk.Label(f, text="Пояснения: project_depth — уровень папки проекта (1 = подпапки корня); extensions — типы файлов; "
                          "spec_keywords — слова в имени файла/папки, по которым файл считается ТЗ; max_file_mb — предел размера для разбора.",
                  wraplength=330, foreground="#555").grid(row=2, column=0, sticky="nw", padx=(0, 8))
        # правила типов
        ttk.Label(f, text="Типы трансформаторов: категория ← регулярное выражение (порядок = приоритет)").grid(row=0, column=1, sticky="w")
        self.rule_tree = ttk.Treeview(f, columns=("pat", "pr"), show="tree headings", height=8)
        self.rule_tree.heading("#0", text="Категория")
        self.rule_tree.heading("pat", text="Шаблон")
        self.rule_tree.heading("pr", text="Приор.")
        self.rule_tree.column("#0", width=230)
        self.rule_tree.column("pat", width=320)
        self.rule_tree.column("pr", width=50)
        self.rule_tree.grid(row=1, column=1, sticky="nsew")
        self._table_buttons(f, 2, 1, "type_rules", self.rule_tree)
        ttk.Label(f, text="Чек-лист ГОСТ Р 52719: параметры, наличие которых проверяется в техзадании").grid(row=3, column=0, columnspan=2, sticky="w", pady=(10, 0))
        self.gost_tree = ttk.Treeview(f, columns=("pat",), show="tree headings", height=7)
        self.gost_tree.heading("#0", text="Параметр")
        self.gost_tree.heading("pat", text="Шаблон поиска (регулярное выражение)")
        self.gost_tree.column("#0", width=300)
        self.gost_tree.column("pat", width=600)
        self.gost_tree.grid(row=4, column=0, columnspan=2, sticky="nsew")
        f.rowconfigure(4, weight=1)
        self._table_buttons(f, 5, 0, "gost_checks", self.gost_tree, span=2)

    def _table_buttons(self, parent, row, col, table, tree, span=1):
        b = ttk.Frame(parent)
        b.grid(row=row, column=col, columnspan=span, sticky="w", pady=4)
        ttk.Button(b, text="Добавить", command=lambda: self.edit_rule(table, tree, False)).pack(side="left", padx=(0, 6))
        ttk.Button(b, text="Изменить", command=lambda: self.edit_rule(table, tree, True)).pack(side="left", padx=(0, 6))
        ttk.Button(b, text="Удалить", command=lambda: self.del_rule(table, tree)).pack(side="left")
        tree.bind("<Double-1>", lambda e: self.edit_rule(table, tree, True))

    def edit_setting(self, _e=None):
        s = self.set_tree.selection()
        if not s:
            return
        key = s[0]
        val = simpledialog.askstring("Параметр", key, initialvalue=self.db.get(key), parent=self)
        if val is not None:
            self.db.set(key, val)
            self.refresh_params()

    def edit_rule(self, table, tree, edit):
        rid, rec = None, None
        if edit:
            s = tree.selection()
            if not s:
                return
            rid = int(s[0])
            rec = self.db.conn.execute("SELECT * FROM %s WHERE id=?" % table, (rid,)).fetchone()
        if table == "type_rules":
            fields = [("category", "Категория", rec["category"] if rec else "", None),
                      ("pattern", "Регулярное выражение", rec["pattern"] if rec else "", None),
                      ("priority", "Приоритет (меньше — раньше)", rec["priority"] if rec else 100, None)]
        else:
            fields = [("title", "Параметр", rec["title"] if rec else "", None),
                      ("pattern", "Регулярное выражение", rec["pattern"] if rec else "", None)]
        dlg = FormDialog(self, "Правило", fields, width=70)
        if not dlg.result or not all(dlg.result.get(k) for k in ("pattern", fields[0][0])):
            return
        import re
        try:
            re.compile(dlg.result["pattern"])
        except re.error as e:
            messagebox.showerror("Регулярное выражение", "Ошибка в шаблоне: %s" % e)
            return
        vals = dict(dlg.result)
        if "priority" in vals:
            try:
                vals["priority"] = int(vals["priority"])
            except ValueError:
                vals["priority"] = 100
        self.db.save_rule(table, rid, **vals)
        self.refresh_params()

    def del_rule(self, table, tree):
        s = tree.selection()
        if s and messagebox.askyesno("Удаление", "Удалить выбранную запись?"):
            self.db.delete_row(table, int(s[0]))
            self.refresh_params()

    def refresh_params(self):
        for t in (self.set_tree, self.rule_tree, self.gost_tree):
            t.delete(*t.get_children())
        for k, v in self.db.all_settings():
            self.set_tree.insert("", "end", iid=k, text=k, values=(v,))
        for r in self.db.type_rules():
            self.rule_tree.insert("", "end", iid=str(r["id"]), text=r["category"], values=(r["pattern"], r["priority"]))
        for r in self.db.gost_checks():
            self.gost_tree.insert("", "end", iid=str(r["id"]), text=r["title"], values=(r["pattern"],))

    # ---------------- база ----------------
    def choose_db(self):
        fn = filedialog.asksaveasfilename(title="Файл базы данных (выберите существующий или введите новое имя)",
                                          defaultextension=".db", confirmoverwrite=False,
                                          filetypes=[("База SQLite", "*.db"), ("Все файлы", "*.*")],
                                          initialdir=os.path.dirname(self.db.path) if self.db else APP_DIR)
        if fn:
            self.open_db(fn)

    def open_db(self, path):
        try:
            if self.db:
                self.db.close()
            self.db = DB(path)
        except Exception as e:
            messagebox.showerror("База данных", "Не удалось открыть базу:\n%s" % e)
            self.db = DB(os.path.join(APP_DIR, "transformers.db"))
            path = self.db.path
        self.db_var.set(path)
        self.cfg["db"] = path
        save_config(self.cfg)
        self.path_var.set(self.db.get("root_path"))
        self.refresh_all()

    def refresh_all(self):
        if not self.db:
            return
        self.refresh_projects()
        self.refresh_summary()
        self.refresh_params()

    def on_close(self):
        if self.scanner and self.thread and self.thread.is_alive():
            self.scanner.cancel = True
        if self.db:
            self.db.close()
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
