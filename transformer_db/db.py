# -*- coding: utf-8 -*-
"""Слой базы данных (SQLite, один файл .db). Совместим с Python 3.8 / Windows 7."""
import sqlite3

DEFAULT_SETTINGS = {
    "root_path": "",
    "project_depth": "1",          # уровень вложенности папки проекта (1 = подпапки корня)
    "extensions": ".docx,.xlsx,.txt,.csv,.pdf,.doc,.xls,.rtf",
    "spec_keywords": "тз,техзад,тех.зад,техническое задание,техническ,специфик,опросн,трансформатор,transformer,tz,spec",
    "gost": "ГОСТ Р 52719",
    "max_file_mb": "30",
    "min_year": "1990",
    "max_year": "2100",
}

# Категории трансформаторов: (название, регулярное выражение по типу/тексту, приоритет)
DEFAULT_TYPE_RULES = [
    ("Автотрансформаторы", r"\b(АТДЦТН|АТДЦН|АОДЦТН|АОДЦН|АТ[ДО]Ц?Т?Н?)[-\s]?\d", 10),
    ("Трёхобмоточные масляные", r"\b(ТДТН|ТДЦТН|ТМТН|ТДТ)[-\s]?\d", 20),
    ("С расщеплённой обмоткой", r"\b(ТРДН[СЦ]?|ТРДЦН|ТРДНС)[-\s]?\d", 30),
    ("Сухие (ТС/ТСЗ)", r"\b(ТСЗН?|ТСЗГЛ|ТСЛ|ТСГЛ|ТСЗЛ|ТСГ|ТС)[-\s]?\d", 40),
    ("Элегазовые/негорючие (ТНЗ)", r"\b(ТНЗ|ТНГ|ТНС)[-\s]?\d", 50),
    ("Масляные герметичные (ТМГ)", r"\b(ТМГ\d*[А-Яа-я]*|ТМГФ|ТМГСУ)[-\s]?\d", 60),
    ("Масляные распределительные (ТМ/ТМН/ТМЗ)", r"\b(ТМН|ТМЗ|ТМФ|ТМ)[-\s]?\d", 70),
    ("Масляные двухобмоточные (ТД/ТДН/ТДЦ)", r"\b(ТДНС|ТДН|ТДЦ|ТД|ТЦ)[-\s]?\d", 80),
    ("Преобразовательные", r"\b(ТМП|ТСП|ТДП|ТЦП)[-\s]?\d", 90),
]

# Чек-лист обязательных технических параметров по ГОСТ Р 52719 (редактируемый)
DEFAULT_GOST_CHECKS = [
    ("Номинальная мощность", r"(мощност|квa|ква|kva|мва|mva)", 1),
    ("Номинальные напряжения", r"(напряжени|\d+\s?/\s?\d+(?:[.,]\d+)?\s?кв|\bкв\b)", 1),
    ("Группа соединения обмоток", r"(групп\w*\s+соединен|[YyУу]н?\s?/\s?[ДдDd]\s?-?\s?11|[Yy]\w?/[Dd]|Y/Yн|Δ|треугольн|звезд)", 1),
    ("Потери холостого хода", r"(потер\w*\s+х\w*\s+х|х\.?х\.?|холост)", 1),
    ("Потери короткого замыкания", r"(потер\w*\s+к\w*\s+з|к\.?з\.?|короткого\s+замыкан)", 1),
    ("Напряжение КЗ (uк)", r"(напряжени\w*\s+к\w*\s+з|uк|u\s?к|%\s*uк)", 1),
    ("Ток холостого хода", r"(ток\w*\s+х\w*\s+х|ток\w*\s+холост)", 1),
    ("Система охлаждения", r"(охлажден|охлаждающ|\bМ\b|\bД\b|\bДЦ\b|\bЦ\b|естественн)", 1),
    ("Регулирование напряжения (РПН/ПБВ)", r"(рпн|пбв|регулирован|ответвлен)", 1),
    ("Климатическое исполнение", r"(климат|\bУ1\b|\bУХЛ\d|\bУ[23]\b|\bТ1\b|15150)", 1),
    ("Уровень изоляции / испытания", r"(изоляци|испытан|импульс|приёмо|приемо)", 1),
    ("Ссылка на ГОСТ Р 52719", r"(52719)", 1),
]

SCHEMA = """
CREATE TABLE IF NOT EXISTS projects(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    path TEXT UNIQUE,
    year INTEGER,
    category TEXT DEFAULT 'Не определён',
    designation TEXT DEFAULT '',
    power_kva REAL,
    voltage_hv REAL,
    voltage_lv REAL,
    cooling TEXT DEFAULT '',
    quantity INTEGER,
    gost_found INTEGER DEFAULT 0,
    gost_total INTEGER DEFAULT 0,
    gost_missing TEXT DEFAULT '',
    notes TEXT DEFAULT '',
    manual INTEGER DEFAULT 0,        -- 1: правка вручную, пересканирование не затирает
    files_count INTEGER DEFAULT 0,
    scanned_at TEXT
);
CREATE TABLE IF NOT EXISTS files(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER REFERENCES projects(id) ON DELETE CASCADE,
    path TEXT UNIQUE,
    size INTEGER,
    mtime TEXT,
    is_spec INTEGER DEFAULT 0,
    designation TEXT DEFAULT ''
);
CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE IF NOT EXISTS type_rules(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    category TEXT NOT NULL, pattern TEXT NOT NULL, priority INTEGER DEFAULT 100);
CREATE TABLE IF NOT EXISTS gost_checks(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    title TEXT NOT NULL, pattern TEXT NOT NULL, required INTEGER DEFAULT 1);
CREATE INDEX IF NOT EXISTS ix_pr_year ON projects(year);
CREATE INDEX IF NOT EXISTS ix_pr_cat ON projects(category);
"""

PROJECT_COLS = ["name", "path", "year", "category", "designation", "power_kva",
                "voltage_hv", "voltage_lv", "cooling", "quantity", "gost_found",
                "gost_total", "gost_missing", "notes", "manual", "files_count", "scanned_at"]


class DB:
    def __init__(self, path):
        self.path = path
        self.conn = sqlite3.connect(path)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys=ON")
        self.conn.executescript(SCHEMA)
        self._seed()

    def _seed(self):
        c = self.conn
        for k, v in DEFAULT_SETTINGS.items():
            c.execute("INSERT OR IGNORE INTO settings(key,value) VALUES(?,?)", (k, v))
        if not c.execute("SELECT 1 FROM type_rules LIMIT 1").fetchone():
            c.executemany("INSERT INTO type_rules(category,pattern,priority) VALUES(?,?,?)",
                          DEFAULT_TYPE_RULES)
        if not c.execute("SELECT 1 FROM gost_checks LIMIT 1").fetchone():
            c.executemany("INSERT INTO gost_checks(title,pattern,required) VALUES(?,?,?)",
                          DEFAULT_GOST_CHECKS)
        c.commit()

    def close(self):
        self.conn.commit()
        self.conn.close()

    # --- настройки ---
    def get(self, key, default=""):
        r = self.conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
        return r["value"] if r else default

    def set(self, key, value):
        self.conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)", (key, str(value)))
        self.conn.commit()

    def all_settings(self):
        return [(r["key"], r["value"]) for r in self.conn.execute("SELECT * FROM settings ORDER BY key")]

    # --- правила ---
    def type_rules(self):
        return self.conn.execute("SELECT * FROM type_rules ORDER BY priority,id").fetchall()

    def gost_checks(self):
        return self.conn.execute("SELECT * FROM gost_checks ORDER BY id").fetchall()

    def save_rule(self, table, rid, **vals):
        cols = list(vals)
        if rid:
            self.conn.execute("UPDATE %s SET %s WHERE id=?" % (table, ",".join(c + "=?" for c in cols)),
                              [vals[c] for c in cols] + [rid])
        else:
            self.conn.execute("INSERT INTO %s(%s) VALUES(%s)" % (table, ",".join(cols), ",".join("?" * len(cols))),
                              [vals[c] for c in cols])
        self.conn.commit()

    def delete_row(self, table, rid):
        self.conn.execute("DELETE FROM %s WHERE id=?" % table, (rid,))
        self.conn.commit()

    # --- проекты ---
    def projects(self, year=None, category=None, text=None):
        q, a = "SELECT * FROM projects WHERE 1=1", []
        if year:
            q += " AND year=?"; a.append(int(year))
        if category:
            q += " AND category=?"; a.append(category)
        if text:
            q += " AND (name LIKE ? OR designation LIKE ? OR notes LIKE ? OR path LIKE ?)"
            a += ["%" + text + "%"] * 4
        q += " ORDER BY year DESC, category, name"
        return self.conn.execute(q, a).fetchall()

    def project(self, pid):
        return self.conn.execute("SELECT * FROM projects WHERE id=?", (pid,)).fetchone()

    def project_files(self, pid):
        return self.conn.execute("SELECT * FROM files WHERE project_id=? ORDER BY is_spec DESC, path", (pid,)).fetchall()

    def upsert_project(self, data):
        """Для проекта с manual=1 пересканирование обновляет только служебные поля."""
        row = self.conn.execute("SELECT id, manual FROM projects WHERE path=?", (data["path"],)).fetchone()
        if row and row["manual"]:
            self.conn.execute("UPDATE projects SET files_count=?, scanned_at=? WHERE id=?",
                              (data["files_count"], data["scanned_at"], row["id"]))
            return row["id"]
        if row:
            sets = [c for c in PROJECT_COLS if c not in ("path", "notes", "manual")]
            self.conn.execute("UPDATE projects SET %s WHERE id=?" % ",".join(c + "=?" for c in sets),
                              [data.get(c) for c in sets] + [row["id"]])
            return row["id"]
        cols = [c for c in PROJECT_COLS if c in data]
        cur = self.conn.execute("INSERT INTO projects(%s) VALUES(%s)" % (",".join(cols), ",".join("?" * len(cols))),
                                [data[c] for c in cols])
        return cur.lastrowid

    def replace_files(self, pid, files):
        self.conn.execute("DELETE FROM files WHERE project_id=?", (pid,))
        self.conn.executemany(
            "INSERT OR REPLACE INTO files(project_id,path,size,mtime,is_spec,designation) VALUES(?,?,?,?,?,?)",
            [(pid, f["path"], f["size"], f["mtime"], f["is_spec"], f["designation"]) for f in files])

    def update_project(self, pid, **vals):
        vals["manual"] = 1
        cols = list(vals)
        self.conn.execute("UPDATE projects SET %s WHERE id=?" % ",".join(c + "=?" for c in cols),
                          [vals[c] for c in cols] + [pid])
        self.conn.commit()

    def add_project(self, **vals):
        vals["manual"] = 1
        vals.setdefault("path", "manual:" + vals.get("name", "") + ":" + str(id(vals)))
        cols = list(vals)
        self.conn.execute("INSERT INTO projects(%s) VALUES(%s)" % (",".join(cols), ",".join("?" * len(cols))),
                          [vals[c] for c in cols])
        self.conn.commit()

    def delete_project(self, pid):
        self.conn.execute("DELETE FROM projects WHERE id=?", (pid,))
        self.conn.commit()

    def remove_missing(self, seen_paths):
        """Удалить автоматические проекты, которых больше нет на диске."""
        for r in self.conn.execute("SELECT id,path FROM projects WHERE manual=0").fetchall():
            if r["path"] not in seen_paths:
                self.conn.execute("DELETE FROM projects WHERE id=?", (r["id"],))
        self.conn.commit()

    def categories(self):
        return [r[0] for r in self.conn.execute("SELECT DISTINCT category FROM projects ORDER BY 1")]

    def years(self):
        return [r[0] for r in self.conn.execute(
            "SELECT DISTINCT year FROM projects WHERE year IS NOT NULL ORDER BY 1 DESC")]

    def stats(self):
        """Категория x год -> количество."""
        return self.conn.execute(
            "SELECT category, year, COUNT(*) n FROM projects GROUP BY category, year ORDER BY year DESC, category").fetchall()
