# -*- coding: utf-8 -*-
"""Сканирование папки: поиск технических заданий, извлечение текста и параметров трансформаторов.
Только стандартная библиотека (docx/xlsx читаются как zip+xml); pdf — при наличии pypdf/PyPDF2."""
import os
import re
import zipfile
import datetime
import xml.etree.ElementTree as ET

DESIG_RE = re.compile(
    r"\b(АТДЦТН|АТДЦН|АОДЦТН|АОДЦН|ТРДЦН|ТРДНС|ТРДН[СЦ]?|ТДЦТН|ТДТН|ТМТН|ТСЗГЛ|ТСЗН|ТСЗЛ|ТСЗ|ТСЛ|ТСГЛ|ТСГ|"
    r"ТНЗ|ТНГ|ТМГФ|ТМГСУ|ТМГ\d*|ТМН|ТМЗ|ТМФ|ТМП|ТСП|ТДП|ТЦП|ТДНС|ТДН|ТДЦ|ТМ|ТД|ТЦ|ТС)"
    r"[-\s]?(\d{2,6})(?:\s?/\s?(\d{1,3}(?:[.,]\d+)?))?(?:\s?[-/]\s?(\d{1,3}(?:[.,]\d+)?))?"
    r"(?:\s?/\s?(\d{1,3}(?:[.,]\d+)?))?([А-ЯA-Zа-яa-z0-9\-]{0,6})?")
YEAR_RE = re.compile(r"(?<!\d)(19[89]\d|20\d\d)(?!\d)")
POWER_RE = re.compile(r"(\d[\d\s]*(?:[.,]\d+)?)\s*(кВ\s?А|kVA|МВ\s?А|MVA)", re.I)
VOLT_RE = re.compile(r"(\d{1,3}(?:[.,]\d+)?)\s*/\s*(\d{1,3}(?:[.,]\d+)?)(?:\s*/\s*(\d{1,3}(?:[.,]\d+)?))?\s*кВ", re.I)
QTY_RE = re.compile(r"(?:количеств\w*|кол-во|в количестве)\s*[:\-–]?\s*(\d{1,4})\s*(?:шт|ед|компл)?", re.I)
COOL_RE = re.compile(r"\b(ONAN|ONAF|ODAF|ДЦ|НДЦ|Ц|Д|М|С)\b")


def _num(s):
    return float(s.replace(",", ".").replace(" ", ""))


# ---------- извлечение текста ----------
def _xml_text(data, tags=("t",)):
    out = []
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return ""
    for el in root.iter():
        name = el.tag.rsplit("}", 1)[-1]
        if name in tags and el.text:
            out.append(el.text)
        elif name in ("p", "tr", "row"):
            out.append("\n")
    return " ".join(out)


def read_docx(path):
    with zipfile.ZipFile(path) as z:
        return _xml_text(z.read("word/document.xml"))


def read_xlsx(path):
    with zipfile.ZipFile(path) as z:
        parts = []
        if "xl/sharedStrings.xml" in z.namelist():
            parts.append(_xml_text(z.read("xl/sharedStrings.xml")))
        for n in z.namelist():
            if n.startswith("xl/worksheets/") and n.endswith(".xml"):
                parts.append(_xml_text(z.read(n), tags=("v", "t")))
        return "\n".join(parts)


def read_pdf(path):
    try:
        try:
            from pypdf import PdfReader
        except ImportError:
            from PyPDF2 import PdfReader
        r = PdfReader(path)
        return "\n".join((p.extract_text() or "") for p in r.pages[:40])
    except Exception:
        return ""


def read_plain(path):
    with open(path, "rb") as f:
        raw = f.read(3000000)
    for enc in ("utf-8-sig", "cp1251", "cp866"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("cp1251", "ignore")


def read_binary_office(path):
    """Старые .doc/.xls: грубое извлечение строк (utf-16le и cp1251)."""
    with open(path, "rb") as f:
        raw = f.read(4000000)
    u = raw.decode("utf-16le", "ignore")
    chunks = re.findall(r"[0-9A-Za-zА-Яа-яЁё\s.,;:%/()\-–]{6,}", u)
    c = raw.decode("cp1251", "ignore")
    chunks += re.findall(r"[0-9A-Za-zА-Яа-яЁё\s.,;:%/()\-–]{6,}", c)
    return "\n".join(chunks)


def extract_text(path):
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".docx":
            return read_docx(path)
        if ext in (".xlsx", ".xlsm"):
            return read_xlsx(path)
        if ext == ".pdf":
            return read_pdf(path)
        if ext in (".doc", ".xls"):
            return read_binary_office(path)
        return read_plain(path)
    except Exception:
        return ""


# ---------- разбор параметров ----------
def parse_text(text):
    """Возвращает dict: designations, power_kva, voltage_hv, voltage_lv, quantity, cooling."""
    info = {"designations": [], "power_kva": None, "voltage_hv": None, "voltage_lv": None,
            "quantity": None, "cooling": ""}
    seen = []
    for m in DESIG_RE.finditer(text):
        d = m.group(0).strip()
        if d not in seen:
            seen.append(d)
    info["designations"] = seen[:10]
    if seen:
        m = DESIG_RE.search(text)
        pw = _num(m.group(2))
        info["power_kva"] = pw if pw > 5 else None
        volt = m.group(3)
        if volt:
            info["voltage_hv"] = _num(volt)
    pm = POWER_RE.search(text)
    if pm:
        try:
            v = _num(pm.group(1))
            info["power_kva"] = v * 1000 if pm.group(2).lower().replace(" ", "") in ("мва", "mva") else v
        except ValueError:
            pass
    vm = VOLT_RE.search(text)
    if vm:
        info["voltage_hv"] = _num(vm.group(1))
        info["voltage_lv"] = _num(vm.group(2))
    qm = QTY_RE.search(text)
    if qm:
        info["quantity"] = int(qm.group(1))
    cm = COOL_RE.search(text)
    if cm:
        info["cooling"] = cm.group(1)
    return info


def classify(text, designations, rules):
    """rules: строки type_rules (category, pattern, priority). Сначала по обозначению, затем по тексту."""
    cands = list(designations) + [text[:20000]]
    for src in cands:
        for r in rules:
            try:
                if re.search(r["pattern"], src, re.I if src is cands[-1] else 0):
                    return r["category"]
            except re.error:
                continue
    return "Не определён"


def check_gost(text, checks):
    found, missing = 0, []
    for c in checks:
        try:
            ok = re.search(c["pattern"], text, re.I) is not None
        except re.error:
            ok = False
        if ok:
            found += 1
        else:
            missing.append(c["title"])
    return found, missing


def guess_year(path, mtime, lo, hi):
    """Год: из имён папок/файла (самое глубокое совпадение), иначе — по дате изменения."""
    parts = path.replace("\\", "/").split("/")
    for p in reversed(parts):
        ys = [int(y) for y in YEAR_RE.findall(p) if lo <= int(y) <= hi]
        if ys:
            return ys[0]
    return mtime.year


# ---------- сканирование ----------
class Scanner:
    def __init__(self, db, log=None, progress=None):
        self.db = db
        self.log = log or (lambda s: None)
        self.progress = progress or (lambda n: None)
        self.cancel = False

    def run(self, root):
        db = self.db
        depth = max(1, int(db.get("project_depth", "1") or 1))
        exts = tuple(e.strip().lower() for e in db.get("extensions").split(",") if e.strip())
        kws = [k.strip().lower() for k in db.get("spec_keywords").split(",") if k.strip()]
        max_size = float(db.get("max_file_mb", "30") or 30) * 1024 * 1024
        lo, hi = int(db.get("min_year", "1990")), int(db.get("max_year", "2100"))
        rules, checks = db.type_rules(), db.gost_checks()
        root = os.path.normpath(root)
        if not os.path.isdir(root):
            raise IOError("Папка недоступна: " + root)

        projects = {}   # путь проекта -> список файлов
        n_files = 0
        for dirpath, dirnames, filenames in os.walk(root, onerror=lambda e: self.log("Нет доступа: %s" % e)):
            if self.cancel:
                self.log("Остановлено пользователем")
                return 0
            rel = os.path.relpath(dirpath, root)
            level = 0 if rel == "." else rel.count(os.sep) + 1
            dirnames.sort()
            for fn in filenames:
                if not fn.lower().endswith(exts) or fn.startswith("~$"):
                    continue
                if level == 0:
                    ppath = root          # файлы прямо в корне — отдельный «проект» корня
                else:
                    ppath = os.path.join(root, *rel.split(os.sep)[:depth])
                full = os.path.join(dirpath, fn)
                projects.setdefault(ppath, []).append(full)
                n_files += 1
                if n_files % 50 == 0:
                    self.progress(n_files)
        self.log("Найдено файлов: %d, проектов: %d" % (n_files, len(projects)))

        seen = set()
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        for i, (ppath, files) in enumerate(sorted(projects.items()), 1):
            if self.cancel:
                self.log("Остановлено пользователем")
                break
            fobjs, texts, desigs = [], [], []
            newest = None
            for f in files:
                try:
                    st = os.stat(f)
                except OSError:
                    continue
                mt = datetime.datetime.fromtimestamp(st.st_mtime)
                newest = mt if newest is None or mt > newest else newest
                name = os.path.basename(f).lower()
                is_spec = any(k in name for k in kws) or any(k in os.path.dirname(f).lower() for k in kws[:4])
                d = ""
                if is_spec and st.st_size <= max_size:
                    t = extract_text(f)
                    if t:
                        texts.append(t)
                        info = parse_text(t)
                        d = ", ".join(info["designations"][:3])
                        desigs += info["designations"]
                fobjs.append({"path": f, "size": st.st_size, "mtime": mt.strftime("%Y-%m-%d"),
                              "is_spec": 1 if is_spec else 0, "designation": d})
            if not texts and not fobjs:
                continue
            text = "\n".join(texts)
            # проект без найденных ТЗ — берём «обозначение» из имён файлов/папок
            if not text:
                text = " ".join(os.path.basename(f["path"]) for f in fobjs) + " " + os.path.basename(ppath)
            info = parse_text(text)
            cat = classify(text, info["designations"], rules)
            gf, gm = check_gost(text, checks) if texts else (0, [c["title"] for c in checks])
            data = {
                "name": os.path.basename(ppath) or ppath, "path": ppath,
                "year": guess_year(ppath, newest or datetime.datetime.now(), lo, hi),
                "category": cat, "designation": ", ".join(info["designations"][:3]),
                "power_kva": info["power_kva"], "voltage_hv": info["voltage_hv"],
                "voltage_lv": info["voltage_lv"], "cooling": info["cooling"],
                "quantity": info["quantity"], "gost_found": gf, "gost_total": len(checks),
                "gost_missing": "; ".join(gm), "manual": 0, "files_count": len(fobjs), "scanned_at": now,
            }
            pid = db.upsert_project(data)
            db.replace_files(pid, fobjs)
            seen.add(ppath)
            self.log("[%d/%d] %s → %s, %s" % (i, len(projects), data["name"], cat, data["year"]))
            self.progress(n_files)
        if not self.cancel:
            db.remove_missing(seen)
        db.conn.commit()
        return len(seen)
