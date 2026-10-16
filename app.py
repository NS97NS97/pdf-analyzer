# -*- coding: utf-8 -*-
"""
Анализатор PDF-документов в папке.

Запуск как скрипта:   python app.py
Запуск как программы: АнализаторPDF.exe
Затем открывается браузер:  http://127.0.0.1:5000
"""

import os
import re
import sys
import glob
import socket
import tempfile
import threading
import traceback
import webbrowser
from urllib.parse import quote
from flask import Flask, render_template, request, jsonify, send_file
from pypdf import PdfReader
from openpyxl import Workbook


# ---------------------------------------------------------------------------
# Пути. Программа работает и как обычный скрипт, и как собранный .exe:
# PyInstaller упаковывает templates и static внутрь файла и распаковывает их
# во временную папку (sys._MEIPASS).
# ---------------------------------------------------------------------------

def resource_dir():
    """Папка с шаблонами и стилями (templates, static)."""
    if getattr(sys, "frozen", False):                 # собранный .exe
        return getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _writable(folder):
    """Можно ли записывать файлы в папку."""
    try:
        test_file = os.path.join(folder, ".write_test")
        with open(test_file, "w", encoding="utf-8") as handle:
            handle.write("")
        os.remove(test_file)
        return True
    except OSError:
        return False


def data_dir():
    """Папка для файла отчёта: рядом с программой, иначе во временной папке."""
    if getattr(sys, "frozen", False):                 # собранный .exe
        near_program = os.path.dirname(sys.executable)
    else:
        near_program = os.path.dirname(os.path.abspath(__file__))

    if _writable(near_program):
        return near_program
    return tempfile.gettempdir()


BASE_DIR = resource_dir()
DATA_DIR = data_dir()

app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, "templates"),
    static_folder=os.path.join(BASE_DIR, "static"),
)


@app.errorhandler(500)
def handle_500(error):
    """Логируем настоящую ошибку и показываем её причину пользователю."""
    original = getattr(error, "original_exception", None) or error
    traceback.print_exc()
    return jsonify({
        "error": "Внутренняя ошибка: " + original.__class__.__name__ + ": " + str(original),
    }), 500

# ---------------------------------------------------------------------------
# Файл отчёта. База данных не нужна.
# ---------------------------------------------------------------------------
REPORT_FILE = os.path.join(DATA_DIR, "report.xlsx")

# Папка последнего анализа — открывать разрешаем только файлы из неё
LAST_FOLDER = ""


# =========================== Вспомогательные функции ========================

def clean_text(text):
    """Убирает служебные символы и лишние пробелы (Excel ищет чистый текст)."""
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", " ", str(text))
    return re.sub(r"\s+", " ", text).strip()


def real_letters(text):
    """Считает обычные буквы (латиница и кириллица)."""
    return len(re.findall(r"[A-Za-zА-Яа-яЁё]", text))


def fix_mojibake(text):
    """Часть PDF отдаёт кириллицу в кодировке latin-1 — пробуем исправить."""
    if not text or real_letters(text) / max(len(text), 1) >= 0.5:
        return text
    try:
        fixed = text.encode("latin-1", "ignore").decode("cp1251", "ignore")
    except Exception:
        return text
    return fixed if real_letters(fixed) > real_letters(text) else text


def is_garbage_text(text):
    """Текст из «сломанных» шрифтов: почти нет обычных букв и знаков."""
    if len(text) < 50:
        return False
    good = len(re.findall(r"[A-Za-zА-Яа-яЁё0-9\s.,;:!?()«»\"'–—-]", text))
    return good / len(text) < 0.75


# Слова-маркеры «это не название»: издательства, институты, адреса и т. п.
JUNK_STEMS = [
    "untitled", ".indd", ".qxp", "layout", "pages_", "http", "www.", "lovepdf",
    "repack", "сформирована", "электроннаябиблиотека", "электроннаякопия",
    "содержаниеэтойстраницы", "министерств", "государствен", "институт",
    "академ", "университет", "эрмитаж", "музей", "исследовательск",
    "историииархеологии", "russianacademy", "instituteof", "academyofsciences",
    "ministryof", "nationalresearch", "stateresearch", "hermitage", "museum",
    "researchcenter", "university", "moscow", "москва", "saintpetersburg",
    "санктпетербург", "смоленск", "великийновгород", "©",
]

JUNK_RE = re.compile(r"\b(удк|ббк|isbn|issn|doi|ран|ниц|оглавление|содержание)\b",
                     re.IGNORECASE)

# Фразы, после которых название уже началось, — обрываем поиск
STOP_WORDS = [
    "ключевые слова", "аннотация", "abstract", "annotation", "summary",
    "в книге", "в статье", "в сборнике", "книга посвящена", "рецензент",
    "редактор", "редакцион", "переводчик", "ответственный", "составител",
]


def looks_like_junk(text):
    """Проверяет, похожа ли строка на технический мусор, а не на название."""
    if not text or len(text.strip()) < 6:
        return True
    if real_letters(text) < 4:
        return True
    flat = text.lower().replace(" ", "")
    if any(stem in flat for stem in JUNK_STEMS):
        return True
    if JUNK_RE.search(text):
        return True
    if re.match(r"^[А-ЯA-Z]{1,2}\d{1,3}\s", text):     # издательский код "Н72"
        return True
    if text[0].islower():                              # название с большой буквы
        return True
    if len(re.findall(r"\d", text)) / len(text) > 0.4:
        return True
    if re.search(r"\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4}", text):
        return True
    if re.search(r"\d+\s*/\s*\d{4}", text):
        return True
    if " " not in text and len(text) > 25:
        return True
    # Разрядка капсом: "И С С ЛЕ Д О В А Н И Я"
    words = text.split()
    if len(words) >= 4 and text.isupper():
        if sum(1 for w in words if len(w) <= 2) / len(words) >= 0.8:
            return True
    return False


def is_author_line(text):
    """Строка похожа на перечисление авторов, а не на название."""
    if re.search(r"([А-ЯЁA-Z]\.\s*){2,}", text):
        return True
    if len(re.findall(r"\b[A-ZА-Я][a-zа-яё]+\s+[A-ZА-Я][a-zа-яё]*\.", text)) >= 2:
        return True
    # Строка целиком из имён с инициалами: "Natalya V. Bykovskaya"
    if re.search(r"[A-ZА-ЯЁ]\.", text):
        tokens = text.replace(",", " ").replace(";", " ").split()
        if 1 <= len(tokens) <= 5:
            name = r"[A-ZА-ЯЁ][a-zа-яё]+\.?|[A-ZА-ЯЁ]\.|[A-ZА-ЯЁ]{2,}\.?"
            if all(re.fullmatch(name, t) for t in tokens):
                return True
    return False


def looks_like_plain_text(text):
    """Строка может продолжать название: текст, буквы, немного цифр."""
    if real_letters(text) < 3:
        return False
    if len(re.findall(r"\d", text)) / len(text) > 0.3:      # почти только цифры
        return False
    if re.search(r"\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4}", text):   # дата
        return False
    flat = text.lower().replace(" ", "")
    if any(stem in flat for stem in JUNK_STEMS) or JUNK_RE.search(text):
        return False
    allowed_punct = set(" -–—.,:;«»()'[]/")
    return all(ch.isalpha() or ch.isdigit() or ch in allowed_punct for ch in text)


def is_stop_line(text):
    """Строка, после которой название уже началось (ключевые слова, авторы)."""
    low = text.lower()
    return any(w in low for w in STOP_WORDS)


def has_ocr_notice(text):
    """Страница с типовой пометкой сканера — пропускаем её."""
    flat = text.lower().replace(" ", "")
    return "добавленоприocr" in flat or "электроннаякопия" in flat


def caps_title_part(line):
    """Из строки «Автор И. О. ЗАГОЛОВОК КАПСОМ» вырезает часть ЗАГОЛОВКА."""
    match = re.search(r"[A-ZА-ЯЁ][A-ZА-ЯЁ0-9\s\-–:,.«»]{19,}", line)
    if match:
        part = clean_text(match.group(0))
        if len(part.split()) >= 3:
            return part
    return line


def build_title_block(lines):
    """Собирает название из нескольких коротких строк в начале страницы."""
    parts = []
    for raw in lines[:15]:
        line = caps_title_part(raw)

        # Строки, на которых название точно заканчивается
        if len(line) > 70 or is_stop_line(line) or is_author_line(line):
            if parts:
                break
            continue

        if not parts:
            # Ищем первую подходящую строку названия
            if looks_like_junk(line):
                continue
            parts.append(line)
        else:
            # Продолжаем, только если строка похожа на продолжение названия
            if looks_like_plain_text(line):
                parts.append(line)
            else:
                break

        if len(parts) >= 5 or len(" ".join(parts)) >= 150:
            break
    return clean_text(" ".join(parts))


def extract_title(page_texts, meta_title):
    """Ищет название статьи в тексте первых страниц PDF."""
    for text in page_texts[:8]:
        text = fix_mojibake(text)
        if has_ocr_notice(text) or is_garbage_text(text):
            continue
        lines = [clean_text(l) for l in text.splitlines() if l.strip()]
        title = build_title_block(lines)
        if title:
            return title
    if meta_title and not looks_like_junk(meta_title):
        return meta_title
    return ""


def extract_author(text, meta_author):
    """Ищет автора: в свойствах PDF, затем шаблонами «Иванов И. И.»."""
    if meta_author and not looks_like_junk(meta_author):
        return meta_author
    patterns = [
        r"[А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.",                        # Иванов И. И.
        r"[А-ЯЁ]\.\s*[А-ЯЁ]\.\s*[А-ЯЁ][а-яё-]+",                       # И. И. Иванов
        r"[A-Z][a-z][a-z-]+\s+[A-Z]\.(?:\s*[A-Z]\.)?\s+[A-Z][a-z][a-z-]+",  # Ivanov I. I.
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            author = clean_text(match.group(0)).strip(" ,;")
            if len(author) <= 60:
                return author
    return ""


def extract_year(text):
    """Первое четырёхзначное число 19xx/20xx — считаем годом выхода."""
    match = re.search(r"\b(19|20)\d{2}\b", text)
    return match.group(0) if match else ""


# ============================== Работа с PDF ================================

def _unreadable(file_path):
    """Файл, который не удалось прочитать (повреждён и т. п.)."""
    return {
        "title": "",
        "author": "",
        "year": "",
        "path": file_path,
        "unreadable": True,
    }


def read_pdf_info(file_path):
    """Читает один PDF и возвращает словарь: название, автор, год, путь.

    Любая ошибка разбора не должна ронять сервис: файл просто помечается
    как нечитаемый, а анализ продолжается.
    """
    try:
        reader = PdfReader(file_path)
    except Exception:
        return _unreadable(file_path)

    try:
        # Файл зашифрован паролем: пробуем открыть пустым паролем.
        # Если не получилось — помечаем как защищённый и не читаем текст.
        try:
            if reader.is_encrypted and not reader.decrypt(""):
                return {
                    "title": "",
                    "author": "",
                    "year": "",
                    "path": file_path,
                    "encrypted": True,
                }
        except Exception:
            return {
                "title": "",
                "author": "",
                "year": "",
                "path": file_path,
                "encrypted": True,
            }

        # 1. Свойства PDF
        meta_title = meta_author = ""
        try:
            meta = reader.metadata
            if meta:
                meta_title = clean_text(meta.title or "")
                meta_author = clean_text(meta.author or "")
        except Exception:
            pass
        if meta_title.lower() in ("untitled", "title", "document"):
            meta_title = ""

        # 2. Текст первых страниц (8 страниц: обложка + начало статьи)
        page_texts = []
        try:
            for index, page in enumerate(reader.pages):
                if index >= 8:
                    break
                try:
                    page_texts.append(page.extract_text() or "")
                except Exception:
                    page_texts.append("")
        except Exception:
            # Не удалось получить страницы — считаем по тому, что успели.
            traceback.print_exc()

        full_text = fix_mojibake("\n".join(page_texts[:4]))

        return {
            "title": extract_title(page_texts, meta_title) or "—",
            "author": extract_author(full_text, meta_author) or "—",
            "year": extract_year(full_text) or "—",
            "path": file_path,
        }
    except Exception:
        # Любая непредвиденная ошибка — помечаем файл, но не роняем анализ.
        traceback.print_exc()
        return _unreadable(file_path)


def analyze_folder(folder_path):
    """Проходит по папке и всем вложенным папкам, собирает информацию о PDF."""
    results = []
    for root, _dirs, files in os.walk(folder_path):
        for name in sorted(files):
            if not name.lower().endswith(".pdf"):
                continue
            full_path = os.path.join(root, name)
            try:
                info = read_pdf_info(full_path)
            except Exception:
                # Ни один файл не должен ронять весь анализ.
                traceback.print_exc()
                info = _unreadable(full_path)
            if not info:
                info = _unreadable(full_path)
            if info.get("encrypted"):
                info["title"] = "защищен паролем"
            elif info.get("unreadable"):
                info["title"] = "не удалось прочитать"
            results.append(info)
    return results


def save_excel(results, file_path):
    """Сохраняет результаты в Excel-файл с таблицей."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Анализ документов"

    headers = ["Название", "Автор", "Год", "Путь к статье"]
    ws.append(headers)

    # Красивое оформление шапки
    from openpyxl.styles import Font, PatternFill, Alignment
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4A6FA5")
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    for item in results:
        ws.append([
            clean_text(item["title"]),
            clean_text(item["author"]),
            clean_text(item["year"]),
            clean_text(item["path"]),
        ])

    # Ширина колонок
    ws.column_dimensions["A"].width = 50
    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 70

    wb.save(file_path)


def save_report(results):
    """Сохраняет отчёт в report.xlsx.

    Если файл открыт в Excel (Windows не даёт перезаписать), сохраняем
    копию с датой в имени и удаляем такие копии при следующем анализе.
    Возвращает имя сохранённого файла.
    """
    global REPORT_FILE

    canonical = os.path.join(DATA_DIR, "report.xlsx")
    try:
        save_excel(results, canonical)
        REPORT_FILE = canonical
        return os.path.basename(canonical)
    except PermissionError:
        pass

    from datetime import datetime
    mask = os.path.join(DATA_DIR, "report_[0-9]*.xlsx")
    for old in glob.glob(mask):                     # чистим прошлые копии
        try:
            os.remove(old)
        except OSError:
            pass

    REPORT_FILE = os.path.join(
        DATA_DIR, "report_" + datetime.now().strftime("%Y%m%d_%H%M%S") + ".xlsx")
    save_excel(results, REPORT_FILE)
    return os.path.basename(REPORT_FILE)


# =============================== Страницы сайта =============================

@app.route("/")
def index():
    return render_template("index.html")


def has_desktop():
    """Есть ли на этой машине рабочий стол (можно ли показать окно выбора)."""
    if sys.platform == "win32" or sys.platform == "darwin":
        return True
    return bool(os.environ.get("DISPLAY") or os.environ.get("WAYLAND_DISPLAY"))


@app.route("/env")
def env_info():
    """Сообщает странице, запущен сайт на своём компьютере или на хостинге."""
    return jsonify({
        "desktop": has_desktop(),
        "separator": os.sep,
        "example": "C:\\Папка с PDF" if os.sep == "\\" else "/home/имя/pdfs",
    })


@app.route("/choose_folder")
def choose_folder():
    """Открывает системное окно выбора папки на компьютере.

    Окно можно показать только там, где запущен сам сайт. Если сервер
    работает без рабочего стола (например, хостинг), сообщаем об этом —
    тогда путь вводят вручную.
    """
    if not has_desktop():
        return jsonify({
            "fallback": True,
            "error": "Сервер работает без рабочего стола (хостинг), "
                     "поэтому открыть окно выбора папки нельзя. "
                     "Введите путь к папке вручную, например: /home/имя/PDF",
        })

    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)      # окно поверх остальных
        folder = filedialog.askdirectory(title="Выберите папку с документами")
        root.destroy()
        return jsonify({"path": folder or ""})
    except Exception as exc:
        return jsonify({
            "fallback": True,
            "error": "Не удалось открыть окно выбора папки (" + str(exc) +
                     "). Введите путь вручную.",
        })


@app.route("/analyze", methods=["POST"])
def analyze():
    """Принимает путь к папке, анализирует PDF и отдаёт результат."""
    global LAST_FOLDER

    folder = (request.json or {}).get("path", "").strip().strip('"')

    if not folder:
        return jsonify({"error": "Путь к папке не указан"}), 400
    if not os.path.isdir(folder):
        return jsonify({"error": "Папка не найдена: " + folder}), 400

    try:
        results = analyze_folder(folder)
        report_name = save_report(results)
    except Exception as exc:
        # Показываем настоящую причину, а не общее «500».
        traceback.print_exc()
        return jsonify({
            "error": "Ошибка при обработке: " + exc.__class__.__name__ + ": " + str(exc),
        }), 500
    LAST_FOLDER = os.path.abspath(folder)

    no_title = sum(1 for item in results if item["title"] == "—")
    return jsonify({
        "count": len(results),
        "items": results,
        "no_title": no_title,
        "report": report_name,
        "report_locked": report_name != "report.xlsx",
    })


@app.route("/open_file")
def open_file():
    """Открывает PDF-файл из проанализированной папки на компьютере."""
    file_path = os.path.abspath(request.args.get("path", ""))

    if not file_path or not os.path.isfile(file_path):
        return jsonify({"error": "Файл не найден: " + file_path}), 404
    if not file_path.lower().endswith(".pdf"):
        return jsonify({"error": "Можно открывать только PDF-файлы"}), 400

    # Открываем только файлы из папки, которую анализировали
    if LAST_FOLDER:
        try:
            inside = os.path.commonpath([file_path, LAST_FOLDER]) == LAST_FOLDER
        except ValueError:                  # разные диски
            inside = False
        if not inside:
            return jsonify({"error": "Файл вне проанализированной папки"}), 403

    if hasattr(os, "startfile"):            # Windows
        os.startfile(file_path)
        return jsonify({"opened": True, "file": file_path})

    # Другая система (сервер без рабочего стола) — отдаём файл на скачивание
    return jsonify({
        "opened": False,
        "file": file_path,
        "download": "/download_file?path=" + quote(file_path),
        "note": "Сервер не может открыть файл на вашем компьютере — файл будет скачан.",
    })


@app.route("/download_file")
def download_file():
    """Отдаёт PDF-файл из проанализированной папки для скачивания."""
    file_path = os.path.abspath(request.args.get("path", ""))

    if not file_path or not os.path.isfile(file_path):
        return jsonify({"error": "Файл не найден: " + file_path}), 404
    if not file_path.lower().endswith(".pdf"):
        return jsonify({"error": "Можно скачивать только PDF-файлы"}), 400

    if LAST_FOLDER:
        try:
            inside = os.path.commonpath([file_path, LAST_FOLDER]) == LAST_FOLDER
        except ValueError:
            inside = False
        if not inside:
            return jsonify({"error": "Файл вне проанализированной папки"}), 403

    return send_file(file_path, as_attachment=True)


@app.route("/report.xlsx")
def report():
    """Скачивает Excel-файл с результатами."""
    if not os.path.exists(REPORT_FILE):
        return jsonify({"error": "Сначала выполните анализ папки"}), 400
    return send_file(REPORT_FILE, as_attachment=True)


@app.route("/open_xls")
def open_xls():
    """Открывает Excel-файл прямо на компьютере (Windows)."""
    if not os.path.exists(REPORT_FILE):
        return jsonify({"error": "Сначала выполните анализ папки"}), 400

    if hasattr(os, "startfile"):                 # Windows
        os.startfile(os.path.abspath(REPORT_FILE))
        return jsonify({"opened": True})

    # Другая система (сервер без рабочего стола) — файл просто скачивается
    return jsonify({
        "opened": False,
        "download": "/report.xlsx",
        "note": "Сервер не может открыть Excel на вашем компьютере — файл будет скачан.",
    })


def find_free_port(start=5000, attempts=50):
    """Ищет свободный порт, начиная с указанного."""
    for port in range(start, start + attempts):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            try:
                sock.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    return start


def main():
    # Порт можно задать при запуске:  app.py 5001  /  АнализаторPDF.exe 5001
    port = None
    for arg in sys.argv[1:]:
        if arg.isdigit():
            port = int(arg)
            break
    if port is None:
        port = find_free_port(5000)

    url = f"http://127.0.0.1:{port}"

    print("=" * 62)
    print("  Анализатор PDF-документов")
    print("=" * 62)
    print(f"  Сайт открывается в браузере:  {url}")
    print(f"  Отчёты сохраняются в:         {DATA_DIR}")
    print("  Пока вы работаете с сайтом — это окно не закрывайте.")
    print("  Чтобы остановить программу — просто закройте это окно.")
    print("=" * 62)

    # Открываем браузер автоматически (чуть позже, когда сервер поднимется)
    if os.environ.get("PDFANALYZER_NO_BROWSER") != "1":
        threading.Timer(1.2, lambda: webbrowser.open(url)).start()

    app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False)


if __name__ == "__main__":
    main()
