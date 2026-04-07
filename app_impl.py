import re

import pymupdf
import streamlit as st

PT_TO_CM = 1 / 28.346

WORK_OPTIONS = [
    "Курсова з БЗВП",
    "Курсова з маркетингу",
    "Звіт з практики (Бакалавр)",
    "Звіт з практики (Магістр)",
    "Кваліфікаційна бакалаврська робота",
    "Кваліфікаційна магістерська робота",
]

TITLE_SAMPLE_MAP = {
    "Курсова з БЗВП": "Тітулка Курсова.pdf",
    "Курсова з маркетингу": "Тітулка Курсова.pdf",
    "Звіт з практики (Бакалавр)": "Тітульний Практика.pdf",
    "Звіт з практики (Магістр)": "Тітульний Практика.pdf",
    "Кваліфікаційна бакалаврська робота": "Тітульна КБР.pdf",
    "Кваліфікаційна магістерська робота": "Тітульна КБР.pdf",
}

CONTENTS_SAMPLE_FILE = "Зразок зміст.pdf"


def normalize_text(text):
    text = text.replace("\xa0", " ")
    return re.sub(r"\s+", " ", text).strip()


def normalize_for_search(text):
    return normalize_text(text).upper().replace("’", "'").replace("`", "'")


def is_bold(font_name, flags):
    return "Bold" in font_name or bool(flags & 16)


def extract_lines(page):
    lines = []
    for block in page.get_text("dict")["blocks"]:
        if "lines" not in block:
            continue
        for line in block["lines"]:
            spans = [span for span in line["spans"] if span["text"].strip()]
            if not spans:
                continue
            text = "".join(span["text"] for span in spans).strip()
            if not text:
                continue
            main_span = max(spans, key=lambda span: len(span["text"].strip()))
            lines.append(
                {
                    "text": text,
                    "normalized": normalize_for_search(text),
                    "x0": line["bbox"][0],
                    "y0": line["bbox"][1],
                    "x1": line["bbox"][2],
                    "y1": line["bbox"][3],
                    "size": round(sum(span["size"] for span in spans) / len(spans), 1),
                    "font": main_span["font"],
                    "flags": main_span.get("flags", 0),
                    "bold": is_bold(main_span["font"], main_span.get("flags", 0)),
                }
            )
    return sorted(lines, key=lambda item: (item["y0"], item["x0"]))


def build_report():
    return {
        "Титульна сторінка": [],
        "Сторінка зі змістом": [],
        "Поля сторінки (Л: 2.5 см, П: 1.0 см, В/Н: 2.0 см)": [],
        "Шрифт (Times New Roman)": [],
        "Розмір шрифту (14)": [],
        "Абзацний відступ (1.5 см)": [],
        "Міжрядковий інтервал (1.5)": [],
        "Оформлення заголовків (ВЕЛИКІ ЛІТЕРИ, по центру)": [],
        "Оформлення підрозділів (Відступ 1.5 см, без крапки в кінці)": [],
        "Підписи до рисунків (Формат: Рисунок X.X - Назва)": [],
        "Підписи до таблиць (Формат: Таблиця X.X - Назва)": [],
        "Межі та розриви таблиць (Не виходять за поля, наявність 'Продовження')": [],
        "Оформлення формул (Номер праворуч у дужках)": [],
        "Список використаних джерел (ДСТУ 8302:2015)": [],
    }


def truncate_report(report):
    for rule, errors in report.items():
        unique_errors = list(dict.fromkeys(errors))
        if len(unique_errors) > 8:
            report[rule] = unique_errors[:8] + [
                f"...та ще <b>{len(unique_errors) - 8} схожих помилок</b> у документі."
            ]
        else:
            report[rule] = unique_errors
    return report


def add_page_error(report, rule, page_number, message):
    report[rule].append(f"<b>Сторінка {page_number}</b>: {message}")


def find_best_line(lines, pattern, expected_x=None, expected_y=None):
    regex = re.compile(pattern, re.IGNORECASE)
    matches = [line for line in lines if regex.search(line["normalized"])]
    if not matches:
        return None

    def score(line):
        value = 0
        if expected_y is not None:
            value += abs(line["y0"] - expected_y) * 3
        if expected_x is not None:
            value += abs(line["x0"] - expected_x)
        return value

    return min(matches, key=score)


def validate_line(lines, rule_errors, page_number, spec):
    line = find_best_line(
        lines,
        spec["pattern"],
        expected_x=spec.get("x"),
        expected_y=spec.get("y"),
    )
    label = spec["label"]
    if line is None:
        rule_errors.append(f"<b>Сторінка {page_number}</b>: Не знайдено елемент '{label}'.")
        return False

    if "x" in spec and abs(line["x0"] - spec["x"]) > spec.get("x_tol", 24):
        rule_errors.append(
            f"<b>Сторінка {page_number}</b>: '{label}' зміщено по горизонталі "
            f"(очікувано x≈{round(spec['x'], 1)}, фактично x≈{round(line['x0'], 1)})."
        )
    if "y" in spec and abs(line["y0"] - spec["y"]) > spec.get("y_tol", 16):
        rule_errors.append(
            f"<b>Сторінка {page_number}</b>: '{label}' зміщено по вертикалі "
            f"(очікувано y≈{round(spec['y'], 1)}, фактично y≈{round(line['y0'], 1)})."
        )
    if "size" in spec and abs(line["size"] - spec["size"]) > spec.get("size_tol", 1.0):
        rule_errors.append(
            f"<b>Сторінка {page_number}</b>: '{label}' має розмір шрифту {line['size']}, "
            f"а за зразком очікується приблизно {spec['size']}."
        )
    if "bold" in spec and line["bold"] != spec["bold"]:
        expected = "жирним" if spec["bold"] else "звичайним"
        rule_errors.append(
            f"<b>Сторінка {page_number}</b>: '{label}' має бути набрано {expected}."
        )
    return True


def get_title_specs(work_type):
    if work_type in {"Курсова з БЗВП", "Курсова з маркетингу"}:
        return {
            "detection_patterns": [
                r"МІНІСТЕРСТВО ОСВІТИ І НАУКИ УКРАЇНИ",
                r"КИЇВСЬКИЙ НАЦІОНАЛЬНИЙ ЕКОНОМІЧНИЙ УНІВЕРСИТЕТ",
                r"КУРСОВА РОБОТА",
                r"ЗДОБУВАЧА ГРУПИ",
            ],
            "required_matches": 4,
            "specs": [
                {"label": "Міністерство", "pattern": r"МІНІСТЕРСТВО ОСВІТИ І НАУКИ УКРАЇНИ", "x": 165.9, "y": 57.3, "size": 14.0, "bold": True},
                {"label": "Університет", "pattern": r"КИЇВСЬКИЙ НАЦІОНАЛЬНИЙ ЕКОНОМІЧНИЙ УНІВЕРСИТЕТ", "x": 101.0, "y": 81.4, "size": 14.0, "bold": True},
                {"label": "Імені Вадима Гетьмана", "pattern": r"ІМЕНІ ВАДИМА ГЕТЬМАНА", "x": 227.4, "y": 105.6, "size": 14.0, "bold": True},
                {"label": "Факультет", "pattern": r"ФАКУЛЬТЕТ МАРКЕТИНГУ", "x": 245.0, "y": 178.0, "size": 14.0, "bold": True},
                {"label": "Кафедра", "pattern": r"КАФЕДРА МАРКЕТИНГУ ІМЕНІ А\\.Ф\\. ПАВЛЕНКА", "x": 185.3, "y": 202.2, "size": 14.0, "bold": True},
                {"label": "Назва роботи", "pattern": r"КУРСОВА РОБОТА", "x": 253.0, "y": 322.9, "size": 14.0, "bold": True},
                {"label": "Дисципліна", "pattern": r"З НАВЧАЛЬНОЇ ДИСЦИПЛІНИ", "x": 195.5, "y": 347.1, "size": 14.0, "bold": True},
                {"label": "Тема", "pattern": r"^НА ТЕМУ", "x": 187.3, "y": 395.2, "size": 14.0, "bold": False, "x_tol": 40},
                {"label": "Блок студента", "pattern": r"ЗДОБУВАЧА ГРУПИ", "x": 318.7, "y": 467.7, "size": 14.0, "bold": False, "x_tol": 40},
                {"label": "Науковий керівник", "pattern": r"НАУКОВИЙ\\s+КЕРІВНИК", "x": 318.7, "y": 564.3, "size": 14.0, "bold": False, "x_tol": 40},
                {"label": "Місто і рік", "pattern": r"КИЇВ", "x": 285.4, "y": 733.3, "size": 14.0, "bold": False, "x_tol": 40},
            ],
        }

    if work_type in {"Звіт з практики (Бакалавр)", "Звіт з практики (Магістр)"}:
        degree_pattern = (
            r"ЗДОБУВАЧА ПЕРШОГО \(БАКАЛАВРСЬКОГО\) РІВНЯ"
            if work_type == "Звіт з практики (Бакалавр)"
            else r"ЗДОБУВАЧА ДРУГОГО \(МАГІСТЕРСЬКОГО\) РІВНЯ"
        )
        return {
            "detection_patterns": [
                r"МІНІСТЕРСТВО ОСВІТИ І НАУКИ УКРАЇНИ",
                r"ЗВІТ З ПРАКТИКИ",
                degree_pattern,
                r"КЕРІВНИКИ ПРАКТИКИ",
            ],
            "required_matches": 4,
            "specs": [
                {"label": "Міністерство", "pattern": r"МІНІСТЕРСТВО ОСВІТИ І НАУКИ УКРАЇНИ", "x": 165.9, "y": 57.2, "size": 13.9, "bold": True},
                {"label": "Університет", "pattern": r"КИЇВСЬКИЙ НАЦІОНАЛЬНИЙ ЕКОНОМІЧНИЙ УНІВЕРСИТЕТ", "x": 100.8, "y": 73.3, "size": 13.9, "bold": True},
                {"label": "Імені Вадима Гетьмана", "pattern": r"ІМЕНІ ВАДИМА ГЕТЬМАНА", "x": 220.9, "y": 89.4, "size": 13.9, "bold": True},
                {"label": "Факультет", "pattern": r"ФАКУЛЬТЕТ МАРКЕТИНГУ", "x": 239.6, "y": 122.5, "size": 15.1, "bold": True},
                {"label": "Кафедра", "pattern": r"КАФЕДРА МАРКЕТИНГУ ІМЕНІ А\\.Ф\\. ПАВЛЕНКА", "x": 185.4, "y": 152.3, "size": 13.9, "bold": True},
                {"label": "Освітня програма", "pattern": r"ОСВІТНЬО-ПРОФЕСІЙНА ПРОГРАМА", "x": 158.7, "y": 200.3, "size": 12.0, "bold": True},
                {"label": "Спеціальність", "pattern": r"СПЕЦІАЛЬНІСТЬ 075", "x": 236.7, "y": 221.0, "size": 12.0, "bold": False},
                {"label": "Галузь знань", "pattern": r"ГАЛУЗЬ ЗНАНЬ 07", "x": 190.9, "y": 241.6, "size": 12.0, "bold": False},
                {"label": "Форма навчання", "pattern": r"ФОРМА НАВЧАННЯ", "x": 241.3, "y": 273.8, "size": 12.0, "bold": False},
                {"label": "Назва звіту", "pattern": r"ЗВІТ З ПРАКТИКИ", "x": 255.5, "y": 331.9, "size": 13.9, "bold": True},
                {"label": "База практики", "pattern": r"^НА ", "x": 231.5, "y": 371.7, "size": 13.9, "bold": False, "x_tol": 45},
                {"label": "Рівень освіти", "pattern": degree_pattern, "x": 106.4, "y": 404.1, "size": 13.9, "bold": False, "x_tol": 60},
                {"label": "Керівники практики", "pattern": r"КЕРІВНИКИ ПРАКТИКИ", "x": 70.8, "y": 504.0, "size": 13.9, "bold": False, "x_tol": 50},
                {"label": "Керівник від кафедри", "pattern": r"ВІД КАФЕДРИ", "x": 70.8, "y": 536.4, "size": 13.9, "bold": False, "x_tol": 50},
                {"label": "Керівник від бази практики", "pattern": r"ВІД БАЗИ ПРАКТИКИ", "x": 70.8, "y": 568.6, "size": 13.9, "bold": False, "x_tol": 50},
                {"label": "Початок практики", "pattern": r"ПОЧАТОК ПРАКТИКИ", "x": 70.8, "y": 632.9, "size": 13.9, "bold": False, "x_tol": 50},
                {"label": "Кінець практики", "pattern": r"КІНЕЦЬ ПРАКТИКИ", "x": 70.8, "y": 665.1, "size": 13.9, "bold": False, "x_tol": 50},
                {"label": "Місто і рік", "pattern": r"КИЇВ", "x": 288.4, "y": 729.7, "size": 13.9, "bold": True, "x_tol": 45},
            ],
        }

    heading_pattern = (
        r"КВАЛІФІКАЦІЙНА БАКАЛАВРСЬКА РОБОТА"
        if work_type == "Кваліфікаційна бакалаврська робота"
        else r"КВАЛІФІКАЦІЙНА МАГІСТЕРСЬКА РОБОТА"
    )
    return {
        "detection_patterns": [
            r"МІНІСТЕРСТВО ОСВІТИ І НАУКИ УКРАЇНИ",
            r"ОСВІТНЬО-ПРОФЕСІЙНА ПРОГРАМА",
            r"КВАЛІФІКАЦІЙНА",
            r"ЗАВІДУВАЧ КАФЕДРИ",
        ],
        "required_matches": 4,
        "specs": [
            {"label": "Міністерство", "pattern": r"МІНІСТЕРСТВО ОСВІТИ І НАУКИ УКРАЇНИ", "x": 176.8, "y": 55.0, "size": 13.0, "bold": True, "x_tol": 70},
            {"label": "Університет", "pattern": r"КИЇВСЬКИЙ НАЦІОНАЛЬНИЙ ЕКОНОМІЧНИЙ УНІВЕРСИТЕТ", "x": 116.6, "y": 69.9, "size": 13.0, "bold": True, "x_tol": 70},
            {"label": "Імені Вадима Гетьмана", "pattern": r"ІМЕНІ ВАДИМА ГЕТЬМАНА", "x": 228.1, "y": 84.9, "size": 13.0, "bold": True},
            {"label": "Факультет", "pattern": r"ФАКУЛЬТЕТ МАРКЕТИНГУ", "x": 239.8, "y": 116.6, "size": 15.0, "bold": True},
            {"label": "Кафедра", "pattern": r"КАФЕДРА МАРКЕТИНГУ ІМЕНІ А\\.Ф\\. ПАВЛЕНКА", "x": 185.5, "y": 146.6, "size": 14.0, "bold": True},
            {"label": "Освітня програма", "pattern": r"ОСВІТНЬО-ПРОФЕСІЙНА ПРОГРАМА", "x": 158.8, "y": 195.0, "size": 12.0, "bold": True},
            {"label": "Галузь знань", "pattern": r"ГАЛУЗЬ ЗНАНЬ 07", "x": 190.9, "y": 215.8, "size": 12.0, "bold": False},
            {"label": "Спеціальність", "pattern": r"СПЕЦІАЛЬНІСТЬ 075", "x": 236.8, "y": 236.4, "size": 12.0, "bold": False},
            {"label": "Форма навчання", "pattern": r"ФОРМА НАВЧАННЯ", "x": 234.8, "y": 284.5, "size": 13.0, "bold": False},
            {"label": "Назва кваліфікаційної роботи", "pattern": heading_pattern, "x": 138.9, "y": 346.1, "size": 16.0, "bold": True, "x_tol": 90},
            {"label": "Тема роботи", "pattern": r"^НА ТЕМУ", "x": 83.2, "y": 382.1, "size": 14.0, "bold": True, "x_tol": 90},
            {"label": "Здобувач", "pattern": r"^ЗДОБУВАЧА", "x": 123.7, "y": 430.4, "size": 14.0, "bold": False, "x_tol": 90},
            {"label": "Науковий керівник", "pattern": r"НАУКОВИЙ КЕРІВНИК", "x": 181.1, "y": 510.9, "size": 13.0, "bold": False, "x_tol": 50},
            {"label": "Допуск до захисту", "pattern": r"РОБОТА ДОПУЩЕНА ДО ЗАХИСТУ", "x": 183.9, "y": 585.5, "size": 13.0, "bold": True, "x_tol": 60},
            {"label": "Завідувач кафедри", "pattern": r"ЗАВІДУВАЧ КАФЕДРИ", "x": 184.3, "y": 631.5, "size": 13.0, "bold": False, "x_tol": 60},
            {"label": "Місто і рік", "pattern": r"КИЇВ", "x": 274.6, "y": 754.4, "size": 14.0, "bold": True, "x_tol": 50},
        ],
    }


def validate_title_page(page, work_type, report):
    lines = extract_lines(page)
    specs = get_title_specs(work_type)
    found = sum(1 for pattern in specs["detection_patterns"] if find_best_line(lines, pattern))
    if found < specs["required_matches"]:
        return False
    for spec in specs["specs"]:
        validate_line(lines, report["Титульна сторінка"], 1, spec)
    return True


def validate_contents_page(page, report):
    lines = extract_lines(page)
    detection_patterns = [
        r"^ЗМІСТ$",
        r"^ВСТУП",
        r"РОЗДІЛ\s+1",
        r"ВИСНОВКИ",
        r"СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ",
    ]
    found = sum(1 for pattern in detection_patterns if find_best_line(lines, pattern))
    if found < 4:
        return False

    specs = [
        {"label": "Заголовок 'ЗМІСТ'", "pattern": r"^ЗМІСТ$", "x": 304.4, "y": 88.2, "size": 13.9, "bold": True, "x_tol": 50},
        {"label": "Рядок 'ВСТУП'", "pattern": r"^ВСТУП", "x": 81.6, "y": 113.2, "size": 13.9, "bold": True, "x_tol": 30},
        {"label": "Розділ 1", "pattern": r"РОЗДІЛ\s+1", "x": 81.6, "y": 137.4, "size": 13.9, "bold": True, "x_tol": 40},
        {"label": "Підрозділ 1.1", "pattern": r"^1\.1\.", "x": 81.6, "y": 209.7, "size": 13.9, "bold": False, "x_tol": 30},
        {"label": "Висновки", "pattern": r"^ВИСНОВКИ", "x": 81.6, "y": 596.4, "size": 13.9, "bold": True, "x_tol": 30},
        {"label": "Список використаних джерел", "pattern": r"^СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ", "x": 81.6, "y": 620.4, "size": 13.9, "bold": True, "x_tol": 30},
        {"label": "Номер сторінки для 'ВСТУП'", "pattern": r"^\d+$", "x": 559.9, "y": 113.2, "size": 13.9, "bold": True, "x_tol": 25},
        {"label": "Номер сторінки для 'Списку джерел'", "pattern": r"^\d+$", "x": 556.3, "y": 620.4, "size": 13.9, "bold": True, "x_tol": 25},
    ]
    for spec in specs:
        validate_line(lines, report["Сторінка зі змістом"], 2, spec)
    return True


def analyze_body_pages(doc, report, start_page=2):
    expected_size = 14.0
    in_bibliography = False

    for page_num in range(start_page, len(doc)):
        page = doc[page_num]
        rect = page.rect
        blocks = page.get_text("dict")["blocks"]
        min_x = rect.width
        max_x = 0
        min_y = rect.height
        max_y = 0

        for block in blocks:
            if "lines" not in block:
                continue

            bbox = block["bbox"]
            min_x = min(min_x, bbox[0])
            min_y = min(min_y, bbox[1])
            max_x = max(max_x, bbox[2])
            max_y = max(max_y, bbox[3])
            lines = block["lines"]

            if len(lines) > 1:
                first_line_x = lines[0]["bbox"][0]
                second_line_x = lines[1]["bbox"][0]
                indent_cm = (first_line_x - second_line_x) * PT_TO_CM
                if indent_cm > 0.5 and abs(indent_cm - 1.5) > 0.3:
                    add_page_error(
                        report,
                        "Абзацний відступ (1.5 см)",
                        page_num + 1,
                        f"Відступ ~{round(indent_cm, 2)} см. <i>'{lines[0]['spans'][0]['text'][:25]}...'</i>",
                    )

                if lines[0]["spans"] and lines[1]["spans"]:
                    prev_y = lines[0]["bbox"][3]
                    curr_y = lines[1]["bbox"][3]
                    fs = lines[0]["spans"][0]["size"]
                    if fs > 10:
                        line_ratio = (curr_y - prev_y) / fs
                        if (line_ratio < 1.35 or line_ratio > 1.7) and line_ratio > 0.9:
                            add_page_error(
                                report,
                                "Міжрядковий інтервал (1.5)",
                                page_num + 1,
                                f"Інтервал ~{round(line_ratio, 1)}. <i>'{lines[0]['spans'][0]['text'][:25]}...'</i>",
                            )

            full_text = ""
            for line in lines:
                for span in line["spans"]:
                    text_strip = span["text"].strip()
                    full_text += text_strip + " "
                    if len(text_strip) <= 5:
                        continue

                    font_name = span["font"]
                    font_size = span["size"]
                    if "Times" not in font_name and "Symbol" not in font_name:
                        add_page_error(
                            report,
                            "Шрифт (Times New Roman)",
                            page_num + 1,
                            f"<code>{font_name}</code>. <i>'{text_strip[:25]}...'</i>",
                        )
                    if abs(font_size - expected_size) > 0.5 and font_size >= 10 and not text_strip.isupper():
                        suffix = " (допустимо лише в таблицях)" if 10 <= font_size <= 12.5 else ""
                        add_page_error(
                            report,
                            "Розмір шрифту (14)",
                            page_num + 1,
                            f"Розмір {round(font_size, 1)}{suffix}. <i>'{text_strip[:25]}...'</i>",
                        )

            full_text_strip = full_text.strip()
            if full_text_strip in ["ВСТУП", "ВИСНОВКИ", "СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ", "ДОДАТКИ"] or full_text_strip.startswith("РОЗДІЛ"):
                if full_text_strip.startswith("РОЗДІЛ") and re.search(r"РОЗДІЛ\s+\d+\.", full_text_strip):
                    add_page_error(report, "Оформлення заголовків (ВЕЛИКІ ЛІТЕРИ, по центру)", page_num + 1, f"Заборонена крапка після номера розділу: '{full_text_strip}'")
                if not full_text_strip.isupper():
                    add_page_error(report, "Оформлення заголовків (ВЕЛИКІ ЛІТЕРИ, по центру)", page_num + 1, f"Має бути великими літерами: '{full_text_strip}'")
                if full_text_strip.endswith("."):
                    add_page_error(report, "Оформлення заголовків (ВЕЛИКІ ЛІТЕРИ, по центру)", page_num + 1, f"Заголовок не повинен мати крапку в кінці: '{full_text_strip}'")

                expected_left_margin_pt = 2.5 / PT_TO_CM
                expected_right_margin_pt = 1.0 / PT_TO_CM
                space_left = bbox[0] - expected_left_margin_pt
                space_right = (rect.width - expected_right_margin_pt) - bbox[2]
                if abs(space_left - space_right) > 35:
                    add_page_error(report, "Оформлення заголовків (ВЕЛИКІ ЛІТЕРИ, по центру)", page_num + 1, f"Заголовок не центрований: '{full_text_strip}'")

            if re.match(r"^\d+\.\d+\s+[А-ЯІЇЄҐ]", full_text_strip):
                if full_text_strip.endswith("."):
                    add_page_error(report, "Оформлення підрозділів (Відступ 1.5 см, без крапки в кінці)", page_num + 1, f"Заборонена крапка в кінці заголовка: '{full_text_strip}'")
                expected_left_margin_pt = 2.5 / PT_TO_CM
                abs_indent = bbox[0] - expected_left_margin_pt
                if abs(abs_indent - (1.5 / PT_TO_CM)) > 15:
                    indent_cm = abs_indent * PT_TO_CM
                    add_page_error(report, "Оформлення підрозділів (Відступ 1.5 см, без крапки в кінці)", page_num + 1, f"Відступ підрозділу ~{round(indent_cm, 1)} см (має бути 1.5 см): '{full_text_strip}'")

            if full_text_strip.startswith("Рисунок"):
                if not re.search(r"Рисунок\s+(\d+\.\d+|[А-Я]\.\d+)\s+[-–]\s+", full_text_strip) and len(full_text_strip) < 150:
                    add_page_error(report, "Підписи до рисунків (Формат: Рисунок X.X - Назва)", page_num + 1, f"Неправильний формат: <i>'{full_text_strip[:45]}...'</i>")

            if full_text_strip.startswith("Таблиця"):
                invalid_format = not re.search(r"Таблиця\s+(\d+\.\d+|[А-Я]\.\d+)\s+[-–]\s+", full_text_strip)
                if invalid_format and "Продовження" not in full_text_strip and "Кінець" not in full_text_strip and len(full_text_strip) < 150:
                    add_page_error(report, "Підписи до таблиць (Формат: Таблиця X.X - Назва)", page_num + 1, f"Неправильний формат: <i>'{full_text_strip[:45]}...'</i>")

            if re.search(r"\(\d+\.\d+\)$", full_text_strip) and len(full_text_strip) < 100:
                expected_right_margin_pt = 1.0 / PT_TO_CM
                space_right = (rect.width - expected_right_margin_pt) - bbox[2]
                if space_right > 35:
                    add_page_error(report, "Оформлення формул (Номер праворуч у дужках)", page_num + 1, f"Номер формули не притиснуто до правого краю: <i>'{full_text_strip[-15:]}'</i>")

            if full_text_strip == "СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ":
                in_bibliography = True
            elif full_text_strip == "ДОДАТКИ":
                in_bibliography = False
            elif in_bibliography and re.match(r"^\d+\.", full_text_strip):
                has_year = re.search(r"20\d\d", full_text_strip) or re.search(r"19\d\d", full_text_strip)
                if not has_year and "URL" not in full_text_strip and "http" not in full_text_strip:
                    add_page_error(report, "Список використаних джерел (ДСТУ 8302:2015)", page_num + 1, f"Можлива помилка ДСТУ (не знайдено року видання): <i>'{full_text_strip[:50]}...'</i>")

        if min_x < rect.width and max_x > 0:
            left_cm = min_x * PT_TO_CM
            right_cm = (rect.width - max_x) * PT_TO_CM
            top_cm = min_y * PT_TO_CM
            if abs(left_cm - 2.5) > 0.35:
                add_page_error(report, "Поля сторінки (Л: 2.5 см, П: 1.0 см, В/Н: 2.0 см)", page_num + 1, f"Ліве поле ~{round(left_cm, 1)} см")
            if abs(right_cm - 1.0) > 0.35:
                add_page_error(report, "Поля сторінки (Л: 2.5 см, П: 1.0 см, В/Н: 2.0 см)", page_num + 1, f"Праве поле ~{round(right_cm, 1)} см")
            if abs(top_cm - 2.0) > 0.35 and top_cm > 1.0:
                add_page_error(report, "Поля сторінки (Л: 2.5 см, П: 1.0 см, В/Н: 2.0 см)", page_num + 1, f"Верхнє поле ~{round(top_cm, 1)} см")


def analyze_pdf(file_bytes, work_type):
    doc = pymupdf.open(stream=file_bytes, filetype="pdf")
    report = build_report()
    if len(doc) < 2:
        return {"report": report, "stop_message": "Завантажте будь ласка роботу з титульною сторінкою та змістом"}

    title_ok = validate_title_page(doc[0], work_type, report)
    contents_ok = validate_contents_page(doc[1], report)
    if not title_ok or not contents_ok:
        return {"report": report, "stop_message": "Завантажте будь ласка роботу з титульною сторінкою та змістом"}

    analyze_body_pages(doc, report, start_page=2)
    return {"report": truncate_report(report), "stop_message": None}


def run_app():
    st.set_page_config(page_title="Перевірка студентських робіт", page_icon="🎓", layout="centered")
    st.markdown(
        """
<style>
    [data-testid="stExpander"] details summary p span { position: absolute; right: 45px; }
    [data-testid="stAlert"] p span { position: absolute; right: 20px; }
    [data-testid="stExpander"] details summary p, [data-testid="stAlert"] p { padding-right: 120px; }
</style>
""",
        unsafe_allow_html=True,
    )

    st.title("🎓 Інспектор оформлення студентських робіт")
    st.write("Застосунок перевіряє титульну сторінку, зміст і основний текст PDF-роботи за правилами оформлення.")
    work_type = st.selectbox("Оберіть напрям вашої роботи", WORK_OPTIONS)
    st.caption(f"Для цього напряму титульна сторінка звіряється зі зразком: {TITLE_SAMPLE_MAP[work_type]}. Сторінка змісту звіряється зі зразком: {CONTENTS_SAMPLE_FILE}.")
    st.info("Завантажуйте будь ласка роботу обов'язково з титульною сторінкою та змістом")
    uploaded_file = st.file_uploader("Завантажте PDF-файл вашої роботи", type=["pdf"])

    if uploaded_file is None:
        return

    with st.spinner("Сканую документ..."):
        result = analyze_pdf(uploaded_file.read(), work_type)

    if result["stop_message"]:
        st.error(result["stop_message"])
        return

    report = result["report"]
    st.markdown("---")
    st.subheader("📊 Результати перевірки")
    all_clear = True
    for rule, errors in report.items():
        if not errors:
            st.success(f"✅ **{rule}**  |  Помилок: **0**  :green[**ІДЕАЛЬНО**]")
        else:
            all_clear = False
            with st.expander(f"❌ **{rule}**  |  Помилок: **{len(errors)}** :red[**РОЗГОРНУТИ**]", expanded=False):
                for error in errors:
                    st.markdown(f"- {error}", unsafe_allow_html=True)

    st.markdown("---")
    if all_clear:
        st.balloons()
        st.success("🏆 Документ виглядає коректно з точки зору правил, які перевіряє застосунок.")
    else:
        st.warning("Будь ласка, виправте зауваження та завантажте файл повторно.")
