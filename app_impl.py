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


@st.cache_data(show_spinner=False)
def load_sample_lines(sample_file):
    doc = pymupdf.open(sample_file)
    try:
        return extract_lines(doc[0])
    finally:
        doc.close()


def build_report():
    return {
        "Титульна сторінка": [],
        "Сторінка зі змістом": [],
        "Нумерація сторінок (вгорі праворуч)": [],
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


TABLE_SOURCE_RULE = "Посилання на джерело даних таблиці"


def truncate_report(report):
    for rule, errors in report.items():
        report[rule] = list(dict.fromkeys(errors))
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


def is_page_number_line(line, page_rect):
    text = normalize_text(line["text"])
    return bool(re.fullmatch(r"\d+", text)) and line["x1"] > page_rect.width * 0.8 and line["y1"] < page_rect.height * 0.12


def validate_page_number(page, report, page_number):
    lines = extract_lines(page)
    page_line = next((line for line in lines if is_page_number_line(line, page.rect)), None)
    if page_line is None:
        add_page_error(report, "Нумерація сторінок (вгорі праворуч)", page_number, "Не знайдено номер сторінки у верхньому правому куті.")


def normalize_font_size(size):
    return int(round(size))


def format_vertical_shift(delta_y, base_size):
    direction = "вниз" if delta_y > 0 else "вгору"
    line_height = max(base_size * 1.5, 12)
    lines = max(1, round(abs(delta_y) / line_height))
    return direction, lines


def format_horizontal_shift(delta_x, base_size):
    direction = "праворуч" if delta_x > 0 else "ліворуч"
    char_width = max(base_size * 0.55, 6)
    chars = max(1, round(abs(delta_x) / char_width))
    return direction, chars


def get_text_blocks_above_table(blocks, table_top):
    texts = []
    for block in blocks:
        if "lines" not in block or block["bbox"][3] > table_top + 10:
            continue
        text = "".join(span["text"] for line in block["lines"] for span in line["spans"]).strip()
        if text:
            texts.append(text)
    return texts


def get_first_text_line(blocks):
    first_line = None
    for block in blocks:
        if "lines" not in block:
            continue
        for line in block["lines"]:
            spans = [span for span in line["spans"] if span["text"].strip()]
            if not spans:
                continue
            text = "".join(span["text"] for span in spans).strip()
            if not text:
                continue
            candidate = {
                "text": text,
                "x0": line["bbox"][0],
                "y0": line["bbox"][1],
                "x1": line["bbox"][2],
                "y1": line["bbox"][3],
            }
            if first_line is None or (candidate["y0"], candidate["x0"]) < (first_line["y0"], first_line["x0"]):
                first_line = candidate
    return first_line


def get_first_meaningful_line(page):
    page_lines = extract_lines(page)
    return next((line for line in page_lines if not is_page_number_line(line, page.rect)), None)


def is_table_title_line(text):
    return bool(re.match(r"^Таблиця\s+(\d+\.\d+|[А-ЯІЇЄҐ]\.\d+)\s+[-–]\s+", normalize_text(text), re.IGNORECASE))


def is_table_end_line(text):
    return bool(re.match(r"^Кінець\s+таблиці", normalize_text(text), re.IGNORECASE))


def is_table_continuation_line(text):
    return bool(re.match(r"^Продовження\s+таблиці", normalize_text(text), re.IGNORECASE))


def caption_has_inline_source(text):
    if not is_table_title_line(text):
        return False
    title_parts = re.split(r"\s+[-–]\s+", normalize_text(text), maxsplit=1)
    if len(title_parts) < 2:
        return False
    return bool(re.search(r"\[[^\[\]]+\]", title_parts[1]))


def is_valid_table_source_line(text):
    normalized = normalize_text(text)
    if not re.match(r"^Джерело\s*:", normalized, re.IGNORECASE):
        return False

    source_body = normalized.split(":", 1)[1].strip()
    if not source_body:
        return False
    if re.fullmatch(r"\[[^\[\]]+\]", source_body):
        return True

    author_based_pattern = r"^(складено автором за даними|розроблено автором на основі)\b.*\[[^\[\]]+\]"
    return bool(re.search(author_based_pattern, source_body, re.IGNORECASE))


def is_valid_table_source_line(text):
    normalized = normalize_text(text)
    if not re.match(r"^Р”Р¶РµСЂРµР»Рѕ\s*:", normalized, re.IGNORECASE):
        return False

    source_body = normalized.split(":", 1)[1].strip()
    if not source_body:
        return False
    if re.fullmatch(r"\[[^\[\]]+\]", source_body):
        return True
    if not re.search(r"\[[^\[\]]+\]", source_body):
        return False

    author_based_pattern = (
        r"^(СЃС‚РІРѕСЂРµРЅРѕ|СЃРєР»Р°РґРµРЅРѕ|СЂРѕР·СЂРѕР±Р»РµРЅРѕ|СѓР·Р°РіР°Р»СЊРЅРµРЅРѕ)"
        r"\s+Р°РІС‚РѕСЂРѕРј\s+(Р·Р°\s+РґР°РЅРёРјРё|РЅР°\s+РѕСЃРЅРѕРІС–)\b.*\[[^\[\]]+\]$"
    )
    return bool(re.match(author_based_pattern, source_body, re.IGNORECASE))


def find_table_header_line(page_lines, table_bbox):
    candidates = [
        line
        for line in page_lines
        if line["y1"] <= table_bbox[1] + 12
        and table_bbox[1] - line["y1"] <= 90
        and (is_table_title_line(line["text"]) or is_table_continuation_line(line["text"]) or is_table_end_line(line["text"]))
    ]
    if not candidates:
        return None
    return max(candidates, key=lambda line: line["y1"])


def find_table_source_line(page_lines, table_bbox):
    candidates = [
        line
        for line in page_lines
        if line["y0"] >= table_bbox[3] - 2
        and line["y0"] - table_bbox[3] <= 110
        and re.match(r"^Джерело\s*:", normalize_text(line["text"]), re.IGNORECASE)
    ]
    if not candidates:
        return None
    return min(candidates, key=lambda line: line["y0"])


def table_looks_real(table, page_rect):
    content_bbox = get_table_content_bbox(table)
    if content_bbox is None:
        return False

    x0, y0, x1, y1 = content_bbox
    width = x1 - x0
    height = y1 - y0
    row_count = getattr(table, "row_count", 0)
    col_count = getattr(table, "col_count", 0)
    if width < page_rect.width * 0.28 or height < 45:
        return False
    if row_count and col_count and (row_count < 2 or col_count < 2):
        return False
    return True


def count_filled_cells(row_values):
    return sum(1 for value in row_values if value and normalize_text(str(value)))


def get_table_content_bbox(table):
    data = table.extract()
    if not data or not getattr(table, "rows", None):
        return None

    meaningful_indexes = [index for index, row in enumerate(data) if count_filled_cells(row) >= 2]
    if not meaningful_indexes:
        return None

    groups = []
    current_group = [meaningful_indexes[0]]
    for index in meaningful_indexes[1:]:
        if index == current_group[-1] + 1:
            current_group.append(index)
        else:
            groups.append(current_group)
            current_group = [index]
    groups.append(current_group)

    best_group = max(groups, key=len)
    if len(best_group) < 2:
        return None

    start_row = table.rows[best_group[0]]
    end_row = table.rows[best_group[-1]]
    return (
        start_row.bbox[0],
        start_row.bbox[1],
        end_row.bbox[2],
        end_row.bbox[3],
    )


def point_in_bbox(x, y, bbox, padding=0):
    return (bbox[0] - padding) <= x <= (bbox[2] + padding) and (bbox[1] - padding) <= y <= (bbox[3] + padding)


def bboxes_intersect(bbox1, bbox2, padding=0):
    return not (
        bbox1[2] < bbox2[0] - padding
        or bbox1[0] > bbox2[2] + padding
        or bbox1[3] < bbox2[1] - padding
        or bbox1[1] > bbox2[3] + padding
    )


def is_measurement_text_line(line):
    text = normalize_text("".join(span["text"] for span in line["spans"]))
    if len(text) < 8:
        return False
    if re.fullmatch(r"[\d\W_]+", text):
        return False
    return True


def block_text_content(block):
    if "lines" not in block:
        return ""
    return normalize_text("".join(span["text"] for line in block["lines"] for span in line["spans"]))


def is_body_margin_text(text):
    if not text or len(text) < 25:
        return False
    upper_text = text.upper()
    if text.isupper():
        return False
    if upper_text.startswith(("РОЗДІЛ", "ДОДАТКИ", "ДОДАТОК", "КІНЕЦЬ ТАБЛИЦІ", "ТАБЛИЦЯ ", "РИСУНОК ")):
        return False
    if re.match(r"^\d+(\.\d+)*\s+[А-ЯІЇЄҐA-Z]", text):
        return False
    return True


def is_figure_like_block(text, bbox, rect):
    if not text:
        return False
    upper_text = text.upper()
    if upper_text.startswith(("РИСУНОК", "РИС.", "ТАБЛИЦЯ", "КІНЕЦЬ ТАБЛИЦІ", "ПРОДОВЖЕННЯ ТАБЛИЦІ")):
        return False
    if upper_text in {"ВСТУП", "ВИСНОВКИ", "ДОДАТКИ"} or upper_text.startswith("РОЗДІЛ"):
        return False
    if re.match(r"^\d+\.\d+\s+[А-ЯІЇЄҐA-Z]", text):
        return False

    width = bbox[2] - bbox[0]
    page_width = rect.width
    centered = abs(bbox[0] - (page_width - bbox[2])) < 45
    narrow = width < page_width * 0.62
    short_text = len(text) < 180
    return short_text and (narrow or centered) and not is_body_margin_text(text)


def collect_margin_bboxes(blocks, rect, table_bboxes):
    top_bboxes = []
    body_bboxes = []
    image_bboxes = []

    for block in blocks:
        bbox = block.get("bbox")
        if not bbox:
            continue

        if block.get("type") == 1:
            image_bboxes.append(bbox)
            top_bboxes.append(bbox)
            continue

        text = block_text_content(block)
        if not text:
            continue

        block_lines = []
        for raw_line in block.get("lines", []):
            line_text = normalize_text("".join(span["text"] for span in raw_line["spans"]))
            if not line_text:
                continue
            block_lines.append({"text": line_text, "x1": raw_line["bbox"][2], "y1": raw_line["bbox"][3]})

        if block_lines and all(is_page_number_line(line, rect) for line in block_lines):
            continue

        top_bboxes.append(bbox)

        if any(bboxes_intersect(bbox, table_bbox, padding=4) for table_bbox in table_bboxes):
            continue

        if is_body_margin_text(text):
            body_bboxes.append(bbox)

    has_large_image = any((bbox[2] - bbox[0]) * (bbox[3] - bbox[1]) > rect.width * rect.height * 0.15 for bbox in image_bboxes)
    return top_bboxes, body_bboxes, has_large_image


def flush_bibliography_entry(report, page_number, entry_parts):
    if not entry_parts:
        return
    entry_text = normalize_text(" ".join(entry_parts))
    has_year = bool(re.search(r"(19|20)\d{2}", entry_text) or re.search(r"(19|20)\d{2}\s*[–-]\s*(19|20)\d{2}", entry_text))
    has_url = "URL" in entry_text.upper() or "HTTP" in entry_text.upper()
    if not has_year and not has_url:
        add_page_error(
            report,
            "Список використаних джерел (ДСТУ 8302:2015)",
            page_number,
            f"Можлива помилка ДСТУ (не знайдено року видання): <i>'{entry_text[:90]}...'</i>",
        )


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
        direction, chars = format_horizontal_shift(line["x0"] - spec["x"], spec.get("size", line["size"]))
        rule_errors.append(
            f"<b>Сторінка {page_number}</b>: '{label}' зміщено по горизонталі {direction} на {chars} знаків."
        )
    if "y" in spec and abs(line["y0"] - spec["y"]) > spec.get("y_tol", 16):
        direction, lines_count = format_vertical_shift(line["y0"] - spec["y"], spec.get("size", line["size"]))
        rule_errors.append(
            f"<b>Сторінка {page_number}</b>: '{label}' зміщено по вертикалі {direction} на {lines_count} строк."
        )
    expected_size = normalize_font_size(spec["size"]) if "size" in spec else None
    actual_size = normalize_font_size(line["size"])
    if expected_size is not None and actual_size != expected_size:
        rule_errors.append(
            f"<b>Сторінка {page_number}</b>: '{label}' має розмір шрифту {actual_size}, "
            f"а за зразком очікується {expected_size}."
        )
    if "bold" in spec and line["bold"] != spec["bold"]:
        expected = "жирним" if spec["bold"] else "звичайним"
        rule_errors.append(
            f"<b>Сторінка {page_number}</b>: '{label}' має бути набрано {expected}."
        )
    return True


def build_specs_from_sample(sample_file, templates):
    sample_lines = load_sample_lines(sample_file)
    specs = []
    for template in templates:
        sample_line = find_best_line(sample_lines, template["sample_pattern"])
        if sample_line is None:
            continue
        specs.append(
            {
                "label": template["label"],
                "pattern": template.get("upload_pattern", template["sample_pattern"]),
                "x": sample_line["x0"],
                "y": sample_line["y0"],
                "size": sample_line["size"],
                "bold": sample_line["bold"],
                "x_tol": template.get("x_tol", 24),
                "y_tol": template.get("y_tol", 16),
                "size_tol": template.get("size_tol", 1.0),
            }
        )
    return specs


def get_title_config(work_type):
    if work_type in {"Курсова з БЗВП", "Курсова з маркетингу"}:
        return {
            "sample_file": TITLE_SAMPLE_MAP[work_type],
            "detection_patterns": [
                r"МІНІСТЕРСТВО ОСВІТИ І НАУКИ УКРАЇНИ",
                r"КИЇВСЬКИЙ НАЦІОНАЛЬНИЙ ЕКОНОМІЧНИЙ УНІВЕРСИТЕТ",
                r"КУРСОВА РОБОТА",
                r"ЗДОБУВАЧА ГРУПИ",
            ],
            "required_matches": 4,
            "templates": [
                {"label": "Міністерство", "sample_pattern": r"МІНІСТЕРСТВО ОСВІТИ І НАУКИ УКРАЇНИ", "upload_pattern": r"МІНІСТЕРСТВО ОСВІТИ І НАУКИ УКРАЇНИ"},
                {"label": "Університет", "sample_pattern": r"КИЇВСЬКИЙ НАЦІОНАЛЬНИЙ ЕКОНОМІЧНИЙ УНІВЕРСИТЕТ", "upload_pattern": r"КИЇВСЬКИЙ НАЦІОНАЛЬНИЙ ЕКОНОМІЧНИЙ УНІВЕРСИТЕТ"},
                {"label": "Імені Вадима Гетьмана", "sample_pattern": r"ІМЕНІ ВАДИМА ГЕТЬМАНА", "upload_pattern": r"ІМЕНІ ВАДИМА ГЕТЬМАНА"},
                {"label": "Факультет", "sample_pattern": r"Факультет маркетингу", "upload_pattern": r"ФАКУЛЬТЕТ МАРКЕТИНГУ"},
                {"label": "Кафедра", "sample_pattern": r"Кафедра маркетингу імені А\\.Ф\\. Павленка", "upload_pattern": r"КАФЕДРА МАРКЕТИНГУ ІМЕНІ А\\.Ф\\. ПАВЛЕНКА"},
                {"label": "Назва роботи", "sample_pattern": r"КУРСОВА РОБОТА", "upload_pattern": r"КУРСОВА РОБОТА"},
                {"label": "Дисципліна", "sample_pattern": r"з навчальної дисципліни", "upload_pattern": r"З НАВЧАЛЬНОЇ ДИСЦИПЛІНИ"},
                {"label": "Тема", "sample_pattern": r"на тему", "upload_pattern": r"^НА ТЕМУ", "x_tol": 40},
                {"label": "Блок студента", "sample_pattern": r"Здобувача групи", "upload_pattern": r"ЗДОБУВАЧА ГРУПИ", "x_tol": 40},
                {"label": "Науковий керівник", "sample_pattern": r"Науковий\\s+керівник", "upload_pattern": r"НАУКОВИЙ\\s+КЕРІВНИК", "x_tol": 40},
                {"label": "Місто і рік", "sample_pattern": r"Київ\\s*-\\s*\\d{4}", "upload_pattern": r"КИЇВ", "x_tol": 40},
            ],
        }

    if work_type in {"Звіт з практики (Бакалавр)", "Звіт з практики (Магістр)"}:
        level_upload = (
            r"ЗДОБУВАЧА ПЕРШОГО \(БАКАЛАВРСЬКОГО\) РІВНЯ"
            if work_type == "Звіт з практики (Бакалавр)"
            else r"ЗДОБУВАЧА ДРУГОГО \(МАГІСТЕРСЬКОГО\) РІВНЯ"
        )
        return {
            "sample_file": TITLE_SAMPLE_MAP[work_type],
            "detection_patterns": [
                r"МІНІСТЕРСТВО ОСВІТИ І НАУКИ УКРАЇНИ",
                r"ЗВІТ З ПРАКТИКИ",
                r"ЗДОБУВАЧА",
                r"КЕРІВНИК(И)? ПРАКТИКИ",
            ],
            "required_matches": 4,
            "templates": [
                {"label": "Міністерство", "sample_pattern": r"МІНІСТЕРСТВО ОСВІТИ І НАУКИ УКРАЇНИ", "upload_pattern": r"МІНІСТЕРСТВО ОСВІТИ І НАУКИ УКРАЇНИ"},
                {"label": "Університет", "sample_pattern": r"КИЇВСЬКИЙ НАЦІОНАЛЬНИЙ ЕКОНОМІЧНИЙ УНІВЕРСИТЕТ", "upload_pattern": r"КИЇВСЬКИЙ НАЦІОНАЛЬНИЙ ЕКОНОМІЧНИЙ УНІВЕРСИТЕТ"},
                {"label": "Імені Вадима Гетьмана", "sample_pattern": r"ІМЕНІ ВАДИМА ГЕТЬМАНА", "upload_pattern": r"ІМЕНІ ВАДИМА ГЕТЬМАНА"},
                {"label": "Факультет", "sample_pattern": r"Факультет маркетингу", "upload_pattern": r"ФАКУЛЬТЕТ МАРКЕТИНГУ"},
                {"label": "Кафедра", "sample_pattern": r"Кафедра маркетингу імені А\\.Ф\\. Павленка", "upload_pattern": r"КАФЕДРА МАРКЕТИНГУ ІМЕНІ А\\.Ф\\. ПАВЛЕНКА"},
                {"label": "Освітня програма", "sample_pattern": r"ОСВІТНЬО-ПРОФЕСІЙНА ПРОГРАМА", "upload_pattern": r"ОСВІТНЬО-ПРОФЕСІЙНА ПРОГРАМА"},
                {"label": "Спеціальність", "sample_pattern": r"Спеціальність 075", "upload_pattern": r"СПЕЦІАЛЬНІСТЬ 075"},
                {"label": "Галузь знань", "sample_pattern": r"Галузь знань 07", "upload_pattern": r"ГАЛУЗЬ ЗНАНЬ 07"},
                {"label": "Форма навчання", "sample_pattern": r"Форма навчання", "upload_pattern": r"ФОРМА НАВЧАННЯ"},
                {"label": "Назва звіту", "sample_pattern": r"ЗВІТ З ПРАКТИКИ", "upload_pattern": r"ЗВІТ З ПРАКТИКИ"},
                {"label": "База практики", "sample_pattern": r"^на ", "upload_pattern": r"^НА ", "x_tol": 45},
                {"label": "Рівень освіти", "sample_pattern": r"здобувача другого \(магістерського\) рівня", "upload_pattern": level_upload, "x_tol": 60},
                {"label": "Керівники практики", "sample_pattern": r"Керівники практики", "upload_pattern": r"КЕРІВНИК(И)? ПРАКТИКИ", "x_tol": 50},
                {"label": "Керівник від кафедри", "sample_pattern": r"від кафедри", "upload_pattern": r"ВІД КАФЕДРИ", "x_tol": 50},
                {"label": "Керівник від бази практики", "sample_pattern": r"від бази практики", "upload_pattern": r"ВІД БАЗИ ПРАКТИКИ", "x_tol": 50},
                {"label": "Початок практики", "sample_pattern": r"Початок практики", "upload_pattern": r"ПОЧАТОК ПРАКТИКИ", "x_tol": 50},
                {"label": "Кінець практики", "sample_pattern": r"Кінець практики", "upload_pattern": r"КІНЕЦЬ ПРАКТИКИ", "x_tol": 50},
                {"label": "Місто і рік", "sample_pattern": r"Київ\\s+\\d{4}", "upload_pattern": r"КИЇВ", "x_tol": 45},
            ],
        }

    heading_upload = (
        r"КВАЛІФІКАЦІЙНА БАКАЛАВРСЬКА РОБОТА"
        if work_type == "Кваліфікаційна бакалаврська робота"
        else r"КВАЛІФІКАЦІЙНА МАГІСТЕРСЬКА РОБОТА"
    )
    return {
        "sample_file": TITLE_SAMPLE_MAP[work_type],
        "detection_patterns": [
            r"МІНІСТЕРСТВО ОСВІТИ І НАУКИ УКРАЇНИ",
            r"ОСВІТНЬО-ПРОФЕСІЙНА ПРОГРАМА",
            r"КВАЛІФІКАЦІЙНА",
            r"ЗАВІДУВАЧ КАФЕДРИ",
        ],
        "required_matches": 4,
        "templates": [
            {"label": "Міністерство", "sample_pattern": r"МІНІСТЕРСТВО ОСВІТИ І НАУКИ УКРАЇНИ", "upload_pattern": r"МІНІСТЕРСТВО ОСВІТИ І НАУКИ УКРАЇНИ", "x_tol": 70},
            {"label": "Університет", "sample_pattern": r"КИЇВСЬКИЙ НАЦІОНАЛЬНИЙ ЕКОНОМІЧНИЙ УНІВЕРСИТЕТ", "upload_pattern": r"КИЇВСЬКИЙ НАЦІОНАЛЬНИЙ ЕКОНОМІЧНИЙ УНІВЕРСИТЕТ", "x_tol": 70},
            {"label": "Імені Вадима Гетьмана", "sample_pattern": r"ІМЕНІ ВАДИМА ГЕТЬМАНА", "upload_pattern": r"ІМЕНІ ВАДИМА ГЕТЬМАНА"},
            {"label": "Факультет", "sample_pattern": r"Факультет маркетингу", "upload_pattern": r"ФАКУЛЬТЕТ МАРКЕТИНГУ"},
            {"label": "Кафедра", "sample_pattern": r"Кафедра маркетингу імені А\\.Ф\\. Павленка", "upload_pattern": r"КАФЕДРА МАРКЕТИНГУ ІМЕНІ А\\.Ф\\. ПАВЛЕНКА"},
            {"label": "Освітня програма", "sample_pattern": r"ОСВІТНЬО-ПРОФЕСІЙНА ПРОГРАМА", "upload_pattern": r"ОСВІТНЬО-ПРОФЕСІЙНА ПРОГРАМА"},
            {"label": "Галузь знань", "sample_pattern": r"Галузь знань 07", "upload_pattern": r"ГАЛУЗЬ ЗНАНЬ 07"},
            {"label": "Спеціальність", "sample_pattern": r"Спеціальність 075", "upload_pattern": r"СПЕЦІАЛЬНІСТЬ 075"},
            {"label": "Форма навчання", "sample_pattern": r"Форма навчання", "upload_pattern": r"ФОРМА НАВЧАННЯ"},
            {"label": "Назва кваліфікаційної роботи", "sample_pattern": r"КВАЛІФІКАЦІЙНА БАКАЛАВРСЬКА РОБОТА", "upload_pattern": heading_upload, "x_tol": 90},
            {"label": "Тема роботи", "sample_pattern": r"^на тему", "upload_pattern": r"^НА ТЕМУ", "x_tol": 90},
            {"label": "Здобувач", "sample_pattern": r"^здобувача", "upload_pattern": r"^ЗДОБУВАЧА", "x_tol": 90},
            {"label": "Науковий керівник", "sample_pattern": r"Науковий керівник", "upload_pattern": r"НАУКОВИЙ КЕРІВНИК", "x_tol": 50},
            {"label": "Допуск до захисту", "sample_pattern": r"Робота допущена до захисту", "upload_pattern": r"РОБОТА ДОПУЩЕНА ДО ЗАХИСТУ", "x_tol": 60},
            {"label": "Завідувач кафедри", "sample_pattern": r"Завідувач кафедри", "upload_pattern": r"ЗАВІДУВАЧ КАФЕДРИ", "x_tol": 60},
            {"label": "Місто і рік", "sample_pattern": r"Київ\\s+\\d{4}", "upload_pattern": r"КИЇВ", "x_tol": 50},
        ],
    }


def get_contents_config():
    return {
        "sample_file": CONTENTS_SAMPLE_FILE,
        "detection_patterns": [
            r"^ЗМІСТ$",
            r"^ВСТУП",
            r"РОЗДІЛ\s+1",
            r"ВИСНОВ(ОК|КИ)",
            r"(СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ|ДЖЕРЕЛА)",
        ],
        "required_matches": 4,
        "templates": [
            {"label": "Заголовок 'ЗМІСТ'", "sample_pattern": r"^ЗМІСТ$", "upload_pattern": r"^ЗМІСТ$", "x_tol": 50},
            {"label": "Рядок 'ВСТУП'", "sample_pattern": r"^ВСТУП", "upload_pattern": r"^ВСТУП", "x_tol": 30},
            {"label": "Розділ 1", "sample_pattern": r"РОЗДІЛ\s+1", "upload_pattern": r"РОЗДІЛ\s+1", "x_tol": 40},
            {"label": "Підрозділ 1.1", "sample_pattern": r"^1\.1\.", "upload_pattern": r"^1\.1\.", "x_tol": 30},
            {"label": "Висновки", "sample_pattern": r"^ВИСНОВКИ", "upload_pattern": r"^ВИСНОВ(ОК|КИ)", "x_tol": 30},
            {"label": "Список використаних джерел", "sample_pattern": r"^СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ", "upload_pattern": r"^(СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ|ДЖЕРЕЛА)", "x_tol": 30},
            {"label": "Номер сторінки для 'ВСТУП'", "sample_pattern": r"^3$", "upload_pattern": r"^\d+$", "x_tol": 25},
            {"label": "Номер сторінки для 'Списку джерел'", "sample_pattern": r"^72$", "upload_pattern": r"^\d+$", "x_tol": 25},
        ],
    }


def count_detection_matches(page, config):
    lines = extract_lines(page)
    return sum(1 for pattern in config["detection_patterns"] if find_best_line(lines, pattern))


def validate_page_against_sample(page, report_key, page_number, config, report, require_detection=True):
    lines = extract_lines(page)
    if require_detection:
        found = sum(1 for pattern in config["detection_patterns"] if find_best_line(lines, pattern))
        if found < config["required_matches"]:
            return False

    specs = build_specs_from_sample(config["sample_file"], config["templates"])
    for spec in specs:
        validate_line(lines, report[report_key], page_number, spec)
    return True


def has_title_page(page):
    page_text = normalize_for_search(page.get_text("text"))
    return "МІНІСТЕРСТВО ОСВІТИ І НАУКИ УКРАЇНИ" in page_text


def validate_title_page(page, work_type, report):
    return validate_page_against_sample(page, "Титульна сторінка", 1, get_title_config(work_type), report, require_detection=False)


def detect_mismatched_work_type(page, selected_work_type):
    best_type = None
    best_found = -1

    for work_type in WORK_OPTIONS:
        if work_type == selected_work_type:
            continue
        config = get_title_config(work_type)
        found = count_detection_matches(page, config)
        if found > best_found:
            best_type = work_type
            best_found = found

    if best_type is None:
        return None

    best_config = get_title_config(best_type)
    if best_found >= best_config["required_matches"]:
        return best_type
    return None


def validate_contents_page(page, report, work_type):
    if work_type in {"Звіт з практики (Бакалавр)", "Звіт з практики (Магістр)"}:
        lines = extract_lines(page)
        return find_best_line(lines, r"^ЗМІСТ$") is not None
    return validate_page_against_sample(page, "Сторінка зі змістом", 2, get_contents_config(), report, require_detection=False)


def has_contents_page(page):
    page_text = normalize_for_search(page.get_text("text"))
    return "ЗМІСТ" in page_text


def analyze_body_pages(doc, report, start_page=2):
    expected_size = 14.0
    in_bibliography = False
    bibliography_entry_parts = []
    bibliography_entry_page = None

    for page_num in range(start_page, len(doc)):
        page = doc[page_num]
        rect = page.rect
        blocks = page.get_text("dict")["blocks"]
        page_lines = [line for line in extract_lines(page) if not is_page_number_line(line, rect)]
        validate_page_number(page, report, page_num + 1)
        table_bboxes = []
        if hasattr(page, "find_tables"):
            tables = page.find_tables()
            table_bboxes = [
                content_bbox
                for table in tables.tables
                if table_looks_real(table, rect)
                for content_bbox in [get_table_content_bbox(table)]
                if content_bbox is not None
            ]
        top_margin_bboxes, body_margin_bboxes, has_large_image = collect_margin_bboxes(blocks, rect, table_bboxes)

        for block in blocks:
            if "lines" not in block:
                continue

            block_lines = []
            for raw_line in block["lines"]:
                spans = [span for span in raw_line["spans"] if span["text"].strip()]
                if not spans:
                    continue
                block_text = "".join(span["text"] for span in spans).strip()
                if not block_text:
                    continue
                block_lines.append(
                    {
                        "text": block_text,
                        "x0": raw_line["bbox"][0],
                        "y0": raw_line["bbox"][1],
                        "x1": raw_line["bbox"][2],
                        "y1": raw_line["bbox"][3],
                    }
                )

            if block_lines and all(is_page_number_line(line, rect) for line in block_lines):
                continue

            bbox = block["bbox"]
            lines = [
                line
                for line in block["lines"]
                if not is_page_number_line(
                    {
                        "text": "".join(span["text"] for span in line["spans"]).strip(),
                        "x1": line["bbox"][2],
                        "y1": line["bbox"][3],
                    },
                    rect,
                )
            ]
            if not lines:
                continue

            full_text = ""
            for line in lines:
                for span in line["spans"]:
                    text_strip = span["text"].strip()
                    full_text += text_strip + " "
            full_text_strip = full_text.strip()

            block_inside_table = any(bboxes_intersect(bbox, table_bbox, padding=4) for table_bbox in table_bboxes)
            is_table_caption_block = full_text_strip.startswith("Таблиця") or full_text_strip.startswith("Кінець таблиці") or full_text_strip.startswith("Продовження таблиці")
            if block_inside_table and not is_table_caption_block:
                continue

            measurable_lines = [line for line in lines if is_measurement_text_line(line)]

            if len(measurable_lines) > 1:
                first_line_x = measurable_lines[0]["bbox"][0]
                second_line_x = measurable_lines[1]["bbox"][0]
                indent_cm = (first_line_x - second_line_x) * PT_TO_CM
                if indent_cm > 0.5 and abs(indent_cm - 1.5) > 0.3:
                    add_page_error(
                        report,
                        "Абзацний відступ (1.5 см)",
                        page_num + 1,
                        f"Відступ ~{round(indent_cm, 2)} см. <i>'{measurable_lines[0]['spans'][0]['text'][:25]}...'</i>",
                    )

                if measurable_lines[0]["spans"] and measurable_lines[1]["spans"]:
                    prev_y = measurable_lines[0]["bbox"][3]
                    curr_y = measurable_lines[1]["bbox"][3]
                    fs = measurable_lines[0]["spans"][0]["size"]
                    if fs > 10:
                        line_ratio = (curr_y - prev_y) / fs
                        if (line_ratio < 1.35 or line_ratio > 1.7) and line_ratio > 0.9:
                            add_page_error(
                                report,
                                "Міжрядковий інтервал (1.5)",
                                page_num + 1,
                                f"Інтервал ~{round(line_ratio, 1)}. <i>'{measurable_lines[0]['spans'][0]['text'][:25]}...'</i>",
                            )

            skip_style_checks = is_figure_like_block(full_text_strip, bbox, rect)

            if not skip_style_checks:
                for line in lines:
                    for span in line["spans"]:
                        text_strip = span["text"].strip()
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
                        normalized_font_size = normalize_font_size(font_size)
                        normalized_expected_size = normalize_font_size(expected_size)
                        if normalized_font_size != normalized_expected_size and font_size >= 10 and not text_strip.isupper():
                            suffix = " (допустимо лише в таблицях)" if 10 <= normalized_font_size <= 12 else ""
                            add_page_error(
                                report,
                                "Розмір шрифту (14)",
                                page_num + 1,
                                f"Розмір {normalized_font_size}{suffix}. <i>'{text_strip[:25]}...'</i>",
                            )

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

            if re.match(r"^(Рисунок|Рис\.)\s*", full_text_strip):
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
                flush_bibliography_entry(report, bibliography_entry_page or (page_num + 1), bibliography_entry_parts)
                bibliography_entry_parts = []
                bibliography_entry_page = None
                in_bibliography = True
            elif full_text_strip == "ДОДАТКИ":
                flush_bibliography_entry(report, bibliography_entry_page or (page_num + 1), bibliography_entry_parts)
                bibliography_entry_parts = []
                bibliography_entry_page = None
                in_bibliography = False
            elif in_bibliography:
                if re.match(r"^\d+\.", full_text_strip):
                    flush_bibliography_entry(report, bibliography_entry_page or (page_num + 1), bibliography_entry_parts)
                    bibliography_entry_parts = [full_text_strip]
                    bibliography_entry_page = page_num + 1
                elif bibliography_entry_parts:
                    bibliography_entry_parts.append(full_text_strip)

        if table_bboxes:
            first_text_line = get_first_text_line(blocks)
            for t_bbox in table_bboxes:
                expected_left = 2.5 / PT_TO_CM
                expected_right = 1.0 / PT_TO_CM
                if t_bbox[0] < expected_left - 5:
                    add_page_error(report, "Межі та розриви таблиць (Не виходять за поля, наявність 'Продовження')", page_num + 1, "Лівий край таблиці перетинає ліве поле 2.5 см")
                if t_bbox[2] > rect.width - expected_right + 5:
                    add_page_error(report, "Межі та розриви таблиць (Не виходять за поля, наявність 'Продовження')", page_num + 1, "Правий край таблиці виходить за межі правого поля 1.0 см")

                header_texts = get_text_blocks_above_table(blocks, t_bbox[1])
                if (
                    first_text_line
                    and first_text_line["text"].startswith("Таблиця")
                    and abs(first_text_line["y0"] - t_bbox[1]) < 40
                    and not header_texts
                ):
                    add_page_error(
                        report,
                        "Межі та розриви таблиць (Не виходять за поля, наявність 'Продовження')",
                        page_num + 1,
                        "У першому рядку сторінки стоїть 'Таблиця ...' без тексту зверху.",
                    )

                header_line = find_table_header_line(page_lines, t_bbox)
                if header_line and is_table_continuation_line(header_line["text"]):
                    continue

                if header_line and caption_has_inline_source(header_line["text"]):
                    continue

                source_line = find_table_source_line(page_lines, t_bbox)
                if source_line and is_valid_table_source_line(source_line["text"]):
                    continue

                if table_continues_on_next_page(doc, page_num, t_bbox):
                    continue

                if source_line:
                    add_page_error(
                        report,
                        TABLE_SOURCE_RULE,
                        page_num + 1,
                        f"Рядок джерела під таблицею знайдено, але формат некоректний: <i>'{normalize_text(source_line['text'])[:90]}...'</i>",
                    )
                else:
                    add_page_error(
                        report,
                        TABLE_SOURCE_RULE,
                        page_num + 1,
                        "Для таблиці не знайдено обов'язкового посилання на джерело ні в підписі, ні під таблицею.",
                    )

        if body_margin_bboxes:
            min_x = min(bbox[0] for bbox in body_margin_bboxes)
            max_x = max(bbox[2] for bbox in body_margin_bboxes)
            left_cm = min_x * PT_TO_CM
            right_cm = (rect.width - max_x) * PT_TO_CM
            if abs(left_cm - 2.5) > 0.35:
                add_page_error(report, "Поля сторінки (Л: 2.5 см, П: 1.0 см, В/Н: 2.0 см)", page_num + 1, f"Ліве поле ~{round(left_cm, 1)} см")
            if abs(right_cm - 1.0) > 0.35:
                add_page_error(report, "Поля сторінки (Л: 2.5 см, П: 1.0 см, В/Н: 2.0 см)", page_num + 1, f"Праве поле ~{round(right_cm, 1)} см")

        if top_margin_bboxes and not (has_large_image and not body_margin_bboxes):
            min_y = min(bbox[1] for bbox in top_margin_bboxes)
            top_cm = min_y * PT_TO_CM
            if abs(top_cm - 2.0) > 0.35 and top_cm > 1.0:
                add_page_error(report, "Поля сторінки (Л: 2.5 см, П: 1.0 см, В/Н: 2.0 см)", page_num + 1, f"Верхнє поле ~{round(top_cm, 1)} см")

    flush_bibliography_entry(report, bibliography_entry_page or start_page + 1, bibliography_entry_parts)


def is_valid_table_source_line(text):
    normalized = normalize_text(text)
    if ":" not in normalized:
        return False

    _, source_body = normalized.split(":", 1)
    source_body = source_body.strip()
    if not source_body:
        return False

    return bool(re.search(r"\[[^\[\]]+\]", source_body))


def table_continues_on_next_page(doc, page_num, table_bbox):
    if page_num + 1 >= len(doc):
        return False

    next_page = doc[page_num + 1]
    first_line = get_first_meaningful_line(next_page)
    if first_line is None:
        return False

    first_text = normalize_text(first_line["text"])
    if not (first_text.startswith("Кінець таблиці") or first_text.startswith("Продовження таблиці")):
        return False

    current_page = doc[page_num]
    bottom_gap = current_page.rect.height - table_bbox[3]
    return bottom_gap <= 95


def analyze_pdf(file_bytes, work_type):
    doc = pymupdf.open(stream=file_bytes, filetype="pdf")
    try:
        report = build_report()
        report.setdefault(TABLE_SOURCE_RULE, [])
        if len(doc) < 2:
            return {"report": report, "stop_message": "Не знайдено титульний лист і сторінку зі змістом."}

        title_present = has_title_page(doc[0])
        contents_present = has_contents_page(doc[1])
        if not title_present or not contents_present:
            if not title_present and not contents_present:
                stop_message = "Не знайдено титульний лист і сторінку зі змістом."
            elif not title_present:
                mismatched_type = detect_mismatched_work_type(doc[0], work_type)
                stop_message = "Обрано невідповідний тип роботи." if mismatched_type else "Не знайдено титульний лист."
            else:
                stop_message = "Не знайдено сторінку зі змістом."
            return {"report": report, "stop_message": stop_message}

        validate_title_page(doc[0], work_type, report)
        validate_contents_page(doc[1], report, work_type)

        analyze_body_pages(doc, report, start_page=2)
        return {"report": truncate_report(report), "stop_message": None}
    finally:
        doc.close()


def run_app():
    st.set_page_config(page_title="Перевірка студентських робіт", page_icon="🎓", layout="centered")
    st.markdown(
        """
<style>
    .stApp {
        background:
            radial-gradient(circle at top left, rgba(196, 224, 255, 0.9), transparent 30%),
            radial-gradient(circle at top right, rgba(255, 226, 196, 0.85), transparent 28%),
            linear-gradient(180deg, #f7f2ea 0%, #eef3f8 100%);
    }
    .main .block-container {
        max-width: 900px;
        padding-top: 2.2rem;
        padding-bottom: 3rem;
    }
    .hero-card {
        padding: 1.6rem 1.7rem;
        border-radius: 28px;
        background: rgba(255, 255, 255, 0.76);
        border: 1px solid rgba(137, 109, 74, 0.14);
        box-shadow: 0 24px 60px rgba(108, 86, 61, 0.12);
        backdrop-filter: blur(10px);
        animation: fadeUp 0.75s ease-out;
    }
    .hero-kicker {
        display: inline-block;
        margin-bottom: 0.8rem;
        padding: 0.35rem 0.7rem;
        border-radius: 999px;
        background: #16324f;
        color: #f6efe6;
        font-size: 0.78rem;
        letter-spacing: 0.08em;
        text-transform: uppercase;
    }
    .hero-title {
        margin: 0;
        color: #14263d;
        font-size: 2.4rem;
        line-height: 1.08;
        font-weight: 800;
    }
    .hero-text {
        margin: 0.9rem 0 0 0;
        color: #4f5e6f;
        font-size: 1.02rem;
        line-height: 1.65;
    }
    .panel-card {
        margin-top: 1rem;
        padding: 1.15rem 1.2rem 0.55rem 1.2rem;
        border-radius: 24px;
        background: rgba(255, 255, 255, 0.82);
        border: 1px solid rgba(137, 109, 74, 0.12);
        box-shadow: 0 18px 40px rgba(108, 86, 61, 0.08);
        animation: fadeUp 0.9s ease-out;
    }
    .panel-note {
        margin: 0 0 0.9rem 0;
        color: #5f6f7f;
        font-size: 0.97rem;
    }
    [data-testid="stSelectbox"] label p {
        color: #16324f !important;
        font-weight: 700;
    }
    [data-testid="stSelectbox"] div[data-baseweb="select"] > div {
        background: #16324f !important;
        border: 1px solid #16324f !important;
        border-radius: 16px !important;
        color: #ffffff !important;
        box-shadow: 0 10px 24px rgba(22, 50, 79, 0.18);
    }
    [data-testid="stSelectbox"] div[data-baseweb="select"] > div span {
        color: #ffffff !important;
    }
    [data-testid="stSelectbox"] svg {
        fill: #ffffff !important;
    }
    [data-testid="stSelectbox"], [data-testid="stFileUploader"] {
        animation: fadeUp 1s ease-out;
    }
    [data-testid="stFileUploader"][aria-disabled="true"] section,
    [data-testid="stFileUploader"] section[disabled] {
        opacity: 0.55;
    }
    [data-testid="stFileUploader"] section,
    [data-testid="stFileUploader"] section * {
        color: #16324f !important;
    }
    [data-testid="stFileUploader"] small {
        color: #5f6f7f !important;
    }
    [data-testid="stExpander"] details summary p,
    [data-testid="stAlert"] p {
        color: #16324f !important;
    }
    [data-testid="stExpander"] details summary p span { position: absolute; right: 45px; }
    [data-testid="stAlert"] p span { position: absolute; right: 20px; }
    [data-testid="stExpander"] details summary p, [data-testid="stAlert"] p { padding-right: 120px; }
    [data-testid="stExpander"] {
        border-radius: 18px;
        overflow: hidden;
        border: 1px solid rgba(137, 109, 74, 0.12);
        box-shadow: 0 12px 26px rgba(108, 86, 61, 0.07);
        background: rgba(255, 255, 255, 0.84);
    }
    [data-testid="stFileUploader"] section {
        border-radius: 18px;
        border: 1.5px dashed rgba(22, 50, 79, 0.28);
        background: rgba(247, 242, 234, 0.72);
    }
    @keyframes fadeUp {
        from {
            opacity: 0;
            transform: translateY(16px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }
</style>
""",
        unsafe_allow_html=True,
    )

    st.markdown(
        """
<div class="hero-card">
    <div class="hero-kicker">PDF Inspector</div>
    <h1 class="hero-title">Інспектор оформлення студентських робіт</h1>
    <p class="hero-text">
        Застосунок перевіряє титульну сторінку, зміст і основний текст PDF-роботи
        за правилами оформлення та показує всі знайдені помилки по сторінках.
    </p>
</div>
""",
        unsafe_allow_html=True,
    )
    st.markdown(
        """
<div class="panel-card">
    <p class="panel-note">Спочатку оберіть напрям роботи, потім завантажте PDF з титульною сторінкою та змістом.</p>
</div>
""",
        unsafe_allow_html=True,
    )

    work_type = st.selectbox(
        "Оберіть напрям вашої роботи",
        WORK_OPTIONS,
        index=None,
        placeholder="Оберіть напрям зі списку",
    )
    st.info("Завантажуйте будь ласка роботу обов'язково з титульною сторінкою та змістом")
    uploaded_file = st.file_uploader(
        "Завантажте PDF-файл вашої роботи",
        type=["pdf"],
        disabled=work_type is None,
    )

    if work_type is None:
        st.caption("Спочатку оберіть вид роботи, після цього стане доступним завантаження PDF.")
        return

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
