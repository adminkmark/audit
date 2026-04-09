import streamlit as st

try:
    from app_impl import run_app
except ModuleNotFoundError:
    run_app = None

if run_app is not None:
    run_app()
    st.stop()

import streamlit as st
import pymupdf
import re

# Конвертація поінтів у сантиметри (1 см = 28.346 pt)
PT_TO_CM = 1 / 28.346

def analyze_pdf(file_bytes):
    doc = pymupdf.open(stream=file_bytes, filetype="pdf")
    
    # Критерії з порожніми списками для збору помилок
    report = {
        "Поля сторінки (Л: 2.5см, П: 1.0см, В/Н: 2.0см)": [],
        "Шрифт (Times New Roman)": [],
        "Розмір шрифту (14)": [],
        "Абзацний відступ (1.5 см)": [],
        "Міжрядковий інтервал (1.5)": [],
        "Оформлення заголовків (ВЕЛИКІ ЛІТЕРИ, по центру)": [],
        "Оформлення підрозділів (Відступ 1.5 см, без крапки в кінці)": [],
        "Підписи до рисунків (Формат: Рисунок X.X – Назва)": [],
        "Підписи до таблиць (Формат: Таблиця X.X – Назва)": [],
        "Межі та розриви таблиць (Не виходять за поля, наявність 'Продовження')": [],
        "Оформлення формул (Номер праворуч у дужках)": [],
        "Список використаних джерел (ДСТУ 8302:2015)": []
    }
    
    expected_size = 14.0
    in_bibliography = False
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        rect = page.rect
        blocks = page.get_text("dict")["blocks"]

        # Знаходження меж тексту на сторінці для розрахунку полів
        min_x = rect.width
        max_x = 0
        min_y = rect.height
        max_y = 0
        
        for b in blocks:
            if "lines" not in b: continue
            
            bbox = b["bbox"]
            if bbox[0] < min_x: min_x = bbox[0]
            if bbox[1] < min_y: min_y = bbox[1]
            if bbox[2] > max_x: max_x = bbox[2]
            if bbox[3] > max_y: max_y = bbox[3]
            
            lines = b["lines"]
            
            # --- Аналіз абзацу та інтервалу ---
            if len(lines) > 1:
                # Відступ першого рядка
                first_line_x = lines[0]["bbox"][0]
                second_line_x = lines[1]["bbox"][0]
                indent_cm = (first_line_x - second_line_x) * PT_TO_CM
                
                # Якщо це потенційно абзац (значний відступ)
                if indent_cm > 0.5:
                    if abs(indent_cm - 1.5) > 0.3:
                         report["Абзацний відступ (1.5 см)"].append(f"<b>Сторінка {page_num+1}</b>: Відступ ~{round(indent_cm, 2)} см. <i>'{lines[0]['spans'][0]['text'][:25]}...'</i>")
                
                # Міжрядковий інтервал (відстань між низом рядків поділено на розмір шрифту)
                y1_prev = lines[0]["bbox"][3]
                y1_curr = lines[1]["bbox"][3]
                dist_pt = y1_curr - y1_prev
                
                if lines[0]["spans"] and lines[1]["spans"]:
                    fs = lines[0]["spans"][0]["size"]
                    if fs > 10:
                        line_ratio = dist_pt / fs
                        # Допустимі межі інтервалу "1.5" в Word (зазвичай 1.35 - 1.6)
                        if line_ratio < 1.35 or line_ratio > 1.7:
                             if line_ratio > 0.9: # Відкидаємо глюки парсингу
                                 report["Міжрядковий інтервал (1.5)"].append(f"<b>Сторінка {page_num+1}</b>: Інтервал ~{round(line_ratio, 1)}. <i>'{lines[0]['spans'][0]['text'][:25]}...'</i>")

            # --- Аналіз шрифтів і ключових слів ---
            full_text = ""
            for l in lines:
                for s in l["spans"]:
                    text = s["text"]
                    text_strip = text.strip()
                    full_text += text_strip + " "
                    
                    if len(text_strip) > 5:
                        font_name = s["font"]
                        font_size = s["size"]
                        
                        if "Times" not in font_name and "Symbol" not in font_name:
                            report["Шрифт (Times New Roman)"].append(f"<b>Сторінка {page_num+1}</b>: <code>{font_name}</code>. <i>'{text_strip[:25]}...'</i>")
                            
                        # Ігноруємо розмір для заголовків (вони можуть бути більшими за 14)
                        if abs(font_size - expected_size) > 0.5 and font_size >= 10 and not text_strip.isupper():
                            if 10 <= font_size <= 12.5:
                                report["Розмір шрифту (14)"].append(f"<b>Сторінка {page_num+1}</b>: Розмір {round(font_size, 1)} (Допустимо <i>тільки</i> в таблицях). <i>'{text_strip[:25]}...'</i>")
                            else:
                                report["Розмір шрифту (14)"].append(f"<b>Сторінка {page_num+1}</b>: Розмір {round(font_size, 1)}. <i>'{text_strip[:25]}...'</i>")
                            
            full_text_strip = full_text.strip()
            
            # --- Заголовки (ЗМІСТ, ВСТУП, РОЗДІЛ тощо) ---
            if full_text_strip in ["ЗМІСТ", "ВСТУП", "ВИСНОВКИ", "СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ", "ДОДАТКИ"] or full_text_strip.startswith("РОЗДІЛ"):
                # Перевірка на крапку після РОЗДІЛ X.
                if full_text_strip.startswith("РОЗДІЛ") and re.search(r"РОЗДІЛ\s+\d+\.", full_text_strip):
                     report["Оформлення заголовків (ВЕЛИКІ ЛІТЕРИ, по центру)"].append(f"<b>Сторінка {page_num+1}</b>: Заборонена крапка після номера розділу: '{full_text_strip}'")
                
                # Перевірка великих літер
                if not full_text_strip.isupper():
                     report["Оформлення заголовків (ВЕЛИКІ ЛІТЕРИ, по центру)"].append(f"<b>Сторінка {page_num+1}</b>: Має бути великими літерами: '{full_text_strip}'")
                if full_text_strip.endswith("."):
                     report["Оформлення заголовків (ВЕЛИКІ ЛІТЕРИ, по центру)"].append(f"<b>Сторінка {page_num+1}</b>: Заголовок не повинен мати крапку в кінці: '{full_text_strip}'")
                     
                # Центрування (враховуючи, що ліве поле 2.5 см, а праве 1.0 см)
                expected_left_margin_pt = 2.5 / PT_TO_CM
                expected_right_margin_pt = 1.0 / PT_TO_CM
                
                space_left = bbox[0] - expected_left_margin_pt
                space_right = (rect.width - expected_right_margin_pt) - bbox[2]
                
                if abs(space_left - space_right) > 35: # 35 pt похибка
                    report["Оформлення заголовків (ВЕЛИКІ ЛІТЕРИ, по центру)"].append(f"<b>Сторінка {page_num+1}</b>: Заголовок не центрований: '{full_text_strip}'")
            
            # --- Підрозділи (напр. "1.1 Назва...") ---
            if re.match(r"^\d+\.\d+\s+[А-ЯІЄЇ]", full_text_strip):
                # Перевірка на відсутність крапки після номера - це задовольняється регекспом (\s+ відразу після \d+),
                # тому перевіримо чи немає крапки в кінці:
                if full_text_strip.endswith("."):
                    report["Оформлення підрозділів (Відступ 1.5 см, без крапки в кінці)"].append(f"<b>Сторінка {page_num+1}</b>: Заборонена крапка в кінці заголовка: '{full_text_strip}'")
                
                # Відступ підрозділу (має бути 1.5 см від лівого поля)
                expected_left_margin_pt =  2.5 / PT_TO_CM
                # Абзацний відступ = 1.5 см = 1.5 / PT_TO_CM
                abs_indent = bbox[0] - expected_left_margin_pt
                if abs(abs_indent - (1.5 / PT_TO_CM)) > 15: # 15 pt похибка
                    indent_cm = abs_indent * PT_TO_CM
                    report["Оформлення підрозділів (Відступ 1.5 см, без крапки в кінці)"].append(f"<b>Сторінка {page_num+1}</b>: Відступ підрозділу ~{round(indent_cm, 1)} см (має бути 1.5 см): '{full_text_strip}'")


            # --- Рисунки ---
            if full_text_strip.startswith("Рисунок"):
                # Перевіряємо наявність тире та формату номера
                if not re.search(r"Рисунок\s+(\d+\.\d+|[А-Я]\.\d+)\s+[–-]\s+", full_text_strip):
                     if len(full_text_strip) < 150: 
                         report["Підписи до рисунків (Формат: Рисунок X.X – Назва)"].append(f"<b>Сторінка {page_num+1}</b>: Неправильний формат: <i>'{full_text_strip[:45]}...'</i>")
                
                # Відступ підпису до рисунку (має бути 1.5 см від лівого поля)
                expected_left_margin_pt =  2.5 / PT_TO_CM
                abs_indent = bbox[0] - expected_left_margin_pt
                if abs(abs_indent - (1.5 / PT_TO_CM)) > 15: 
                    report["Підписи до рисунків (Формат: Рисунок X.X – Назва)"].append(f"<b>Сторінка {page_num+1}</b>: Немає абзацного відступу 1.5 см: <i>'{full_text_strip[:45]}...'</i>")

            # --- Таблиці ---
            if full_text_strip.startswith("Таблиця"):
                if not re.search(r"Таблиця\s+(\d+\.\d+|[А-Я]\.\d+)\s+[–-]\s+", full_text_strip) and "Продовження" not in full_text_strip and "Кінець" not in full_text_strip:
                    if len(full_text_strip) < 150:
                         report["Підписи до таблиць (Формат: Таблиця X.X – Назва)"].append(f"<b>Сторінка {page_num+1}</b>: Неправильний формат: <i>'{full_text_strip[:45]}...'</i>")
                         
                # Відступ підпису до таблиці (має бути 1.5 см)
                expected_left_margin_pt =  2.5 / PT_TO_CM
                abs_indent = bbox[0] - expected_left_margin_pt
                if abs(abs_indent - (1.5 / PT_TO_CM)) > 15 and "Продовження" not in full_text_strip and "Кінець" not in full_text_strip: 
                    report["Підписи до таблиць (Формат: Таблиця X.X – Назва)"].append(f"<b>Сторінка {page_num+1}</b>: Немає абзацного відступу 1.5 см: <i>'{full_text_strip[:45]}...'</i>")

            # --- Формули ---
            if re.search(r"\(\d+\.\d+\)$", full_text_strip) and len(full_text_strip) < 100:
                expected_right_margin_pt = 1.0 / PT_TO_CM
                space_right = (rect.width - expected_right_margin_pt) - bbox[2]
                if space_right > 35: # номер має бути притиснутий до правого краю
                    report["Оформлення формул (Номер праворуч у дужках)"].append(f"<b>Сторінка {page_num+1}</b>: Номер формули не притиснуто до правого краю: <i>'{full_text_strip[-15:]}'</i>")

            if full_text_strip.startswith("де ") and len(full_text_strip) < 150:
                 if "-" not in full_text_strip and "–" not in full_text_strip:
                      report["Оформлення формул (Номер праворуч у дужках)"].append(f"<b>Сторінка {page_num+1}</b>: В розшифровці 'де' пропущено тире між символом та описом: <i>'{full_text_strip[:45]}...'</i>")

            # --- Список джерел ---
            if full_text_strip == "СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ":
                in_bibliography = True
            elif full_text_strip == "ДОДАТКИ":
                in_bibliography = False
                
            if in_bibliography and re.match(r"^\d+\.", full_text_strip):
                # Дуже базова перевірка (принаймні на рік, який зазвичай з 4 цифр)
                if not re.search(r"20\d\d", full_text_strip) and not re.search(r"19\d\d", full_text_strip):
                     if "URL" not in full_text_strip and "http" not in full_text_strip: # Якщо це не суто інтернет посилання без року
                         report["Список використаних джерел (ДСТУ 8302:2015)"].append(f"<b>Сторінка {page_num+1}</b>: Можлива помилка ДСТУ (не знайдено року видання): <i>'{full_text_strip[:50]}...'</i>")

        # --- Аналіз фізичних таблиць (сітки) ---
        if hasattr(page, 'find_tables'):
            tables = page.find_tables()
            for tab in tables.tables:
                t_bbox = tab.bbox
                # Перевірка на вихід за межі полів
                expected_left = 2.5 / PT_TO_CM
                expected_right = 1.0 / PT_TO_CM
                # Допуск 5 pt
                if t_bbox[0] < expected_left - 5:
                    report["Межі та розриви таблиць (Не виходять за поля, наявність 'Продовження')"].append(f"<b>Сторінка {page_num+1}</b>: Лівий край таблиці перетинає ліве поле 2.5 см")
                if t_bbox[2] > rect.width - expected_right + 5:
                    report["Межі та розриви таблиць (Не виходять за поля, наявність 'Продовження')"].append(f"<b>Сторінка {page_num+1}</b>: Правий край таблиці виходить за межі правого поля 1.0 см")
                
                # Перевірка наявності "Продовження" якщо таблиця розірвана
                has_header = False
                for b in blocks:
                    # Шукаємо блоки тексту, які вище або злегка на рівні з верхньою межею таблиці
                    if "lines" in b and b["bbox"][3] <= t_bbox[1] + 10: 
                        text = "".join([s["text"] for l in b["lines"] for s in l["spans"]]).strip()
                        if "Таблиця" in text or "Продовження" in text or "Кінець" in text:
                            has_header = True
                            break
                            
                # Якщо таблиця досить високо на сторінці і над нею немає відповідних підписів
                if not has_header and t_bbox[1] < (2.0 / PT_TO_CM) + 150:
                    report["Межі та розриви таблиць (Не виходять за поля, наявність 'Продовження')"].append(f"<b>Сторінка {page_num+1}</b>: Розірвана таблиця без обов'язкового підпису 'Продовження таблиці...' або 'Кінець таблиці...' зверху")

        # --- Перевірка полів ---
        if min_x < rect.width and max_x > 0 and page_num > 0: # пропускаємо титулку (індекс 0)
            left_cm = min_x * PT_TO_CM
            right_cm = (rect.width - max_x) * PT_TO_CM
            top_cm = min_y * PT_TO_CM
            bottom_cm = (rect.height - max_y) * PT_TO_CM
            
            # Допустима похибка 0.35 см
            if abs(left_cm - 2.5) > 0.35:
                report["Поля сторінки (Л: 2.5см, П: 1.0см, В/Н: 2.0см)"].append(f"<b>Сторінка {page_num+1}</b>: Ліве поле ~{round(left_cm, 1)} см")
            if abs(right_cm - 1.0) > 0.35:
                report["Поля сторінки (Л: 2.5см, П: 1.0см, В/Н: 2.0см)"].append(f"<b>Сторінка {page_num+1}</b>: Праве ~{round(right_cm, 1)} см")
            if abs(top_cm - 2.0) > 0.35 and top_cm > 1.0:
                report["Поля сторінки (Л: 2.5см, П: 1.0см, В/Н: 2.0см)"].append(f"<b>Сторінка {page_num+1}</b>: Верхнє ~{round(top_cm, 1)} см")

    # Збираємо унікальні помилки і обрізаємо до топ-8 для читабельності
    for rule in list(report.keys()):
        # Використовуємо dict.fromkeys для видалення дублікатів зі збереженням порядку
        unique_errors = list(dict.fromkeys(report[rule]))
        if len(unique_errors) > 8:
            report[rule] = unique_errors[:8] + [f"...та ще <b>{len(unique_errors)-8} схожих помилок</b> у документі."]
        else:
            report[rule] = unique_errors

    return report


# ========== STREAMLIT UI ==========

st.set_page_config(page_title="Перевірка студентських робіт", page_icon="🎓", layout="centered")

st.markdown("""
<style>
    /* Фіксуємо текст РОЗГОРНУТИ / ІДЕАЛЬНО строго з правої сторони */
    [data-testid="stExpander"] details summary p span {
        position: absolute;
        right: 45px;
    }
    [data-testid="stAlert"] p span {
        position: absolute;
        right: 20px;
    }
    /* Забороняємо переносу тексту обрізатися під абсолютним елементом */
    [data-testid="stExpander"] details summary p, [data-testid="stAlert"] p {
        padding-right: 120px; 
    }
</style>
""", unsafe_allow_html=True)

st.title("🎓 Інспектор оформлення (PDF)")
st.write("Скрипт перевіряє документ за всіма загальними правилами методички КНЕУ (підходить для курсових, звітів з практики, дипломних тощо).")

uploaded_file = st.file_uploader("Завантажте PDF-файл вашої роботи", type=["pdf"])

if uploaded_file is not None:
    bytes_data = uploaded_file.read()
    
    with st.spinner('Сканую документ...'):
        report = analyze_pdf(bytes_data)
        
    st.markdown("---")
    st.subheader("📊 Результати перевірки:")
    
    all_clear = True
    
    # Виводимо кожне правило окремо з його статусом
    for rule, errors in report.items():
        if len(errors) == 0:
            st.success(f"✅ **{rule}**  |  Помилок: **0** :green[**ІДЕАЛЬНО**]")
        else:
            all_clear = False
            # За замовчуванням згорнуто (expanded=False)
            with st.expander(f"❌ **{rule}**  |  Помилок: **{len(errors)}** :red[**РОЗГОРНУТИ**]", expanded=False):
                for err in errors:
                    st.markdown(f"- {err}", unsafe_allow_html=True)
                    
    st.markdown("---")
    if all_clear:
        st.balloons()
        st.success("🏆 Блискуче! Цей документ виглядає ідеально з точки зору форматування.")
    else:
        st.warning("⚠️ Будь ласка, виправте виділені зауваження та завантажте файл знову.")
