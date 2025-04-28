import tkinter
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import os

# === НАСТРОЙКИ ===

A4_WIDTH_MM, A4_HEIGHT_MM = 210, 297       # Размер листа A4 в мм
PAGE_SCALING = 10                          # Масштабирование страницы
PAGE_WIDTH = A4_WIDTH_MM * PAGE_SCALING    # Ширина листа
PAGE_HEIGHT = A4_HEIGHT_MM * PAGE_SCALING  # Высота листа

FONT_NAME = "Arial Bold"                   # Название основного шрифта
FONT_FILE_NAME = "arialbd.ttf"             # Название файла основного шрифта
FONT_SIZE_TITLE = 32                       # Размер шрифта для названия товара
FONT_SIZE_PRICE = 120                      # Размер шрифта для цены товара
WIDTH_TAG_FRAME = 4                        # Ширина рамки ценника

# === ФУНКЦИИ ===

def choose_file():
    filepath = filedialog.askopenfilename(
        title="Выберите Excel файл",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if filepath:
        # Скрытие стартового экрана
        start_screen_hide()
        root.update()
        # Показ экрана прогресса
        progress_screen_show()
        root.update()

        # Основной алгоритм генерации страниц с ценниками
        generate_tags(filepath)

        # Скрытие экрана прогресса
        progress_screen_hide()
        root.update()
        # Показываем кнопку
        start_screen_show()
        root.update()

def draw_wrapped_title_text(draw, text, max_width, max_height, start_x, start_y, margin, line_spacing=5):
    # Учёт отступов
    max_width -= 2 * margin
    max_height -= 2 * margin
    start_x += margin
    start_y += margin

    # Шрифт для названия товара
    font = ImageFont.truetype(FONT_FILE_NAME, FONT_SIZE_TITLE)

    # Линии/строки текста
    lines = []         # Список строк
    current_line = ""  # Текущая строка

    # Позиция по Y относительно ценника
    y = 0

    # Генерация списка строк текста
    for idx, word in enumerate(text.split()):
        # Тестовая версия строки
        test_line = current_line + (" " if current_line else "") + word

        # Ширина и высота получившейся строки
        bbox = draw.textbbox((0, 0), test_line, font=font)
        line_width = bbox[2] - bbox[0]
        line_height = bbox[3] - bbox[1]

        # Продвижение позиции по Y для проверки
        y += line_height

        # Если строка вышла за рамки по высоте
        if y > max_height and lines:
            # Последняя строка
            last_line = lines[-1]

            # Ширина последней строки с троеточием
            bbox = draw.textbbox((0, 0), last_line + "...", font=font)
            last_line_width = bbox[2] - bbox[0]

            # Если ширина последней строки с троеточием подходящая
            if last_line_width <= max_width:
                # Обновляем прошлую строку
                lines[-1] = last_line + "..."
            else:
                # Обновляем прошлую строку
                lines[-1] = (" ".join(lines[-1].split()[0:-1]) if lines[-1] else "") + "..."

            # Останавливаем генерацию списка
            break

        # Если строка не вышла за рамки по ширин
        if line_width <= max_width and idx < len(text.split()) - 1:
            # Обновляем текущую строку
            current_line = test_line
            # Откатываем продвижение позиции по Y
            y -= line_height
        else:
            # Добавляем текущую версию строки
            if current_line:
                lines.append(current_line)

            # Записываем не поместившееся слово на новую строку
            current_line = word

            # Продвижение позиции по Y
            y += line_spacing

    # Обнуление позиции по Y относительно страницы

    y = start_y

    # Рисование строк текста
    for line in lines:
        # Расчёт ширины и высоты строки
        bbox = draw.textbbox((0, 0), line, font=font)
        # line_width = bbox[2] - bbox[0] # Пока не нужно
        line_height = bbox[3] - bbox[1]

        # Позиция по X
        # x = start_x + (max_width - line_width) / 2 # Центрирование
        # x = start_x + (max_width - line_width) # Выравнивание по правому краю
        x = start_x # Выравнивание по левому краю

        # Рисуем строку на изображении
        draw.text((x, y), line, fill="black", font=font)

        # Продвижение позиции по Y
        y += line_height + line_spacing

# def draw_price_text(draw, price_text, max_width, max_height, start_x, start_y, margin): # ⚠️ ₽ и шт
#     # Учёт отступов
#     max_width -= 2 * margin
#     max_height -= 2 * margin
#     start_x += margin
#     start_y += margin
#
#     # Базовый шрифт
#     font_size = FONT_SIZE_PRICE
#     font = ImageFont.truetype(FONT_FILE_NAME, font_size)
#
#     # Проверка, помещается ли цена
#     bbox = draw.textbbox((0, 0), price_text, font=font)
#     w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
#
#     # Если не помещается — уменьшаем шрифт
#     while (w > max_width or h > max_height) and font_size > 10:
#         font_size -= 2  # Плавно уменьшаем размер
#         font = ImageFont.truetype(FONT_FILE_NAME, font_size)
#         bbox = draw.textbbox((0, 0), price_text, font=font)
#         w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
#
#     # Рисуем цену
#     draw.text(
#         (start_x + (max_width - w) / 2, start_y + (max_height - h) / 2),
#         price_text,
#         fill="black",
#         font=font
#     )

def generate_tags(filepath):
    try:
        # Чтение Excel файла
        df = pd.read_excel(filepath, usecols=[0, 2, 6])
        items = df.values.tolist()

        # Список страниц
        pages = []
        # Создание первой страницы
        page = Image.new("RGB", (PAGE_WIDTH, PAGE_HEIGHT), "white")
        draw = ImageDraw.Draw(page)

        # Получение данных
        tags_per_row, tags_per_col, margin = get_user_settings()

        # Размеры
        max_tags_per_page = tags_per_row * tags_per_col              # Максимальное количество ценников на страницу
        page_margin = margin * PAGE_SCALING                          # Размер отступа
        tag_width = (PAGE_WIDTH - 2 * page_margin) / tags_per_row    # Ширина ценника
        tag_height = (PAGE_HEIGHT - 2 * page_margin) / tags_per_col  # Длина ценника

        # Шрифты
        # font_price = ImageFont.truetype(FONT_FILE_NAME, FONT_SIZE_PRICE)  # Шрифт для цены товара

        # Прочее
        current_tag = 0           # Количество записанных ценников
        total_items = len(items)  # Количество товаров в таблице

        for idx, (name, unit, price) in enumerate(items):
            # Расположение ценника
            col = current_tag % tags_per_row                    # Столбец
            row = (current_tag // tags_per_row) % tags_per_col  # Ряд

            # Координатное расположение ценника
            x = page_margin + col * tag_width
            y = page_margin + row * tag_height

            # Рамка ценника
            draw.rectangle([x, y, x + tag_width, y + tag_height], outline="black", width=WIDTH_TAG_FRAME)
            # Дополнительная рамка ценников для одинаковой ширины всех линий
            draw.rectangle(
                [x - WIDTH_TAG_FRAME, y - WIDTH_TAG_FRAME,
                 x + tag_width + WIDTH_TAG_FRAME, y + tag_height + WIDTH_TAG_FRAME],
                outline="black", width=WIDTH_TAG_FRAME
            )

            # -=-=-=-=-=-=-=-=-=-=-=-=-=-=- Отрисовка названия товара -=-=-=-=-=-=-=-=-=-=-=-=-=-=-
            draw_wrapped_title_text(
                draw,            #
                str(name),       # Название товара
                tag_width,       # Максимальная ширина под название
                tag_height / 2,  # Половина высоты ценника под название
                x,               # Начальная позиция X
                y,               # Начальная позиция Y
                10               # Отступ от краёв ценника
            )

            # -=-=-=-=-=-=-=-=-=-=-=-=-=-=- Отрисовка цены товара -=-=-=-=-=-=-=-=-=-=-=-=-=-=-
            # draw_price_text(
            #     draw,                #
            #     str(price),          # Цена товара
            #     tag_width,           # Максимальная ширина под цену
            #     tag_height / 2,      # Половина высоты ценника под цену
            #     x,                   # Начальная позиция X
            #     y + tag_height / 2,  # Начальная позиция Y
            #     10                   # Отступ от краёв ценника
            # )

            # +1 созданный ценник
            current_tag += 1

            # Проверка заполненности страницы
            if current_tag % max_tags_per_page == 0:
                # Сохранение страницы в итоговый список
                pages.append(page)
                # Создание следующей страницы
                page = Image.new("RGB", (PAGE_WIDTH, PAGE_HEIGHT), "white")
                draw = ImageDraw.Draw(page)

            # Обновляем прогресс
            progress = int((idx + 1) / total_items * 100)
            progress_label.config(text=f"Прогресс: {progress}%")
            root.update_idletasks()

        # Добавляем последнюю страницу
        if current_tag % max_tags_per_page != 0:
            # Сохранение страницы в итоговый список
            pages.append(page)

        # Спросить папку для сохранения
        save_folder = filedialog.askdirectory(title="Выберите папку для сохранения")
        if not save_folder:
            messagebox.showwarning("Отмена", "Сохранение отменено.")
            return

        # Сохранение страниц в виде png файла
        for i, page in enumerate(pages, start=1):
            page.save(os.path.join(save_folder, f"Ценники {i}.png"))

        messagebox.showinfo("Успех", f"Сохранено {len(pages)} страниц!") # Сообщение об успехе
    except Exception as e:
        messagebox.showerror("Ошибка", str(e)) # Сообщение об ошибке

def get_user_settings():
    return (
        tags_per_row_var.get(),
        tags_per_col_var.get(),
        margin_var.get()
    )

def start_screen_show():
    btn.pack(pady=(20, 0))

    rows_label.pack(side="left", padx=(0, 0))
    rows_spinbox.pack(side="left", padx=(0, 30))

    columns_label.pack(side="left", padx=(0, 0))
    columns_spinbox.pack(side="left", padx=(0, 30))

    margin_label.pack(side="left", padx=(0, 0))
    margin_combo.pack(side="left", padx=(0, 0))

def start_screen_hide():
    btn.pack_forget()
    rows_spinbox.pack_forget()
    rows_label.pack_forget()
    columns_spinbox.pack_forget()
    columns_label.pack_forget()
    margin_combo.pack_forget()
    margin_label.pack_forget()

def progress_screen_show():
    progress_label.pack(anchor="center")

def progress_screen_hide():
    progress_label.pack_forget()

# === GUI ===

# Основное окно приложения
root = tkinter.Tk()
root.title("Генератор ценников для магазина Алмаз")
root.geometry("600x140")
root.configure(bg="#f0f0f0")
root.iconbitmap('Almaz.ico')
root.resizable(False, False)

# Общие настройки кнопок и полей
min_rows_and_columns = 1
max_rows_and_columns = 16
margin_options = [0, 5, 10]
width_of_boxes = len(str(max_rows_and_columns))
main_font = (FONT_NAME, 16, "bold")
progress_label_font = (FONT_NAME, 20, "bold")
font_of_boxes = (FONT_NAME, 12, "bold")
font_of_labels_of_boxes = (FONT_NAME, 10)

# Кнопка для выбора таблицы товаров
btn = tkinter.Button(
    root, text="Выбрать Excel таблицу с товарами", command=choose_file, font=main_font, bg="black", fg="white")

# Данные от пользователя
tags_per_row_var = tkinter.IntVar(value=4)  # Количество ценников на строку
tags_per_col_var = tkinter.IntVar(value=4)  # Количество ценников на столбец
margin_var = tkinter.IntVar(value=10)       # Размер отступа в мм

# Фрейм для спинбоксов и комбобокса
fields_frame = tkinter.Frame(root)
fields_frame.pack(pady=(20, 0))

# Поле выбора количества ценников в строку
rows_label = tkinter.Label(fields_frame, text="Количество строк:", font=font_of_labels_of_boxes)
rows_spinbox = tkinter.Spinbox(
    fields_frame, from_=min_rows_and_columns, to=max_rows_and_columns, textvariable=tags_per_row_var, state="readonly",
    font=font_of_boxes, width=width_of_boxes)

# Поле выбора количества ценников в столбец
columns_label = tkinter.Label(fields_frame, text="Количество столбцов:", font=font_of_labels_of_boxes)
columns_spinbox = tkinter.Spinbox(
    fields_frame, from_=min_rows_and_columns, to=max_rows_and_columns, textvariable=tags_per_col_var, state="readonly",
    font=font_of_boxes, width=width_of_boxes)

# Поле выбора отступа в мм
margin_label = tkinter.Label(fields_frame, text="Отступы (мм):", font=font_of_labels_of_boxes)
margin_combo = ttk.Combobox(
    fields_frame, textvariable=margin_var, values=margin_options, state="readonly",
    font=font_of_boxes, width=width_of_boxes)

# Добавление стартового экрана
start_screen_show()

# Надпись прогресса
progress_label = tkinter.Label(root, text="Прогресс: 0%", font=progress_label_font)

# Зацикливание
root.mainloop()
