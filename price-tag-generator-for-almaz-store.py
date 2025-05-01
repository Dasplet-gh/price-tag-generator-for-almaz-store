import tkinter
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import threading
import sys
import os
# Для старого PyInstaller
import openpyxl

"""
# === Version 1.4 (stable version) === #
"""

# === ОСНОВНЫЕ НАСТРОЙКИ ===

# Окно
WINDOW_NAME = "Генератор ценников для магазина Алмаз"  # Название окна
WINDOW_SIZE = "800x140"                                # Фиксированный размер окна
WINDOW_ICO = "almaz.ico"                               # Путь к иконке окна

# Настройки чтения таблицы
COLUMN_NUMBERS_IN_PRODUCT_TABLE = [0, 2, 6]            # Номера нужных столбцов в таблице товаров

# Размеры страниц в пикселях
A4_WIDTH_MM, A4_HEIGHT_MM = 210, 297                   # Размер листа A4 в мм
PAGE_SCALING = 20                                      # Масштабирование страницы
PAGE_WIDTH = A4_WIDTH_MM * PAGE_SCALING                # Ширина листа
PAGE_HEIGHT = A4_HEIGHT_MM * PAGE_SCALING              # Высота листа

# Названия основного шрифта
FONT_NAME = "Arial Bold"                               # Название основного шрифта
FONT_FILE_NAME = "arial_bold.ttf"                      # Название файла основного шрифта

# Отступы и промежутки в тексте ценника в %
MARGIN_TITLE_PART_PERC = 0.05                          # Отступ в % от краёв ценника в именной части
MARGIN_PRICE_PART_PERC = 0.05                          # Отступ в % от краёв ценника в ценовой части
LINE_SPACING_PROC = 0.15                               # Промежуток между строками названия товара в %
SPACING_PRICE_PART_PERC = 0.10                         # Промежуток между ценой и правым блоком с копейками в %

# Размер шрифтов в % от шрифта цены
FONT_SIZE_PERC_FRAC_AND_RUB = 0.45                     # Размер шрифта копеек в % от шрифта цены
FONT_SIZE_PERC_UNIT = 0.25                             # Размер шрифта единиц измерения в % от шрифта цены

# === НАСТРОЙКИ КНОПОК И ПОЛЕЙ ===

# Настройки для данных полей
FIELD_OPTIONS = {
    "SPINBOX_ROWS_AND_COLS": {"MIN": 1, "MAX": 12},    # Минимальное и максимальное количество строк и столбцов
    "SPINBOX_FRAME_WIDTH": {"MIN": 0, "MAX": 16},      # Минимальный и максимальный размер рамки ценника
    "COMBOBOX_MARGIN_OPTIONS": [0, 5, 10, 15, 20]      # Варианты отступов
}

# Шрифты для кнопок и полей
FONT_STYLES = {
    "BUTTON": (FONT_NAME, 16, "bold"),                 # Основной шрифт кнопок
    "BOX_LABEL": (FONT_NAME, 10),                      # Шрифт названий полей с пользовательскими данными
    "BOX": (FONT_NAME, 12, "bold"),                    # Шрифт внутри полей с пользовательскими данными
    "PROGRESS": (FONT_NAME, 20, "bold"),               # Шрифт надписи прогресса
}

# === БАЗОВЫЕ РАЗМЕРЫ ЭТАЛОННОГО ЦЕННИКА ===

# Предустановленные базовые размеры
BASE_CONFIG = {
    "FONT_SIZE_TITLE": 4,                              # Базовый размер шрифта именной части ценника
    "FONT_SIZE_PRICE": 18,                             # Базовый размер шрифта ценовой части ценника
    "ROWS": 4,                                         # Базовое количество строк
    "COLUMNS": 4,                                      # Базовое количество столбцов
    "WIDTH_FRAME_TAG": 1,                              # Базовый размер рамки ценника
    "PAGE_MARGIN": 5,                                  # Базовый размер отступа в мм
}

# Расчёт базовых размеров эталонного ценника при базовых настройках
BASE_TAG_4X4_WIDTH = (PAGE_WIDTH - 2 * BASE_CONFIG["PAGE_MARGIN"] * PAGE_SCALING) // BASE_CONFIG["ROWS"]
BASE_TAG_4X4_HEIGHT = (PAGE_HEIGHT - 2 * BASE_CONFIG["PAGE_MARGIN"] * PAGE_SCALING) // BASE_CONFIG["COLUMNS"]

# Расчёт базовых размеров именной части эталонного ценника
BASE_TITLE_PART_MARGIN_PX = int(MARGIN_TITLE_PART_PERC * BASE_TAG_4X4_WIDTH)
BASE_TITLE_PART_WIDTH = BASE_TAG_4X4_WIDTH - 2 * BASE_TITLE_PART_MARGIN_PX
BASE_TITLE_PART_HEIGHT = (BASE_TAG_4X4_HEIGHT // 2 - 1 * BASE_TITLE_PART_MARGIN_PX) # Учёт только верхнего отступа

# Расчёт базовых размеров ценовой части эталонного ценника
BASE_PRICE_PART_MARGIN_PX = int(MARGIN_PRICE_PART_PERC * BASE_TAG_4X4_WIDTH)
BASE_PRICE_PART_WIDTH = BASE_TAG_4X4_WIDTH - 2 * BASE_PRICE_PART_MARGIN_PX
BASE_PRICE_PART_HEIGHT = (BASE_TAG_4X4_HEIGHT // 2 - 2 * BASE_PRICE_PART_MARGIN_PX)

# === ОСНОВНЫЕ ФУНКЦИИ ===

# -=-=-=- Открытие файла и начало генерации -=-=-=-
def choose_file():
    # Открытие Excel файла
    filepath = filedialog.askopenfilename(
        title="Выберите Excel файл",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    # Если файл был выбран
    if filepath:
        # Обновления GUI
        start_screen_hide()
        progress_screen_show()
        # Запуск основного алгоритма в другом потоке
        thread = threading.Thread(target=generate_tags_wrapper, args=(filepath,))
        thread.start()

# -=-=-=- Начало генерации -=-=-=-
def generate_tags_wrapper(filepath):
    # Запуск основного алгоритма
    generate_tags(filepath)
    # Возврат в основной поток для обновления GUI
    root.after(0, generation_complete)

# -=-=-=- Завершение генерации -=-=-=-
def generation_complete():
    # Обновления GUI
    progress_screen_hide()
    start_screen_show()

# -=-=-=- Отрисовка области с названием товара на ценнике -=-=-=-
def draw_title_text(draw, text, max_width, max_height, start_x, start_y, margin_perc):
    # Расчёт отступа в пикселях
    margin_px = int(margin_perc * max_width)

    # Учёт отступов
    max_width -= 2 * margin_px
    max_height -= 1 * margin_px # Учёт только верхнего отступа для большей вместительности текста
    start_x += margin_px
    start_y += margin_px

    # Вычисление размера шрифта
    font_title = int(BASE_CONFIG["FONT_SIZE_TITLE"] * max_width / BASE_TITLE_PART_WIDTH * PAGE_SCALING)

    # Шрифт
    font = ImageFont.truetype(resource_path(FONT_FILE_NAME), font_title)

    # Расчёт высоты строки
    ascent, descent = font.getmetrics()
    line_height = ascent + descent

    # Расчёт промежутка между строками в пикселях
    line_spacing_px = int(line_height * LINE_SPACING_PROC)

    # Линии/строки текста
    lines = []         # Список строк
    current_line = ""  # Текущая строка

    # Позиция по Y относительно ценника
    y = 0

    # Генерация списка строк текста
    for idx, word in enumerate(text.split()):
        # Новая версия строки
        new_line_version = current_line + (" " if current_line else "") + word

        # Ширина и высота получившейся строки
        bbox = draw.textbbox((0, 0), new_line_version, font=font)
        new_line_width = bbox[2] - bbox[0]

        # Продвижение позиции по Y для проверки
        y += line_height

        # Если строка вышла за рамки по высоте, заканчиваем список строк
        if y > max_height and lines:
            # Последняя строка
            last_line = lines[-1]

            # Ширина последней строки с троеточием
            bbox = draw.textbbox((0, 0), last_line + "...", font=font)
            last_line_width = bbox[2] - bbox[0]

            # Если ширина последней строки с троеточием подходящая
            if last_line_width <= max_width:
                # Обновляем прошлую строку добавляя троеточие
                lines[-1] = last_line + "..."
            else:
                # Обновляем прошлую строку заменяя последнее слово троеточием
                lines[-1] = (" ".join(lines[-1].split()[0:-1]) if lines[-1] else "") + "..."

            # Останавливаем генерацию списка
            break

        # Если строка не вышла за рамки по ширине
        if new_line_width <= max_width:
            # Если это последнее слово
            if idx == len(text.split()) - 1:
                # Записываем новую строку в список
                lines.append(new_line_version)

                # Останавливаем генерацию списка
                break
            else:
                # Заменяем текущую строку на новую
                current_line = new_line_version
                # Откатываем продвижение позиции по Y
                y -= line_height
        else:
            # Добавляем текущую строку
            lines.append(current_line)

            # Если это последнее слово, записываем его в список, если помещается по высоте
            if idx == len(text.split()) - 1 and y + line_spacing_px + line_height <= max_height:
                lines.append(word)
            else:
                # Записываем не поместившееся слово на новую строку
                current_line = word
                # Продвижение позиции по Y
                y += line_spacing_px

    # Координатное расположение по X и Y
    x = start_x  # Выравнивание по левому краю
    y = start_y  # Обнуление позиции по Y относительно страницы

    # Рисование строк текста
    for line in lines:
        # Рисуем строку на изображении
        draw.text((x, y), line, fill="black", font=font)
        # Продвижение позиции по Y
        y += line_height + line_spacing_px

# -=-=-=- Отрисовка области с ценой товара на ценнике -=-=-=-
def draw_price_text(page, price_text, unit, max_width, max_height, start_x, start_y, margin_perc):
    # Расчёт отступа в пикселях
    margin_px = int(margin_perc * max_width)

    # Учёт отступов
    max_width -= 2 * margin_px
    max_height -= 2 * margin_px
    start_x += margin_px
    start_y += margin_px

    # Разбор строки с ценой на часть с рублями и часть с копейками
    if "." in price_text or "," in price_text:
        price_part, frac_part = price_text.split("." if "." in price_text else ",")  # Разделяем на рубли и копейки
        frac_part = frac_part.ljust(2, '0')[:2]  # Гарантируем 2 символа (на всякий случай)
        if frac_part == '00':
            frac_part = None
    else:
        price_part = price_text
        frac_part = None

    # Текст с символом рублей
    rub_part = "₽"

    # Текст единицы измерения
    unit_part = f"/{normalize_unit(unit)}" if unit else ""

    # Создаём временное изображение
    temp_width, temp_height = max_width, max_height
    temp_image = Image.new("RGBA", (temp_width, temp_height), (255, 255, 255, 0))
    temp_draw = ImageDraw.Draw(temp_image)

    # Вычисление размера основного шрифта
    font_size_price_part = int(BASE_CONFIG["FONT_SIZE_PRICE"] * max_width / BASE_PRICE_PART_WIDTH * PAGE_SCALING)

    # Расчёт меньших шрифтов относительно основного
    font_size_frac_and_rub_part = int(font_size_price_part * FONT_SIZE_PERC_FRAC_AND_RUB)
    font_size_unit_part = int(font_size_price_part * FONT_SIZE_PERC_UNIT)

    # Шрифты
    font_price_part = ImageFont.truetype(resource_path(FONT_FILE_NAME), font_size_price_part)
    font_frac_and_rub_part = ImageFont.truetype(resource_path(FONT_FILE_NAME), font_size_frac_and_rub_part)
    font_unit_part = ImageFont.truetype(resource_path(FONT_FILE_NAME), font_size_unit_part)

    # Получаем ascent и descent шрифтов для правильной высоты по координатам
    ascent_price_part, descent_price_part = font_price_part.getmetrics()
    ascent_frac_and_rub_part, descent_frac_and_rub_part = font_frac_and_rub_part.getmetrics()
    ascent_unit_part, _ = font_unit_part.getmetrics()

    # Ширина и высота текста основной цены
    bbox_price_part = temp_draw.textbbox((0, 0), price_part, font=font_price_part)
    price_part_width = bbox_price_part[2] - bbox_price_part[0]
    price_part_height = ascent_price_part + descent_price_part

    # Ширина текста с символом рублей
    bbox_rub_part = temp_draw.textbbox((0, 0), rub_part, font=font_frac_and_rub_part)
    rub_part_width = bbox_rub_part[2] - bbox_rub_part[0]

    # Ширина текста единицы измерения
    bbox_unit_part = temp_draw.textbbox((0, 0), unit_part, font=font_unit_part)
    unit_part_width = bbox_unit_part[2] - bbox_unit_part[0]

    # Расчёт промежутка между ценой и блоком с копейками в пикселях
    spacing = rub_part_width * SPACING_PRICE_PART_PERC

    # Увеличиваем размер холста если цена выходит за рамки
    while (price_part_width + spacing + rub_part_width + unit_part_width > temp_width or
           price_part_height > temp_height):
        # Пересоздание холста
        temp_width *= 2
        temp_height *= 2
        temp_image = Image.new("RGBA", (temp_width, temp_height), (255, 255, 255, 0))
        temp_draw = ImageDraw.Draw(temp_image)

    # Рисуем основную часть цены
    temp_draw.text((0, 0), price_part, font=font_price_part, fill="black")

    # Рисуем копейки справа от основной цены, если они есть
    if frac_part:
        temp_draw.text(
            (price_part_width + spacing, descent_price_part - descent_frac_and_rub_part),
            frac_part, font=font_frac_and_rub_part, fill="black"
        )

    # Рисуем символ рублей справа от основной цены
    temp_draw.text(
        (price_part_width + spacing, ascent_price_part - ascent_frac_and_rub_part),
        rub_part, font=font_frac_and_rub_part, fill="black"
    )

    # Рисуем единицу измерения справа от основной цены и символа рублей
    temp_draw.text(
        (price_part_width + spacing + rub_part_width, ascent_price_part - ascent_unit_part),
        unit_part, font=font_unit_part, fill="black"
    )

    # Обрезаем по содержимому
    temp_bbox = temp_image.getbbox()
    temp_image = temp_image.crop(temp_bbox)

    # Уменьшаем до нужного размера, если выходит за рамки (на всякий случай)
    if temp_image.width > max_width or temp_image.height > max_height:
        temp_height = temp_image.height # Временное сохранение temp_image.height
        temp_image.thumbnail((max_width, max_height), Image.Resampling.LANCZOS)
        # Если изначально не поместилось по высоте, то перезаписывает высоту цены для дальнейших расчётов
        if temp_height > max_height:
            price_part_height = temp_image.getbbox()[3]

    # Центральная линия, на которую должны "садиться" все цены (независимо от их высоты и ширины)
    price_center_y = start_y + price_part_height // 2

    # Вставляем по правому краю, центрируя по вертикали
    page.paste(
        temp_image,
        (int(start_x + (max_width - temp_image.width)), int(price_center_y - (temp_image.height // 2))),
        temp_image
    )

# -=-=-=- Генерация и сохранение страниц с ценниками -=-=-=-
def generate_tags(filepath):
    try:
        # Чтение полученного Excel файла
        df = pd.read_excel(filepath, usecols=COLUMN_NUMBERS_IN_PRODUCT_TABLE)
        items = df.values.tolist()

        # Список страниц
        pages = []

        # Создание первой страницы
        page = Image.new("RGB", (PAGE_WIDTH, PAGE_HEIGHT), "white")
        draw = ImageDraw.Draw(page)

        # Получение пользовательских данных
        rows, columns, width_tag_frame, page_margin_mm = get_user_settings()

        # Расчёт размеров
        max_tags_per_page = rows * columns                         # Максимальное количество ценников на страницу
        page_margin_px = page_margin_mm * PAGE_SCALING             # Размер отступа в пикселях (изначально в мм)
        width_tag_frame_px = width_tag_frame * PAGE_SCALING // 10  # Размер рамки в пикселях
        tag_width = (PAGE_WIDTH - 2 * page_margin_px) // columns   # Ширина ценника
        tag_height = (PAGE_HEIGHT - 2 * page_margin_px) // rows    # Длина ценника

        # Проверка на слишком узкие/тонкие ценники
        assert abs(rows - columns) <= 8, "Ценники слишком узкие/тонкие"

        # Новое вычисление отступов для ровного расположения (из-за целочисленного счёта)
        start_x = (PAGE_WIDTH - tag_width * columns) // 2
        start_y = (PAGE_HEIGHT - tag_height * rows) // 2

        # Количество записанных ценников
        current_tag = 0

        # Генерация ценников для всех товаров
        for idx, (name, unit, price) in enumerate(items):
            # Расположение ценника
            col = current_tag % columns            # Столбец
            row = (current_tag // columns) % rows  # Ряд

            # Координатное расположение ценника по X и Y
            x = start_x + col * tag_width
            y = start_y + row * tag_height

            # Рамка ценника
            draw.rectangle(
                [x, y, x + tag_width - 1,
                 y + tag_height - 1],
                outline="black", width=width_tag_frame_px
            )
            # Дополнительная рамка ценников для одинаковой ширины всех линий
            draw.rectangle(
                [x - width_tag_frame_px, y - width_tag_frame_px,
                 x + tag_width - 1 + width_tag_frame_px, y + tag_height - 1 + width_tag_frame_px],
                outline="black", width=width_tag_frame_px
            )

            # -=-=-=-=-=-=-=-=-=-=-=-=-=-=- Отрисовка названия товара -=-=-=-=-=-=-=-=-=-=-=-=-=-=-
            draw_title_text(
                draw,                   # Контекст рисования
                str(name),              # Название товара
                tag_width,              # Максимальная ширина под название
                tag_height // 2,        # Максимальная высота под название
                x,                      # Начальная позиция X
                y,                      # Начальная позиция Y (верхняя половина ценника)
                MARGIN_TITLE_PART_PERC  # Отступ от краёв ценника в %
            )

            # -=-=-=-=-=-=-=-=-=-=-=-=-=-=- Отрисовка цены товара -=-=-=-=-=-=-=-=-=-=-=-=-=-=-
            draw_price_text(
                page,                   # Текущая страница
                str(price),             # Цена товара
                str(unit),              # Единица измерения товара
                tag_width,              # Максимальная ширина под цену
                tag_height // 2,        # Максимальная высота под цену
                x,                      # Начальная позиция X
                y + tag_height // 2,    # Начальная позиция Y (нижняя половина ценника)
                MARGIN_PRICE_PART_PERC  # Отступ от краёв ценника в %
            )

            # +1 созданный ценник
            current_tag += 1

            # Сохранение страницы в итоговый список если страница заполнена
            if current_tag % max_tags_per_page == 0:
                pages.append(page)
                # Создание следующей страницы
                page = Image.new("RGB", (PAGE_WIDTH, PAGE_HEIGHT), "white")
                draw = ImageDraw.Draw(page)

            # Обновляем экран прогресса
            progress_label.config(text=f"Прогресс: {int((idx + 1) / len(items) * 100)}%")
            root.update_idletasks()

        # Добавляем последнюю страницу в итоговый список
        if current_tag % max_tags_per_page != 0:
            pages.append(page)

        # Запрос папки для сохранения страниц
        save_folder = filedialog.askdirectory(title="Выберите папку для сохранения")
        if not save_folder:
            messagebox.showwarning("Отмена", "Сохранение отменено.")
            return

        # Установка курсора в "ожидание"
        root.config(cursor="watch")
        root.update()

        # Сохранение страниц в виде png файлов
        for i, page in enumerate(pages, start=1):
            page.save(os.path.join(save_folder, f"Ценники {i}.png"))

        # Возвращаем обычный курсор
        root.config(cursor="")
        root.update()

        messagebox.showinfo("Успех", f"Сохранено {len(pages)} страниц!")
    except Exception as e:
        messagebox.showerror("Ошибка", str(e))

# === ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ===

# -=-=-=- Приведение к нужному виду данных о единице измерения -=-=-=-
def normalize_unit(unit_text):
    # Словарь сокращений
    mapping = {
        "килограмм": "кг",
        "штука": "шт",
        "грамм": "г",
        "литр": "л",
        "миллилитр": "мл",
        "метр": "м",
    }

    # Нормализация данного текста
    unit_text = unit_text.strip().lower()

    # Поиск в тексте единицы измерения из словаря для сокращения
    for key in mapping:
        if key in unit_text:
            return mapping[key]

    # Если ничего не нашли, вернуть как есть
    return unit_text

# -=-=-=- Получение пользовательских данных из полей -=-=-=-
def get_user_settings():
    return (
        rows_var.get(),             # Количество строк
        cols_var.get(),             # Количество столбцов
        width_tag_frame_var.get(),  # Размер рамки ценника
        margin_var.get()            # Размер отступа в мм
    )

# -=-=-=- Включение/показ стартового экрана -=-=-=-
def start_screen_show(boxes_pad=30):
    btn.pack(pady=(20, 0))

    rows_label.pack(side="left", padx=(0, 0))
    rows_spinbox.pack(side="left", padx=(0, boxes_pad))

    columns_label.pack(side="left", padx=(0, 0))
    columns_spinbox.pack(side="left", padx=(0, boxes_pad))

    width_tag_frame_label.pack(side="left", padx=(0, 0))
    width_tag_frame_spinbox.pack(side="left", padx=(0, boxes_pad))

    margin_label.pack(side="left", padx=(0, 0))
    margin_combo.pack(side="left", padx=(0, 0))

    root.update()

# -=-=-=- Отключение/скрытие стартового экрана -=-=-=-
def start_screen_hide():
    btn.pack_forget()

    rows_label.pack_forget()
    rows_spinbox.pack_forget()

    columns_label.pack_forget()
    columns_spinbox.pack_forget()

    width_tag_frame_label.pack_forget()
    width_tag_frame_spinbox.pack_forget()

    margin_label.pack_forget()
    margin_combo.pack_forget()

    root.update()

# -=-=-=- Включение/показ экрана прогресса -=-=-=-
def progress_screen_show():
    progress_label.pack(anchor="center")
    root.update()

# -=-=-=- Отключение/скрытие экрана прогресса -=-=-=-
def progress_screen_hide():
    progress_label.pack_forget()
    root.update()

# -=-=-=- Возвращает абсолютный путь к ресурсу при запуске из .exe или .py -=-=-=-
def resource_path(relative_path: str) -> str:
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# === GUI: ОСНОВНОЕ ОКНО ===

# Основное окно приложения
root = tkinter.Tk()
root.title(WINDOW_NAME)
root.geometry(WINDOW_SIZE)
root.configure(bg="#f0f0f0")
root.iconbitmap(resource_path(WINDOW_ICO))
root.resizable(False, False)

# === GUI: СТАРТОВЫЙ ЭКРАН ===

# Кнопка для выбора таблицы товаров
btn = tkinter.Button(
    root,
    text="Выбрать Excel таблицу с товарами",
    command=choose_file,
    font=FONT_STYLES["BUTTON"],
    bg="black",
    fg="white"
)

# Данные от пользователя которые изменяются в полях ниже
rows_var = tkinter.IntVar(value=BASE_CONFIG["ROWS"])                        # Количество строк
cols_var = tkinter.IntVar(value=BASE_CONFIG["COLUMNS"])                     # Количество столбцов
width_tag_frame_var = tkinter.IntVar(value=BASE_CONFIG["WIDTH_FRAME_TAG"])  # Размер рамки ценника
margin_var = tkinter.IntVar(value=BASE_CONFIG["PAGE_MARGIN"])               # Размер отступа в мм

# Фрейм для спинбоксов и комбобокса
fields_frame = tkinter.Frame(root)
fields_frame.pack(pady=(20, 0))

# Поле выбора количества строк
rows_label = tkinter.Label(fields_frame, text="Количество строк:", font=FONT_STYLES["BOX_LABEL"])
rows_spinbox = tkinter.Spinbox(
    fields_frame,
    from_=FIELD_OPTIONS["SPINBOX_ROWS_AND_COLS"]["MIN"],
    to=FIELD_OPTIONS["SPINBOX_ROWS_AND_COLS"]["MAX"],
    textvariable=rows_var,
    state="readonly",
    font=FONT_STYLES["BOX"],
    width=len(str(FIELD_OPTIONS["SPINBOX_ROWS_AND_COLS"]["MAX"]))
)

# Поле выбора количества столбцов
columns_label = tkinter.Label(fields_frame, text="Количество столбцов:", font=FONT_STYLES["BOX_LABEL"])
columns_spinbox = tkinter.Spinbox(
    fields_frame,
    from_=FIELD_OPTIONS["SPINBOX_ROWS_AND_COLS"]["MIN"],
    to=FIELD_OPTIONS["SPINBOX_ROWS_AND_COLS"]["MAX"],
    textvariable=cols_var,
    state="readonly",
    font=FONT_STYLES["BOX"],
    width=len(str(FIELD_OPTIONS["SPINBOX_ROWS_AND_COLS"]["MAX"]))
)

# Поле выбора размера рамки ценника
width_tag_frame_label = tkinter.Label(fields_frame, text="Размер рамки:", font=FONT_STYLES["BOX_LABEL"])
width_tag_frame_spinbox = tkinter.Spinbox(
    fields_frame,
    from_=FIELD_OPTIONS["SPINBOX_FRAME_WIDTH"]["MIN"],
    to=FIELD_OPTIONS["SPINBOX_FRAME_WIDTH"]["MAX"],
    textvariable=width_tag_frame_var,
    state="readonly",
    font=FONT_STYLES["BOX"],
    width=len(str(FIELD_OPTIONS["SPINBOX_FRAME_WIDTH"]["MAX"]))
)

# Поле выбора отступа примерно в мм
margin_label = tkinter.Label(fields_frame, text="Отступы ~(мм):", font=FONT_STYLES["BOX_LABEL"])
margin_combo = ttk.Combobox(
    fields_frame,
    textvariable=margin_var,
    values=FIELD_OPTIONS["COMBOBOX_MARGIN_OPTIONS"],
    state="readonly",
    font=FONT_STYLES["BOX"],
    width=len(str(max(FIELD_OPTIONS["COMBOBOX_MARGIN_OPTIONS"])))
)

# Включение группы кнопок и полей "Стартовый экран"
start_screen_show()

# === GUI: ЭКРАН ПРОГРЕССА ===

# Надпись прогресса
progress_label = tkinter.Label(root, text="Прогресс: 0%", font=FONT_STYLES["PROGRESS"], width=20)

# === GUI: ПРОЧЕЕ ===

# Зацикливание
root.mainloop()
