import tkinter
from tkinter import filedialog, messagebox
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import os

# === НАСТРОЙКИ ===

PAGE_WIDTH, PAGE_HEIGHT = 2100, 2970             # Размер A4 * 10

MARGIN_X = 100                                   # Отступ по x
MARGIN_Y = 100                                   # Отступ по y
TAGS_PER_ROW = 4                                 # Количество ценников в ряд
TAGS_PER_COL = 4                                 # Количество ценников на столбец
TAG_WIDTH = (PAGE_WIDTH - 2 * MARGIN_X) / 4      # Ширина ценника
TAG_HEIGHT = (PAGE_WIDTH - 2 * MARGIN_X) / 4     # Длина ценника
MAX_TAGS_PER_PAGE = TAGS_PER_ROW * TAGS_PER_COL  # Максимальное количество ценников на страницу

FONT_NAME = "arialbd.ttf"                        # Основной шрифт
FONT_SIZE_TITLE = 40                             # Размер шрифта для названия товара
FONT_SIZE_PRICE = 120                            # Размер шрифта для цены товара
WIDTH_TAG_FRAME = 4                              # Ширина рамки ценника

# === ФУНКЦИИ ===

def choose_file():
    filepath = filedialog.askopenfilename(
        title="Выберите Excel файл",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if filepath:
        # Скрываем кнопку
        btn.pack_forget()
        # Показываем надпись прогресса
        progress_label_pack()

        # Основной алгоритм генерации страниц с ценниками
        generate_tags(filepath)

        # Скрываем надпись прогресса
        progress_label.pack_forget()
        # Показываем кнопку
        btn_pack()

def generate_tags(filepath):
    try:
        # Чтение Excel файла
        df = pd.read_excel(filepath)
        items = df[['Название', 'Цена']].values.tolist() # ⚠️

        # Список страниц
        pages = []
        # Создание первой страницы
        page = Image.new("RGB", (PAGE_WIDTH, PAGE_HEIGHT), "white")
        draw = ImageDraw.Draw(page)

        font_title = ImageFont.truetype(FONT_NAME, FONT_SIZE_TITLE)  # Шрифт для названия товара
        font_price = ImageFont.truetype(FONT_NAME, FONT_SIZE_PRICE)  # Шрифт для цены товара

        current_tag = 0           # Количество записанных ценников
        total_items = len(items)  # Количество товаров в таблице

        for idx, (name, price) in enumerate(items):
            # Расположение ценника
            col = current_tag % TAGS_PER_ROW                    # Столбец
            row = (current_tag // TAGS_PER_ROW) % TAGS_PER_COL  # Ряд

            # Координатное расположение ценника
            x = MARGIN_X + col * TAG_WIDTH
            y = MARGIN_Y + row * TAG_HEIGHT

            # Рамка ценника
            draw.rectangle([x, y, x + TAG_WIDTH, y + TAG_HEIGHT], outline="black", width=WIDTH_TAG_FRAME)

            # Название товара
            text_name = str(name)
            bbox_name = draw.textbbox((0, 0), text_name, font=font_title)
            w_name = bbox_name[2] - bbox_name[0]
            draw.text(
                (x + (TAG_WIDTH - w_name) / 2, y + 50),
                text_name,
                fill="black",
                font=font_title
            )

            # Цена товара
            text_price = f"{price}₽"
            bbox_price = draw.textbbox((0, 0), text_price, font=font_price)
            w_price = bbox_price[2] - bbox_price[0]
            draw.text(
                (x + (TAG_WIDTH - w_price) / 2, y + TAG_HEIGHT / 2),
                text_price,
                fill="black",
                font=font_price
            )

            # +1 созданный ценник
            current_tag += 1

            # Проверка заполненности страницы
            if current_tag % MAX_TAGS_PER_PAGE == 0:
                # Дополнительная рамка ценников для одинаковой ширины всех линий
                draw.rectangle([
                    MARGIN_X - WIDTH_TAG_FRAME,
                    MARGIN_Y - WIDTH_TAG_FRAME,
                    PAGE_WIDTH - MARGIN_X + WIDTH_TAG_FRAME,
                    PAGE_HEIGHT - MARGIN_Y + WIDTH_TAG_FRAME
                ], outline="black", width=WIDTH_TAG_FRAME)
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
        if current_tag % MAX_TAGS_PER_PAGE != 0:
            pages.append(page)

        # Спросить папку для сохранения
        save_folder = filedialog.askdirectory(title="Выберите папку для сохранения")
        if not save_folder:
            messagebox.showwarning("Отмена", "Сохранение отменено.")
            return

        # Сохранение страниц в виде png файла
        for i, page in enumerate(pages, start=1):
            page.save(os.path.join(save_folder, f"Ценники {i}.png"))

        messagebox.showinfo("Успех", f"✅ Сохранено {len(pages)} страниц!") # Сообщение об успехе
    except Exception as e:
        messagebox.showerror("Ошибка", str(e)) # Сообщение об ошибке


def btn_pack():
    btn.pack(expand=True, anchor="center")

def progress_label_pack():
    progress_label.pack(expand=True, anchor="center")

# === GUI ===

# Основное окно приложения
root = tkinter.Tk()
root.title("Генератор ценников для магазина Алмаз")
root.geometry("500x100")
root.configure(bg="#f0f0f0")
root.iconbitmap('Almaz.ico')
root.resizable(False, False)

# Кнопка для выбора таблицы товаров
btn = tkinter.Button(
    root,
    text="Выбрать Excel таблицу с товарами", command=choose_file, font=("Arial Bold", 16, "bold"),
    bg="black", fg="white"
)
btn_pack()

# Надпись прогресса
progress_label = tkinter.Label(root, text="Прогресс: 0%", font=("Arial Bold", 16, "bold"))
progress_label_pack()
progress_label.pack_forget()

# Зацикливание
root.mainloop()
