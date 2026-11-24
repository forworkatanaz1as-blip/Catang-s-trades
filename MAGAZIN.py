# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import random
import os
import re
import json
import shutil
import uuid
import zipfile

# -------------------------
# Редкости предметов и множители цен
# -------------------------
ITEM_RARITIES = ["Обычный", "Необычный", "Редкий", "Очень редкий"]
PRICE_MULTIPLIERS = {
    "Обычный": 1.0,
    "Необычный": 1.0,
    "Редкий": 1.0,
    "Очень редкий": 1.0
}

# Распределение редкостей товаров в зависимости от редкости продавца
# Ключ: редкость продавца, значение: список вероятностей для ITEM_RARITIES в том же порядке
RARITY_DISTRIBUTION = {
    "Обычный":     [1.0, 0.0, 0.0, 0.0],   # 100% обычные
    "Необычный":   [0.8, 0.2, 0.0, 0.0],   # 70% обычные, 30% необычные
    "Редкий":      [0.2, 0.6, 0.2, 0.0],   # 70% необычные, 30% редкие
    "Очень редкий":[0.3, 0.3, 0.3, 0.1]    # 70% редкие, 30% очень редкие
}

# Диапазоны капитала в зависимости от редкости продавца
CAPITAL_RANGES_BY_RARITY = {
    "Обычный":     (1000, 10000),
    "Необычный":   (10000, 100000),
    "Редкий":      (100000, 1000000),
    "Очень редкий":(1000000, 10000000)
}

# -------------------------
# Данные: предметы с тегами-специальностями и редкостью
# -------------------------
# -------------------------
# Загрузка базы товаров из Excel
# -------------------------
def load_items_db():
    """Загружает товары из файла База_товаров.xlsx рядом с программой."""
    from openpyxl import load_workbook

    db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "База_товаров.xlsx")
    if not os.path.exists(db_path):
        raise FileNotFoundError(
            f"Не найден файл базы товаров:\n{db_path}\n\n"
            "Положите База_товаров.xlsx в одну папку с MAGAZIN.py."
        )

    try:
        workbook = load_workbook(db_path, read_only=True, data_only=True)
        if "Товары" not in workbook.sheetnames:
            raise ValueError("В Excel-файле отсутствует лист «Товары».")
        sheet = workbook["Товары"]
        rows = sheet.iter_rows(values_only=True)
        headers = next(rows, None)
        expected = ["ID", "Название", "Базовая цена", "Макс. количество", "Специализации", "Редкость"]
        if not headers or list(headers[:len(expected)]) != expected:
            raise ValueError("Неверные заголовки столбцов на листе «Товары». Не переименовывайте обязательные столбцы.")

        result = []
        seen_ids = set()
        valid_rarities = set(ITEM_RARITIES)
        for row_number, row in enumerate(rows, start=2):
            if not row or all(value is None for value in row[:6]):
                continue
            item_id, name, base_price, stack_max, tags, rarity = row[:6]
            item_id = str(item_id).strip() if item_id is not None else ""
            name = str(name).strip() if name is not None else ""
            if not item_id or not name:
                raise ValueError(f"Строка {row_number}: ID и название обязательны.")
            if item_id in seen_ids:
                raise ValueError(f"Строка {row_number}: повторяется ID «{item_id}».")
            try:
                base_price = int(base_price)
                stack_max = int(stack_max)
            except (TypeError, ValueError):
                raise ValueError(f"Строка {row_number}: цена и макс. количество должны быть целыми числами.")
            if base_price <= 0 or stack_max <= 0:
                raise ValueError(f"Строка {row_number}: цена и макс. количество должны быть больше нуля.")
            tags_list = [tag.strip() for tag in str(tags or "").replace(",", ";").split(";") if tag.strip()]
            if not tags_list:
                raise ValueError(f"Строка {row_number}: у товара должна быть хотя бы одна специализация.")
            rarity = str(rarity).strip() if rarity is not None else ""
            if rarity not in valid_rarities:
                raise ValueError(f"Строка {row_number}: неизвестная редкость «{rarity}».")
            seen_ids.add(item_id)
            result.append({"id": item_id, "name": name, "base_price": base_price,
                           "stack_max": stack_max, "tags": tags_list, "rarity": rarity})
        workbook.close()
        if not result:
            raise ValueError("На листе «Товары» нет ни одного товара.")
        return result
    except Exception as exc:
        raise RuntimeError(f"Ошибка чтения базы товаров Excel:\n{exc}") from exc


ITEMS_DB = load_items_db()



# -------------------------
# Словарь специальностей по редкости (обновлённый)
# -------------------------
SPECIALIZATIONS_BY_RARITY = {
    "Обычный": ["торгаш","бакалейщик","овощной","мясник","рыбак","разливщик","кок","плотник","ткач"],
    "Необычный": ["литейщик","стеклодув","ботаник","доктор","бронник","клинковщик","стрелок","навигатор"],
    "Редкий": ["ювелир","полимат","антиквар","музыкант","книжник"],
    "Очень редкий": ["корабельщик","зверолов","делец"]
}

# Плоский список всех специальностей для выбора пользователем
ALL_SPECIALIZATIONS = []
for rar_list in SPECIALIZATIONS_BY_RARITY.values():
    ALL_SPECIALIZATIONS.extend(rar_list)
ALL_SPECIALIZATIONS = sorted(set(ALL_SPECIALIZATIONS))
ALL_SPECS = ["Случайно"] + ALL_SPECIALIZATIONS

# -------------------------
# Остальные настройки
# -------------------------
NAMES = [
"Арден","Борис","Валерий","Гаррет","Дариус","Елена","Жанна","Зорян","Иларион","Кай",
"Лира","Марк","Нора","Оскар","Петра","Рен","Саймон","Талия","Ульрик","Фиона",

"Аарон","Абигаил","Адель","Азиз","Айлин","Акира","Алекс","Алексей","Алина","Алисия",
"Альберт","Альма","Амир","Ана","Андрей","Анита","Антон","Ари","Арина","Арман",
"Артур","Ася","Бела","Бен","Бенедикт","Беатрис","Богдан","Борислав","Бруно","Ванда",
"Василий","Вера","Виктор","Виктория","Вильям","Виола","Влад","Владимир","Габи","Габриэль",
"Галина","Ганс","Гектор","Георгий","Грета","Давид","Даниил","Даниэль","Дара","Дарья",
"Джеймс","Джон","Диана","Диего","Дмитрий","Доминик","Ева","Егор","Екатерина","Жак",
"Жасмин","Захар","Зейн","Злата","Иван","Игорь","Изабель","Илья","Инес","Инга",
"Ирина","Исак","Карим","Карл","Катарина","Кевин","Кира","Кирилл","Клара","Клаус",
"Крис","Кристина","Лев","Лео","Лейла","Леонид","Лиам","Лилия","Лин","Лора",
"Лука","Лукас","Майя","Максим","Мария","Марио","Марта","Матвей","Медина","Мелисса",
"Мика","Мила","Мирон","Мия","Мохаммед","Надя","Назар","Наталья","Натан","Нелли",
"Никита","Николь","Нильс","Олег","Оливия","Омар","Офелия","Павел","Паола","Патрик",
"Пьер","Радмир","Рамон","Рафаэль","Регина","Рик","Рита","Роберт","Роза","Руслан",
"Сабина","Саид","Сакура","Самир","Сандра","Сара","Сергей","Сильвия","София","Степан",
"Сулейман","Тамер","Тарас","Тимур","Томас","Тони","Умар","Фарид","Федор","Феликс",
"Флоренс","Ханна","Хасан","Хлоя","Цезарь","Чингиз","Шон","Эдвард","Элина","Элиот",
"Эльза","Эмиль","Эмма","Эрик","Юлия","Юсуф","Яков","Яна","Ярослав",

"Adrian","Aisha","Akio","Alvaro","Amara","Anya","Arjun","Ariel","Arman","Asher",
"Bao","Bianca","Carlos","Carmen","Chen","Chloe","Daichi","Dalia","Dante","Elias",
"Elina","Enzo","Fatima","Farah","Felix","Hiro","Imani","Iris","Jamal","Jin",
"Jorge","Kamal","Kenji","Lana","Leila","Malik","Mei","Nina","Noah","Omar",
"Ravi","Sami","Sora","Tariq","Yara","Zane",

"Акаину","Аокидзи","Кизару","Акира","Аяме","Асука","Аой","Ацуси","Акио","Арата",
"Даити","Дайске","Дзюн","Дзюнко","Дзётаро","Дзюнпэй","Дзюнъити","Дзиро","Дзюнсукэ","Дзюри",
"Итиро","Итиро","Икуми","Исао","Исаму","Идзуми","Итиро","Иори","Ицки","Итиро",
"Кайто","Каори","Канэки","Кадзуки","Кадзуо","Кадзухиро","Кадзухиса","Кадзуми","Кадзуё","Кадзуко",
"Кен","Кендзи","Кэнта","Кэнсин","Кэнъити","Кэйко","Кэйта","Кэйсукэ","Кэй","Киёми",
"Киёси","Кирито","Кохару","Кота","Котаро","Куми","Кумико","Куро","Куросаки","Кёко",
"Кё","Кёта","Кёхэй","Май","Маи","Маки","Макото","Мамору","Марико","Масато",
"Масуми","Маюми","Мегуми","Мидори","Микаса","Мики","Минато","Мио","Мирай","Мисато",
"Митсуо","Мицуки","Мицуру","Момо","Момока","Нао","Наоки","Наоми","Наруто","Нобору",
"Нодзоми","Ран","Рей","Рэйко","Рэн","Рин","Рицу","Рю","Рюдзи","Рюноскэ",
"Сабуро","Саки","Сакура","Сатоси","Сатору","Саэ","Саюри","Синдзи","Сора","Сосукэ",
"Субару","Суми","Сумико","Сюта","Сюдзи","Сюн","Сюнсукэ","Таити","Такахиро","Такаси",
"Такеши","Такуми","Таро","Тацухиро","Тэцуя","Томоко","Тору","Умэ","Хана","Ханако",
"Хару","Харука","Харухи","Хаято","Хикари","Хината","Хироки","Хироси","Хитоми","Хокто",
"Хотару","Юи","Юка","Юки","Юкико","Юко","Юма","Юми","Юмэ","Юри","Юрико",
"Юта","Юто","Юдзи","Яёи","Ясуо","Ясухиро","Ясуко","Ямато","Яёй","Ясин"
]

# Варианты для редкости специализации (с весами)
SPEC_RARITY_OPTIONS = [
    ("Случайно", None),
    ("Обычный", 0.40),
    ("Необычный", 0.30),
    ("Редкий", 0.20),
    ("Очень редкий", 0.10)
]

# Варианты для редкости продавца (с весами)
SELLER_RARITY_OPTIONS = [
    ("Случайно", None),
    ("Обычный", 0.40),
    ("Необычный", 0.30),
    ("Редкий", 0.20),
    ("Очень редкий", 0.10)
]

ASSORTMENT_SIZES = [
    ("Случайно", None, None),
    ("Маленький (1-4)", 0.40, (1,4)),
    ("Средний (4-8)", 0.30, (4,8)),
    ("Большой (8-12)", 0.20, (8,12)),
    ("Огромный (12-16)", 0.10, (12,16)),
]

# -------------------------
# Вспомогательные функции
# -------------------------
def weighted_choice(options_with_weights):
    filtered = [(v,w) for (v,w) in options_with_weights if w is not None]
    if not filtered:
        return random.choice([v for (v,w) in options_with_weights])
    total = sum(w for (_,w) in filtered)
    r = random.random() * total
    upto = 0
    for val,w in filtered:
        if upto + w >= r:
            return val
        upto += w
    return filtered[-1][0]

def choose_range(options):
    filtered = [(lab,prob,rng) for (lab,prob,rng) in options if prob and rng]
    if filtered:
        lab,prob,rng = random.choices(filtered, weights=[p for (_,p,_) in filtered], k=1)[0]
        return rng
    valid = [rng for (lab,prob,rng) in options if rng]
    return random.choice(valid)

def pick_assort(v):
    if v and v != "Случайно":
        for lab,prob,rng in ASSORTMENT_SIZES:
            if lab == v:
                return random.randint(rng[0], rng[1])
    rng = choose_range(ASSORTMENT_SIZES)
    return random.randint(rng[0], rng[1])

def pick_spec_rarity(selected):
    """Выбирает редкость специализации (строку) на основе выбора пользователя"""
    if selected and selected != "Случайно":
        return selected
    return weighted_choice(SPEC_RARITY_OPTIONS)

def pick_specialization(spec_rarity, selected_spec):
    """Выбирает специализацию. Если selected_spec == 'Случайно', то выбирает из списка соответствующей редкости"""
    if selected_spec and selected_spec != "Случайно":
        return selected_spec
    if spec_rarity is None:
        spec_rarity = pick_spec_rarity("Случайно")
    if spec_rarity in SPECIALIZATIONS_BY_RARITY:
        return random.choice(SPECIALIZATIONS_BY_RARITY[spec_rarity])
    else:
        return random.choice(ALL_SPECIALIZATIONS)

def pick_seller_rarity(selected):
    if selected and selected != "Случайно":
        return selected
    return weighted_choice(SELLER_RARITY_OPTIONS)

def pick_capital(seller_rarity):
    low, high = CAPITAL_RANGES_BY_RARITY[seller_rarity]
    return random.randint(low, high)

def pick_item_rarity(seller_rarity):
    """Выбирает редкость товара согласно распределению для данного продавца"""
    probs = RARITY_DISTRIBUTION.get(seller_rarity, [1.0,0.0,0.0,0.0])
    return random.choices(ITEM_RARITIES, weights=probs, k=1)[0]

def pick_qty(stack_max):
    if stack_max <= 1:
        return 1
    r = random.random()
    qty = max(1, int((r ** 1.5) * stack_max))
    return min(qty, stack_max)

def build_pool(spec, desired_rarity=None):
    pool = [item for item in ITEMS_DB if spec in item["tags"]]
    if desired_rarity:
        pool = [item for item in pool if item["rarity"] == desired_rarity]
    return pool

def format_price(n):
    return f"{n:,}".replace(",", ".")

def generate(settings):
    # Редкость специализации (для выбора специализации, если не указана)
    spec_rarity = pick_spec_rarity(settings["spec_rarity"])
    # Специализация
    spec = pick_specialization(spec_rarity, settings["spec"])
    # Редкость продавца
    seller_rarity = pick_seller_rarity(settings["seller_rarity"])
    # Капитал
    capital = pick_capital(seller_rarity)
    # Размер ассортимента
    assort = pick_assort(settings["assort"])
    # Имя
    name = random.choice(NAMES)

    chosen = []
    attempts = 0
    trader_price_mod = random.uniform(0.9, 1.1)
    while len(chosen) < assort and attempts < assort * 10:
        attempts += 1
        item_rarity = pick_item_rarity(seller_rarity)
        pool = build_pool(spec, item_rarity)
        if not pool:
            pool = build_pool(spec)
        if not pool:
            break
        it = random.choice(pool)
        if any(x["id"] == it["id"] for x in chosen):
            continue
        qty = pick_qty(it["stack_max"])
        price = int(round(it["base_price"] * PRICE_MULTIPLIERS[item_rarity] * random.uniform(0.7,1.3) * trader_price_mod / 10)) * 10
        chosen.append({
            "id": it["id"],
            "name": it["name"],
            "qty": qty,
            "price": price,
            "rarity": item_rarity
        })

    return {
        "id": f"T-{random.randint(100,999)}",
        "name": name,
        "spec": spec,
        "seller_rarity": seller_rarity,
        "capital": capital,
        "items": chosen
    }

# -------------------------
# GUI
# -------------------------
class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Генератор торговцев One Piece")
        self.root.geometry("1280x720")
        self.apply_pirate_style()
        self.current_merchant = None
        self.library_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Библиотека_торговцев")
        self.library_images_dir = os.path.join(self.library_dir, "Изображения")
        self.library_file = os.path.join(self.library_dir, "торговцы.json")
        os.makedirs(self.library_images_dir, exist_ok=True)
        self.merchant_library = self.load_merchant_library()
        self.library_thumbnails = []
        self.build_ui()

    def apply_pirate_style(self):
        self.bg_dark = "#2b1e12"
        self.bg_light = "#3d2a1a"
        self.pergament = "#ecd5a4"
        self.pergament_dark = "#c8a87c"
        self.gold = "#d4af37"
        self.ink = "#2c1a0c"
        self.combobox_bg = "#5c3a21"  # тёмно-коричневый для выпадающих списков
        self.root.configure(bg=self.bg_dark)
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure(".", background=self.bg_dark, foreground=self.pergament, fieldbackground=self.pergament, font=("Segoe UI", 10))
        style.configure("TFrame", background=self.bg_dark)
        style.configure("TLabelframe", background=self.bg_dark, foreground=self.gold)
        style.configure("TLabelframe.Label", background=self.bg_dark, foreground=self.gold)
        style.configure("TButton", background=self.bg_light, foreground=self.gold, borderwidth=2, padding=5)
        style.map("TButton", background=[("active", self.pergament_dark)], foreground=[("active", self.ink)])
        style.configure("TCombobox", fieldbackground=self.combobox_bg, background=self.combobox_bg,
                        foreground=self.gold, arrowcolor=self.gold)
        style.map("TCombobox", fieldbackground=[("readonly", self.combobox_bg)], foreground=[("readonly", self.gold)])
        style.configure("TLabel", background=self.bg_dark, foreground=self.pergament)
        style.configure("Treeview", background=self.pergament, foreground=self.ink, fieldbackground=self.pergament, rowheight=30, font=("Segoe UI", 15))
        style.map("Treeview", background=[("selected", self.gold)], foreground=[("selected", self.ink)])
        style.configure("Treeview.Heading", background=self.bg_light, foreground=self.gold, font=("Segoe UI", 11, "bold"), borderwidth=1, relief="raised")
        style.configure("Vertical.TScrollbar", background=self.bg_light, troughcolor=self.bg_dark, arrowcolor=self.gold)

    def build_ui(self):
        # Единая палитра вкладок: тёмная основа, светлый текст, золотая активная вкладка.
        self.ui_style = ttk.Style()
        try:
            self.ui_style.theme_use("clam")
        except tk.TclError:
            pass
        self.ui_style.configure(
            "App.TNotebook",
            background=self.bg_dark,
            borderwidth=0,
            tabmargins=(4, 4, 4, 0)
        )
        self.ui_style.configure(
            "App.TNotebook.Tab",
            background="#4b3524",
            foreground="#f4e6cb",
            padding=(18, 10),
            font=("Segoe UI", 10, "bold"),
            borderwidth=1
        )
        self.ui_style.map(
            "App.TNotebook.Tab",
            background=[("selected", "#d2a653"), ("active", "#765536")],
            foreground=[("selected", "#24180e"), ("active", "#fff4df")],
            expand=[("selected", (2, 2, 2, 0))]
        )

        self.notebook = ttk.Notebook(self.root, style="App.TNotebook")
        self.notebook.pack(fill="both", expand=True)
        generator_tab = tk.Frame(self.notebook, bg=self.bg_dark)
        self.notebook.add(generator_tab, text="Генератор")
        self.library_tab = tk.Frame(self.notebook, bg=self.bg_dark)
        self.notebook.add(self.library_tab, text="Библиотека торговцев")
        self.items_db_tab = tk.Frame(self.notebook, bg=self.bg_dark)
        self.notebook.add(self.items_db_tab, text="База всех товаров")

        # Главная вкладка сохраняет исходную компоновку на весь экран:
        # слева настройки, справа информация, список и закреплённая форма товара.
        main_paned = tk.Frame(generator_tab, bg=self.bg_dark)
        main_paned.pack(fill="both", expand=True, padx=10, pady=10)

        # Левая панель
        left_frame = tk.Frame(main_paned, bg=self.bg_light, relief="raised", bd=3, width=300)
        left_frame.pack(side="left", fill="y", padx=(0,10))
        left_frame.pack_propagate(False)

        # Правая панель
        right_frame = tk.Frame(main_paned, bg=self.bg_dark)
        right_frame.pack(side="right", fill="both", expand=True)

        # ---------- Левая панель ----------
        tk.Label(left_frame, text="Настройки генерации", font=("Segoe UI", 14, "bold"),
                 bg=self.bg_light, fg=self.gold).pack(pady=(10,20))

        # Редкость специализации
        tk.Label(left_frame, text="Редкость специализации:", bg=self.bg_light, fg=self.pergament,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=(5,0))
        self.spec_rarity = tk.StringVar(value="Случайно")
        self.spec_rarity_cb = ttk.Combobox(left_frame, values=[r[0] for r in SPEC_RARITY_OPTIONS],
                                           textvariable=self.spec_rarity, state="readonly", width=25)
        self.spec_rarity_cb.pack(padx=10, pady=(0,10), fill="x")

        # Редкость продавца
        tk.Label(left_frame, text="Редкость продавца:", bg=self.bg_light, fg=self.pergament,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=(5,0))
        self.seller_rarity = tk.StringVar(value="Случайно")
        self.seller_rarity_cb = ttk.Combobox(left_frame, values=[r[0] for r in SELLER_RARITY_OPTIONS],
                                             textvariable=self.seller_rarity, state="readonly", width=25)
        self.seller_rarity_cb.pack(padx=10, pady=(0,10), fill="x")

        # Специализация (Listbox с прокруткой)
        tk.Label(left_frame, text="Специализация:", bg=self.bg_light, fg=self.pergament,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=(5,0))
        spec_frame = tk.Frame(left_frame, bg=self.bg_light)
        spec_frame.pack(padx=10, pady=(0,10), fill="both", expand=True)
        scrollbar = tk.Scrollbar(spec_frame, bg=self.bg_light, troughcolor=self.bg_dark)
        scrollbar.pack(side="right", fill="y")
        self.spec_listbox = tk.Listbox(spec_frame, yscrollcommand=scrollbar.set,
                                       bg=self.pergament, fg=self.ink,
                                       font=("Segoe UI", 10), height=12,
                                       exportselection=False)
        self.spec_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.spec_listbox.yview)
        for spec in ALL_SPECS:
            self.spec_listbox.insert(tk.END, spec)
        self.spec_listbox.selection_set(0)

        # Ассортимент
        tk.Label(left_frame, text="Ассортимент:", bg=self.bg_light, fg=self.pergament,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=(5,0))
        self.assort = tk.StringVar(value="Случайно")
        self.assort_cb = ttk.Combobox(left_frame, values=[a[0] for a in ASSORTMENT_SIZES],
                                      textvariable=self.assort, state="readonly", width=25)
        self.assort_cb.pack(padx=10, pady=(0,20), fill="x")

        # Кнопка
        gen_btn = tk.Button(left_frame, text="Сгенерировать", command=self.generate_one,
                            bg=self.bg_light, fg=self.gold, font=("Segoe UI", 11, "bold"),
                            relief="raised", bd=3, padx=15, pady=8,
                            activebackground=self.pergament_dark, activeforeground=self.ink)
        gen_btn.pack(pady=(10, 6))

        self.export_btn = None  # PDF-экспорт доступен отдельно от главной вкладки.\n
        self.save_library_btn = tk.Button(left_frame, text="Сохранить торговца",
                            command=self.save_current_to_library,
                            bg=self.bg_light, fg=self.gold, font=("Segoe UI", 10, "bold"),
                            relief="raised", bd=3, padx=12, pady=6,
                            state="disabled")
        self.save_library_btn.pack(pady=(0, 8))

        credits_frame = tk.Frame(left_frame, bg=self.bg_light)
        credits_frame.pack(anchor="w", fill="x", padx=10, pady=(0, 10))
        tk.Label(credits_frame, text="Разработка: Catang", bg=self.bg_light, fg=self.pergament,
                 font=("Segoe UI", 9), anchor="w").pack(anchor="w")
        tk.Label(credits_frame, text="Доработка: Atanazi", bg=self.bg_light, fg=self.pergament,
                 font=("Segoe UI", 9), anchor="w").pack(anchor="w")

        # ---------- Правая панель ----------
        info_frame = tk.Frame(right_frame, bg=self.bg_light, relief="sunken", bd=2)
        info_frame.pack(fill="x", pady=(0,10))

        self.name_lbl = tk.Label(info_frame, text="Имя: —", font=("Segoe UI", 18, "bold"),
                                 bg=self.bg_light, fg=self.gold)
        self.name_lbl.pack(side="left", padx=15, pady=5)

        self.spec_lbl = tk.Label(info_frame, text="Специализация: —", font=("Segoe UI", 14, "bold"),
                                 bg=self.bg_light, fg=self.pergament)
        self.spec_lbl.pack(side="left", padx=15, pady=5)

        self.cap_lbl = tk.Label(info_frame, text="Капитал:", font=("Segoe UI", 14, "bold"),
                                bg=self.bg_light, fg=self.pergament)
        self.cap_lbl.pack(side="left", padx=(15, 5), pady=5)
        self.capital_var = tk.StringVar(value="—")
        self.capital_entry = tk.Entry(info_frame, textvariable=self.capital_var, width=12,
                                      bg=self.pergament, fg=self.ink,
                                      insertbackground=self.ink, font=("Segoe UI", 12, "bold"))
        self.capital_entry.pack(side="left", pady=5)
        self.capital_entry.bind("<Return>", self.apply_capital_edit)
        self.capital_entry.bind("<FocusOut>", self.apply_capital_edit)
        tk.Label(info_frame, text="белли", font=("Segoe UI", 12, "bold"),
                 bg=self.bg_light, fg=self.pergament).pack(side="left", padx=(5, 0), pady=5)

        # Список занимает всё свободное место справа, а форма добавления
        # закреплена непосредственно под ним и не сжимает всю компоновку.
        # Таблицу можно прокручивать собственной полосой прокрутки.
        # Таблица товаров
        tree_frame = tk.Frame(right_frame, bg=self.bg_dark)

        cols = ("qty", "name", "price", "remove")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=9)
        self.tree.heading("qty", text="Кол-во")
        self.tree.heading("name", text="Название")
        self.tree.heading("price", text="Цена/ед.")
        self.tree.heading("remove", text="")
        self.tree.column("qty", width=80, anchor="center", stretch=False)
        self.tree.column("name", width=360, minwidth=180, anchor="w", stretch=True)
        self.tree.column("price", width=130, minwidth=100, anchor="e", stretch=False)
        self.tree.column("remove", width=42, minwidth=42, anchor="center", stretch=False)
        self.tree.tag_configure("remove_cell", foreground=self.ink)
        self.tree.bind("<Button-1>", self.on_product_tree_click, add="+")
        self.tree.bind("<Double-1>", self.edit_product_cell)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True, side="left")

        self.tree.tag_configure("odd", background="#f7e5c2")
        self.tree.tag_configure("even", background="#ecd5a4")

        # Закреплённая нижняя панель добавления товаров справа под таблицей.
        add_panel = tk.LabelFrame(right_frame, text="Добавить товар в ассортимент",
                                  bg=self.bg_dark, fg=self.gold,
                                  font=("Segoe UI", 10, "bold"), padx=10, pady=8)
        add_panel.pack(fill="x", pady=(8, 0))
        tk.Label(add_panel, text="Название", bg=self.bg_dark, fg=self.pergament).grid(row=0, column=0, sticky="w")
        tk.Label(add_panel, text="Кол-во", bg=self.bg_dark, fg=self.pergament).grid(row=0, column=1, sticky="w", padx=(8, 0))
        tk.Label(add_panel, text="Цена/ед.", bg=self.bg_dark, fg=self.pergament).grid(row=0, column=2, sticky="w", padx=(8, 0))
        tk.Label(add_panel, text="Редкость", bg=self.bg_dark, fg=self.pergament).grid(row=0, column=3, sticky="w", padx=(8, 0))
        self.new_item_name = tk.Entry(add_panel, bg=self.pergament, fg=self.ink, insertbackground=self.ink)
        self.new_item_name.grid(row=1, column=0, sticky="ew", pady=(3, 0))
        self.new_item_qty = tk.Entry(add_panel, width=8, bg=self.pergament, fg=self.ink)
        self.new_item_qty.insert(0, "1")
        self.new_item_qty.grid(row=1, column=1, padx=(8, 0), pady=(3, 0))
        self.new_item_price = tk.Entry(add_panel, width=12, bg=self.pergament, fg=self.ink)
        self.new_item_price.insert(0, "0")
        self.new_item_price.grid(row=1, column=2, padx=(8, 0), pady=(3, 0))
        self.new_item_rarity = ttk.Combobox(add_panel, values=["Обычный", "Необычный", "Редкий", "Эпический", "Легендарный", "—"],
                                             state="readonly", width=14)
        self.new_item_rarity.set("Обычный")
        self.new_item_rarity.grid(row=1, column=3, padx=(8, 0), pady=(3, 0))
        self.add_item_btn = tk.Button(add_panel, text="Добавить", command=self.add_custom_item_to_merchant,
                                      bg=self.bg_light, fg=self.gold, font=("Segoe UI", 9, "bold"),
                                      relief="raised", bd=2, state="disabled")
        self.add_item_btn.grid(row=1, column=4, padx=(10, 0), pady=(3, 0))
        self.add_from_db_btn = tk.Button(
            add_panel, text="Выбрать из базы", command=self.open_item_database_picker,
            bg=self.bg_light, fg=self.gold, font=("Segoe UI", 9, "bold"),
            relief="raised", bd=2, state="disabled")
        self.add_from_db_btn.grid(row=1, column=5, padx=(8, 0), pady=(3, 0))
        add_panel.grid_columnconfigure(0, weight=1)
        add_panel.pack(side="bottom", fill="x", pady=(8, 0))
        tree_frame.pack(side="top", fill="both", expand=True)

        self.build_library_tab()
        self.build_items_db_tab()

    def build_items_db_tab(self):
        """Полный редактор исходной базы товаров Excel."""
        top = tk.Frame(self.items_db_tab, bg=self.bg_dark, padx=12, pady=10)
        top.pack(fill="x")
        tk.Label(top, text="Полная база товаров", bg=self.bg_dark, fg=self.gold,
                 font=("Segoe UI", 16, "bold")).pack(side="left")
        tk.Label(top, text="Изменения автоматически записываются в База_товаров.xlsx",
                 bg=self.bg_dark, fg=self.pergament).pack(side="right")

        search_bar = tk.Frame(self.items_db_tab, bg=self.bg_dark, padx=10, pady=8)
        search_bar.pack(fill="x")
        tk.Label(search_bar, text="Поиск по названию:", bg=self.bg_dark,
                 fg=self.pergament).pack(side="left", padx=(0, 8))
        self.db_search_var = tk.StringVar()
        search_entry = tk.Entry(search_bar, textvariable=self.db_search_var,
                                bg=self.pergament, fg=self.ink, width=42)
        search_entry.pack(side="left", fill="x", expand=True)
        self.db_search_var.trace_add("write", lambda *_: self.refresh_db_tree())
        tk.Button(search_bar, text="Очистить", command=lambda: self.db_search_var.set(""),
                  bg=self.bg_light, fg=self.gold).pack(side="left", padx=8)

        frame = tk.Frame(self.items_db_tab, bg=self.bg_dark, padx=10)
        frame.pack(fill="both", expand=True)
        cols = ("id", "name", "base_price", "stack_max", "tags", "rarity")
        self.db_columns = cols
        self.db_sort_state = {c: 0 for c in cols}
        self.db_tree = ttk.Treeview(frame, columns=cols, show="headings")
        headings = {"id":"ID", "name":"Название", "base_price":"Базовая цена",
                    "stack_max":"Макс. количество", "tags":"Специализации", "rarity":"Редкость"}
        widths = {"id":110, "name":330, "base_price":130, "stack_max":145, "tags":380, "rarity":140}
        for c in cols:
            self.db_tree.heading(c, text=headings[c],
                                 command=lambda col=c: self.sort_db_column(col))
            self.db_tree.column(c, width=widths[c], minwidth=80,
                                stretch=(c in ("name", "tags")), anchor="w")
        y = ttk.Scrollbar(frame, orient="vertical", command=self.db_tree.yview)
        x = ttk.Scrollbar(frame, orient="horizontal", command=self.db_tree.xview)
        self.db_tree.configure(yscrollcommand=y.set, xscrollcommand=x.set)
        y.pack(side="right", fill="y"); x.pack(side="bottom", fill="x")
        self.db_tree.pack(fill="both", expand=True)
        self.db_tree.bind("<Double-1>", self.edit_db_cell)

        add = tk.LabelFrame(self.items_db_tab, text="Добавить товар в исходную базу",
                            bg=self.bg_dark, fg=self.gold, padx=8, pady=8,
                            font=("Segoe UI", 10, "bold"))
        add.pack(fill="x", padx=10, pady=10)
        self.db_new_vars = {}
        fields = [("name", "Название"), ("base_price", "Базовая цена"),
                  ("stack_max", "Макс. количество"), ("tags", "Специализации (; через точку с запятой)")]
        next_id = self.next_db_item_id()
        self.db_next_id_label = tk.Label(add, text=f"ID будет присвоен автоматически: {next_id}",
                                         bg=self.bg_dark, fg=self.pergament)
        self.db_next_id_label.grid(row=0, column=0, columnspan=4, sticky="w", padx=4, pady=(0, 4))
        for col, (key, label) in enumerate(fields):
            box = tk.Frame(add, bg=self.bg_dark); box.grid(row=1, column=col, sticky="ew", padx=4)
            tk.Label(box, text=label, bg=self.bg_dark, fg=self.pergament).pack(anchor="w")
            var = tk.StringVar(); self.db_new_vars[key] = var
            tk.Entry(box, textvariable=var, bg=self.pergament, fg=self.ink, width=18).pack(fill="x")
            add.grid_columnconfigure(col, weight=1)
        box = tk.Frame(add, bg=self.bg_dark); box.grid(row=1, column=4, padx=4, sticky="ew")
        tk.Label(box, text="Редкость", bg=self.bg_dark, fg=self.pergament).pack(anchor="w")
        self.db_new_vars["rarity"] = tk.StringVar(value=ITEM_RARITIES[0])
        ttk.Combobox(box, textvariable=self.db_new_vars["rarity"], values=ITEM_RARITIES,
                     state="readonly", width=15).pack(fill="x")
        tk.Button(add, text="Добавить в базу", command=self.add_db_item,
                  bg=self.bg_light, fg=self.gold, font=("Segoe UI", 10, "bold")).grid(row=1, column=5, padx=8)
        self.refresh_db_tree()

    def next_db_item_id(self):
        """Следующий числовой ID; новые записи всегда добавляются в конец базы."""
        numeric_ids = []
        for item in ITEMS_DB:
            try:
                numeric_ids.append(int(str(item.get("id", "")).strip()))
            except (TypeError, ValueError):
                continue
        return str(max(numeric_ids, default=0) + 1)

    def sort_db_column(self, column):
        # Цикл из трёх состояний: без сортировки -> по возрастанию -> по убыванию -> без сортировки.
        state = self.db_sort_state.get(column, 0)
        state = {0: 1, 1: -1, -1: 0}[state]
        self.db_sort_state = {c: 0 for c in self.db_columns}
        self.db_sort_state[column] = state
        self.db_sort_column = column if state else None
        headings = {"id":"ID", "name":"Название", "base_price":"Базовая цена",
            "stack_max":"Макс. количество", "tags":"Специализации", "rarity":"Редкость"}
        for col, title in headings.items():
            marker = " ▲" if col == column and state == 1 else " ▼" if col == column and state == -1 else ""
            self.db_tree.heading(col, text=title + marker)
        self.refresh_db_tree()

    def refresh_db_tree(self):
        if not hasattr(self, "db_tree"): return
        self.db_tree.delete(*self.db_tree.get_children())
        query = self.db_search_var.get().strip().casefold() if hasattr(self, "db_search_var") else ""
        rows = [(i, item) for i, item in enumerate(ITEMS_DB)
                if not query or query in str(item.get("name", "")).casefold()]
        sort_col = getattr(self, "db_sort_column", None)
        if sort_col:
            def sort_key(pair):
                value = pair[1].get(sort_col, "")
                if sort_col in ("base_price", "stack_max"):
                    try: return (0, float(value))
                    except (TypeError, ValueError): return (1, str(value).casefold())
                if sort_col == "id":
                    try: return (0, int(value))
                    except (TypeError, ValueError): return (1, str(value).casefold())
                if sort_col == "tags": value = "; ".join(value) if isinstance(value, list) else value
                return str(value).casefold()
            rows.sort(key=sort_key, reverse=self.db_sort_state.get(sort_col, 0) == -1)
        for i, item in rows:
            self.db_tree.insert("", "end", iid=f"db_{i}", values=(item["id"], item["name"],
                format_price(item["base_price"]), item["stack_max"], "; ".join(item["tags"]), item["rarity"]))
        if hasattr(self, "db_next_id_label"):
            self.db_next_id_label.config(text=f"ID будет присвоен автоматически: {self.next_db_item_id()}")

    def persist_items_db(self):
        from openpyxl import load_workbook
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "База_товаров.xlsx")
        wb = load_workbook(path)
        ws = wb["Товары"]
        if ws.max_row > 1: ws.delete_rows(2, ws.max_row - 1)
        for item in ITEMS_DB:
            ws.append([item["id"], item["name"], item["base_price"], item["stack_max"],
                       "; ".join(item["tags"]), item["rarity"]])
        wb.save(path); wb.close()

    def edit_db_cell(self, event):
        row = self.db_tree.identify_row(event.y); col = self.db_tree.identify_column(event.x)
        if not row or not col: return
        idx = int(row.split("_")[1]); field = self.db_columns[int(col[1:])-1]
        item = ITEMS_DB[idx]; x, y, w, h = self.db_tree.bbox(row, col)
        old = "; ".join(item["tags"]) if field == "tags" else str(item[field])
        if field == "rarity":
            editor = ttk.Combobox(self.db_tree, values=ITEM_RARITIES, state="readonly"); editor.set(old)
        else:
            editor = tk.Entry(self.db_tree, bg=self.pergament, fg=self.ink); editor.insert(0, old)
        editor.place(x=x, y=y, width=w, height=h); editor.focus_set()
        if isinstance(editor, tk.Entry): editor.select_range(0, "end")
        def commit(_=None):
            value = editor.get().strip()
            try:
                if field in ("base_price", "stack_max"):
                    value = int(value.replace(" ", "").replace(".", "").replace(",", ""))
                    if value <= 0: raise ValueError
                elif field in ("id", "name") and not value: raise ValueError
                elif field == "tags":
                    value = [v.strip() for v in value.replace(",", ";").split(";") if v.strip()]
                    if not value: raise ValueError
                elif field == "rarity" and value not in ITEM_RARITIES: raise ValueError
                if field == "id" and any(j != idx and z["id"] == value for j,z in enumerate(ITEMS_DB)):
                    raise ValueError
            except ValueError:
                messagebox.showwarning("База товаров", "Проверьте значение. Числа должны быть положительными, ID и название — непустыми; ID должен быть уникальным.", parent=self.root)
                editor.destroy(); return "break"
            item[field] = value
            try:
                self.persist_items_db(); self.refresh_db_tree()
            except Exception as exc:
                messagebox.showerror("Ошибка записи", str(exc), parent=self.root)
            editor.destroy(); return "break"
        editor.bind("<Return>", commit); editor.bind("<FocusOut>", commit)
        return "break"

    def add_db_item(self):
        d = {k:v.get().strip() for k,v in self.db_new_vars.items()}
        try:
            if not d["name"]: raise ValueError("Укажите название товара.")
            price, stack = int(d["base_price"]), int(d["stack_max"])
            if price <= 0 or stack <= 0: raise ValueError("Цена и максимальное количество должны быть больше нуля.")
            tags = [x.strip() for x in d["tags"].replace(",", ";").split(";") if x.strip()]
            if not tags: raise ValueError("Укажите хотя бы одну специализацию.")
            if d["rarity"] not in ITEM_RARITIES: raise ValueError("Выберите корректную редкость.")
            new_id = self.next_db_item_id()
            ITEMS_DB.append({"id":new_id, "name":d["name"], "base_price":price,
                             "stack_max":stack, "tags":tags, "rarity":d["rarity"]})
            self.persist_items_db(); self.refresh_db_tree()
            for k,v in self.db_new_vars.items():
                if k != "rarity": v.set("")
            messagebox.showinfo("База товаров", f"Товар добавлен в конец исходной базы. Присвоен ID: {new_id}", parent=self.root)
        except Exception as exc:
            messagebox.showwarning("Не удалось добавить товар", str(exc), parent=self.root)

    def apply_capital_edit(self, _event=None):
        """Применяет новое значение капитала прямо в строке информации о торговце."""
        if not self.current_merchant:
            return
        raw_value = self.capital_var.get().strip().replace(" ", "").replace(".", "").replace(",", "")
        try:
            capital = int(raw_value)
            if capital < 0:
                raise ValueError
        except (TypeError, ValueError):
            messagebox.showwarning("Капитал", "Введите целое неотрицательное число белли.", parent=self.root)
            self.capital_var.set(format_price(self.current_merchant.get("capital", 0)))
            return "break" if _event is not None else None
        self.current_merchant["capital"] = capital
        self.capital_var.set(format_price(capital))
        return "break" if _event is not None else None

    def add_custom_item_to_merchant(self):
        """Добавляет пользовательский товар в текущего торговца на главной вкладке."""
        if not self.current_merchant:
            messagebox.showinfo("Добавление товара", "Сначала сгенерируйте торговца.", parent=self.root)
            return
        name = self.new_item_name.get().strip()
        try:
            qty = int(self.new_item_qty.get().strip())
            price = int(self.new_item_price.get().strip())
            if not name or qty <= 0 or price < 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Товар", "Укажите название, положительное количество и неотрицательную цену.", parent=self.root)
            return
        rarity = self.new_item_rarity.get() or "—"
        self.current_merchant.setdefault("items", []).append({
            "name": name, "qty": qty, "price": price, "rarity": rarity
        })
        self.tree.delete(*self.tree.get_children())
        for idx, item in enumerate(self.current_merchant["items"]):
            tag = "even" if idx % 2 == 0 else "odd"
            self.tree.insert("", "end", iid=f"merchant_item_{idx}",
                             values=(item.get("qty", 1), item.get("name", ""),
                                     format_price(item.get("price", 0)), "✕"),
                             tags=(tag, "remove_cell"))
        self.auto_resize_columns()
        self.new_item_name.delete(0, "end")
        self.new_item_qty.delete(0, "end")
        self.new_item_qty.insert(0, "1")
        self.new_item_price.delete(0, "end")
        self.new_item_price.insert(0, "0")
        self.new_item_rarity.set("Обычный")
        self.new_item_name.focus_set()

    def open_item_database_picker(self):
        """Открывает каталог товаров с поиском и множественным выбором."""
        if not self.current_merchant:
            messagebox.showinfo("База товаров", "Сначала сгенерируйте или загрузите торговца.", parent=self.root)
            return

        win = tk.Toplevel(self.root)
        win.title("Выбрать товары из базы")
        win.geometry("900x650")
        win.minsize(650, 420)
        win.configure(bg=self.bg_dark)
        win.transient(self.root)
        win.grab_set()

        header = tk.Frame(win, bg=self.bg_dark, padx=12, pady=10)
        header.pack(fill="x")
        tk.Label(header, text="Поиск по названию:", bg=self.bg_dark, fg=self.pergament,
                 font=("Segoe UI", 10, "bold")).pack(side="left")
        search_var = tk.StringVar()
        search_entry = tk.Entry(header, textvariable=search_var, bg=self.pergament,
                                fg=self.ink, insertbackground=self.ink, font=("Segoe UI", 10))
        search_entry.pack(side="left", fill="x", expand=True, padx=10)

        selected_ids = set()
        body = tk.Frame(win, bg=self.bg_dark, padx=12)
        body.pack(fill="both", expand=True)
        columns = ("selected", "id", "name", "price", "maxqty", "tags", "rarity")
        table = ttk.Treeview(body, columns=columns, show="headings", selectmode="browse")
        headings = {
            "selected": "✓", "id": "ID", "name": "Название", "price": "Цена",
            "maxqty": "Макс. кол-во", "tags": "Специализации", "rarity": "Редкость"
        }
        widths = {"selected": 42, "id": 65, "name": 230, "price": 90,
                  "maxqty": 95, "tags": 260, "rarity": 110}
        for col in columns:
            table.heading(col, text=headings[col])
            table.column(col, width=widths[col], minwidth=40,
                         anchor="center" if col in ("selected", "id", "price", "maxqty", "rarity") else "w",
                         stretch=(col in ("name", "tags")))
        table.tag_configure("odd", background="#f7e5c2")
        table.tag_configure("even", background="#ecd5a4")
        vsb = ttk.Scrollbar(body, orient="vertical", command=table.yview)
        hsb = ttk.Scrollbar(body, orient="horizontal", command=table.xview)
        table.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        table.pack(side="left", fill="both", expand=True)

        count_var = tk.StringVar(value="Выбрано: 0")
        tk.Label(win, textvariable=count_var, bg=self.bg_dark, fg=self.gold,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=14, pady=(6, 0))

        def render(*_):
            query = search_var.get().strip().casefold()
            table.delete(*table.get_children())
            for idx, item in enumerate(ITEMS_DB):
                if query and query not in item["name"].casefold():
                    continue
                iid = str(idx)
                mark = "☑" if item["id"] in selected_ids else "☐"
                tag = "even" if idx % 2 == 0 else "odd"
                table.insert("", "end", iid=iid, values=(
                    mark, item["id"], item["name"], format_price(item["base_price"]),
                    item["stack_max"], ", ".join(item["tags"]), item["rarity"]
                ), tags=(tag,))
            count_var.set(f"Выбрано: {len(selected_ids)}")

        def toggle(_event=None):
            row = table.identify_row(_event.y) if _event else ""
            col = table.identify_column(_event.x) if _event else ""
            if not row:
                return
            item = ITEMS_DB[int(row)]
            if item["id"] in selected_ids:
                selected_ids.remove(item["id"])
            else:
                selected_ids.add(item["id"])
            render()
            return "break"

        table.bind("<Button-1>", toggle)
        search_var.trace_add("write", render)

        footer = tk.Frame(win, bg=self.bg_dark, padx=12, pady=12)
        footer.pack(fill="x")

        def add_selected():
            chosen = [item for item in ITEMS_DB if item["id"] in selected_ids]
            if not chosen:
                messagebox.showinfo("Выбор товаров", "Отметьте галочками хотя бы один товар.", parent=win)
                return
            merchant_items = self.current_merchant.setdefault("items", [])
            for item in chosen:
                merchant_items.append({
                    "name": item["name"], "qty": 1, "price": item["base_price"],
                    "rarity": item["rarity"], "base_id": item["id"]
                })
            self.tree.delete(*self.tree.get_children())
            for idx, merchant_item in enumerate(merchant_items):
                tag = "even" if idx % 2 == 0 else "odd"
                self.tree.insert("", "end", iid=f"merchant_item_{idx}", values=(
                    merchant_item.get("qty", 1), merchant_item.get("name", ""),
                    format_price(merchant_item.get("price", 0)), "✕"
                ), tags=(tag, "remove_cell"))
            self.auto_resize_columns()
            win.destroy()

        tk.Button(footer, text="Добавить выбранные", command=add_selected,
                  bg=self.bg_light, fg=self.gold, font=("Segoe UI", 10, "bold"),
                  relief="raised", bd=2, padx=14, pady=6).pack(side="right")
        tk.Button(footer, text="Отмена", command=win.destroy,
                  bg=self.bg_light, fg=self.pergament, relief="raised", bd=2,
                  padx=14, pady=6).pack(side="right", padx=(0, 8))
        search_entry.focus_set()
        render()

    def edit_product_cell(self, event):
        """Редактирование количества, названия и цены двойным щелчком."""
        if not self.current_merchant:
            return
        region = self.tree.identify("region", event.x, event.y)
        column = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        if region != "cell" or not row_id or column not in ("#1", "#2", "#3"):
            return
        try:
            index = int(row_id.replace("merchant_item_", ""))
            item = self.current_merchant["items"][index]
        except (ValueError, IndexError, KeyError):
            return
        field = {"#1": "qty", "#2": "name", "#3": "price"}[column]
        x, y, width, height = self.tree.bbox(row_id, column)
        old_value = str(item.get(field, ""))
        if field == "price":
            old_value = str(item.get("price", 0))
        editor = tk.Entry(self.tree, bg=self.pergament, fg=self.ink,
                          insertbackground=self.ink, relief="solid", bd=1)
        editor.place(x=x, y=y, width=width, height=height)
        editor.insert(0, old_value)
        editor.select_range(0, "end")
        editor.focus_set()

        def commit(_event=None):
            value = editor.get().strip()
            try:
                if field == "qty":
                    parsed = int(value)
                    if parsed <= 0:
                        raise ValueError
                elif field == "price":
                    parsed = int(value.replace(" ", "").replace(",", "").replace(".", ""))
                    if parsed < 0:
                        raise ValueError
                else:
                    parsed = value
                    if not parsed:
                        raise ValueError
            except ValueError:
                messagebox.showwarning("Редактирование товара",
                    "Количество должно быть положительным целым числом, цена — неотрицательным целым числом, название не должно быть пустым.",
                    parent=self.root)
                editor.destroy()
                return "break"
            item[field] = parsed
            self.tree.set(row_id, field, format_price(parsed) if field == "price" else parsed)
            editor.destroy()
            self.auto_resize_columns()
            return "break"

        # Явно поддерживаем копирование выделенного текста в полях редактирования.
        def copy_selection(_event=None):
            try:
                selected = editor.selection_get()
            except tk.TclError:
                return "break"
            editor.clipboard_clear()
            editor.clipboard_append(selected)
            return "break"

        def paste_clipboard(_event=None):
            try:
                pasted = editor.clipboard_get()
            except tk.TclError:
                return "break"
            try:
                start = editor.index("sel.first")
                end = editor.index("sel.last")
                editor.delete(start, end)
                editor.insert(start, pasted)
            except tk.TclError:
                editor.insert(editor.index(tk.INSERT), pasted)
            return "break"

        editor.bind("<Control-c>", copy_selection)
        editor.bind("<Control-C>", copy_selection)
        editor.bind("<Control-v>", paste_clipboard)
        editor.bind("<Control-V>", paste_clipboard)
        editor.bind("<Return>", commit)
        editor.bind("<FocusOut>", commit)
        editor.bind("<Escape>", lambda _e: editor.destroy())

    def on_product_tree_click(self, event):
        """Удаляет товар только из текущего ассортимента торговца, не из базы ITEMS_DB."""
        region = self.tree.identify("region", event.x, event.y)
        column = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        if region != "cell" or column != "#4" or not row_id:
            return
        try:
            item_index = int(row_id.replace("merchant_item_", ""))
        except ValueError:
            return
        if not self.current_merchant:
            return
        items = self.current_merchant.get("items", [])
        if 0 <= item_index < len(items):
            removed_name = items[item_index].get("name", "Товар")
            del items[item_index]
            self.tree.delete(row_id)
            # Перерисовываем оставшиеся строки, чтобы индексы соответствовали списку.
            self.tree.delete(*self.tree.get_children())
            for idx, item in enumerate(items):
                tag = "even" if idx % 2 == 0 else "odd"
                self.tree.insert(
                    "", "end", iid=f"merchant_item_{idx}",
                    values=(item.get("qty", 1), item.get("name", ""),
                            format_price(item.get("price", 0)), "✕"),
                    tags=(tag, "remove_cell")
                )
            self.auto_resize_columns()

    def generate_one(self):
        selection = self.spec_listbox.curselection()
        if selection:
            spec_value = self.spec_listbox.get(selection[0])
        else:
            spec_value = "Случайно"

        settings = {
            "spec_rarity": self.spec_rarity.get(),
            "seller_rarity": self.seller_rarity.get(),
            "spec": spec_value,
            "assort": self.assort.get()
        }
        try:
            m = generate(settings)
        except Exception as e:
            messagebox.showerror("Ошибка генерации", f"Не удалось создать торговца:\n{str(e)}")
            return

        self.current_merchant = m
        self.save_library_btn.configure(state="normal")
        self.add_item_btn.configure(state="normal")
        self.add_from_db_btn.configure(state="normal")
        self.name_lbl.config(text=f"Имя: {m['name']}")
        self.spec_lbl.config(text=f"Специализация: {m['spec']}")
        self.capital_var.set(format_price(m['capital']))

        for i in self.tree.get_children():
            self.tree.delete(i)

        if not m["items"]:
            self.tree.insert("", "end", values=("", "Товары отсутствуют", "", ""), tags=("odd",))
        else:
            for idx, it in enumerate(m["items"]):
                tag = "even" if idx % 2 == 0 else "odd"
                self.tree.insert("", "end", iid=f"merchant_item_{idx}", values=(
                    it["qty"],
                    it["name"],
                    format_price(it["price"]),
                    "✕"
                ), tags=(tag, "remove_cell"))

        self.auto_resize_columns()

    def open_export_dialog(self):
        """Окно настройки карточки торговца перед экспортом в PDF."""
        if not self.current_merchant:
            messagebox.showinfo("Экспорт", "Сначала сгенерируйте торговца.", parent=self.root)
            return

        merchant = self.current_merchant
        win = tk.Toplevel(self.root)
        win.title("Экспорт торговца в PDF")
        win.configure(bg=self.bg_dark)
        win.transient(self.root)
        win.grab_set()
        # Окно можно растягивать, чтобы длинный комментарий было удобно редактировать.
        win.resizable(True, True)
        win.minsize(660, 700)

        # Основная часть окна прокручивается целиком: после выбора большого арта
        # превью не вытеснит поля и кнопки за пределы видимой области.
        shell = tk.Frame(win, bg=self.bg_dark)
        shell.pack(fill="both", expand=True)

        content_area = tk.Frame(shell, bg=self.bg_dark)
        content_area.pack(side="top", fill="both", expand=True)

        canvas = tk.Canvas(content_area, bg=self.bg_dark, highlightthickness=0, bd=0)
        page_scroll = ttk.Scrollbar(content_area, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=page_scroll.set)
        page_scroll.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        outer = tk.Frame(canvas, bg=self.bg_dark, padx=18, pady=16)
        content_window = canvas.create_window((0, 0), window=outer, anchor="nw")

        def update_scroll_region(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def fit_content_width(event):
            canvas.itemconfigure(content_window, width=event.width)

        outer.bind("<Configure>", update_scroll_region)
        canvas.bind("<Configure>", fit_content_width)

        # Колёсико мыши прокручивает всё содержимое окна экспорта.
        def on_mousewheel(event):
            if getattr(event, "delta", 0):
                canvas.yview_scroll(int(-event.delta / 120), "units")
            elif getattr(event, "num", None) == 4:
                canvas.yview_scroll(-3, "units")
            elif getattr(event, "num", None) == 5:
                canvas.yview_scroll(3, "units")

        def bind_wheel(_event=None):
            canvas.bind_all("<MouseWheel>", on_mousewheel)
            canvas.bind_all("<Button-4>", on_mousewheel)
            canvas.bind_all("<Button-5>", on_mousewheel)

        def unbind_wheel(_event=None):
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        canvas.bind("<Enter>", bind_wheel)
        canvas.bind("<Leave>", unbind_wheel)
        def cleanup_export_window(event):
            if event.widget is win:
                unbind_wheel()
        win.bind("<Destroy>", cleanup_export_window)

        tk.Label(outer, text="Оформление карточки торговца", font=("Segoe UI", 16, "bold"),
                 bg=self.bg_dark, fg=self.gold).pack(anchor="w", pady=(0, 14))

        tk.Label(outer, text="Имя в PDF:", bg=self.bg_dark, fg=self.pergament,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")
        name_var = tk.StringVar(value=merchant["name"])
        name_entry = tk.Entry(outer, textvariable=name_var, width=54, font=("Segoe UI", 11),
                              bg=self.pergament, fg=self.ink, insertbackground=self.ink)
        name_entry.pack(fill="x", pady=(4, 12))

        tk.Label(outer, text="Иллюстрация / арт:", bg=self.bg_dark, fg=self.pergament,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")
        image_var = tk.StringVar(value="")
        image_row = tk.Frame(outer, bg=self.bg_dark)
        image_row.pack(fill="x", pady=(4, 12))
        image_entry = tk.Entry(image_row, textvariable=image_var, width=43, font=("Segoe UI", 10),
                               bg=self.pergament, fg=self.ink)
        image_entry.pack(side="left", fill="x", expand=True)
        preview_label = tk.Label(outer, text="Изображение не выбрано", bg=self.bg_dark,
                                 fg=self.pergament_dark, font=("Segoe UI", 9),
                                 anchor="w", justify="left")
        preview_label.pack(anchor="w", fill="x", pady=(0, 10))
        preview_label.image = None

        def choose_image():
            path = filedialog.askopenfilename(parent=win, title="Выберите изображение торговца",
                filetypes=[("Изображения", "*.png *.jpg *.jpeg *.webp *.bmp *.gif"), ("Все файлы", "*.*")])
            if not path:
                return
            image_var.set(path)
            try:
                from PIL import Image, ImageTk
                im = Image.open(path)
                im.thumbnail((220, 260))
                photo = ImageTk.PhotoImage(im)
                preview_label.configure(image=photo, text="")
                preview_label.image = photo
            except Exception:
                preview_label.configure(image="", text=os.path.basename(path))
                preview_label.image = None
            # Обновляем область прокрутки после изменения размера превью.
            win.after_idle(update_scroll_region)

        tk.Button(image_row, text="Выбрать…", command=choose_image, bg=self.bg_light, fg=self.gold,
                  font=("Segoe UI", 9, "bold"), relief="raised", bd=2).pack(side="left", padx=(8, 0))

        tk.Label(outer, text="Краткий комментарий:", bg=self.bg_dark, fg=self.pergament,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")
        comment_frame = tk.Frame(outer, bg=self.bg_dark)
        comment_frame.pack(fill="x", expand=False, pady=(4, 14))
        comment_scroll = ttk.Scrollbar(comment_frame, orient="vertical")
        comment = tk.Text(comment_frame, width=72, height=10, wrap="word", font=("Segoe UI", 11),
                          undo=True, padx=8, pady=8, spacing1=2,
                          bg=self.pergament, fg=self.ink, insertbackground=self.ink,
                          yscrollcommand=comment_scroll.set)
        comment_scroll.configure(command=comment.yview)
        comment.pack(side="left", fill="both", expand=True)
        comment_scroll.pack(side="right", fill="y")

        def save_pdf():
            display_name = name_var.get().strip()
            if not display_name:
                messagebox.showwarning("Не указано имя", "Введите имя торговца для PDF.", parent=win)
                return
            image_path = image_var.get().strip()
            if image_path and not os.path.isfile(image_path):
                messagebox.showwarning("Изображение не найдено", "Выбранный файл изображения недоступен.", parent=win)
                return
            suggested = "Торговец_" + "_".join(display_name.split()) + ".pdf"
            output_path = filedialog.asksaveasfilename(parent=win, title="Сохранить карточку торговца",
                defaultextension=".pdf", initialfile=suggested,
                filetypes=[("PDF-документ", "*.pdf")])
            if not output_path:
                return
            try:
                merchant_for_pdf = dict(merchant)
                merchant_for_pdf["items"] = list(merchant.get("items", []))
                self.create_merchant_pdf(merchant_for_pdf, output_path, display_name,
                                          comment.get("1.0", "end-1c").strip(), image_path or None)
            except Exception as exc:
                messagebox.showerror("Ошибка экспорта", f"Не удалось создать PDF:\\n{exc}", parent=win)
                return
            win.destroy()
            messagebox.showinfo("Экспорт завершён", f"Карточка торговца сохранена:\\n{output_path}", parent=self.root)

        # Панель действий закреплена внизу и остаётся доступной при прокрутке.
        buttons = tk.Frame(shell, bg=self.bg_dark, padx=18, pady=12)
        buttons.pack(side="bottom", fill="x")
        tk.Frame(shell, bg=self.gold, height=1).pack(side="bottom", fill="x")
        tk.Button(buttons, text="Сохранить и экспортировать", command=save_pdf, bg=self.gold, fg=self.ink,
                  font=("Segoe UI", 10, "bold"), padx=16, pady=9, relief="raised", bd=2).pack(side="right")
        tk.Button(buttons, text="Отмена", command=win.destroy, bg=self.bg_light, fg=self.pergament,
                  font=("Segoe UI", 10), padx=14, pady=7, relief="raised", bd=2).pack(side="right", padx=8)

        name_entry.focus_set()
        win.update_idletasks()
        win.geometry("760x700")
        win.minsize(620, 480)
        x = max(0, self.root.winfo_rootx() + (self.root.winfo_width() - win.winfo_width()) // 2)
        y = max(0, self.root.winfo_rooty() + (self.root.winfo_height() - win.winfo_height()) // 2)
        win.geometry(f"+{x}+{y}")

    def load_merchant_library(self):
        """Читает локальную библиотеку сохранённых торговцев."""
        if not os.path.isfile(self.library_file):
            return []
        try:
            with open(self.library_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, list) else []
        except Exception as exc:
            messagebox.showwarning("Библиотека торговцев",
                                   f"Не удалось прочитать файл библиотеки:\n{exc}")
            return []

    def save_merchant_library_file(self):
        os.makedirs(self.library_dir, exist_ok=True)
        temp_file = self.library_file + ".tmp"
        with open(temp_file, "w", encoding="utf-8") as f:
            json.dump(self.merchant_library, f, ensure_ascii=False, indent=2)
        os.replace(temp_file, self.library_file)

    def build_library_tab(self):
        top = tk.Frame(self.library_tab, bg=self.bg_dark, padx=16, pady=12)
        top.pack(fill="x")
        tk.Label(top, text="Библиотека торговцев", bg=self.bg_dark, fg=self.gold,
                 font=("Segoe UI", 18, "bold")).pack(side="left")
        tk.Label(top, text="Сохранённые персонажи и их ассортимент",
                 bg=self.bg_dark, fg=self.pergament,
                 font=("Segoe UI", 10)).pack(side="left", padx=18)

        body = tk.Frame(self.library_tab, bg=self.bg_dark, padx=14, pady=6)
        body.pack(fill="both", expand=True)
        self.library_canvas = tk.Canvas(body, bg="#342416", highlightthickness=1,
                                        highlightbackground=self.gold)
        scrollbar = ttk.Scrollbar(body, orient="vertical", command=self.library_canvas.yview)
        self.library_canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.library_canvas.pack(side="left", fill="both", expand=True)
        self.library_cards_frame = tk.Frame(self.library_canvas, bg="#342416")
        self.library_canvas_window = self.library_canvas.create_window(
            (0, 0), window=self.library_cards_frame, anchor="nw")
        self.library_cards_frame.bind(
            "<Configure>",
            lambda _e: self.library_canvas.configure(
                scrollregion=self.library_canvas.bbox("all")))
        self.library_canvas.bind(
            "<Configure>",
            lambda e: self.library_canvas.itemconfigure(self.library_canvas_window, width=e.width))

        actions = tk.Frame(self.library_tab, bg=self.bg_light, padx=12, pady=10)
        actions.pack(fill="x", side="bottom")
        self.library_selected_id = None
        tk.Button(actions, text="Загрузить в генератор", command=self.load_selected_merchant,
                  bg=self.pergament_dark, fg=self.ink, font=("Segoe UI", 10, "bold"),
                  padx=12, pady=7).pack(side="left", padx=4)
        tk.Button(actions, text="Редактировать", command=self.edit_selected_merchant,
                  bg=self.bg_light, fg=self.gold, font=("Segoe UI", 10, "bold"),
                  padx=12, pady=7).pack(side="left", padx=4)
        tk.Button(actions, text="Экспортировать файл", command=self.export_selected_merchant_file,
                  bg=self.bg_light, fg=self.gold, font=("Segoe UI", 10, "bold"),
                  padx=12, pady=7).pack(side="left", padx=4)
        tk.Button(actions, text="Экспорт в PDF", command=self.export_selected_merchant,
                  bg=self.gold, fg=self.ink, font=("Segoe UI", 10, "bold"),
                  padx=12, pady=7).pack(side="left", padx=4)
        tk.Button(actions, text="Импортировать файл", command=self.import_merchant_file,
                  bg=self.gold, fg=self.ink, font=("Segoe UI", 10, "bold"),
                  padx=12, pady=7).pack(side="left", padx=4)
        tk.Button(actions, text="Удалить", command=self.delete_selected_merchant,
                  bg=self.bg_light, fg="#b85b45", font=("Segoe UI", 10, "bold"),
                  padx=12, pady=7).pack(side="right", padx=4)
        self.refresh_merchant_library()

    def refresh_merchant_library(self):
        if not hasattr(self, "library_cards_frame"):
            return
        for child in self.library_cards_frame.winfo_children():
            child.destroy()
        self.library_thumbnails.clear()
        self.library_selected_id = None

        if not self.merchant_library:
            tk.Label(self.library_cards_frame,
                     text="Здесь пока пусто.\nСоздайте торговца и нажмите «Сохранить торговца».",
                     bg="#342416", fg=self.pergament, font=("Segoe UI", 13),
                     justify="center").grid(row=0, column=0, padx=30, pady=45)
            return

        columns = 5
        for index, record in enumerate(self.merchant_library):
            card = tk.Frame(self.library_cards_frame, bg="#4a3422",
                            highlightthickness=2, highlightbackground="#7b5b36",
                            padx=9, pady=9, cursor="hand2")
            card.grid(row=index // columns, column=index % columns,
                      padx=10, pady=10, sticky="n")
            image_widget = tk.Label(card, bg="#2b1e12", fg=self.gold,
                                    text="NO ART", width=16, height=8,
                                    font=("Segoe UI", 10, "bold"))
            image_widget.pack()
            image_path = record.get("image_path")
            if image_path and os.path.isfile(image_path):
                try:
                    from PIL import Image, ImageTk
                    im = Image.open(image_path)
                    im.thumbnail((132, 112))
                    thumb = ImageTk.PhotoImage(im)
                    image_widget.configure(image=thumb, text="", width=132, height=112)
                    self.library_thumbnails.append(thumb)
                except Exception:
                    image_widget.configure(text="ART")
            title = tk.Label(card, text=record.get("name", "Без имени"),
                             bg="#4a3422", fg=self.pergament,
                             font=("Segoe UI", 10, "bold"), wraplength=145,
                             justify="center")
            title.pack(pady=(7, 2))
            subtitle = tk.Label(card,
                text=f"{record.get('spec', '—')} · {len(record.get('items', []))} товаров",
                bg="#4a3422", fg="#d8bd8d", font=("Segoe UI", 8),
                wraplength=145, justify="center")
            subtitle.pack()

            def select_card(_event=None, rid=record["id"]):
                self.library_selected_id = rid
                for child in self.library_cards_frame.winfo_children():
                    child.configure(highlightbackground="#7b5b36", highlightthickness=2)
                card_by_id = next((w for w in self.library_cards_frame.winfo_children()
                                   if getattr(w, "_merchant_id", None) == rid), None)
                if card_by_id:
                    card_by_id.configure(highlightbackground=self.gold, highlightthickness=3)

            card._merchant_id = record["id"]
            for widget in (card, image_widget, title, subtitle):
                widget.bind("<Button-1>", select_card)
                widget.bind("<Double-Button-1>",
                            lambda _e, rid=record["id"]: self.load_merchant_by_id(rid))
        self.library_canvas.configure(scrollregion=self.library_canvas.bbox("all"))

    def get_selected_merchant_record(self):
        if not self.library_selected_id:
            messagebox.showinfo("Библиотека", "Сначала выберите торговца.")
            return None
        return next((r for r in self.merchant_library
                     if r.get("id") == self.library_selected_id), None)

    def save_current_to_library(self):
        if not self.current_merchant:
            return
        self.open_library_editor(dict(self.current_merchant), None)

    def edit_selected_merchant(self):
        record = self.get_selected_merchant_record()
        if record:
            self.open_library_editor(record, record.get("id"))

    def open_library_editor(self, merchant, record_id=None):
        existing = next((r for r in self.merchant_library if r.get("id") == record_id), None)
        win = tk.Toplevel(self.root)
        win.title("Сохранить торговца" if existing is None else "Редактировать торговца")
        win.geometry("820x760")
        win.minsize(620, 500)
        win.resizable(True, True)
        win.configure(bg=self.bg_dark)
        win.transient(self.root)

        # Область с формой прокручивается целиком; кнопки остаются закреплёнными снизу.
        window_body = tk.Frame(win, bg=self.bg_dark)
        window_body.pack(fill="both", expand=True)

        scroll_area = tk.Frame(window_body, bg=self.bg_dark)
        scroll_area.pack(side="top", fill="both", expand=True)

        form_canvas = tk.Canvas(scroll_area, bg=self.bg_dark, highlightthickness=0, bd=0)
        form_scrollbar = ttk.Scrollbar(scroll_area, orient="vertical", command=form_canvas.yview)
        form_canvas.configure(yscrollcommand=form_scrollbar.set)
        form_scrollbar.pack(side="right", fill="y")
        form_canvas.pack(side="left", fill="both", expand=True)

        shell = tk.Frame(form_canvas, bg=self.bg_dark, padx=16, pady=14)
        shell_window = form_canvas.create_window((0, 0), window=shell, anchor="nw")

        def update_form_scroll(_event=None):
            form_canvas.configure(scrollregion=form_canvas.bbox("all"))

        def resize_form_width(event):
            form_canvas.itemconfigure(shell_window, width=event.width)

        shell.bind("<Configure>", update_form_scroll)
        form_canvas.bind("<Configure>", resize_form_width)

        def scroll_form(event):
            if getattr(event, "delta", 0):
                form_canvas.yview_scroll(int(-event.delta / 120), "units")
            elif getattr(event, "num", None) == 4:
                form_canvas.yview_scroll(-3, "units")
            elif getattr(event, "num", None) == 5:
                form_canvas.yview_scroll(3, "units")

        def bind_form_wheel(_event=None):
            form_canvas.bind_all("<MouseWheel>", scroll_form)
            form_canvas.bind_all("<Button-4>", scroll_form)
            form_canvas.bind_all("<Button-5>", scroll_form)

        def unbind_form_wheel(_event=None):
            form_canvas.unbind_all("<MouseWheel>")
            form_canvas.unbind_all("<Button-4>")
            form_canvas.unbind_all("<Button-5>")

        form_canvas.bind("<Enter>", bind_form_wheel)
        form_canvas.bind("<Leave>", unbind_form_wheel)

        # Полноэкранный режим переключается отдельной кнопкой.
        fullscreen_state = {"enabled": False}
        def toggle_fullscreen():
            fullscreen_state["enabled"] = not fullscreen_state["enabled"]
            win.attributes("-fullscreen", fullscreen_state["enabled"])
            if not fullscreen_state["enabled"]:
                win.geometry("820x760")

        def close_editor():
            unbind_form_wheel()
            win.destroy()

        win.bind("<Escape>", lambda _e: toggle_fullscreen() if fullscreen_state["enabled"] else close_editor())
        win.bind("<Destroy>", lambda event: unbind_form_wheel() if event.widget is win else None)
        tk.Label(shell, text="Карточка торговца", bg=self.bg_dark, fg=self.gold,
                 font=("Segoe UI", 16, "bold")).pack(anchor="w", pady=(0, 10))

        name_var = tk.StringVar(value=(existing or {}).get("name", merchant.get("name", "")))
        tk.Label(shell, text="Имя:", bg=self.bg_dark, fg=self.pergament).pack(anchor="w")
        tk.Entry(shell, textvariable=name_var, bg=self.pergament, fg=self.ink,
                 font=("Segoe UI", 11)).pack(fill="x", pady=(3, 10))

        image_var = tk.StringVar(value=(existing or {}).get("image_path", ""))
        tk.Label(shell, text="Изображение торговца:", bg=self.bg_dark,
                 fg=self.pergament).pack(anchor="w")
        image_row = tk.Frame(shell, bg=self.bg_dark)
        image_row.pack(fill="x", pady=(3, 8))
        tk.Entry(image_row, textvariable=image_var, bg=self.pergament,
                 fg=self.ink).pack(side="left", fill="x", expand=True)

        preview = tk.Label(shell, bg=self.bg_dark, fg=self.pergament,
                           text="Изображение не выбрано")
        preview.pack(anchor="w", pady=(0, 8))
        preview.image = None

        def show_preview(path):
            if not path or not os.path.isfile(path):
                preview.configure(image="", text="Изображение не выбрано")
                preview.image = None
                return
            try:
                from PIL import Image, ImageTk
                im = Image.open(path)
                im.thumbnail((150, 150))
                photo = ImageTk.PhotoImage(im)
                preview.configure(image=photo, text="")
                preview.image = photo
            except Exception:
                preview.configure(image="", text=os.path.basename(path))
                preview.image = None

        show_preview(image_var.get())

        def choose_art():
            path = filedialog.askopenfilename(parent=win, title="Выберите арт",
                filetypes=[("Изображения", "*.png *.jpg *.jpeg *.webp *.bmp *.gif"),
                           ("Все файлы", "*.*")])
            if path:
                image_var.set(path)
                show_preview(path)

        tk.Button(image_row, text="Выбрать…", command=choose_art,
                  bg=self.bg_light, fg=self.gold).pack(side="left", padx=6)

        tk.Label(shell, text="Описание / заметка:", bg=self.bg_dark,
                 fg=self.pergament).pack(anchor="w")
        comment = tk.Text(shell, height=5, wrap="word", font=("Segoe UI", 10),
                          bg=self.pergament, fg=self.ink, padx=7, pady=6)
        comment.pack(fill="x", pady=(3, 10))
        comment.insert("1.0", (existing or {}).get("comment", ""))

        # Дополнительные товары больше не редактируются в окне сохранения.
        # Уже сохранённые позиции сохраняем без изменений.
        custom_items = [dict(x) for x in (existing or {}).get("custom_items", [])]

        def save_record():
            display_name = name_var.get().strip()
            if not display_name:
                messagebox.showwarning("Имя", "Введите имя торговца.", parent=win)
                return
            source_image = image_var.get().strip()
            saved_image = (existing or {}).get("image_path")
            if source_image and os.path.isfile(source_image):
                ext = os.path.splitext(source_image)[1].lower() or ".png"
                saved_image = os.path.join(self.library_images_dir,
                                           f"{record_id or uuid.uuid4().hex}{ext}")
                if os.path.abspath(source_image) != os.path.abspath(saved_image):
                    shutil.copy2(source_image, saved_image)
            elif source_image and not os.path.isfile(source_image):
                messagebox.showwarning("Изображение", "Выбранный файл не найден.", parent=win)
                return

            source_base_items = (existing or {}).get("base_items", merchant.get("items", []))
            base_items = [dict(x) for x in source_base_items]
            combined_items = base_items + [dict(x) for x in custom_items]
            new_record = {
                "id": record_id or uuid.uuid4().hex,
                "name": display_name,
                "spec": merchant.get("spec", "—"),
                "seller_rarity": merchant.get("seller_rarity", "—"),
                "capital": merchant.get("capital", 0),
                "items": combined_items,
                "base_items": base_items,
                "custom_items": [dict(x) for x in custom_items],
                "comment": comment.get("1.0", "end-1c").strip(),
                "image_path": saved_image or ""
            }
            if existing:
                pos = self.merchant_library.index(existing)
                self.merchant_library[pos] = new_record
            else:
                self.merchant_library.append(new_record)
            try:
                self.save_merchant_library_file()
            except Exception as exc:
                messagebox.showerror("Сохранение", f"Не удалось сохранить библиотеку:\n{exc}", parent=win)
                return
            close_editor()
            self.refresh_merchant_library()
            self.notebook.select(self.library_tab)

        bottom = tk.Frame(window_body, bg=self.bg_light, padx=12, pady=10)
        bottom.pack(side="bottom", fill="x")
        tk.Button(bottom, text="Сохранить в библиотеку", command=save_record,
                  bg=self.gold, fg=self.ink, font=("Segoe UI", 10, "bold"),
                  padx=14, pady=8).pack(side="right")
        tk.Button(bottom, text="Отмена", command=close_editor,
                  bg=self.bg_light, fg=self.pergament, padx=12, pady=8).pack(side="right", padx=8)
        tk.Button(bottom, text="Во весь экран", command=toggle_fullscreen,
                  bg=self.bg_light, fg=self.gold, padx=12, pady=8).pack(side="left")

    def load_merchant_by_id(self, record_id):
        record = next((r for r in self.merchant_library if r.get("id") == record_id), None)
        if not record:
            return
        merchant = {k: record.get(k) for k in ("name", "spec", "seller_rarity", "capital")}
        merchant["items"] = [dict(x) for x in record.get("items", [])]
        self.current_merchant = merchant
        self.name_lbl.config(text=f"Имя: {merchant['name']}")
        self.spec_lbl.config(text=f"Специализация: {merchant['spec']}")
        self.capital_var.set(format_price(merchant['capital']))
        self.save_library_btn.configure(state="normal")
        self.add_item_btn.configure(state="normal")
        self.add_from_db_btn.configure(state="normal")
        self.tree.delete(*self.tree.get_children())
        for i, item in enumerate(merchant["items"]):
            self.tree.insert("", "end", iid=f"merchant_item_{i}",
                values=(item.get("qty", 1), item.get("name", ""),
                format_price(item.get("price", 0)), "✕"),
                tags=(("even" if i % 2 == 0 else "odd"), "remove_cell"))
        self.auto_resize_columns()
        self.notebook.select(0)

    def load_selected_merchant(self):
        record = self.get_selected_merchant_record()
        if record:
            self.load_merchant_by_id(record["id"])

    def export_selected_merchant_file(self):
        """Экспортирует торговца в переносимый .merchant-файл (данные + арт)."""
        record = self.get_selected_merchant_record()
        if not record:
            return

        safe_name = re.sub(r'[^\w\- ]+', '', record.get("name", "Торговец")).strip()
        suggested = (safe_name or "Торговец") + ".merchant"
        output_path = filedialog.asksaveasfilename(
            parent=self.root, title="Экспортировать торговца",
            defaultextension=".merchant", initialfile=suggested,
            filetypes=[("Файл торговца", "*.merchant"), ("ZIP-архив", "*.zip")])
        if not output_path:
            return

        payload = {k: record.get(k) for k in (
            "name", "spec", "seller_rarity", "capital", "items",
            "base_items", "custom_items", "comment"
        )}
        payload["format"] = "MAGAZIN_MERCHANT"
        payload["version"] = 1
        image_path = record.get("image_path") or ""
        image_name = ""
        try:
            with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
                if image_path and os.path.isfile(image_path):
                    image_name = "portrait" + (os.path.splitext(image_path)[1] or ".png")
                    archive.write(image_path, image_name)
                payload["image_file"] = image_name
                archive.writestr("merchant.json",
                                 json.dumps(payload, ensure_ascii=False, indent=2))
        except Exception as exc:
            messagebox.showerror("Экспорт", f"Не удалось экспортировать торговца:\n{exc}")
            return
        messagebox.showinfo("Экспорт", f"Файл торговца сохранён:\n{output_path}")

    def import_merchant_file(self):
        """Импортирует .merchant-файл в локальную библиотеку."""
        input_path = filedialog.askopenfilename(
            parent=self.root, title="Импортировать торговца",
            filetypes=[("Файл торговца", "*.merchant *.zip"), ("Все файлы", "*.*")])
        if not input_path:
            return
        try:
            with zipfile.ZipFile(input_path, "r") as archive:
                names = archive.namelist()
                if "merchant.json" not in names:
                    raise ValueError("В файле отсутствует merchant.json.")
                payload = json.loads(archive.read("merchant.json").decode("utf-8"))
                if payload.get("format") != "MAGAZIN_MERCHANT":
                    raise ValueError("Формат файла не распознан.")
                if not isinstance(payload.get("items", []), list):
                    raise ValueError("Список товаров повреждён.")
                # Изображение берём только из архива, запрещая пути за его пределы.
                image_member = payload.get("image_file", "")
                imported_image = ""
                if image_member:
                    base_name = os.path.basename(image_member)
                    if base_name != image_member or base_name not in names:
                        raise ValueError("Некорректная ссылка на изображение в файле.")
                    ext = os.path.splitext(base_name)[1].lower() or ".png"
                    new_image = os.path.join(self.library_images_dir, f"{uuid.uuid4().hex}{ext}")
                    with archive.open(base_name) as src_file, open(new_image, "wb") as dst_file:
                        shutil.copyfileobj(src_file, dst_file)
                    imported_image = new_image

            imported_record = {
                "id": uuid.uuid4().hex,
                "name": str(payload.get("name") or "Импортированный торговец"),
                "spec": str(payload.get("spec") or "—"),
                "seller_rarity": str(payload.get("seller_rarity") or "—"),
                "capital": payload.get("capital", 0),
                "items": payload.get("items", []),
                "base_items": payload.get("base_items", payload.get("items", [])),
                "custom_items": payload.get("custom_items", []),
                "comment": str(payload.get("comment") or ""),
                "image_path": imported_image
            }
            self.merchant_library.append(imported_record)
            self.save_merchant_library_file()
            self.refresh_merchant_library()
            self.notebook.select(self.library_tab)
            messagebox.showinfo("Импорт завершён",
                                f"Торговец «{imported_record['name']}» добавлен в библиотеку.")
        except Exception as exc:
            messagebox.showerror("Импорт", f"Не удалось импортировать файл:\n{exc}")

    def export_selected_merchant(self):
        record = self.get_selected_merchant_record()
        if not record:
            return
        suggested = "Торговец_" + "_".join(record.get("name", "Без имени").split()) + ".pdf"
        output_path = filedialog.asksaveasfilename(parent=self.root, title="Экспорт торговца в PDF",
            defaultextension=".pdf", initialfile=suggested,
            filetypes=[("PDF-документ", "*.pdf")])
        if not output_path:
            return
        merchant = {k: record.get(k) for k in ("name", "spec", "seller_rarity", "capital")}
        merchant["items"] = [dict(x) for x in record.get("items", [])]
        try:
            self.create_merchant_pdf(merchant, output_path, record.get("name", "Торговец"),
                                     record.get("comment", ""), record.get("image_path") or None)
        except Exception as exc:
            messagebox.showerror("Экспорт", f"Не удалось создать PDF:\n{exc}")
            return
        messagebox.showinfo("Экспорт", f"PDF сохранён:\n{output_path}")

    def delete_selected_merchant(self):
        record = self.get_selected_merchant_record()
        if not record:
            return
        if not messagebox.askyesno("Удаление", f"Удалить «{record.get('name', 'Без имени')}» из библиотеки?"):
            return
        self.merchant_library = [r for r in self.merchant_library if r.get("id") != record["id"]]
        self.save_merchant_library_file()
        self.refresh_merchant_library()

    def create_merchant_pdf(self, merchant, output_path, display_name, comment, image_path=None):
        """Создаёт оформленную PDF-карточку торговца с кириллическим шрифтом."""
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_CENTER, TA_LEFT
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import mm
        from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                        TableStyle, Image as PDFImage, KeepTogether)
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        from xml.sax.saxutils import escape

        # Подбираем шрифт с поддержкой кириллицы на Windows и Linux.
        font_pairs = [
            (r"C:\Windows\Fonts\arial.ttf", r"C:\Windows\Fonts\arialbd.ttf"),
            ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
            ("/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf", "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf"),
        ]
        regular_path = bold_path = None
        for regular, bold in font_pairs:
            if os.path.isfile(regular):
                regular_path = regular
                bold_path = bold if os.path.isfile(bold) else regular
                break
        if not regular_path:
            raise RuntimeError("Не найден шрифт с поддержкой кириллицы. Установите Arial или DejaVu Sans.")
        if "MerchantRegular" not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont("MerchantRegular", regular_path))
            pdfmetrics.registerFont(TTFont("MerchantBold", bold_path))

        parchment = colors.HexColor("#F3E4C5")
        dark = colors.HexColor("#2B1E12")
        brown = colors.HexColor("#5A3B22")
        gold = colors.HexColor("#C49A35")
        ink = colors.HexColor("#2C1A0C")

        doc = SimpleDocTemplate(output_path, pagesize=A4, rightMargin=17*mm, leftMargin=17*mm,
                                topMargin=15*mm, bottomMargin=16*mm,
                                title=f"Торговец — {display_name}", author="Генератор торговцев")
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle("MerchantTitle", parent=styles["Title"], fontName="MerchantBold",
            fontSize=23, leading=28, textColor=gold, alignment=TA_LEFT, spaceAfter=3*mm)
        subtitle_style = ParagraphStyle("MerchantSubtitle", parent=styles["Normal"], fontName="MerchantRegular",
            fontSize=10, leading=15, textColor=parchment)
        section_style = ParagraphStyle("MerchantSection", parent=styles["Heading2"], fontName="MerchantBold",
            fontSize=13, leading=17, textColor=dark, spaceBefore=5*mm, spaceAfter=3*mm)
        body_style = ParagraphStyle("MerchantBody", parent=styles["BodyText"], fontName="MerchantRegular",
            fontSize=9, leading=13, textColor=ink)
        small_style = ParagraphStyle("MerchantSmall", parent=body_style, fontSize=8, leading=11)

        story = []
        heading = [Paragraph(escape(display_name), title_style),
                   Paragraph(f"{escape(merchant['spec'])} · {escape(merchant['seller_rarity'])} торговец", subtitle_style),
                   Paragraph(f"Капитал: {format_price(merchant['capital'])} белли", subtitle_style)]
        if image_path:
            try:
                from PIL import Image as PILImage
                with PILImage.open(image_path) as im:
                    im.verify()
                art = PDFImage(image_path, width=42*mm, height=52*mm, kind="proportional", hAlign="CENTER")
                header_data = [[art, heading]]
                header_widths = [48*mm, 126*mm]
            except Exception as exc:
                raise ValueError(f"Не удалось открыть изображение: {exc}") from exc
        else:
            header_data = [[heading]]
            header_widths = [174*mm]
        header = Table(header_data, colWidths=header_widths)
        header.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), dark), ("BOX", (0,0), (-1,-1), 1.2, gold),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"), ("LEFTPADDING", (0,0), (-1,-1), 7*mm),
            ("RIGHTPADDING", (0,0), (-1,-1), 7*mm), ("TOPPADDING", (0,0), (-1,-1), 6*mm),
            ("BOTTOMPADDING", (0,0), (-1,-1), 6*mm),
        ]))
        story.extend([header, Spacer(1, 5*mm)])

        if comment:
            story.append(Paragraph("Заметка", section_style))
            story.append(Paragraph(escape(comment).replace("\n", "<br/>"), body_style))

        story.append(Paragraph("Ассортимент", section_style))
        rows = [["Кол-во", "Товар", "Редкость", "Цена за ед."]]
        total_value = 0
        for item in merchant.get("items", []):
            rows.append([str(item["qty"]), Paragraph(escape(str(item["name"])), body_style),
                         str(item.get("rarity", "—")), f"{format_price(item['price'])} белли"])
            total_value += int(item["qty"]) * int(item["price"])
        if len(rows) == 1:
            rows.append(["—", "Товары отсутствуют", "—", "—"])
        products = Table(rows, colWidths=[19*mm, 83*mm, 31*mm, 41*mm], repeatRows=1, hAlign="LEFT")
        products.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), brown), ("TEXTCOLOR", (0,0), (-1,0), colors.white),
            ("FONTNAME", (0,0), (-1,0), "MerchantBold"), ("FONTNAME", (0,1), (-1,-1), "MerchantRegular"),
            ("FONTSIZE", (0,0), (-1,-1), 8), ("LEADING", (0,0), (-1,-1), 11),
            ("BACKGROUND", (0,1), (-1,-1), parchment), ("ROWBACKGROUNDS", (0,1), (-1,-1), [parchment, colors.HexColor("#EAD7B1")]),
            ("TEXTCOLOR", (0,1), (-1,-1), ink), ("GRID", (0,0), (-1,-1), 0.45, colors.HexColor("#B69A6B")),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"), ("ALIGN", (0,0), (0,-1), "CENTER"),
            ("ALIGN", (2,1), (3,-1), "CENTER"), ("LEFTPADDING", (0,0), (-1,-1), 7),
            ("RIGHTPADDING", (0,0), (-1,-1), 7), ("TOPPADDING", (0,0), (-1,-1), 7),
            ("BOTTOMPADDING", (0,0), (-1,-1), 7),
        ]))
        story.append(products)
        story.append(Spacer(1, 4*mm))
        summary = Table([[Paragraph(f"Позиций: {len(merchant.get('items', []))}", body_style),
                          Paragraph(f"Стоимость всего количества: {format_price(total_value)} белли", body_style)]],
                        colWidths=[55*mm, 119*mm])
        summary.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#E4C98F")),
                                     ("BOX", (0,0), (-1,-1), 0.7, gold), ("PADDING", (0,0), (-1,-1), 8),
                                     ("VALIGN", (0,0), (-1,-1), "MIDDLE")]))
        story.append(summary)

        def paint_page(canvas, document):
            canvas.saveState()
            width, height = A4
            canvas.setFillColor(dark)
            canvas.rect(0, height-7*mm, width, 7*mm, fill=1, stroke=0)
            canvas.setStrokeColor(gold)
            canvas.setLineWidth(0.7)
            canvas.line(17*mm, 11*mm, width-17*mm, 11*mm)
            canvas.setFont("MerchantRegular", 7)
            canvas.setFillColor(brown)
            canvas.drawString(17*mm, 7*mm, "ТОРГОВЦЫ МОРЕЙ · КАРТОЧКА ПЕРСОНАЖА")
            canvas.drawRightString(width-17*mm, 7*mm, f"{document.page}")
            canvas.restoreState()

        doc.build(story, onFirstPage=paint_page, onLaterPages=paint_page)

    def auto_resize_columns(self):
        # Цена рассчитывается по самой длинной строке с запасом; при нехватке
        # места доступна горизонтальная прокрутка таблицы.
        self.tree.column("qty", width=80, minwidth=70, stretch=False)
        self.tree.column("remove", width=42, minwidth=42, stretch=False)
        price_width = max(110, len(self.tree.heading("price")["text"]) * 9 + 24)
        for row in self.tree.get_children():
            price_width = max(price_width, len(str(self.tree.set(row, "price"))) * 10 + 28)
        self.tree.column("price", width=price_width, minwidth=110, stretch=False)
        name_width = max(220, len(self.tree.heading("name")["text"]) * 10 + 24)
        for row in self.tree.get_children():
            name_width = max(name_width, min(700, len(str(self.tree.set(row, "name"))) * 9 + 24))
        self.tree.column("name", width=name_width, minwidth=180, stretch=True)


    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    try:
        App().run()
    except Exception as e:
        # Показываем ошибку загрузки базы, даже если окно приложения ещё не создано.
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Ошибка запуска", f"Не удалось запустить программу:\n{e}\n\nУбедитесь, что установлен openpyxl и файл База_товаров.xlsx лежит рядом с MAGAZIN.py.")
        root.destroy()