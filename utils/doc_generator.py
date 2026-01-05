import os
import json
from docx import Document
from datetime import datetime, timedelta

# ==========================================
# 📁 Пути проекта
# ==========================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(BASE_DIR)

TEMPLATE_DIR = os.path.join(PROJECT_DIR, "templates")
DATA_DIR = os.path.join(PROJECT_DIR, "data")
OUTPUT_DIR = os.path.join(PROJECT_DIR, "output")

CARS_FILE = os.path.join(DATA_DIR, "cars.json")
COUNTER_FILE = os.path.join(DATA_DIR, "contract_counter.json")

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ==========================================
# 🔧 Замена текста в DOCX с сохранением стиля
# ==========================================
def replace_text_in_doc(doc, replacements):

    def process_paragraph(paragraph):
        if not paragraph.runs:
            return

        full_text = "".join(run.text for run in paragraph.runs)
        new_text = full_text

        for key, value in replacements.items():
            new_text = new_text.replace(key, str(value))

        if new_text != full_text:
            ref_run = paragraph.runs[0]

            font_name = ref_run.font.name
            font_size = ref_run.font.size
            font_bold = ref_run.font.bold
            font_italic = ref_run.font.italic

            for r in paragraph.runs:
                r._element.getparent().remove(r._element)

            new_run = paragraph.add_run(new_text)
            new_run.font.name = font_name
            new_run.font.size = font_size
            new_run.font.bold = font_bold
            new_run.font.italic = font_italic

    for p in doc.paragraphs:
        process_paragraph(p)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    process_paragraph(p)

# ==========================================
# 🔢 Контрактный номер
# ==========================================
def load_contract_number():
    if not os.path.exists(COUNTER_FILE):
        return 1
    with open(COUNTER_FILE, "r", encoding="utf-8") as f:
        return json.load(f).get("last_number", 0) + 1

def save_contract_number(number):
    with open(COUNTER_FILE, "w", encoding="utf-8") as f:
        json.dump({"last_number": number}, f, ensure_ascii=False, indent=4)

# ==========================================
# 📝 Генерация документов
# ==========================================
def generate_docs(data, client_name):
    for filename in os.listdir(TEMPLATE_DIR):
        if not filename.endswith(".docx") or filename.startswith("~$"):
            continue

        template_path = os.path.join(TEMPLATE_DIR, filename)

        if "contract" in filename.lower():
            out_name = f"contract_{client_name}.docx"
        elif "poa" in filename.lower():
            out_name = f"poa_{client_name}.docx"
        elif "waybill" in filename.lower():
            out_name = f"waybill_{client_name}.docx"
        else:
            out_name = f"{client_name}_{filename}"

        doc = Document(template_path)
        replace_text_in_doc(doc, data)

        out_path = os.path.join(OUTPUT_DIR, out_name)
        doc.save(out_path)
        print(f"✅ Создан файл: {out_path}")

# ==========================================
# 🚗 Загрузка авто
# ==========================================
def load_cars():
    with open(CARS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

# ==========================================
# 📆 Даты
# ==========================================
def calculate_days(start, end):
    d1 = datetime.strptime(start, "%d.%m.%Y")
    d2 = datetime.strptime(end, "%d.%m.%Y")
    return (d2 - d1).days + 1

def extend_date(date_str, days=3):
    d = datetime.strptime(date_str, "%d.%m.%Y")
    return (d + timedelta(days=days)).strftime("%d.%m.%Y")

# ==========================================
# 🛣️ Типы дорог (Permission Kazakhstan)
# ==========================================
def choose_road_types():
    options = ["Asphalt", "Gravel", "Dirt", "Off-road"]
    print("\nTypes of Roads:")
    for i, o in enumerate(options, 1):
        print(f"{i}. {o}")

    choice = input("Выберите (через запятую): ").strip()
    result = []

    for c in choice.split(","):
        if c.strip().isdigit():
            idx = int(c.strip())
            if 1 <= idx <= len(options):
                result.append(options[idx - 1])

    return ", ".join(result) if result else "Asphalt"

# ==========================================
# 🌍 Доп. страны
# ==========================================
def choose_additional_countries():
    options = {
        1: "Кыргызстан",
        2: "Узбекистан",
        3: "Таджикистан"
    }

    print("\nДополнительные страны:")
    for i, name in options.items():
        print(f"{i}. {name}")

    choice = input("Выбор (через запятую или Enter): ").strip()
    result = []

    for c in choice.split(","):
        if c.strip().isdigit() and int(c.strip()) in options:
            result.append(options[int(c.strip())])

    return result

# ==========================================
# 🚀 MAIN
# ==========================================
if __name__ == "__main__":

    cars = load_cars()
    for i, c in enumerate(cars, 1):
        print(f"{i}. {c['make']} {c['model']} ({c['plate']})")
    selected_car = cars[int(input("Выбор авто: ")) - 1]

    client_name = input("ФИО клиента: ")
    date_of_birth = input("Дата рождения: ")
    address = input("Адрес: ")
    phone = input("Телефон: ")
    email = input("Email: ")

    passport_number = input("Паспорт №: ")
    passport_issue_date = input("Дата выдачи: ")
    passport_issue_by = input("Кем выдан: ")
    license_num = input("ВУ №: ")

    start_date = input("Начало аренды: ")
    end_date = input("Конец аренды: ")

    rental_rate = float(input("Цена за сутки USD: "))
    days = calculate_days(start_date, end_date)
    total = rental_rate * days
    deposit = float(input("Залог USD: "))

    # ===== ДОП ВОДИТЕЛИ =====
    drivers = {}
    for i in range(1, 4):
        drivers[f"{{{{DRIVER{i}_NAME}}}}"] = ""
        drivers[f"{{{{DRIVER{i}_LICENSE}}}}"] = ""
        drivers[f"{{{{PASSPORT{i}_NUMBER}}}}"] = ""
        drivers[f"{{{{PASSPORT{i}_ISSUE_DATE}}}}"] = ""
        drivers[f"{{{{PASSPORT{i}_ISSUE_BY}}}}"] = ""

    if input("Доп. водители? (да/нет): ").lower() == "да":
        count = int(input("Сколько (1–3): "))
        for i in range(count):
            drivers[f"{{{{DRIVER{i+1}_NAME}}}}"] = input("Имя: ")
            drivers[f"{{{{DRIVER{i+1}_LICENSE}}}}"] = input("ВУ: ")
            drivers[f"{{{{PASSPORT{i+1}_NUMBER}}}}"] = input("Паспорт №: ")
            drivers[f"{{{{PASSPORT{i+1}_ISSUE_DATE}}}}"] = input("Дата выдачи: ")
            drivers[f"{{{{PASSPORT{i+1}_ISSUE_BY}}}}"] = input("Кем выдан: ")

    road_types = choose_road_types()
    extra_countries = choose_additional_countries()

    allowed_countries = "Казахстан"
    if extra_countries:
        allowed_countries += ", " + ", ".join(extra_countries)

    outside_kz_block = (
        f"и за ее пределами ({', '.join(extra_countries)})"
        if extra_countries else ""
    )

    contract_number = load_contract_number()
    save_contract_number(contract_number)

    data = {
        "{{CONTRACT_DATE}}": datetime.now().strftime("%d.%m.%Y"),
        "{{CONTRACT_NUMBER}}": contract_number,
        "{{CLIENT_NAME}}": client_name,
        "{{DATE_OF_BIRTH}}": date_of_birth,
        "{{ADDRESS}}": address,
        "{{PHONE}}": phone,
        "{{EMAIL}}": email,
        "{{PASSPORT_NUMBER}}": passport_number,
        "{{PASSPORT_ISSUE_DATE}}": passport_issue_date,
        "{{PASSPORT_ISSUE_BY}}": passport_issue_by,
        "{{DRIVER_LICENSE}}": license_num,

        "{{RENTAL_START}}": start_date,
        "{{RENTAL_END}}": end_date,
        "{{RENTAL_END_EXTENDED}}": extend_date(end_date),

        "{{RENTAL_RATE}}": f"{rental_rate:.2f}",
        "{{TOTAL_AMOUNT}}": f"{total:.2f}",
        "{{SECURITY_DEPOSIT}}": f"{deposit:.2f}",

        "{{ALLOWED_COUNTRIES}}": allowed_countries,
        "{{TYPES_OF_ROADS}}": road_types,
        "{{OUTSIDE_KZ_BLOCK}}": outside_kz_block,

        "{{CAR_MAKE}}": selected_car["make"],
        "{{CAR_MODEL}}": selected_car["model"],
        "{{CAR_NAME}}": f"{selected_car['make']} {selected_car['model']}",
        "{{CAR_YEAR}}": selected_car["year"],
        "{{CAR_COLOR}}": selected_car["color"],
        "{{CAR_PLATE}}": selected_car["plate"],
        "{{CAR_VIN}}": selected_car["vin"]
    }

    data.update(drivers)
    generate_docs(data, client_name.replace(" ", "_"))
