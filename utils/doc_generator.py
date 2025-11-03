import os
import json
from docx import Document
from datetime import datetime

# === Пути ===
TEMPLATE_DIR = r"E:\training\Python\IHNscript\templates"
OUTPUT_DIR = r"E:\training\Python\IHNscript\output"
CARS_FILE = r"E:\training\Python\IHNscript\data\cars.json"

# === Создание выходной папки ===
os.makedirs(OUTPUT_DIR, exist_ok=True)

# === Функция замены текста ===
def replace_text_in_doc(doc, replacements):
    for paragraph in doc.paragraphs:
        full_text = ''.join(run.text for run in paragraph.runs)
        new_text = full_text
        for key, value in replacements.items():
            new_text = new_text.replace(key, str(value))
        if new_text != full_text:
            for i in range(len(paragraph.runs) - 1, -1, -1):
                p = paragraph.runs[i]._element
                p.getparent().remove(p)
            paragraph.add_run(new_text)

    # Таблицы
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    full_text = ''.join(run.text for run in paragraph.runs)
                    new_text = full_text
                    for key, value in replacements.items():
                        new_text = new_text.replace(key, str(value))
                    if new_text != full_text:
                        for i in range(len(paragraph.runs) - 1, -1, -1):
                            p = paragraph.runs[i]._element
                            p.getparent().remove(p)
                        paragraph.add_run(new_text)

# === Генерация документов ===
def generate_docs(data_dict, client_name):
    if not os.path.exists(TEMPLATE_DIR):
        print(f"❌ Папка шаблонов не найдена: {TEMPLATE_DIR}")
        return

    for filename in os.listdir(TEMPLATE_DIR):
        if filename.endswith(".docx"):
            template_path = os.path.join(TEMPLATE_DIR, filename)

            if "contract" in filename.lower():
                new_name = f"contract_{client_name}.docx"
            elif "poa" in filename.lower():
                new_name = f"poa_{client_name}.docx"
            elif "waybill" in filename.lower():
                new_name = f"waybill_{client_name}.docx"
            else:
                new_name = f"{client_name}_{filename}"

            output_path = os.path.join(OUTPUT_DIR, new_name)
            print(f"📄 Обрабатываю шаблон: {filename}")

            doc = Document(template_path)
            replace_text_in_doc(doc, data_dict)
            doc.save(output_path)

            print(f"✅ Файл сохранён: {output_path}\n")

# === Загрузка данных машин ===
def load_cars():
    if not os.path.exists(CARS_FILE):
        print(f"❌ Не найден cars.json: {CARS_FILE}")
        return []
    with open(CARS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

# === Подсчёт количества дней ===
def calculate_days(start_date, end_date):
    fmt = "%d.%m.%Y"
    d1 = datetime.strptime(start_date, fmt)
    d2 = datetime.strptime(end_date, fmt)
    return (d2 - d1).days + 1  # включая последний день

# === Выбор покрытия дорог ===
def choose_road_types():
    road_options = [
        "Paved",
        "Gravel",
        "Dirt Tracks",
        "Off-Road",
        "Asphalt"
    ]
    print("\n🛣️ Доступные типы дорог:")
    for i, r in enumerate(road_options, 1):
        print(f"{i}. {r}")

    print("\nМожно выбрать несколько (через запятую, например: 1,3,5)")
    choice = input("Выберите типы дорог: ").strip()
    selected = []
    if choice:
        for ch in choice.split(","):
            ch = ch.strip()
            if ch.isdigit() and 1 <= int(ch) <= len(road_options):
                selected.append(road_options[int(ch) - 1])
    return ", ".join(selected) if selected else "Paved"

# === Основной запуск ===
if __name__ == "__main__":
    cars = load_cars()
    if not cars:
        print("🚫 Нет данных об автомобилях!")
        exit()

    print("🚗 Доступные автомобили:")
    for i, car in enumerate(cars, 1):
        print(f"{i}. {car['make']} {car['model']} ({car['plate']})")

    choice = int(input("Выбери номер машины: ")) - 1
    selected_car = cars[choice]

    print("\nВведите данные клиента:")
    client_name = input("Имя клиента (Фамилия Имя): ")
    date_of_birth = input("Дата рождения (ДД.ММ.ГГГГ): ")
    address = input("Адрес проживания: ")
    phone = input("Телефон: ")
    email = input("Email: ")
    passport = input("Паспорт: ")
    license_num = input("Вод. удостоверение: ")
    start_date = input("Дата начала аренды (ДД.ММ.ГГГГ): ")
    end_date = input("Дата конца аренды (ДД.ММ.ГГГГ): ")

    rental_rate = float(input("Цена за сутки (USD): "))
    days = calculate_days(start_date, end_date)
    total_amount = rental_rate * days
    security_deposit = float(input("Сумма залога (USD): "))

    print("\nЕсть ли дополнительные водители? (да/нет)")
    add_drivers = input().strip().lower()
    driver_data = {"{{DRIVER1_NAME}}": "", "{{DRIVER1_LICENSE}}": "",
                   "{{DRIVER2_NAME}}": "", "{{DRIVER2_LICENSE}}": "",
                   "{{DRIVER3_NAME}}": "", "{{DRIVER3_LICENSE}}": ""}

    if add_drivers == "да":
        num = int(input("Сколько дополнительных водителей (до 3)? "))
        for i in range(num):
            name = input(f"Имя водителя {i+1}: ")
            lic = input(f"Номер прав водителя {i+1}: ")
            driver_data[f"{{{{DRIVER{i+1}_NAME}}}}"] = name
            driver_data[f"{{{{DRIVER{i+1}_LICENSE}}}}"] = lic

    road_types = choose_road_types()

    data = {
        "{{CLIENT_NAME}}": client_name,
        "{{DATE_OF_BIRTH}}": date_of_birth,
        "{{ADDRESS}}": address,
        "{{PHONE}}": phone,
        "{{EMAIL}}": email,
        "{{PASSPORT_NUMBER}}": passport,
        "{{DRIVER_LICENSE}}": license_num,
        "{{RENTAL_START}}": start_date,
        "{{RENTAL_END}}": end_date,
        "{{RENTAL_RATE}}": rental_rate,
        "{{TOTAL_AMOUNT}}": f"{total_amount:.2f}",
        "{{SECURITY_DEPOSIT}}": f"{security_deposit:.2f}",
        # --- Машина ---
        "{{CAR_MAKE}}": selected_car["make"],
        "{{CAR_MODEL}}": selected_car["model"],
        "{{CAR_YEAR}}": selected_car["year"],
        "{{CAR_COLOR}}": selected_car["color"],
        "{{CAR_PLATE}}": selected_car["plate"],
        "{{CAR_VIN}}": selected_car["vin"],
        # --- Остальные ---
        "{{ALLOWED_TERRITORIES}}": "KZ, KGZ, UZ, TJ",
        "{{TYPES_OF_ROADS}}": road_types
    }

    data.update(driver_data)

    generate_docs(data, client_name.replace(" ", "_"))
