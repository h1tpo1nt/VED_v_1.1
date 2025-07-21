import pandas as pd
import re
import os

# === Функция поиска колонки по префиксу ===
def find_column(df, prefix):
    """
    Находит первую колонку, имя которой начинается с заданного префикса.
    """
    for col in df.columns:
        if col.startswith(prefix):
            return col
    raise KeyError(f"❌ В таблице отсутствует колонка с префиксом '{prefix}'")


# === Пути ===
SOURCE_FOLDER = './input'
OUTPUT_FOLDER = './output'

# === Префиксы колонок ===
tnved_col_prefix = "G33"
desc_col_prefix = "G31_1"

# === Загружаем справочник ТН ВЭД -> Вид МУ ===
product_file = './Products.xlsx'
df_product = pd.read_excel(product_file, sheet_name='ВЭД')

# Находим правильные имена колонок в справочнике
tnved_col_real = find_column(df_product, tnved_col_prefix)
product_map = dict(zip(df_product[tnved_col_real], df_product['Вид МУ']))

# === Создаём папку для готовых файлов, если её нет ===
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


# === Функция extract_npk (без изменений) ===
def extract_npk(description):
    desc = str(description).lower().strip()
    desc = re.sub(r'[\s\xa0\u3000]+', ' ', desc)

    npk_match = re.search(
        r'\b(?:npk\s*)?(\d+(?:\.\d+)?)\s*[:-]\s*(\d+(?:\.\d+)?)\s*[:-]\s*(\d+(?:\.\d+)?)',
        desc, re.IGNORECASE)
    if npk_match:
        n = float(npk_match.group(1))
        p = float(npk_match.group(2))
        k = float(npk_match.group(3))

        n = int(n) if n <= 100 and n == int(n) else (n if n <= 100 else 0)
        p = int(p) if p <= 100 and p == int(p) else (p if p <= 100 else 0)
        k = int(k) if k <= 100 and k == int(k) else (k if k <= 100 else 0)

        return {
            'N': {'value': n},
            'P': {'value': p},
            'K': {'value': k}
        }

    elements = {
        'N': {
            'keywords': [
                r'\bазот', r'\bnитрат', r'\bn\s*содержащие', r'\bсодержание\s*азота',
                r'\bаммонийный\s*азот', r'\bнитрат', r'\bn\s*общий', r'\bаммиачный\s*азот'
            ],
            'value': 0
        },
        'P': {
            'keywords': [
                r'\bфосфор', r'\bp2o5', r'\bп2о5', r'\bphosphorus',
                r'\bсодержание\s*фосфора', r'\bфосфаты'
            ],
            'value': 0
        },
        'K': {
            'keywords': [
                r'\bкали[йяие]', r'\bk2o', r'\bкалийные', r'\bсодержание\s*калия'
            ],
            'value': 0
        },
        'Ca': {
            'keywords': [
                r'\bкальций', r'\bcao', r'\bca\s*содержащие', r'\bизвесть',
                r'\bкарбонат\s*кальца', r'\bсодержание\s*кальция', r'\bcacо3'
            ],
            'value': 0
        }
    }

    for el_key, data in elements.items():
        for keyword in data['keywords']:
            pattern = rf'{keyword}\D*?(\d+(?:[,.]\d+)?)(?=\s*(?:%|мас|в пересчёте|гост|п/п|кг|л|литров|литра|мешк|пакет|упаковк|порошок|гранулы|таблетк|вес|брутто|нетто|пластик|бумажн|поддон|паллет|предназначен|используется|входит|содержит|состав|марка|не более|не менее|не превышает|минимум|максимум|,|\.|;|:|$))'
            match = re.search(pattern, desc, re.IGNORECASE)
            if match:
                try:
                    value = float(match.group(1).replace(',', '.'))
                    if value > 100:
                        value = 0
                    data['value'] = int(value) if value == int(value) else value
                except ValueError:
                    continue
                break

    k2o_match = re.search(r'в\s*пересч[ёе]те.*?k2o\D*(\d+(?:[,.]\d+)?)', desc, re.IGNORECASE)
    if k2o_match:
        try:
            k_value = float(k2o_match.group(1).replace(',', '.'))
            elements['K']['value'] = int(k_value) if k_value == int(k_value) else k_value
        except ValueError:
            pass

    p2o5_match = re.search(r'в\s*пересч[ёе]те.*?p2o5\D*(\d+(?:[,.]\d+)?)', desc, re.IGNORECASE)
    if p2o5_match:
        try:
            p2o5_value = float(p2o5_match.group(1).replace(',', '.'))
            elements['P']['value'] = int(p2o5_value * 0.436) if (p2o5_value * 0.436) == int(p2o5_value * 0.436) else p2o5_value * 0.436
        except ValueError:
            pass

    total_n_match = re.search(r'(?:содержание|содержит|общий|содержание азота).*?(\d+(?:[,.]\d+)?)', desc, re.IGNORECASE)
    if total_n_match:
        try:
            n_value = float(total_n_match.group(1).replace(',', '.'))
            elements['N']['value'] = int(n_value) if n_value == int(n_value) else n_value
        except ValueError:
            pass

    return {k: v for k, v in elements.items()}


# === Функция determine_grade (без изменений) ===
def determine_grade(description, product):
    result = extract_npk(description)

    n = result['N']['value']
    p = result['P']['value']
    k = result['K']['value']

    n = n if isinstance(n, (int, float)) and n > 0 and n <= 100 else 0
    p = p if isinstance(p, (int, float)) and p > 0 and p <= 100 else 0
    k = k if isinstance(k, (int, float)) and k > 0 and k <= 100 else 0

    if product == 'Калий':
        n = 0
        p = 0
    elif product == 'NP':
        k = 0
    elif product == 'PK':
        n = 0
    elif product == 'NS':
        p = 0
        k = 0
    elif product == 'Ca':
        n = 0
        p = 0
        k = 0

    n = int(n) if isinstance(n, float) and n == int(n) else n
    p = int(p) if isinstance(p, float) and p == int(p) else p
    k = int(k) if isinstance(k, float) and k == int(k) else k

    grade = f"{n}-{p}-{k}"

    return "X-X-X" if grade == "0-0-0" else grade


# === Список разрешённых типов продуктов ===
allowed_product_types = {
    "НПК",
    "МАФ",
    "Карбамид",
    "Прочие удобрения животного или растительного происхождения",
    "Прочие фосфорные удобрения",
    "PK",
    "CAN",
    "AN",
    "Прочие NP/NPK",
    "НПК в таблетках или упаковке менее 10 кг",
    "AS",
    "КАС",
    "Калий",
    "SOP",
    "ДАФ",
    "NP",
    "Нитрат натрия",
    "NS",
    "CN",
    "Прочие калийные удобрения",
    "Удобрения животного или растительного происхождения",
    "Прочие суперфосфаты",
    "Суперфосфаты"
}

# === ОСНОВНОЙ ЦИКЛ ПО ФАЙЛАМ ===
for filename in os.listdir(SOURCE_FOLDER):
    if filename.endswith('.xlsx'):
        source_file = os.path.join(SOURCE_FOLDER, filename)
        output_file = os.path.join(OUTPUT_FOLDER, f"{os.path.splitext(filename)[0]} SORTING.xlsx")

        try:
            # Загружаем исходные данные
            df_source = pd.read_excel(source_file)

            # Находим нужные колонки по префиксу
            tnved_col_real = find_column(df_source, tnved_col_prefix)
            desc_col_real = find_column(df_source, desc_col_prefix)

            # Добавляем колонку Product
            df_new = df_source.copy()
            df_new['Product'] = df_new[tnved_col_real].map(product_map)

            # Добавляем Grade
            df_new['Grade'] = df_new.apply(
                lambda row: determine_grade(row[desc_col_real], row['Product']), axis=1
            )

            # Очищаем Grade, если Product не в списке разрешённых
            df_new['Grade'] = df_new.apply(
                lambda row: row['Grade'] if row['Product'] in allowed_product_types else '',
                axis=1
            )

            # Сохраняем в новый файл
            df_new.to_excel(output_file, sheet_name='Лист 1', index=False)
            print(f"✅ Обработано: {filename} → {os.path.basename(output_file)}")

        except Exception as e:
            print(f"❌ Ошибка при обработке файла {filename}: {e}")
