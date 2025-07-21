!pip show pandas openpyxl


import pandas as pd
import re

# Пути к файлам
source_file = './ВЭД гр 31 март 2025.xlsx'
product_file = './Products.xlsx'
output_file = './output.xlsx'

# Название колонки с ТН ВЭД в исходном файле
tnved_col = "G33 (код товара по ТН ВЭД РФ)"

# Загружаем исходные данные
df_source = pd.read_excel(source_file)

# Загружаем справочник ТН ВЭД -> Вид МУ
df_product = pd.read_excel(product_file, sheet_name='ВЭД')

# Создаём словарь соответствий ТН ВЭД -> Вид МУ
product_map = dict(zip(df_product[tnved_col], df_product['Вид МУ']))

# Добавляем колонку Product
df_new = df_source.copy()
df_new['Product'] = df_new[tnved_col].map(product_map)


# ======================================
# Функции для определения Grade по описанию товара
# ======================================
def extract_npk(description):
    desc = str(description).lower().strip()
    desc = re.sub(r'[\s\xa0\u3000]+', ' ', desc)

    # Попытка найти явный формат NPK: 12:32:16 или 16-16-16 или NPK 16:16:16
    npk_match = re.search(
        r'\b(?:npk\s*)?(\d+(?:\.\d+)?)\s*[:-]\s*(\d+(?:\.\d+)?)\s*[:-]\s*(\d+(?:\.\d+)?)',
        desc, re.IGNORECASE)
    if npk_match:
        n = float(npk_match.group(1))
        p = float(npk_match.group(2))
        k = float(npk_match.group(3))

        # Проверяем на > 100
        n = int(n) if n <= 100 and n == int(n) else (n if n <= 100 else 0)
        p = int(p) if p <= 100 and p == int(p) else (p if p <= 100 else 0)
        k = int(k) if k <= 100 and k == int(k) else (k if k <= 100 else 0)

        return {
            'N': {'value': n},
            'P': {'value': p},
            'K': {'value': k}
        }

    # === ОПРЕДЕЛЕНИЕ elements ===
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

    # === ДАЛЬШЕ ИДЁТ ОСНОВНОЙ ЦИКЛ ПО elements ===
    for el_key, data in elements.items():
        for keyword in data['keywords']:
            pattern = rf'{keyword}\D*?(\d+(?:[,.]\d+)?)(?=\s*(?:%|мас|в пересчёте|марка|гост|п/п|кг|л|литров|литра|мешк|пакет|упаковк|порошок|гранулы|таблетк|вес|брутто|нетто|пластик|бумажн|поддон|паллет|предназначен|используется|входит|содержит|состав|марка|не более|не менее|не превышает|минимум|максимум|,|\.|;|:|$))'
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

    # Обработка "в пересчёте на K2O"
    k2o_match = re.search(r'в\s*пересч[ёе]те.*?k2o\D*(\d+(?:[,.]\d+)?)', desc, re.IGNORECASE)
    if k2o_match:
        try:
            k_value = float(k2o_match.group(1).replace(',', '.'))
            elements['K']['value'] = int(k_value) if k_value == int(k_value) else k_value
        except ValueError:
            pass

    # Обработка "в пересчёте на P2O5"
    p2o5_match = re.search(r'в\s*пересч[ёе]те.*?p2o5\D*(\d+(?:[,.]\d+)?)', desc, re.IGNORECASE)
    if p2o5_match:
        try:
            p2o5_value = float(p2o5_match.group(1).replace(',', '.'))
            elements['P']['value'] = int(p2o5_value * 0.436) if (p2o5_value * 0.436) == int(p2o5_value * 0.436) else p2o5_value * 0.436
        except ValueError:
            pass

    # Обработка "содержание азота"
    total_n_match = re.search(r'(?:содержание|содержит|общий|содержание азота).*?(\d+(?:[,.]\d+)?)', desc, re.IGNORECASE)
    if total_n_match:
        try:
            n_value = float(total_n_match.group(1).replace(',', '.'))
            elements['N']['value'] = int(n_value) if n_value == int(n_value) else n_value
        except ValueError:
            pass

    return {k: v for k, v in elements.items()}


def determine_grade(description, product):
    """Возвращает строку вида X-X-X на основе описания и типа Product"""
    result = extract_npk(description)

    # Извлекаем значения
    n = result['N']['value']
    p = result['P']['value']
    k = result['K']['value']

    # Проверяем на корректность
    n = n if isinstance(n, (int, float)) and n > 0 and n <= 100 else 0
    p = p if isinstance(p, (int, float)) and p > 0 and p <= 100 else 0
    k = k if isinstance(k, (int, float)) and k > 0 and k <= 100 else 0

    # Фильтруем по типу продукта
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

    # Округляем, если целое
    n = int(n) if isinstance(n, float) and n == int(n) else n
    p = int(p) if isinstance(p, float) and p == int(p) else p
    k = int(k) if isinstance(k, float) and k == int(k) else k

    grade = f"{n}-{p}-{k}"

    if grade == "0-0-0":
        return "X-X-X"
    return grade


# ======================================
# Применяем функции к данным
# ======================================

# Проверяем наличие нужного столбца
desc_col = "G31_1 (Описание и характеристика товара)"
if desc_col not in df_new.columns:
    raise KeyError(f"❌ В таблице отсутствует колонка: '{desc_col}'")

# Добавляем Grade
def apply_determine_grade(row):
    return determine_grade(row[desc_col], row['Product'])

df_new['Grade'] = df_new.apply(apply_determine_grade, axis=1)

# Список разрешённых типов Product
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

# Очищаем Grade, если Product не в списке разрешённых
df_new['Grade'] = df_new.apply(
    lambda row: row['Grade'] if row['Product'] in allowed_product_types else '',
    axis=1
)

# Сохраняем в новый файл Excel с новым листом "Лист 1"
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df_new.to_excel(writer, sheet_name='Лист 1', index=False)

print("✅ Файл успешно обработан и сохранён как:", output_file)
