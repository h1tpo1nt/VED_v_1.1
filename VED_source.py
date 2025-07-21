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

    # Формат NPK: например, NPK 16-16-16 или NPK(16:16:16)
    npk_match = re.search(
        r'\bnpk\s*(?:$s$)?\s*(\d+(?:\.\d+)?)\s*[-:/]\s*(\d+(?:\.\d+)?)\s*[-:/]\s*(\d+(?:\.\d+)?)',
        desc, re.IGNORECASE)
    if npk_match:
        n = float(npk_match.group(1))
        p = float(npk_match.group(2))
        k = float(npk_match.group(3))
        return {
            'N': int(n) if n.is_integer() else n,
            'P': int(p) if p.is_integer() else p,
            'K': int(k) if k.is_integer() else k,
            'full_match': True
        }

    # Словарь элементов
    elements = {
        'N': {'keywords': [r'\bазот', r'\bnитрат', r'\bn\s*содержащие'], 'value': 0},
        'P': {'keywords': [r'\bфосфор', r'\bp2o5', r'\bп2о5', r'\bphosphorus'], 'value': 0},
        'K': {'keywords': [r'\bкали[йяие]', r'\bk2o', r'\bкалийные'], 'value': 0}
    }

    for el_key, data in elements.items():
        for keyword in data['keywords']:
            pattern = rf'{keyword}\D*?(\d+(?:[,.]\d+)?)%?'
            match = re.search(pattern, desc, re.IGNORECASE)
            if match:
                try:
                    value = float(match.group(1).replace(',', '.'))
                    data['value'] = int(value) if value.is_integer() else value
                except ValueError:
                    continue
                break

    return {
        'N': elements['N']['value'],
        'P': elements['P']['value'],
        'K': elements['K']['value'],
        'full_match': False
    }


def determine_grade(description):
    """Возвращает строку вида 'N-P-K' на основе описания товара"""
    result = extract_npk(description)

    n = result['N']
    p = result['P']
    k = result['K']

    # Приводим к строке, сохраняя формат (целое или дробное)
    brand = f"{n}-{p}-{k}"

    # Если все нули — возвращаем 'NPK'
    if brand == "0-0-0":
        return "NPK"
    else:
        return brand


# ======================================
# Применяем функции к данным
# ======================================

# Проверяем наличие нужного столбца
desc_col = "G31_1 (Описание и характеристика товара)"
if desc_col not in df_new.columns:
    raise KeyError(f"❌ В таблице отсутствует колонка: '{desc_col}'")

# Добавляем Grade
df_new['Grade'] = df_new[desc_col].apply(determine_grade)

# Определяем, для каких типов Product нужно оставлять Grade
allowed_product_types = {
    "НПК",
    "NP",
    "НПК в таблетках или упаковке менее 10 кг",
    "Прочие NP/NPK",
    "PK"
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
