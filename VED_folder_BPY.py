!pip show pandas openpyxl

import pandas as pd
import re
import os

# Папки
input_folder = './input'
output_folder = './output'
os.makedirs(output_folder, exist_ok=True)

# Файл со справочником
product_file = './Products.xlsx'
tnved_col = "G33 (код товара по ТН ВЭД РФ)"
df_product = pd.read_excel(product_file, sheet_name='ВЭД')
product_map = dict(zip(df_product[tnved_col], df_product['Вид МУ']))

# ==== ФУНКЦИИ ====

def extract_npk(description):
    desc = str(description).lower().strip()
    desc = re.sub(r'[\s\xa0\u3000]+', ' ', desc)

    # --- удаляем паттерны ТУ ---
    desc = re.sub(r'ту\s*\d{2}\.\d{2}\.\d{2}-\d{3}-\d{4}', '', desc)
    desc = re.sub(r'ту\s*\d{2}\.\d{2}\.\d{2}-\d{4}-\d{4}', '', desc)

    # формат NPK x:x:x или x-x-x
    npk_match = re.search(
        r'\b(?:npk\s*)?(\d+(?:\.\d+)?)\s*[:-]\s*(\d+(?:\.\d+)?)\s*[:-]\s*(\d+(?:\.\d+)?)',
        desc, re.IGNORECASE
    )
    if npk_match:
        n = float(npk_match.group(1))
        p = float(npk_match.group(2))
        k = float(npk_match.group(3))

        n = int(n) if n <= 100 and n == int(n) else (n if n <= 100 else 0)
        p = int(p) if p <= 100 and p == int(p) else (p if p <= 100 else 0)
        k = int(k) if k <= 100 and k == int(k) else (k if k <= 100 else 0)

        return {'N': {'value': n}, 'P': {'value': p}, 'K': {'value': k}}

    # поиск по ключам
    elements = {
        'N': {'keywords': [
            r'\bазот', r'\bnитрат', r'\bn\s*содержащие', r'\bсодержание\s*азота',
            r'\bаммонийный\s*азот', r'\bнитрат', r'\bn\s*общий', r'\bаммиачный\s*азот'
        ], 'value': 0},
        'P': {'keywords': [
            r'\bфосфор', r'\bp2o5', r'\bп2о5', r'\bphosphorus',
            r'\bсодержание\s*фосфора', r'\bфосфаты'
        ], 'value': 0},
        'K': {'keywords': [
            r'\bкали[йяие]', r'\bk2o', r'\bкалийные', r'\bсодержание\s*калия'
        ], 'value': 0},
        'Ca': {'keywords': [
            r'\bкальций', r'\bcao', r'\bca\s*содержащие', r'\bизвесть',
            r'\bкарбонат\s*кальца', r'\bсодержание\s*кальция', r'\bcacо3'
        ], 'value': 0}
    }

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

def determine_grade(description, product):
    result = extract_npk(description)
    n = result['N']['value']
    p = result['P']['value']
    k = result['K']['value']

    n = n if isinstance(n, (int, float)) and 0 < n <= 100 else 0
    p = p if isinstance(p, (int, float)) and 0 < p <= 100 else 0
    k = k if isinstance(k, (int, float)) and 0 < k <= 100 else 0

    if product == 'Калий':
        n = 0; p = 0
    elif product == 'NP':
        k = 0
    elif product == 'PK':
        n = 0
    elif product == 'NS':
        p = 0; k = 0
    elif product == 'Ca':
        n = 0; p = 0; k = 0

    n = int(n) if isinstance(n, float) and n == int(n) else n
    p = int(p) if isinstance(p, float) and p == int(p) else p
    k = int(k) if isinstance(k, float) and k == int(k) else k

    grade = f"{n}-{p}-{k}"
    return "X-X-X" if grade == "0-0-0" else grade

def check_all_less_than_one(grade):
    if not grade or grade == 'X-X-X':
        return grade
    parts = grade.split('-')
    try:
        nums = [float(x) for x in parts]
        if all(x < 1 for x in nums):
            return ''
    except ValueError:
        pass
    return grade

def check_product_type(row, desc_col):
    if row['Product'] in ['НПК', 'Прочие NP/NPK']:
        if pd.notna(row[desc_col]) and re.search(r'водорастворим\w*', str(row[desc_col]).lower()):
            return 'ВРУ'
    return ''

allowed_product_types = {
    "НПК","МАФ","Карбамид","Прочие удобрения животного или растительного происхождения",
    "Прочие фосфорные удобрения","PK","CAN","AN","Прочие NP/NPK",
    "НПК в таблетках или упаковке менее 10 кг","AS","КАС","Калий",
    "SOP","ДАФ","NP","Нитрат натрия","NS","CN",
    "Прочие калийные удобрения","Удобрения животного или растительного происхождения",
    "Прочие суперфосфаты","Суперфосфаты"
}

# ==== ЦИКЛ ====
files = [f for f in os.listdir(input_folder) if f.lower().endswith('.xlsx')]
for i, fname in enumerate(files, 1):
    in_path = os.path.join(input_folder, fname)
    out_path = os.path.join(output_folder, f"{os.path.splitext(fname)[0]} SORTING.xlsx")

    df_source = pd.read_excel(in_path)
    if "G31_1 (Описание и характеристика товара)" not in df_source.columns:
        raise KeyError("❌ Нет колонки 'G31_1 (Описание и характеристика товара)'")

    df_new = df_source.copy()
    df_new['Product'] = df_new[tnved_col].map(product_map)

    df_new['Grade'] = df_new.apply(lambda r: determine_grade(r["G31_1 (Описание и характеристика товара)"], r['Product']), axis=1)
    df_new['Grade'] = df_new.apply(lambda r: r['Grade'] if r['Product'] in allowed_product_types else '', axis=1)
    df_new['Grade'] = df_new['Grade'].apply(check_all_less_than_one)

    df_new['Product Type'] = df_new.apply(check_product_type, axis=1, desc_col="G31_1 (Описание и характеристика товара)")

    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        df_new.to_excel(writer, sheet_name='Лист 1', index=False)

    print(f"✅ {i}/{len(files)} готово → {out_path}")

print("🎯 Все файлы обработаны!")
