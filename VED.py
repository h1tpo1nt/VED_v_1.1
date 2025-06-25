import pandas as pd
import os

# Конфигурация
input_file = "исходный_файл.xlsx"
output_folder = "результаты/"

# Соответствие: {имя_выходного_файла: [листы для обновления]}
file_sheet_mapping = {
    "01_Калиевая селитра.xlsx": ["Данные"],
    "02_Формиаты.xlsx": ["Данные"],
    "03_Нитрат кальция.xlsx": ["Данные"],
    "05_Нитрит натрия.xlsx": ["Данные"],
    "06_МАФ.xlsx": ["Данные"],
    "07_NPK(S) ВРУ.xlsx": ["Данные для объемов", "Данные для цен"],
    "08_Сульфат магния.xlsx": ["Данные"],
    "09_Монокалийфосфат.xlsx": ["Данные"],
    "Экспорт NPK ВРУ.xlsx": ["Данные"],
    "Экспорт МАФ 12-61.xlsx": ["Данные"],
    "Экспорт Монокалийфосфата.xlsx": ["Лист1"],
    "Экспорт Сульфата магния.xlsx": ["Данные"]
}

# Полное соответствие: {код ТН ВЭД -> имя файла}
code_file_map = {
    "01_Калиевая селитра.xlsx": ["2834210000"],
    "02_Формиаты.xlsx": ["2915120000"],
    "03_Нитрат кальция.xlsx": ["3102600000", "3102900000", "3105902000", "2834298000"],
    "05_Нитрит натрия.xlsx": ["3102500000", "2834100000"],
    "06_МАФ.xlsx": ["3105400000"],
    "07_NPK(S) ВРУ.xlsx": [
        "3105100000", "3105200000", "3105201000", "3105209000",
        "3105510000", "3105590000", "3105600000", "3105902000",
        "3105908000", "3105909100", "3105909900"
    ],
    "08_Сульфат магния.xlsx": ["2833210000"],
    "09_Монокалийфосфат.xlsx": ["2835240000", "3105600000"],
    "Экспорт NPK ВРУ.xlsx": ["3105100", "3105200", "3105201", "3105209", "3105908"],
    "Экспорт МАФ 12-61.xlsx": ["3105400000", "3105400001"],
    "Экспорт Монокалийфосфата.xlsx": ["2835240000", "3105600000"],
    "Экспорт Сульфата магния.xlsx": ["2833210000"]
}

# Заполняем словарь соответствий код -> файл
codes_to_files = {}
for filename, codes in code_file_map.items():
    for code in codes:
        codes_to_files[code] = filename


def process_files():
    print("Начало обработки...")

    # Создаем папку для результатов
    os.makedirs(output_folder, exist_ok=True)

    # Загружаем исходные данные
    try:
        source_df = pd.read_excel(input_file)
    except Exception as e:
        print(f"Ошибка при чтении исходного файла: {e}")
        return

    print(f"Исходный файл загружен. Найдено {len(source_df)} записей.")

    # Группируем данные по файлам и листам
    file_sheet_data = {
        filename: {sheet: [] for sheet in sheets}
        for filename, sheets in file_sheet_mapping.items()
    }

    # Получаем список всех колонок из исходного файла
    source_columns = set(source_df.columns)

    # Распределяем строки по файлам и листам
    for _, row in source_df.iterrows():
        code = str(row['G33 (код товара по ТН ВЭД РФ)']).strip()
        if code in codes_to_files:
            filename = codes_to_files[code]
            row_dict = row.to_dict()

            # Получаем целевые листы
            target_sheets = file_sheet_mapping.get(filename, [])
            if not target_sheets:
                continue

            output_path = os.path.join(output_folder, filename)

            try:
                with pd.ExcelFile(output_path) as xls:
                    for sheet_name in target_sheets:
                        if sheet_name in xls.sheet_names:
                            # Читаем только заголовки целевого листа
                            target_cols = set(pd.read_excel(xls, sheet_name=sheet_name, nrows=0).columns)

                            # Фильтруем данные по совпадению заголовков
                            filtered_row = {
                                col: row_dict[col]
                                for col in row_dict
                                if col in target_cols and col in source_columns
                            }

                            if filtered_row:
                                file_sheet_data[filename][sheet_name].append(filtered_row)
            except FileNotFoundError:
                # Если файл не существует — пропускаем, он будет создан позже
                pass

    # Обрабатываем каждый файл
    for filename, sheets_data in file_sheet_data.items():
        output_path = os.path.join(output_folder, filename)

        new_dfs = {
            sheet: pd.DataFrame(data)
            for sheet, data in sheets_data.items() if data
        }

        if not new_dfs:
            continue

        # Подгружаем существующие данные
        existing_dfs = {}
        try:
            with pd.ExcelFile(output_path) as xls:
                for sheet, df in new_dfs.items():
                    if sheet in xls.sheet_names:
                        existing_dfs[sheet] = pd.read_excel(xls, sheet_name=sheet)
                    else:
                        existing_dfs[sheet] = pd.DataFrame(columns=df.columns)
        except FileNotFoundError:
            existing_dfs = {sheet: pd.DataFrame(columns=df.columns) for sheet, df in new_dfs.items()}

       # Объединяем старые и новые данные
final_dfs = {}
for sheet, new_df in new_dfs.items():
    old_df = existing_dfs.get(sheet, pd.DataFrame())
    combined_df = pd.concat([old_df, new_df], ignore_index=True)
    final_dfs[sheet] = combined_df

        # Сохраняем в Excel
        with pd.ExcelWriter(output_path, mode='w', engine='openpyxl') as writer:
            for sheet_name, df in final_dfs.items():
                df.to_excel(writer, index=False, sheet_name=sheet_name)
            print(f"Добавлены данные в файл '{filename}' в листы: {list(final_dfs.keys())}")

    print(f"Готово! Результаты в папке '{output_folder}'")


if __name__ == "__main__":
    process_files()
