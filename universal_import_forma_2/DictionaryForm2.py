import pandas as pd
from tqdm import tqdm

# Загрузим данные из первого Excel файла
excel1 = pd.read_excel('1.xlsx')

# Поиск фразы "Общий объем финансирования, в том числе за счет:" во всех ячейках 4 строки
columns_to_transfer = []
target_phrase = "иных источников финансирования"
found_phrases = []

for column in tqdm(excel1.columns, desc="Ищем столбцы"):
    for cell in excel1[column][2:3]:
        if pd.notna(cell) and target_phrase.lower() in str(cell).lower():
            columns_to_transfer.append(column)
            found_phrases.append(f"Фраза '{target_phrase}' найдена в столбце '{column}', ячейка: {cell}")

# Вывод найденных фраз
print("Найденные фразы:")
for phrase in found_phrases:
    print(phrase)

# Если найдены столбцы, переносим их во второй Excel файл
if columns_to_transfer:
    # Загрузим данные из второго Excel файла
    excel2 = pd.read_excel('2.xlsx')

    # Переносим выбранные столбцы
    for column in tqdm(columns_to_transfer, desc="Переносим столбцы"):
        excel2[column] = excel1[column]

    # Сохраняем изменения во втором Excel файле
    excel2.to_excel('2.xlsx', index=False)
    print("Столбцы успешно перенесены.")
else:
    print(f"Фраза '{target_phrase}' не найдена в 3 строке первого файла.")