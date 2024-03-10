# matches = re.search(r"([-+]?\d*\.\d+)$", line)

import re
from decimal import Decimal, getcontext

# Установка точности для модуля decimal
getcontext().prec = 28

def extract_numbers(line):
    # Используем re.findall чтобы получить все числа из строки
    matches = re.findall(r"[-+]?\d*\.\d+|\d+", line)
    
    # Мы хотим получить только последнее число в строке, предполагая, что оно является значением
    if matches:
        return Decimal(matches[-1])  # Берем последнее число
    return None

def calculate_difference(num1, num2):
    try:
        return (abs(num1 - num2) / abs(num1)).quantize(Decimal('0.00000000000001'))
    except ZeroDivisionError:
        return Decimal('Inf')

def process_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as file:
        current_project = None
        file_numbers = []

        for line in file:
            line = line.strip()
            if 'Проект' in line:
                
                # Проверяем предыдущий проект, если таковой был
                if current_project is not None and len(file_numbers) == 2:
                    
                    # Проверка и вывод разницы для предыдущего проектаS
                    difference = calculate_difference(*file_numbers)
                    if difference > Decimal('0.01'):
                        print(f'Проект {current_project} отличается более чем на 1%. Разница: {difference * 100:.2f}%')
                        print(difference)
                
                # Получаем идентификатор проекта, исключая слово "отличается"
                current_project = ' '.join(line.split()[1:-1])
                file_numbers = []  # Сброс чисел для нового проекта

            elif 'Файл' in line:
                number = extract_numbers(line)
                if number is not None:  # Если число найдено, добавляем его в список
                    file_numbers.append(number)

        # Проверка для последнего проекта
        if current_project is not None and len(file_numbers) == 2:
            difference = calculate_difference(*file_numbers)
            if difference > Decimal('0.001'):
                print(difference)
                print(f'Проект {current_project} отличается более чем на 1%. Разница: {difference * 100:.2f}%')

    print("Обработка файла завершена.")

# Замените 'path_to_your_file.txt' на путь к вашему файлу
process_file('differences_report.txt')


