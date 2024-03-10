import os
import pandas as pd

# Определение пути к папке, где находится данный скрипт
script_dir = os.path.dirname(os.path.abspath(__file__))

column_index = 20

def get_total_values(df):
    """Функция для извлечения итоговых значений проектов."""
    total_values = {}
    itogo_rows = df[df.iloc[:, 4] == "Итого объем финансовых потребностей по инвестиционному проекту, тыс. рублей"]
    for _, row in itogo_rows.iterrows():
        project_identifier = row.get('Идентификатор инвестиционного проекта', None)
        if project_identifier:
            total_values[project_identifier] = row.iloc[column_index]
    return total_values

def save_total_differences_to_file(differences, file1, file2):
    """Функция сохраняет различия итоговых значений в файл."""
    filename = os.path.join(script_dir, "total_comparison_results.txt")
    with open(filename, "w", encoding="utf-8") as f:
        for identifier, (val1, val2) in differences.items():
            f.write(f"Идентификатор проекта: {identifier}\n")
            f.write(f"Данные в {os.path.basename(file1)}: {val1}\n")
            f.write(f"Данные в {os.path.basename(file2)}: {val2}\n")
            f.write('-' * 80 + '\n')

def compare_total_values(file1, file2, sheets):
    differences = {}
    for sheet in sheets:
        df1 = pd.read_excel(file1, sheet_name=sheet)
        df2 = pd.read_excel(file2, sheet_name=sheet)
        
        totals1 = get_total_values(df1)
        totals2 = get_total_values(df2)
        
        for identifier, value1 in totals1.items():
            if identifier in totals2:
                value2 = totals2[identifier]
                if value1 != value2:
                    differences[identifier] = (value1, value2)
    save_total_differences_to_file(differences, file1, file2)

if __name__ == "__main__":
    file1 = os.path.join(script_dir, 'forma_20_Gipro.xlsx')
    file2 = os.path.join(script_dir, 'forma_20_tumen.xlsx')
    sheets = ['20.1', '20.2', '20.3', '20.4']

    compare_total_values(file1, file2, sheets)