import os
import pandas as pd

script_dir = os.path.dirname(os.path.abspath(__file__))

def compare_project_identifiers(df1, df2):
    """Сравнение идентификаторов проектов между двумя DataFrame."""
    ids1 = set(df1.iloc[:, 4])
    ids2 = set(df1.iloc[:, 4])

    
    missing_in_file1 = ids2 - ids1
    missing_in_file2 = ids1 - ids2

    return missing_in_file1, missing_in_file2

def save_report_to_file(missing_in_file1, missing_in_file2, sheet, output_file):
    """Сохранение отчета в файл."""
    with open(output_file, "a", encoding="utf-8") as f:
        f.write(f"Вкладка: {sheet}\n")
        f.write("Отсутствующие в первом файле проекты:\n")
        for identifier in missing_in_file1:
            f.write(f"  {identifier}\n")
        
        f.write("Отсутствующие во втором файле проекты:\n")
        for identifier in missing_in_file2:
            f.write(f"  {identifier}\n")
        f.write('-' * 80 + '\n')

if __name__ == "__main__":
    file1 = os.path.join(script_dir, 'Форма 20', 'forma_20_Gipro.xlsx')
    file2 = os.path.join(script_dir, 'Форма 20', 'forma_20_tumen.xlsx')
    sheets = ['20.1', '20.2', '20.3', '20.4']
    
    output_file = os.path.join(script_dir, "projects_comparison_report.txt")

    # Убедитесь, что файл существует перед записью в него.
    if os.path.exists(output_file):
        os.remove(output_file)

    for sheet in sheets:
        df1 = pd.read_excel(file1, sheet_name=sheet)
        df2 = pd.read_excel(file2, sheet_name=sheet)

        missing_in_file1, missing_in_file2 = compare_project_identifiers(df1, df2)
        save_report_to_file(missing_in_file1, missing_in_file2, sheet, output_file)

    print(f"Отчет сохранен в {output_file}")