import pandas as pd

def compare_excel_sheets(file1, file2, sheet_name):
    df1 = pd.read_excel(file1, sheet_name=sheet_name)
    df2 = pd.read_excel(file2, sheet_name=sheet_name)

    if df1.shape != df2.shape:
        return f"Размеры таблиц на вкладке '{sheet_name}' различны."

    different_rows = []
    for i in range(df1.shape[0]):
        if not df1.iloc[i].equals(df2.iloc[i]):
            different_rows.append(i + 1)

    if different_rows:
        return f"Отличия в строках на вкладке '{sheet_name}': {', '.join(map(str, different_rows))}."
    else:
        return f"На вкладке '{sheet_name}' различий не найдено."

def compare_excels(file1, file2, sheets):
    results = []
    for sheet in sheets:
        result = compare_excel_sheets(file1, file2, sheet)
        results.append(result)
    
    with open("comparison_result.txt", "w", encoding="utf-8") as f:
        for line in results:
            f.write(line + '\n')

if __name__ == "__main__":
    file1 = 'forma_20_Gipro.xlsx'
    file2 = 'path_to_file2.xlsx'
    sheets = ['20.1', '20.2', '20.3', '20.4']
    compare_excels(file1, file2, sheets)