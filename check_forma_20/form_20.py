import pandas as pd

def compare_sheets_by_column(file1, file2, sheet_name, id_col_num, param_col_num):
    # Load the data into dataframes
    df1 = pd.read_excel(file1, sheet_name=sheet_name)
    df2 = pd.read_excel(file2, sheet_name=sheet_name)

    # Get the column names from the dataframe
    id_column_name = df1.columns[id_col_num]
    param_column_name = df1.columns[param_col_num]

    # Convert the index and parameter columns to strings to ensure proper comparison
    df1[id_column_name] = df1[id_column_name].astype(str)
    df2[id_column_name] = df2[id_column_name].astype(str)
    df1[param_column_name] = df1[param_column_name].astype(str)
    df2[param_column_name] = df2[param_column_name].astype(str)

    # Set the index to the identifier column and sort
    df1.set_index(id_column_name, inplace=True)
    df2.set_index(id_column_name, inplace=True)
    df1.sort_index(inplace=True)
    df2.sort_index(inplace=True)

    differences = []

    # Check for differences
    for project_id in df1.index.intersection(df2.index):
        values_df1 = df1.loc[project_id, param_column_name]
        values_df2 = df2.loc[project_id, param_column_name]

        # If the result is a Series, compare elementwise
        if isinstance(values_df1, pd.Series) or isinstance(values_df2, pd.Series):
            for val1, val2 in zip(values_df1, values_df2):
                if val1 != val2:
                    differences.append((project_id, val1, val2))
        else:
            # If the result is not a Series, directly compare the values
            if values_df1 != values_df2:
                differences.append((project_id, values_df1, values_df2))

    # Format the differences into strings
    differences_formatted = [
        f'Проект {project_id} отличается.\nФайл 1: {value_df1}\nФайл 2: {value_df2}\n'
        for project_id, value_df1, value_df2 in differences
    ]

    return differences_formatted  # Return the formatted list of differences

# Пути к файлам
file1_path = 'forma_20_Gipro.xlsx'
file2_path = 'forma_20_tumen.xlsx'

# Вызов функции сравнения и вывод результатов
differences_report = compare_sheets_by_column(file1_path, file2_path, '20.3', 3, 6)

# Вывод отчета в консоль и запись в файл
with open('differences_report.txt', 'w', encoding='utf-8') as file:
    for difference in differences_report:
        print(difference)
        file.write(difference + '\n')