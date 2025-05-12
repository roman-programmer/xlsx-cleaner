import pandas as pd
import numpy as np


# Вказати шлях до вхідного файлу .xls
input_file = "qq.xlsx"
output_file = "CLEAR_Перелік33.xlsx"


def clean_data(df):
    """
    Очищає дані в DataFrame:
    - Видаляє зайві пробіли перед/після тексту
    - Замінює кілька пробілів на один
    - Видаляє символи переносу рядка
    - Замінює повністю порожні значення на "null"
    - Замінює текст у передостанньому стовпці на стиль "Привіт" (тобто з великої букви та інші маленькі)
    """
   
    # Очищення текстових даних
    cleaned_df = df.applymap(
        lambda x: ' '.join(str(x).strip().replace('\n', ' ').split())
        if isinstance(x, str) else x
    )
    # Замінюємо порожні значення на "null"
    #cleaned_df = cleaned_df.replace([None, "", "-", 0, "0", np.nan], "null")


    # Зміна тексту в передостанньому стовпці
    if len(cleaned_df.columns) > 1:  # Перевірка, чи є хоча б 2 стовпці
        last_col_index = cleaned_df.columns[-2]  # Індекс передостаннього стовпця
        cleaned_df[last_col_index] = cleaned_df[last_col_index].apply(
            lambda x: x.title() if isinstance(x, str) else x
        )


    # Форматування 3-го стовпця як int
    if len(cleaned_df.columns) >= 3:
        third_col = cleaned_df.columns[2]
        cleaned_df[third_col] = pd.to_numeric(cleaned_df[third_col], errors='coerce').fillna(0).astype(int)


    # Форматування 5-го стовпця як дата (РРРР-ММ-ДД)
    if len(cleaned_df.columns) >= 5:
        fifth_col = cleaned_df.columns[4]
        cleaned_df[fifth_col] = pd.to_datetime(cleaned_df[fifth_col], errors='coerce').dt.strftime('%Y-%m-%d')
        cleaned_df[fifth_col] = cleaned_df[fifth_col].fillna("null")  # Замінюємо недати на "null"


    # Інші стовпці залишаються текстовими
    for col in cleaned_df.columns:
        if col not in [third_col, fifth_col]:
            cleaned_df[col] = cleaned_df[col].astype(str)


    cleaned_df = cleaned_df.replace([None, "", "-", 0, "0", np.nan, "nan"], "null")


    return cleaned_df


try:
    # Зчитування даних з .xls
    data = pd.read_excel(input_file, sheet_name=None)  # Зчитати всі аркуші


    # Якщо вхідний файл має лише один аркуш:
    if len(data) == 1:
        sheet_name = list(data.keys())[0]
        df = data[sheet_name]


        # Очищення даних
        df_cleaned = clean_data(df)


        # Додаткове очищення першого рядка, першої комірки
        if isinstance(df_cleaned.iat[0, 0], str):
            df_cleaned.iat[0, 0] = ' '.join(df_cleaned.iat[0, 0].strip().split())


        # Збереження у форматі .xlsx
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_cleaned.to_excel(writer, index=False, sheet_name=sheet_name)
        print(f"Файл успішно відформатовано та збережено: {output_file}")
    else:
        print("Файл має кілька аркушів. Очищення та збереження кожного.")
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for sheet_name, df in data.items():
                # Очищення даних
                df_cleaned = clean_data(df)


                # Додаткове очищення першого рядка, першої комірки
                if isinstance(df_cleaned.iat[0, 0], str):
                    df_cleaned.iat[0, 0] = ' '.join(df_cleaned.iat[0, 0].strip().split())


                # Збереження кожного аркуша у форматі .xlsx
                df_cleaned.to_excel(writer, index=False, sheet_name=sheet_name)
        print(f"Усі аркуші збережено у файл: {output_file}")
except Exception as e:
    print(f"Помилка під час обробки файлу: {e}")