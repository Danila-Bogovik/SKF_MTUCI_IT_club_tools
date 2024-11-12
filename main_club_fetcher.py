import pandas as pd
import os
from tkinter import filedialog, Tk


def process_excel_file(file_path):
    try:
        df = pd.read_excel(file_path)
        if df.shape[1] < 3:
            print("Ошибка: Excel файл должен содержать как минимум три столбца.")
            return

        output_file_path = os.path.splitext(file_path)[0] + '_output.txt'

        with open(output_file_path, 'w', encoding='utf-8') as f:
            for index, row in df.iterrows():
                line = f"{index + 1}. {row[1]}, {row[1].upper()}, {row[2]}"
                f.write(line + '\n')

        print(f"Данные успешно сохранены в файл: {output_file_path}")

    except Exception as e:
        print(f"Произошла ошибка: {e}")


def get_file_path():
    choice = input("Выберите режим: введите '1' для ввода пути вручную или '2' для выбора файла через окно: ")

    if choice == "1":
        res = input("Введите путь к Excel файлу: ")
        return res
    elif choice == "2":
        file_path = filedialog.askopenfilename (title="Выберите Excel файл",            initialdir='/', filetypes=[("Excel files", "*.xlsx;*.xls")])
        return file_path
    else:
        print("Неверный выбор. Попробуйте снова.")
        return get_file_path()

root = Tk()

resulted_file_path = get_file_path()
if resulted_file_path:
    process_excel_file(resulted_file_path)
    root.destroy()

root.mainloop()