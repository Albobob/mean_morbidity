import string
from statistics import mean
import time

from openpyxl import load_workbook

HEADER_SMP_OUTLIERS = 'CМП (выброс макс мин)'
HEADER_SMP_MR = 'CМП (по МР)'
HEADER_OUTLIERS_FLAG = 'Есть выскакивающий показатель'

class Menu:
    def __init__(self) -> None:
        self.excel_file = None

    def input_file_name(self) -> None:
        file_name = input("Введите название файла:")
        self.excel_file = ExcelFile(file_name)
    
    def input_start_cell(self) -> None:
        col = str(input('Введите букву столбца начала данных: '))
        row = int(input('Введите номер строки начала данных: '))
        self.start_cell = f'{col}{row}'
    
    def input_smp_col(self) 
    
    def quit(self) -> None:
        if self.excel_file:
            self.excel_file.close()
        print('Выход из  программы.')
        for i in range(1, 4):
            print(i)
            time.sleep(0.08)
        print('До новых встреч...')
        exit()


class ExcelFile:
    def __init__(self, file_name: str) -> None:
        self.file_name = file_name
        self.wb = load_workbook(filename=file_name, data_only=True)

    def reed_cell(self, sheet_name: str, col: str, row: int) -> float:
        sheet = self.wb[sheet_name]
        return sheet[f'{col}{row}'].value

    def write_cell(self, sheet_name: str, row: int, col: str,
                   value: float) -> None:
        sheet = self.wb[sheet_name]
        sheet[f'{col}{row}'] = value

    def save(self) -> None:
        self.wb.save(f'!{self.file_name}')
    
    def close(self) -> None:
        self.wb.close()


def get_data(excel_file: ExcelFile, sheet_name: str, row: int) -> list[float]:
    coll: list = list(string.ascii_lowercase)
    data = []

    for c in coll[START_END_COll[0] - 1:START_END_COll[1] + 1]:
        cell_value = excel_file.reed_cell(sheet_name, row, c)
        data.append(cell_value)

    return data


def main():
    menu = Menu()
    choices = {
        '1': menu.input_file_name,
        '2': menu.input_start_cell,
        'quit': menu.quit,
        'albert': menu.quit
    }
    while True:
        print("""
Меню:
1. Ввод названия файла
2. Ввод координат ячейки начала данных
Введите 'quit' для выхода из программы.
        """)
        choice = input("Выберите действие: ")
        action = choices.get(choice)
        if action:
            action()
            if choice == '1':
                book = ExcelFile(menu.excel_file)
                sheet: str = book.wb.sheetnames
            if choice  == '2':
                pass
        else:
            print("Некорректный выбор, попробуйте еще раз.")

main()
