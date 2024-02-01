import string
from statistics import mean
import time
import os
import platform

from openpyxl import load_workbook


def clear_console() -> None:
    os_type = platform.system()
    if os_type == 'Windows':
        os.system('cls')
    elif os_type == "Linux" or os_type == "Darwin":
        os.system('clear')


HEADER_SMP_OUTLIERS = 'CМП (выброс макс мин)'
HEADER_SMP_MR = 'CМП (по МР)'
HEADER_OUTLIERS_FLAG = 'Есть выскакивающий показатель'


class Menu:
    def __init__(self) -> None:
        self.excel_file = None

    def input_file_info(self) -> None:
        clear_console()
        file_name = input("Введите название файла:")
        col = str(input('Введите букву столбца начала данных: '))
        row = int(input('Введите номер строки начала данных: '))
        
        self.file = file_name
        self.start_cell = f'{col}{row}'
        self.col = col
        self.row = row

        self.info_file = (f'\n\n----INFO----\n'
                          f'Файл: {file_name}\n'
                          f'Стартовая позиця: {self.start_cell}\n\n')
    
    def input_smp_col(self):
        clear_console()
        smp_col = str(input('Введите букву столбца для записи СМП' 
                                     'посчитоного по старой методике: '))
        smp_mr_col = str(input('Введите букву столбца для записи СМП' 
                                     'посчитоного по Методичке: '))
        
        self.write_to_smp_mr_col = smp_mr_col
        self.write_to_smp_col = smp_col

        self.info_writing_smp = (f'\n-*--*--*-info_writing_smp-*--*--*--*-\n'
                                 f'{HEADER_SMP_OUTLIERS} - '
                                 f'{self.write_to_smp_col}\n'
                                 f'{HEADER_SMP_MR} - '
                                 f'{self.write_to_smp_mr_col}\n\n')
    
    def info(self):
        clear_console()
        try:
            if self.info_file or self.info_writing_smp:
                print(f'{self.info_file}\n{self.info_writing_smp}')
            else:
                print('\n\n*#*#*#*#*#*#*#*#*#\n'
                      'Кажется вы еще ничего не выбрали...'
                      '*#*#*#*#*#*#*#*#*#\n\n')
        except AttributeError: 
            print('\nКажется вы еще ничего не выбрали...\nпопробуйте снвоа')
    
    def quit(self) -> None:
        clear_console()
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
        '1': menu.input_file_info,
        '2': menu.input_smp_col,
        '3': menu.info,
        'quit': menu.quit,
        'albert': menu.quit
    }
    while True:
        clear_console()
        print("""
Меню:
1. Ввод названия файла и координат ячейки начала данных...
2. Ввод коордиант для записи СМП:
3. INFO:
Введите 'quit' для выхода из программы.
        """)
        choice = input("Выберите действие: ")
        action = choices.get(choice)
        if action:
            action()
            if choice == '1':
                
                # book = ExcelFile(menu.file)
                # sheet: str = book.wb.sheetnames
                pass
            if choice == '2':
                pass
            
            if choice == '3':
                pass    
                          
        else:
            print("Некорректный выбор, попробуйте еще раз.")

main()
