import json
import string
from statistics import mean
from openpyxl import load_workbook

# TODO Сделать инпутами, вывести в отдельынй файл
FILE: str = 'Ф1.xlsx'
FORM: str = 'ф 1_субъекты_рост-сниж_июнь_2023.xlsx'
REC_COL: str = 'L'
REC_COL_MR: str = 'M'
OUTLIERS_FLAG: str = 'N'
START_END_COll: list[int, int] = [2, 11]
START_END_ROW: list[int, int] = [12, 105]


def decode(work_book: str, sheet: str, coll: str, rows: int) -> float:
    """ Функция для чтения из EXEL. На вход подается:
 work_book = файл EXEL,
 sheet = страница,
 коардинаты ячейки(колонка и строчка)
 coll = колонка,
 rows = строчка
    """
    sheet: str = work_book[f'{sheet}']
    value_cell: float = sheet[f'{coll}{rows}'].value
    return value_cell


def record(work_book: str, sheet: str, coll: str, rows: int,
           value: float) -> None:
    """ Функция для записи в EXEL. На вход подается:
        work_book = файл EXEL,
        sheet = страница,
        коардинаты ячейки(колонка и строчка)
        coll = колонка,
        rows = строчка
        значение = value (value имеет тип float)
    """
    sheet: str = work_book[f'{sheet}']
    sheet[f'{coll}{rows}'] = value


wb_write = load_workbook(filename=FORM)
sheet_write: str = wb_write.sheetnames

wb_read = load_workbook(filename=f'!{FILE}')
sheet_read: str = wb_read.sheetnames

with open('nod_2.json', 'r', encoding='utf-8') as file:
    data = json.load(file)

form_1_data = data['form_1']

for sheet in sheet_read:
    for item in form_1_data:
        page: str = item['page']
        code: str = item['code']
        name: str = item['name']
        if sheet == code:
            for row in range(START_END_ROW[0], START_END_ROW[1] + 1):
                mr_mean = decode(wb_read, sheet, REC_COL_MR, row)
                name_region_read = decode(wb_read, sheet, 'A', row)

                row_write = row - 2
                name_region_write = decode(wb_write, page, 'B', row_write)

                if name_region_write == name_region_read:
                    print(f'----------\n{page} | {code} | {name}\n'
                          f'{name_region_write} {name_region_read} '
                          f'{mr_mean}\n----------\n\n')
                    record(wb_write, page, 'AK', row_write, mr_mean)

wb_write.save(f'_{FORM}')
