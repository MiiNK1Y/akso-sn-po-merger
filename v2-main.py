#!./env/bin/python3

import openpyxl
import configparser

class excel_workbook():
    def __init__(self, path: str) -> None:
        self.xlsx_file = path
        self.workbook = openpyxl.load_workbook(self.xlsx_files)
        self.sheet = self.workbook.active
        self.sheet_height = self.sheet.max_row
        self.sheet_width = self.sheet.max_column

    def get_column_header(self) -> list:
        column_header = []
        for header in self.sheet.iter_cols(1, self.sheet_width, values_only=True):
            column_headers.append(header[0])
        return column_headers

    def map_data_pair(self, column_0: int, column_1: int) -> list:
        column_data_pair = []
        for row in self.sheet.iter_rows(1, self.sheet_height, values_only=True):
            data_pair.append(row[column_0] + "." + row[column_1])
        return column_data_pair

    def delete_column(self, columns_delete: list) -> None:
        for column in columns_delete:
            self.sheet.delete_cols(column)

    def save_workbook(self, filename: str) -> None:
        self.workbook.save(filename)

    def close_workbook(self) -> None:
        self.workbook.close() 

    def match_and_insert(self, column_lane_index: int, column_insert_index: int, replacement_0: str, replacement_1: str, data: list) -> None:
        row_count = 1
        data_SNs = [i.split(".")[0] for i in data]
        for row in self.sheet.iter_rows(1, self.sheet_height, values_only=True):
            found_cell_value = row[column_lane_index]
            if found_cell_value == None:
                self.sheet.cell(row=row_count, column=column_insert_index, value=replacement_0)
                row_count += 1
                continue
            if found_cell_value in data_SNs:
                for value in data:
                    data_pair = value.split(".")
                    serial = data_pair[0]
                    po = data_pair[1]
                    if serial == found_cell_value:
                        self.sheet.cell(row=row_count, column=column_insert_index, value=po)
                        row_count += 1
            else:
                self.sheet.cell(row=row_count, column=column_insert_index, value=replacement_1)
                row_count += 1

def date_is_valid(date: str) -> bool:
    days = int(date[0:2])
    months = int(date[3:5])
    years = int(date[6:])
    if (days <= 31) and (months <= 12) and ((years < 2100) and (years > 1970)) :
        return True
    else:
        return False

def get_date_from_str(file_name: str) -> str:
    char_list = ".0123456789" #instead of trying to convert every character to int
    date_format = "dd.mm.yyyy"
    for index, char in enumerate(file_name):
        if char in char_list:
            date_format_lenght = len(date_format) + 1
            for x in range(date_format_lenght):
                x_val = file_name[index + x]
                if (x == 2 or x == 5) and (x_val != '.'):
                    break
                elif x == (date_format_lenght - 1):
                    date_found = file_name[index:(index + x)]
                    if date_is_valid(date_found):
                        return date_found
                    else:
                        break
                elif x_val not in char_list:
                    break
                else:
                    continue
        else:
            continue

# THIS ONE IS BIG BRAIN UNIQUE BRO.
def get_newest_date(d1: str, d2: str) -> str:
    if d1 == d2:
        return None
    d1_formated = d1.split('.')[::-1]
    d2_formated = d2.split('.')[::-1]
    if d1_formated > d2_formated:
        return d1
    else:
        return d2

def main() -> None:
    config = configparser.ConfigParser()
    config.read('config.ini')
    default_config = config['DEFAULT']

    old_sheet_path = default_config['old_sheet_path']
    new_sheet_path = default_config['new_sheet_path']
    final_sheet_path = default_config['final_sheet_path']
    columns_to_delete = default_config['columns_to_delete']
    serial_column_text = default_config['serial_column_text']
    po_column_text = default_config['po_column_text']
    final_sheet_column_insert_po = default_config['final_sheet_column_insert_po']
    none_to_match_replacement = default_config['none_to_match_replacement']
    no_match_replacement = default_config['no_match_replacement']

    print(columns_to_delete)

if __name__ == '__main__':
    main()
