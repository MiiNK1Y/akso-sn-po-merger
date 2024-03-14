import openpyxl

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
            data_pair.append(row[column_0] + . + row[column_1])

        return column_data_pair

    def delete_column(self, columns_delete: list) -> None:
        #The 'delete_cols' works as 'in-range' deletion. 
        #When giving it 2 int args, it will delete the cols in range (example: 10 - 12 will delete 11 aswell).
        #Range 1 - 1 or 'delete_cols(1, 1)' just deletes column A.
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

def main():
    #implement a way to seek the oldest and newest filename-dates, then work on those files instead of manualy naming them.
    old_path = "./old.xlsx"
    new_path = "./new.xlsx"
    final_path = "./final.xlsx"

if __name__ == '__main__':
    main()
