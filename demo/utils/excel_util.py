import openpyxl

def create_excel(file_name, headers):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Products"
    sheet.append(headers)
    workbook.save(file_name)

def add_row_to_excel(file_name, row_data):
    workbook = openpyxl.load_workbook(file_name)
    sheet = workbook.active
    sheet.append(row_data)
    workbook.save(file_name)
