import openpyxl

path = "Dummymappe1.xlsx"

wb_obj = openpyxl.load_workbook(path)

sheet_obj = wb_obj.active

row_number = sheet_obj.max_row
column_number = sheet_obj.max_column

print("Total Rows:", row_number)
print("Total Columns:", column_number)


def people_number(x):

    for i in range(2, row_number+1):
        name_a = sheet_obj.cell(row=i, column=3)
        name_b = sheet_obj.cell(row=i + 1, column=3)
        if name_a.value == name_b.value:
            i = i + 1
        else:
            x = x + 1
            i = i + 1
    print("Number of people is:", x)


people_number(0)
