import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import rows_from_range
import datetime
import csv
import os

workbook = Workbook()

print("Please select a file in format: Name.xlsx")
# path = filename = input()
path = "Dummymappe2csv.xlsx"
wb_obj = openpyxl.load_workbook(path, data_only=True)
sheet_obj = wb_obj.active

row_number = sheet_obj.max_row
column_number = sheet_obj.max_column

list_fall_30 = ["ABT SFN st/sv-pfl",
                "AG-Zu.lfd.3(63)",
                "Ant.Listenperis P",
                "Ant.SFN st/sv-pfl",
                "Ant.SFN sv-pfl.",
                "Ant.Wo.-Arbeit PK",
                "BAV St.lfd. 3/63",
                "DBA/ATE lfd ST",
                "EFZ-Entgelt DS (3",
                "Fahrtkosten stpfl",
                "Grundvergütung",
                "GV pro Stunde",
                "GV-Aushilfen Zahl",
                "GV-Prakt. Ind.",
                "Inflationsausglei",
                "Kleidergeld",
                "Kleidergeld HR",
                "Kontoführungsgeb.",
                "Krankentgelt DP",
                "Krankentgelt DS",
                "LFZ Zuschlag",
                "Listenpreis PKW",
                "MA Zuschlag allg.",
                "Mehrbereichzul.",
                "Netto-Hochr.",
                "persönl. Zulage",
                "PKW Zahlung gwV",
                "Prämie NHR",
                "Reinigungspauscha",
                "Reinigungspsch.",
                "Saalzschl.",
                "Saalzschl. Kasse",
                "Saalzschlag",
                "Saalzschlag Kasse",
                "Stunden",
                "Taetigkeitszulage",
                "Urlaubentgelt DS",
                "VL-AG-Zuschu",
                "Weihnachtsgeld",
                "Zlg MuSchu Zuschuß",
                "Zlg var Zulagen",
                "Zulage",
                "Zuschlag Feiertag",
                "Zuschlag Nacht",
                "Zuschlag Sonntag"]

with open('Dummymappe2csv.csv', 'w', newline="") as file_handle:
    csv_writer = csv.writer(file_handle, delimiter=";")
    for row in sheet_obj.iter_rows(min_row=2):
        csv_writer.writerow([cell.value for cell in row])

with open("Dummymappe2csv.csv", "r", encoding="utf-8-sig", newline="") as f:
    csv_rows = list(csv.reader(f, delimiter=';'))
    people_dict = {int: list}
    for line in csv_rows:
        if line[1] not in people_dict:
            people_dict[line[1]] = [line]
        else:
            people_dict[line[1]].append(line)

all_the_people = list(dict.fromkeys(people_dict))


class Person:
    def __init__(self, i_in_people):
        current_person = all_the_people[i_in_people]
        self.abrk = people_dict[current_person][0][0]
        self.name = people_dict[current_person][0][2]
        self.PN = people_dict[current_person][0][1]
        self.month = []
        self.lohn = []
        self.fall30 = False


def people_classes(x):
    person = Person(x)
    local_month = []
    local_lohn = []
    for i in people_dict[all_the_people[x]]:
        local_month.append(i[3])
        local_lohn.append(i[11])
    local_month = list(dict.fromkeys(local_month))
    local_lohn = list(dict.fromkeys(local_lohn))

    person.fall30 = any(k in list_fall_30 for k in local_lohn)

    person.month = local_month
    person.lohn = local_lohn
    return person


list_of_People = [Person]
for i in range(1, len(all_the_people)):
    list_of_People.append(people_classes(i))


def final_report():
    sheet = workbook.active
    month_tuple = datetime.date.today().replace(day=1) - datetime.timedelta(days=1)
    last_month = month_tuple.strftime("%m")
    current_year = month_tuple.strftime("%y")
    workbook.save(filename=(f"Qualität_{last_month}_{current_year}.xlsx"))

    asrow = 1000+len(all_the_people)

    file_name_source = 'End_of_Report.xlsx'
    wb_source = openpyxl.load_workbook(file_name_source)
    sheet_source = wb_source.active

    range_str = 'A1:R64'

    for row in rows_from_range(range_str):
        for cell in row:
            dest_sheet_cell = sheet[cell].offset(row=asrow-1)
            source_sheet_cell = sheet_source[cell]

            is_merged = False
            for merged_range in sheet_source.merged_cells.ranges:
                if merged_range.min_row <= source_sheet_cell.row <= merged_range.max_row and merged_range.min_col <= source_sheet_cell.column <= merged_range.max_col:
                    is_merged = True
                    if (source_sheet_cell.row == merged_range.min_row) and (source_sheet_cell.column == merged_range.min_col):
                        first_merged_source = sheet_source.cell(row=merged_range.min_row, column=merged_range.min_col)
                        first_merged_dest = sheet[first_merged_source.coordinate].offset(row=asrow-1)

                        first_merged_dest.value = first_merged_source.value
                        first_merged_dest.font = Font(bold=first_merged_source.font.bold,
                                                      color=first_merged_source.font.color)

                        sheet.merge_cells(start_row=first_merged_dest.row,
                                          start_column=first_merged_dest.column,
                                          end_row=first_merged_dest.row + merged_range.max_row - merged_range.min_row,
                                          end_column=first_merged_dest.column + merged_range.max_col - merged_range.min_col)
                    break

            if not is_merged:
                dest_sheet_cell.value = source_sheet_cell.value
                dest_sheet_cell.font = Font(bold=source_sheet_cell.font.bold,
                                            color=source_sheet_cell.font.color)

    for i in range(0, 12):
        row_gsmterg = str(asrow)
        column_gsmterg = chr(64 + 2 + i)
        sheet[column_gsmterg + row_gsmterg] = f'=SUM({column_gsmterg}2:{column_gsmterg}{str(len(all_the_people))})'

    cl = sheet.cell(row=asrow+11, column=3)
    cl.value = f'=COUNT(B2:M{str(asrow)})'

    c = sheet['A1']
    c.value = "Zeilenbeschriftungen"
    sheet.column_dimensions['A'].width = 20
    c1 = sheet['N1']
    c1.value = "RR A."
    c2 = sheet['O1']
    c2.value = "Grund"

    for i in range(0, 12):
        sheet_cell = sheet.cell(row=1, column=13 - i)
        if int(last_month) - i >= 1:
            sheet_cell.value = str(int(last_month) - i) + "/" + current_year
        else:
            sheet_cell.value = str((int(last_month) + 12) - i) + "/" + str(int(current_year) - 1)

    for i in range(1, len(all_the_people)):
        person = list_of_People[i]
        sheet_cell = sheet.cell(row=i+1, column=1)
        sheet_cell.value = (person.abrk + "/"
                            + person.PN + "/" + person.name)

        if person.fall30:
            sheet_cell1 = sheet.cell(row=i+1, column=15)
            sheet_cell1.value = 30

        for k in range(0, len(person.month)):
            month_position = int(last_month) - int(person.month[k])
            if month_position >= 0:
                sheet_cell = sheet.cell(row=i+1, column=13-month_position)
                sheet_cell.value = 1
            else:
                sheet_cell = sheet.cell(row=i+1, column=1-month_position)
                sheet_cell.value = 1
        row_sum = str(i + 1)
        sheet["N" + row_sum] = f'=SUM(B{row_sum}:M{row_sum})'
    workbook.save(filename=(f"Qualität_{last_month}_{current_year}.xlsx"))


final_report()
os.remove("Dummymappe2csv.csv")
