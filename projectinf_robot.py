import openpyxl
from openpyxl import Workbook
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

ListFall30 = ["ABT SFN st/sv-pfl",
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

End_of_finalreport = ["Grund", 10, 13, 19, 25, 26, 27, 30, 31, 0, "Summe"]

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

All_the_People = list(dict.fromkeys(people_dict))


class Person:
    def __init__(self, i_in_people):
        Current_person = All_the_People[i_in_people]
        self.Abrk = people_dict[Current_person][0][0]
        self.Name = people_dict[Current_person][0][2]
        self.PN = people_dict[Current_person][0][1]
        self.Month = []
        self.Lohn = []
        self.Fall30 = False


def People_classes(x):
    Local_month = []
    Local_lohn = []
    for i in people_dict[All_the_People[x]]:
        Local_month.append(i[3])
        Local_lohn.append(i[11])
    Local_month = list(dict.fromkeys(Local_month))
    Local_lohn = list(dict.fromkeys(Local_lohn))

    for k in Local_lohn:
        for j in ListFall30:
            if k == j:
                Fall30 = True
                break

    Person(x).Month = Local_month
    Person(x).Lohn = Local_lohn
    Person(x).Fall30 = Fall30
    return Person(x)


List_of_People = [Person]
for i in range(1, len(All_the_People)):
    List_of_People.append(People_classes(i))


def End_of_report():
    sheet = workbook.active
    active_space_row = 1000+len(All_the_People)
    c3 = sheet.cell(row=active_space_row, column=1)
    c3.value = "Gesamtergebnis"
    for i in range(0, 12):
        row_gsmterg = str(active_space_row)
        column_gsmterg = chr(64 + 2 + i)
        sheet[column_gsmterg + row_gsmterg] = f'=SUM({column_gsmterg}2:{column_gsmterg}{str(len(All_the_People))})'

    c4 = sheet.cell(row=active_space_row+11, column=2)
    c4.value = "RR="
    c5 = sheet.cell(row=active_space_row+11, column=3)
    c5.value = f'=COUNT(B2:M{str(active_space_row)})'
    for i in range(0, len(End_of_finalreport)):
        sheet_cell = sheet.cell(row=active_space_row+14+i, column=2)
        sheet_cell.value = End_of_finalreport[i]
    for i in range(0, len(End_of_finalreport)):
        sheet_cell = sheet.cell(row=active_space_row+15+len(End_of_finalreport)+i, column=2)
        sheet_cell.value = End_of_finalreport[i]
    c6 = sheet.cell(row=active_space_row+14, column=1)
    c6.value = "Qualität Streamline:"
    c7 = sheet.cell(row=active_space_row+14+len(End_of_finalreport), column=1)
    c7.value = "Faktura"
    c8 = sheet.cell(row=active_space_row+15+len(End_of_finalreport), column=1)
    c8.value = "Qualität Intern:"
    c9 = sheet.cell(row=active_space_row+18+2*len(End_of_finalreport), column=1)
    c9.value = "Echt:"


def Final_report():
    month_tuple = datetime.date.today().replace(day=1) - datetime.timedelta(days=1)
    last_month = month_tuple.strftime("%m")
    current_year = month_tuple.strftime("%y")
    workbook.save(filename=(("Qualität_" + last_month + "_" + current_year + ".xlsx")))

    sheet = workbook.active
    c = sheet['A1']
    c.value = "Zeilenbeschriftungen"
    sheet.column_dimensions['A'].width = 20
    c1 = sheet['N1']
    c1.value = "RR A."
    c2 = sheet['O1']
    c2.value = "Grund"
    End_of_report()

    for i in range(0, 12):
        sheet_cell = sheet.cell(row=1, column=13 - i)
        if int(last_month) - i >= 1:
            sheet_cell.value = str(int(last_month) - i) + "/" + current_year
        else:
            sheet_cell.value = str((int(last_month) + 12) - i) + "/" + str(int(current_year) - 1)

    for i in range(1, len(All_the_People)):
        sheet_cell = sheet.cell(row=i+1, column=1)
        sheet_cell.value = (Person(i).Abrk + "/"
                            + Person(i).PN + "/" + Person(i).Name)

        if Person(i).Fall30:
            sheet_cell1 = sheet.cell(row=i+1, column=15)
            sheet_cell1.value = 30

        for k in range(0, len(Person(i).Month)):
            month_position = int(last_month) - int(Person(i).Month[k])
            if month_position >= 0:
                sheet_cell = sheet.cell(row=i+1, column=13-month_position)
                sheet_cell.value = 1
            else:
                sheet_cell = sheet.cell(row=i+1, column=1-month_position)
                sheet_cell.value = 1

        row_sum = str(i + 1)
        column_sum = chr(64 + 14)
        sheet[column_sum + row_sum] = f'=SUM(B{row_sum}:M{row_sum})'
    workbook.save(filename=(("Qualität_" + last_month + "_" + current_year + ".xlsx")))


Final_report()
os.remove("Dummymappe2csv.csv")
