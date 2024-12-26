import openpyxl
from openpyxl import Workbook
import time

workbook = Workbook()

path = "Dummymappe1.xlsx"

wb_obj = openpyxl.load_workbook(path)

sheet_obj = wb_obj.active

row_number = sheet_obj.max_row
column_number = sheet_obj.max_column

Column_Name = 0
Column_Lohn = 0

Fall30 = ["ABT SFN st/sv-pfl",
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

print("Total Rows:", row_number)
print("Total Columns:", column_number)


def Namecolumn(x):
    for i in range(1, column_number + 1):
        name_column = sheet_obj.cell(row=1, column=i)
        if name_column.value == "Name" or name_column == "name":
            x = i
            break
        else:
            i = i + 1
    Column_Name = x
    if Column_Name == 0:
        print("Column with Namen not found. Please change the title of the column to 'Name or name'.")
    return Column_Name


def Lohncolumn(x):
    for i in range(1, column_number + 1):
        lohn_column = sheet_obj.cell(row=1, column=i)
        if lohn_column.value == "Lohnartbeschreibung" or lohn_column == "lohnartbeschreibung":
            x = i
            break
        else:
            i = i + 1
    Column_Lohn = x
    if Column_Lohn == 0:
        print("Column with Lohnartbeschreibungen not found. Please change the title of the column to 'Lohnartbeschreibung or lohnartbeschreibung'.")
    return Column_Lohn


def people_number(x):

    for i in range(2, row_number+1):
        name_a = sheet_obj.cell(row=i, column=Namecolumn(Column_Name))
        name_b = sheet_obj.cell(row=i + 1, column=Namecolumn(Column_Name))
        if name_a.value == name_b.value:
            i = i + 1
        else:
            x = x + 1
            i = i + 1
    print("Number of people is:", x)


def Final_report():
    named_tuple = time.localtime()
    current_month = time.strftime("%m", named_tuple)
    current_year = time.strftime("%y", named_tuple)
    workbook.save(filename=(("Qualität_" + current_month + "_" + current_year + ".xlsx")))

    sheet = workbook.active
    c = sheet['A1']
    c.value = "Zeilenbeschriftungen"
    sheet.column_dimensions['A'].width = 20
    c1 = sheet['N1']
    c1.value = "RR A."
    c2 = sheet['O1']
    c2.value = "Grund"

    for i in range(1, 13):
        sheet_cell = sheet.cell(row=1, column=i+1)
        if not i == current_month and i < 10:
            j = str(i)
            sheet_cell.value = "0" + j + "/" + current_year
            i = i + 1
        elif not i == current_month and i >= 10:
            j = str(i)
            sheet_cell.value = j + "/" + current_year
            i = i + 1
        else:
            sheet_cell.value = current_month + "/" + current_year
            break

    for i in range(0, len(list_names)):
        sheet_cell = sheet.cell(row=i+2, column=1)
        j = str(list_AbrK[i])
        k = str(list_PN[i])
        sheet_cell.value = (j + "/" + k + "/" + list_names[i])
        sheet_cell1 = sheet.cell(row=i+2, column=15)
        sheet_cell1.value = 30
        i = i + 1

    workbook.save(filename=(("Qualität_" + current_month + "_" + current_year + ".xlsx")))


def fall_30(x, y):
    remembered_names = list(())
    remembered_AbrK = list(())
    remembered_PN = list(())
    for i in range(2, row_number+1):
        Lohnartbeschreibung = sheet_obj.cell(row=i, column=Lohncolumn(Column_Lohn))
        name_a = sheet_obj.cell(row=i, column=Namecolumn(Column_Name))
        data_a = sheet_obj.cell(row=i, column=Namecolumn(Column_Name)-1)
        data_b = sheet_obj.cell(row=i, column=Namecolumn(Column_Name)-2)
        name_b = sheet_obj.cell(row=i + 1, column=Namecolumn(Column_Name))
        if name_a.value == name_b.value and y == 1:
            for j in range(0, len(Fall30)):
                if Lohnartbeschreibung.value == Fall30[j]:
                    y = y + 1
                    print("Fall 30 detected")
                    remembered_names.insert(x, name_a.value)
                    remembered_PN.insert(x, data_a.value)
                    remembered_AbrK.insert(x, data_b.value)
                    x = x + 1
                    break
                else:
                    continue
        elif name_a.value == name_b.value and y == 2:
            i = i + 1
        elif not name_a.value == name_b.value and y == 2:
            y = y - 1
            i = i + 1
        else:
            i = i + 1
    print("Number of Fall 30 detected is:", x)
    print(remembered_names)
    return remembered_names, remembered_PN, remembered_AbrK

list_names, list_PN, list_AbrK = fall_30(0, 1)

Namecolumn(Column_Name)
print("Column with Namen is letter:", chr(64 + Namecolumn(Column_Name)))

Lohncolumn(Column_Lohn)
print("Column with Lohnartbeschreibungen is letter:", chr(64 + Lohncolumn(Column_Lohn)))

people_number(0)
fall_30(0, 1)
Final_report()
