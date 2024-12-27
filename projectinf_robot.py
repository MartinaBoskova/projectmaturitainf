import openpyxl
from openpyxl import Workbook
import datetime

workbook = Workbook()

path = "Dummymappe1.xlsx"

wb_obj = openpyxl.load_workbook(path)

sheet_obj = wb_obj.active

row_number = sheet_obj.max_row
column_number = sheet_obj.max_column

Column_Name = 0
Column_Lohn = 0
Column_Month = 0

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
        exit()
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
        exit()
    return Column_Lohn


def Monthcolumn(x):
    for i in range(1, column_number + 1):
        month_column = sheet_obj.cell(row=1, column=i)
        if month_column.value == "Monat" or month_column == "monat":
            x = i
            break
        else:
            i = i + 1
    Column_Month = x
    if Column_Month == 0:
        print("Column with Months not found. Please change the title of the column to 'Monat or monat'.")
        exit()
    return Column_Month


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

    for i in range(0, 12):
        sheet_cell = sheet.cell(row=1, column=13 - i)
        if int(last_month) - i >= 1:
            sheet_cell.value = str(int(last_month) - i) + "/" + current_year
        else:
            sheet_cell.value = str((int(last_month) + 12) - i) + "/" + str(int(current_year) - 1)
#       if not i == last_month and i < 10:
#            j = str(i)
#            sheet_cell.value = "0" + j + "/" + current_year
#            i = i + 1
#        elif not i == last_month and i >= 10:
#            j = str(i)
#            sheet_cell.value = j + "/" + current_year
#            i = i + 1
#        else:
#            sheet_cell.value = last_month + "/" + current_year
#            break

    for i in range(0, len(list_names)):
        sheet_cell = sheet.cell(row=i+2, column=1)
        AbrK = str(list_AbrK[i])
        PN = str(list_PN[i])
        sheet_cell.value = (AbrK + "/" + PN + "/" + list_names[i])
        for j in range(0, len(list_names30)):
            if list_names30[j] == list_names[i]:
                sheet_cell1 = sheet.cell(row=i+2, column=15)
                sheet_cell1.value = 30
                break
            else:
                j = j + 1
        i = i + 1

    for i in range(0, len(list_months)):
        monthvalue = list_months[i]
        if isinstance(monthvalue, list):
            for j in range(0, len(monthvalue)):
                moremonthvalue = monthvalue[j]
                sheet_cell = sheet.cell(row=i+2, column=moremonthvalue+1)
                sheet_cell.value = 1
            i = i + 1
        else:
            sheet_cell = sheet.cell(row=i+2, column=monthvalue+1)
            sheet_cell.value = 1
            i = i + 1
    workbook.save(filename=(("Qualität_" + last_month + "_" + current_year + ".xlsx")))


def fall_30(x, y):
    remembered_names_30 = list(())
    remembered_names = list(())
    remembered_AbrK = list(())
    remembered_PN = list(())
    for i in range(2, row_number+1):
        Lohnartbeschreibung = sheet_obj.cell(row=i, column=Lohncolumn(Column_Lohn))
        name_a = sheet_obj.cell(row=i, column=Namecolumn(Column_Name))
        data_a = sheet_obj.cell(row=i, column=Namecolumn(Column_Name)-1)
        data_b = sheet_obj.cell(row=i, column=Namecolumn(Column_Name)-2)
        name_b = sheet_obj.cell(row=i + 1, column=Namecolumn(Column_Name))

        remembered_names.insert(i-2, name_a.value)
        remembered_AbrK.insert(i-2, data_b.value)
        remembered_PN.insert(i-2, data_a.value)
        if name_a.value == name_b.value and y == 1:
            for j in range(0, len(Fall30)):
                if Lohnartbeschreibung.value == Fall30[j]:
                    y = y + 1
                    print("Fall 30 detected")
                    remembered_names_30.insert(x, name_a.value)
                    x = x + 1
                    break
                else:
                    continue
            i = i + 1
        elif not name_a.value == name_b.value and y == 1:
            for j in range(0, len(Fall30)):
                if Lohnartbeschreibung.value == Fall30[j]:
                    y = y + 1
                    print("Fall 30 detected")
                    remembered_names_30.insert(x, name_a.value)
                    x = x + 1
                    break
                else:
                    continue
            i = i + 1
        elif not name_a.value == name_b.value and y == 2:
            y = y - 1
            i = i + 1
        else:
            i = i + 1
    print("Number of Fall 30 detected is:", x)
    print(remembered_names_30)
    print(remembered_names)
    remembered_names = list(dict.fromkeys(remembered_names))
    print(remembered_names)
    return remembered_names, remembered_PN, remembered_AbrK, remembered_names_30


list_names, list_PN, list_AbrK, list_names30 = fall_30(0, 1)


def month_count(x, y):
    remembered_months = list(())
    for i in range(2, row_number):
        name_a = sheet_obj.cell(row=i, column=Namecolumn(Column_Name))
        name_b = sheet_obj.cell(row=i + 1, column=Namecolumn(Column_Name))
        month_a = sheet_obj.cell(row=i, column=Monthcolumn(Column_Month))
        month_b = sheet_obj.cell(row=i + 1, column=Monthcolumn(Column_Month))

        if name_a.value == name_b.value and month_a.value == month_b.value and y == 1:
            remembered_months.insert(x, month_a.value)
            x = x + 1
            y = 2
            i = i + 1
        elif name_a.value == name_b.value and not month_a.value == month_b.value and y == 1:
            remembered_months.insert(x, list((month_a.value, month_b.value)))
            x = x + 1
            y = 2
            i = i + 1
        elif not name_a.value == name_b.value and y == 1:
            remembered_months.insert(x, month_a.value)
            x = x + 1
            remembered_months.insert(x, month_b.value)
            x = x + 1
            y = 2
            i = i + 1
#        elif name_a.value == name_b.value and month_a.value == month_b.value and y > 2:
#            i = i + 1
        elif name_a.value == name_b.value and not month_a.value == month_b.value and y > 2:
            remembered_months.insert(y - 1, list.insert(y-1, month_b.value))
            i = i + 1
        elif not name_a.value == name_b.value and y > 2:
            remembered_months.insert(x, month_b.value)
            x = x + 1
            y = 2
            i = i + 1
        elif name_a.value == name_b.value and not month_a.value == month_b.value and y == 2:
            remembered_months.pop(x - 1)
            remembered_months.insert(x - 1, list((month_a.value, month_b.value)))
            y = y + 1
            i = i + 1
#        elif name_a.value == name_b.value and month_a.value == month_b.value and y == 2:
#            i = i + 1
        elif not name_a.value == name_b.value and y == 2:
            remembered_months.insert(x, month_b.value)
            x = x + 1
            i = i + 1
        else:
            i = i + 1
    print(remembered_months)
    return remembered_months


list_months = month_count(0, 1)


Namecolumn(Column_Name)
if Namecolumn(Column_Name) <= 26:
    print("Column with Namen is letter:", chr(64 + Namecolumn(Column_Name)))
else:
    first_chr = int(Namecolumn(Column_Name) / 26)
    second_chr = Namecolumn(Column_Name) - (26 * first_chr)
    print("Column with Namen is letter:", chr(64 + first_chr), chr(64 + second_chr))

Lohncolumn(Column_Lohn)
if Lohncolumn(Column_Lohn) <= 26:
    print("Column with Lohnartbeschreibungen is letter:", chr(64 + Lohncolumn(Column_Lohn)))
else:
    firs1_chr = int(Lohncolumn(Column_Lohn) / 26)
    second1_chr = Lohncolumn(Column_Lohn) - (26 * firs1_chr)
    print("Column with Lohnartbeschreibungen is letter:", chr(64 + firs1_chr), chr(64 + second1_chr))
Monthcolumn(Column_Month)

people_number(0)
fall_30(0, 1)
month_count(0, 1)
Final_report()
