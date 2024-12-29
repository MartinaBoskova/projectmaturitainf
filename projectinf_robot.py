import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font
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

end_of_final_report = ["Grund", 10, 13, 19, 25, 26, 27, 30, 31, 0, "Summe"]

legende_final_report = ["10 = fehlende / falsche Eingabe",
                        "11 = durch Kunde reklamierte Fehler",
                        "12 = frei",
                        "13 = frei",
                        "14 = frei",
                        "15 = Verständnisproblem",
                        "16 = falsche Berechnung",
                        "17 = Fehler aus Setup-Übernahme",
                        "18 = Programmfehler",
                        "19 = sonstige Fehlergründe",
                        "20 = Unterlagen unrichtig",
                        "21 = Lieferung nach Abgebetermin",
                        "22 = Beleg nicht eindeutig verständlich",
                        "23 = frei",
                        "24 = frei",
                        "25 = Nachzahlungen",
                        "26 = rückwirkende Ein-/Austritte",
                        "27 = ELStAM-Korrektur",
                        "28 = masch. Zahlstellenverfahren",
                        "29 = sonstige Fehlergründe",
                        "30 = vespätete Vorlage von Unterlagen",
                        "31 = Korrekturen Unterbrechung/Zeitwirtschaft",
                        "32 = fehlerhafte Datenübermittlung"]

legend_of_report = ["10 = missing / false input",
                    "11 = Fault claimed by client",
                    "12 = other KB’s mistake",
                    "13 = ADP Dresden mistake",
                    "14 = free",
                    "15 = comprehensive problem",
                    "16 = false statement of account",
                    "17 = Fault caused by Setup-Upload",
                    "18 = Program’s fault",
                    "19 = anorher fault reasons",
                    "20 = uncorrect records",
                    "21 = late data delivery",
                    "22 = uncomplete and unclear document",
                    "23 = conjuncture pile II",
                    "24 = free",
                    "25 = additional payment",
                    "26 = backward accession/leaving ",
                    "27 = free",
                    "28 = free",
                    "29 = anorher fault reasons ",
                    "30 = delayed documents submission",
                    "31 = corrections of interrupts/time-management",
                    "32 = uncorrect datatransfer"]

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


def end_of_report():
    sheet = workbook.active
    asrow = 1000+len(all_the_people)
    c1 = sheet.cell(row=asrow, column=1)
    c1.value = "Gesamtergebnis"
    for i in range(0, 12):
        row_gsmterg = str(asrow)
        column_gsmterg = chr(64 + 2 + i)
        sheet[column_gsmterg + row_gsmterg] = f'=SUM({column_gsmterg}2:{column_gsmterg}{str(len(all_the_people))})'

    c2 = sheet.cell(row=asrow+11, column=2)
    c2.value = "RR="
    c3 = sheet.cell(row=asrow+11, column=3)
    c3.value = f'=COUNT(B2:M{str(asrow)})'
    for i in range(0, len(end_of_final_report)):
        sheet_cell = sheet.cell(row=asrow+14+i, column=2)
        sheet_cell.value = end_of_final_report[i]
    for i in range(0, len(end_of_final_report)):
        sheet_cell = sheet.cell(row=asrow+15+len(end_of_final_report)+i, column=2)
        sheet_cell.value = end_of_final_report[i]
    c4 = sheet.cell(row=asrow+14, column=1)
    c4.value = "Qualität Streamline:"
    c5 = sheet.cell(row=asrow+14+len(end_of_final_report), column=1)
    c5.value = "Faktura"
    c6 = sheet.cell(row=asrow+15+len(end_of_final_report), column=1)
    c6.value = "Qualität Intern:"
    c7 = sheet.cell(row=asrow+18+2*len(end_of_final_report), column=1)
    c7.value = "Echt:"
    c8 = sheet.cell(row=asrow+20+2*len(end_of_final_report), column=1)
    c8.value = "Legende:"
    sheet.cell(row=asrow+20+2*len(end_of_final_report), column=1).font = Font(bold=True)
    c9 = sheet.cell(row=asrow+32+2*len(end_of_final_report), column=1)
    c9.value = "Legend"
    c10 = sheet.cell(row=asrow+24+2*len(end_of_final_report), column=14)
    c10.value = "Bemerkungen:"
    sheet.cell(row=asrow+24+2*len(end_of_final_report), column=14).font = Font(bold=True)

    for i in range(0, 9):
        sheet_cell = sheet.cell(row=asrow+21+i+2*len(end_of_final_report), column=1)
        sheet_cell.value = legende_final_report[i]
        sheet.merge_cells(f"A{asrow+21+i+2*len(end_of_final_report)}:E{asrow+21+i+2*len(end_of_final_report)}")
    for i in range(10, 19):
        sheet_cell = sheet.cell(row=asrow+11+i+2*len(end_of_final_report), column=7)
        sheet_cell.value = legende_final_report[i]
        sheet.merge_cells(f"G{asrow+11+i+2*len(end_of_final_report)}:L{asrow+11+i+2*len(end_of_final_report)}")
    for i in range(20, len(legende_final_report)):
        sheet_cell = sheet.cell(row=asrow+1+i+2*len(end_of_final_report), column=14)
        sheet_cell.value = legende_final_report[i]
        sheet.merge_cells(f"N{asrow+1+i+2*len(end_of_final_report)}:R{asrow+1+i+2*len(end_of_final_report)}")
    for i in range(0, 9):
        sheet_cell = sheet.cell(row=asrow+33+i+2*len(end_of_final_report), column=1)
        sheet_cell.value = legend_of_report[i]
        sheet.merge_cells(f"A{asrow+33+i+2*len(end_of_final_report)}:E{asrow+33+i+2*len(end_of_final_report)}")
    for i in range(10, 19):
        sheet_cell = sheet.cell(row=asrow+23+i+2*len(end_of_final_report), column=7)
        sheet_cell.value = legend_of_report[i]
        sheet.merge_cells(f"G{asrow+23+i+2*len(end_of_final_report)}:L{asrow+23+i+2*len(end_of_final_report)}")
    for i in range(20, len(legend_of_report)):
        sheet_cell = sheet.cell(row=asrow+13+i+2*len(end_of_final_report), column=14)
        sheet_cell.value = legend_of_report[i]
        sheet.merge_cells(f"N{asrow+13+i+2*len(end_of_final_report)}:R{asrow+13+i+2*len(end_of_final_report)}")

    sheet.cell(row=asrow+26+2*len(end_of_final_report), column=7).font = Font(color="00FF0000")
    sheet.cell(row=asrow+27+2*len(end_of_final_report), column=7).font = Font(color="00FF0000")
    sheet.cell(row=asrow+28+2*len(end_of_final_report), column=7).font = Font(color="00FF0000")
    sheet.cell(row=asrow+21+2*len(end_of_final_report), column=14).font = Font(color="00FF0000")
    sheet.cell(row=asrow+22+2*len(end_of_final_report), column=14).font = Font(color="00FF0000")


def final_report():
    sheet = workbook.active
    month_tuple = datetime.date.today().replace(day=1) - datetime.timedelta(days=1)
    last_month = month_tuple.strftime("%m")
    current_year = month_tuple.strftime("%y")
    workbook.save(filename=(f"Qualität_{last_month}_{current_year}.xlsx"))

    c = sheet['A1']
    c.value = "Zeilenbeschriftungen"
    sheet.column_dimensions['A'].width = 20
    c1 = sheet['N1']
    c1.value = "RR A."
    c2 = sheet['O1']
    c2.value = "Grund"
    end_of_report()

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
        print(person.month)
        row_sum = str(i + 1)
        sheet["N" + row_sum] = f'=SUM(B{row_sum}:M{row_sum})'
    workbook.save(filename=(f"Qualität_{last_month}_{current_year}.xlsx"))


final_report()
os.remove("Dummymappe2csv.csv")
