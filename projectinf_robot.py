import openpyxl

path = "Dummymappe1.xlsx"

wb_obj = openpyxl.load_workbook(path)

sheet_obj = wb_obj.active

row_number = sheet_obj.max_row
column_number = sheet_obj.max_column

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


def fall_30(x):
    for i in range(2, row_number+1):
        Lohnartbeschreibung = sheet_obj.cell(row=i, column=12)
        for j in range(0, len(Fall30)):
            if Lohnartbeschreibung.value == Fall30[j]:
                print("Fall 30 detected")
                x = x + 1
                break
            else:
                continue
        i = i + 1
    print("Number of 30 detected is:", x)


people_number(0)
fall_30(0)
