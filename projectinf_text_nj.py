import openpyxl
import csv
import os
import sys

l_fall30_vers = ["AN Arbeitslosenve",
                 "AN Krankenvers.",
                 "AN Plegeversich.",
                 "AN Rentenversich."]

l_fall_27 = ["AG Pauschsteuer",
             "AN Pauschsteuer",
             "Kirchensteuer",
             "Lohnsteuer"]


def not_valid(end):
    for i in range(5):
        print(f"Please select a file in format: C:/example/of/path/Name{end}")
        end_path = input()
        try:
            if not end_path.endswith(end):
                raise ValueError()
            with open(end_path, "r"):
                pass
            return end_path
        except ValueError:
            print("Invalid input given")
        if i == 4:
            print("Error: Invalid input given five times - Program ends.")
            exit()


# path = "C:/Users/Martina/Desktop/škola/informatika/git.projectinf/Qualität_01_25.xlsx"
# path = not_valid(".xlsx")

# Öffnen einer Excel-Datei im Format "Endbericht (unvollständig)
path = sys.argv[1]
wb_obj = openpyxl.load_workbook(path, data_only=True)
sheet_obj = wb_obj.active
copy_obj = wb_obj.copy_worksheet(sheet_obj)
# text_path = "C:/Users/Martina/Desktop/škola/informatika/git.projectinf/DataQuali.txt"

# Öffnen einer Textdatei im Format "Gehaltsabrechnung"
text_path = not_valid(".txt")

# Umwandlung einer Excel-Datei in CSV
with open(f"{path}.csv", "w", newline="") as file_handle:
    csv_writer = csv.writer(file_handle, delimiter=";")
    for row in sheet_obj.iter_rows():
        csv_writer.writerow([cell.value for cell in row])

# Erstellen eines Sheets aus den gegebenen Fällen (Fall 30)
with open("Fall30.txt", "r") as fall_30:
    lines_from_text = fall_30.readlines()
    for i in range(0, len(lines_from_text)):
        lines_fall_30 = lines_from_text[i].replace("\n", "")
        lines_fall_30 = [line.strip() for line in lines_from_text if line.strip()]

# Erstellen eines Dictionaries mit allen Personen im Dokument
with open(f"{path}.csv", "r", encoding="latin-1", newline="") as f:
    csv_rows = list(csv.reader(f, delimiter=';'))
    people_dict = {}
    months_list = csv_rows[0]
    for line in csv_rows:
        if line[0].isspace() or line[0] == "":
            break
        else:
            if line[0] not in people_dict:
                people_dict[line[0]] = [line]
            else:
                people_dict[line[0]].append(line)

all_the_people = list(people_dict.keys())


# Erstellen eines Strings im Suchformat für eine TXT-Datei
def what_month(person):
    months = []
    for i in range(1, 13):
        if person[i].isspace() or person[i] == "":
            pass
        elif person[i] == "1":
            separate = months_list[i].split("/")
            month = int(separate[0])
            "{:02d}".format(month)
            month_in_format = ".".join(separate) + "/"
            months.append(month_in_format)
        else:
            print("Error has ocured program will end. (check values in months)")
            exit()
    months.reverse()
    return months


# Klasse für jede Person mit wichtigen Informationen
class Person:
    def __init__(self, i_in_people):
        current_person = all_the_people[i_in_people]
        separate = people_dict[current_person][0][0].split("/")
        self.abrk = separate[0]
        self.name = separate[2]
        self.PN = separate[1]
        self.find = f"{separate[0]}/{separate[1]}"
        self.found = False
        self.fall30 = False
        self.fall27 = False
        self.month = what_month(people_dict[current_person][0])


# Loop durch alle Personen
list_of_People = []
for i in range(1, len(all_the_people)):
    list_of_People.append(Person(i))

# Durchsuchen der Textdatei für jede Person
with open(f"{text_path}", "r") as f:
    dash = "-----"
    text_rows = f.readlines()
    for person in list_of_People:
        for line in range(0, len(text_rows)):
            if person.found is False:
                if f"{person.find}" in text_rows[line]:
                    if f"{person.month[0]}" in text_rows[line - 4]:
                        person.found = True
                        while dash not in text_rows[line]:
                            line = line + 1
                        frst_dash = line + 1
                        line = line - 1
                        # Bei gefundener Gehaltsabrechnung Überprüfung des spezifischen Falls 30
                        while person.fall30 is False and dash not in text_rows[line]:
                            lohnart = text_rows[line][:17]
                            person.fall30 = any(k in lohnart for k in lines_fall_30)
                            line = line - 1
                        # Bei gefundener Gehaltsabrechnung Überprüfung von Fall 30 und Fall 27
                        while person.fall30 is False and dash not in text_rows[frst_dash]:
                            frst_dash = frst_dash + 1
                        scnd_dash = frst_dash + 1
                        while person.fall30 is False and dash not in text_rows[scnd_dash]:
                            lohnart = text_rows[scnd_dash][:17]
                            vers = any(k in lohnart for k in l_fall30_vers)
                            steuer = any(k in lohnart for k in l_fall_27)
                            if vers is True:
                                if text_rows[scnd_dash][38].isspace() or text_rows[scnd_dash][38] == "":
                                    scnd_dash = scnd_dash + 1
                                else:
                                    person.fall30 = True
                            if steuer is True:
                                if text_rows[scnd_dash][45].isspace() or text_rows[scnd_dash][45] == "":
                                    scnd_dash = scnd_dash + 1
                                else:
                                    person.fall27 = True
                                    scnd_dash = scnd_dash + 1
                            else:
                                scnd_dash = scnd_dash + 1
                        continue
            else:
                continue

# Eintragen der Ergebnisse in den Endbericht
for i in range(0, len(list_of_People)):
    grund_cell = copy_obj[f"O{i+2}"]
    person = list_of_People[i]

    if person.fall30 is True:
        grund_cell.value = "30"
    if person.fall30 is False and person.fall27 is True:
        grund_cell.value = "27"

# Ausgabe der nicht gefundenen Personen
for person in list_of_People:
    if person.found is False:
        print(f"Error. Person {person.PN} or payroll with last month not found in text file.")

wb_obj.save(path)
print("You should be able to find the finished final report in your excel file.")
os.remove(f"{path}.csv")
