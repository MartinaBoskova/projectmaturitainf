import openpyxl
import csv
import os

l_fall30_vers = ["AN Arbeitslosenve",
                 "AN Krankenvers.",
                 "AN Plegeversich.",
                 "AN Rentenversich."]

l_fall_27 = ["AG Pauschsteuer",
             "AN Pauschsteuer",
             "Kirchensteuer",
             "Lohnsteuer"]


def not_valid():
    print("Invalid input given")
    if i == 4:
        print("Invalid input given five times - Program ends.")
        exit()


# Otvírání excel souboru ve formátu konečný report nevyplněný
project_path = "C:/Users/Martina/Desktop/škola/informatika/git.projectinf/"
path = "C:/Users/Martina/Desktop/škola/informatika/git.projectinf/Qualität_01_25.xlsx"
#for i in range(5):
#    print(f"Please select a file in format: {project_path}Name.xlsx")
#    path = input()
#    try:
#        if not path.endswith(".xlsx"):
#            raise ValueError()
#        with open(path, "r") as f:
#            pass
#        break
#    except ValueError:
#        not_valid()
wb_obj = openpyxl.load_workbook(path, data_only=True)
sheet_obj = wb_obj.active

# Otvírání textového souboru ve formátu výplatnice
text_path = "C:/Users/Martina/Desktop/škola/informatika/git.projectinf/DataQuali.txt"
#for i in range(5):
#    print(f"Please select a file with payrolls in format: {project_path}Name.txt")
#    text_path = input()
#    try:
#        if not text_path.endswith(".txt"):
#            raise ValueError()
#        with open(text_path, "r") as f:
#            pass
#        break
#    except ValueError:
#        not_valid()

# Převedení excelu na csv
with open(f"{path}.csv", "w", newline="") as file_handle:
    csv_writer = csv.writer(file_handle, delimiter=";")
    for row in sheet_obj.iter_rows():
        csv_writer.writerow([cell.value for cell in row])

# Vytváření listu z daných Fallů 30
with open("Fall30.txt", "r") as fall_30:
    lines_from_text = fall_30.readlines()
    for i in range(0, len(lines_from_text)):
        lines_fall_30 = lines_from_text[i].replace("\n", "")
        lines_fall_30 = [line.strip() for line in lines_from_text if line.strip()]

# Vytvoření dictionary ze všech lidí v dokumentu
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


# Vytvoření stringu ve formátu na hledání v txt. souboru
def what_month(person):
    months = []
    for i in range(1, 13):
        if person[i].isspace() or person[i] == "":
            pass
        elif person[i] == "1":
            separate = months_list[i].split("/")
            month = int(separate[0])
            if month < 10:
                separate.pop(0)
                separate.insert(0, f"0{month}")
            month_in_format = ".".join(separate) + "/"
            months.append(month_in_format)
        else:
            print("Error has ocured program will end. (check values in months)")
            exit()
    months.reverse()
    return months


# Třída každého člověka s důležitým info
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


# Loop skrz všechny lidi
list_of_People = []
for i in range(1, len(all_the_people)):
    list_of_People.append(Person(i))

# Projíždění skrz textový soubor pro každého člověka
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
                        first_dash = line + 1
                        line = line - 1
                        # Pro nalezenou výplatnici prohlížení jasného fallu 30
                        while person.fall30 is False and dash not in text_rows[line]:
                            lohnart = text_rows[line][:17]
                            person.fall30 = any(k in lohnart for k in lines_fall_30)
                            line = line - 1
                        # Pro nalezenou výplatnici prohlížení fallu 30 a fallu 27
                        while person.fall30 is False and dash not in text_rows[first_dash]:
                            first_dash = first_dash + 1
                        scnd_dash = first_dash + 1
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
                                if text_rows[scnd_dash][38].isspace() or text_rows[scnd_dash][38] == "":
                                    scnd_dash = scnd_dash + 1
                                else:
                                    person.fall27 = True
                                    scnd_dash = scnd_dash + 1
                            else:
                                scnd_dash = scnd_dash + 1
                        continue
            else:
                continue

# Zapsání výsledků do konečného reportu
for i in range(0, len(list_of_People)):
    grund_cell = sheet_obj[f"O{i+2}"]
    person = list_of_People[i]

    if person.fall30 is True:
        grund_cell.value = "30"
    if person.fall30 is False and person.fall27 is True:
        grund_cell.value = "27"

# Vypsání nenalezených lidí
for person in list_of_People:
    if person.found is False:
        print(f"Error. Person {person.PN} not found in text file.")

wb_obj.save(path)
os.remove(f"{path}.csv")
