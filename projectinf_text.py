import openpyxl
from openpyxl import Workbook
import csv
import os

workbook = Workbook()

# Otvírání excel souboru ve formátu konečný report nevyplněný
project_path = "C:/Users/Martina/Desktop/škola/informatika/git.projectinf/"
print(f"Please select a file in format: {project_path}Name.xlsx")
# path = filename = input()
path = "C:/Users/Martina/Desktop/škola/informatika/git.projectinf/Qualität_01_25.xlsx"
wb_obj = openpyxl.load_workbook(path, data_only=True)
sheet_obj = wb_obj.active

print(f"Please select a file in format: {project_path}Name.txt")
text_path = "C:/Users/Martina/Desktop/škola/informatika/git.projectinf/DataQuali.txt"
# text_path = filename = input()

# Převedení na csv
with open(f"{path}.csv", "w", newline="") as file_handle:
    csv_writer = csv.writer(file_handle, delimiter=";")
    for row in sheet_obj.iter_rows():
        csv_writer.writerow([cell.value for cell in row])

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
print(f"{people_dict}")
print(f"{months_list}")

all_the_people = list(people_dict.keys())


def what_month(person):
    months = []
    for i in range(1, 13):
        if person[i].isspace() or person[i] == "":
            pass
        elif person[i] == "1":
            local_month = months_list[i]
            months.append(local_month)
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

with open(f"{text_path}", "r") as f:
    text_rows = f.readlines()
    for person in list_of_People:
        for line in text_rows:
            if f"{person.find}" in line:
                print("yay")
                person.found = True

print(f"{list_of_People[1].found}")
os.remove(f"{project_path}Qualität_01_25.xlsx.csv")
