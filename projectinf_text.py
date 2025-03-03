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


# Třída každého člověka s důležitým info
class Person:
    def __init__(self, i_in_people):
        current_person = all_the_people[i_in_people]
        self.abrk = people_dict[current_person][0][0]
        self.name = people_dict[current_person][0][2]
        self.PN = people_dict[current_person][0][1]
        self.month = []
        self.lohn = []
        self.fall30 = False


os.remove(f"{project_path}Qualität_01_25.xlsx.csv")
