import openpyxl
import os
import subprocess


# Beenden des Programms bei falscher Eingabe
def not_valid(end):
    for i in range(5):
        print(f"Please select a file in format: C:/example/of/path/Name_of_file{end}")
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


# Überprüfung der Zelle A1 in der gegebenen Excel-Datei zur Initiierung des nächsten Skripts
def decision():
    if sheet_obj['A1'].value == "Zeilenbeschriftungen":
        return f"{project_path}/projectinf_text_nj.py"
    elif sheet_obj['A1'].value == "Firma":
        return f"{project_path}/projectinf_robot_nj.py"
    else:
        print("Error. File given in wrong format - Program ends.")
        exit()


print("Hi! You've started my project.")
project_path = os.getcwd()

path = not_valid(".xlsx")
wb_obj = openpyxl.load_workbook(path, data_only=True)
sheet_obj = wb_obj.active

# Start des nächsten Skripts – Übergabe der Eingabe
decided_script = decision()
subprocess.run([f"{project_path}/venv/scripts/python.exe", decided_script, path])