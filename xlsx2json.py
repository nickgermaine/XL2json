from openpyxl import load_workbook
import re
import simplejson as json


# The final results will be here
result = {}

# Make a list of all individual sheets we want to parse
file_name = input("Type xlsx file location: ")
if not file_name:
    file_name = "imports/translations.xlsx"
sheet_names = input("What sheets do you want to parse? (ex. common, sheet2, another sheet): ")
if not sheet_names:
    sheet_names = "Common, Login, Dashboard, Reports, Distribution"

column_names = input("Which two columns would you like to parse? (ex. B=en, C=de, D=cz): ")
if not column_names:
    column_names = "B=en, C=de"

sheets = sheet_names.split(", ")

cn = column_names.split(", ")

languages = {}
ck = ""
cv = ""

langs = []

for c in cn:
    ck = c.split("=")[0]
    cv = c.split("=")[1]
    languages[ck] = cv

    langs.append(cv)

print("languages:", languages)
wb = load_workbook(filename=file_name, read_only=False)

for sv in sheets:
    ws = wb[sv]  # ws is now an IterableWorksheet

    for row in ws.rows:
        values = []

        for cell in row:
            col = re.findall('[A-Za-z]+', cell.coordinate)[0]
            row = re.findall('[0-9]+', cell.coordinate)[0]

            for k, v in languages.items():
                if col == k:
                    values.append(cell.value)

        length = len(langs)
        count = 0
        count2 = 0
        rs = {}
        for l in langs:
            rs[l] = values[count]
            if count > length:
                count = 0
            else:
                count += 1

        result[values[count2]] = rs

with open("lang.json", "w+") as f:
    f.write(json.dumps(result))
    f.close()

print(result)