from openpyxl import load_workbook
import re
import simplejson as json


# The final results will be here
result = {}

# Make a list of all individual sheets we want to parse, column=lang associations and the file location
file_name = input("Type xlsx file location: ")
sheet_names = input("What sheets do you want to parse? (ex. common, sheet2, another sheet): ")
column_names = input("Which two columns would you like to parse? (ex. B=en, C=de, D=cz): ")

# Split the comma separated inputs into lists
sheets = sheet_names.split(", ")
cn = column_names.split(", ")

# Begin creating language control variable
languages = {}
ck = ""
cv = ""

langs = []

for c in cn:
    ck = c.split("=")[0]
    cv = c.split("=")[1]
    languages[ck] = cv

    langs.append(cv)


#Open Spreadsheet
wb = load_workbook(filename=file_name, read_only=False)

for sv in sheets:
    ws = wb[sv]  # ws is now an IterableWorksheet

    for row in ws.rows:
        values = []

        for cell in row:
            col = re.findall('[A-Za-z]+', cell.coordinate)[0]
            row = re.findall('[0-9]+', cell.coordinate)[0]

            # Basically here we just check if the column is in our language associations and if so, build the lang
            # object from its value
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

        # Compile the final results variable from all parsed and organized data
        result[values[count2]] = rs


# Finally, save the results to a json file, and print out a success message.
with open("lang.json", "w+") as f:
    f.write(json.dumps(result))
    f.close()

print("Excel document has been successfully transpiled to json, and saved in the current directory as lang.json")