from openpyxl import load_workbook
DATA_PATH = ".\\ossz.xlsx"
book = load_workbook(DATA_PATH,data_only=True)
sheet = book["Egybe"]
rows = 0
#Sorok felszámlálása
end = False
for row in sheet["D2":"D100000"]:
    for cell in row:
        if (cell.value == "END"):
            end = True
            break
        rows += 1
    if (end):
        break

hdoszlista = []
for i in range(2,rows+2):
    if (sheet[f"C{i}"].value != None):
        if sheet[f"C{i}"].value.rfind("\xa0") >=0:
                while True:
                    if sheet[f"C{i}"].value.rfind("\xa0") < 0:
                        break
                    sheet[f"C{i}"].value = sheet[f"C{i}"].value[:sheet[f"C{i}"].value.rfind("\xa0")] + sheet[f"C{i}"].value[sheet[f"C{i}"].value.rfind("\xa0")+4:]
        try:
            list.index(hdoszlista,sheet[f"C{i}"].value)
        except ValueError:
            hdoszlista.append(sheet[f"C{i}"].value)
print(hdoszlista)