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
        try:
            list.index(hdoszlista,sheet[f"C{i}"].value)
        except ValueError:
            hdoszlista.append(sheet[f"C{i}"].value)
print(hdoszlista)
out = open(".\\C.txt","w")
for line in hdoszlista:
    out.write(f"{line}\n")
out.close()
nevek = []
for i in range(2,rows+2):
    if (sheet[f"G{i}"].value != None):
        try:
            list.index(nevek,sheet[f"G{i}"].value)
        except ValueError:
            nevek.append(sheet[f"G{i}"].value)
print(nevek)
out = open(".\\Nevek.txt","w")
for line in nevek:
    out.write(f"{line}\n")
out.close()
