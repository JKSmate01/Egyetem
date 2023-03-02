from openpyxl import load_workbook
book = load_workbook('.\egyetem\ossz.xlsx',data_only=True)
sheet = book["Munka2"]
rows = 0
for row in sheet:
    for cell in row:
        if(cell.value != None):
            rows += 1


# i = 0
# n = 0
# line = sheet["D3"].value
# while True:
#     if sheet["D3"].value[i] == ".":
#         n = i+1
#         space = sheet["D2"].value[n]
#         while True:
#             print(sheet["D3"].value[n])
#             if sheet["D3"].value[n] == space:
#                 print(sheet["D3"].value[n])
#                 n+=1
#             else:
#                 break
#         break
#     i+=1
# print(i,n)
# after = line[:i+1] + line[n+1:]
# print(after)
#book.save('.\egyetem\ossz.xlsx')
#for i in range(2,rows+1):