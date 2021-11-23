# code pour importer les qualification afin de les entrer automatiquement dans la BD

import openpyxl
from pathlib import Path

xlsx_file = Path('qualification_LCP.xlsx')
wb_object = openpyxl.load_workbook(xlsx_file)

print("---------------------")
print("wb_object", wb_object)
print("---------------------")

activeSheet = wb_object.active 


print("activeSheet : ", activeSheet)
print("---------------------")

rowNb = activeSheet.max_row
colNb = activeSheet.max_column

print("nombre des lignes = ", rowNb)
print("nombre des colonnes = ", colNb)
print("---------------------")

print("A1 : ", activeSheet["A1"].value)
print("---------------------")

cell = "A" + str(2)
print("cell : ", cell)
print("cell type : ", type(cell))
print("---------------------")
couter = 0


###################################################################
#
# passer cellule par cellule et chercher un mot clefs spepicifique
#
###################################################################
for row in activeSheet.iter_rows(max_row=rowNb):
    for cell in row:
        print("cell.value = ", cell.value)
    print(" ") 