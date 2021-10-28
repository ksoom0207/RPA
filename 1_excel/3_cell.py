from openpyxl import Workbook
wb = Workbook()

ws = wb.active
ws.title = "SOOMSheet"

ws["A1"] = 1
ws["A2"] = 2
ws["A3"] = 3
ws["B1"] = 1
ws["B2"] = 2
ws["B3"] = 3

print(ws["A1"])

wb.save("sample2.xlsx")
