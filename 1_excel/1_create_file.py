from openpyxl import Workbook
wb = Workbook()  # 새 워크북 생성
ws = wb.active  # 현재 활성화 시트 가져옴
ws.title = "Test1"  # 워크시트의 이름을 변경
wb.save("sample.xlsx")
wb.close()
