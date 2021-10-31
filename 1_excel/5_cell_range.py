from openpyxl.utils.cell import coordinate_from_string
from openpyxl import Workbook
from random import *
wb = Workbook()
ws = wb.active

# 1 줄씩 데이터 넣기
ws.append(["번호", "영어", "수학"])  # A, B, C컬럼
# 데이터
for i in range(1, 11):
    ws.append([i, randint(0, 100), randint(0, 100)])

# 영어점수만 가져오기
# col_B = ws["B"]  # 영어컬럼나 가지고 오기
# # print(col_B) # <Cell 'Sheet'.B1>, <Cell 'Sheet'.B2> 이런식으로 출력
# for cell in col_B:
#     print(cell.value)

# 함께 가져올떄

# col_range = ws["B:C"]  # 컬럼함께 가지고오기

# for cols in col_range:
#     for cols2 in cols:
#         print(cols2.value)

# # Q. 2개를 갖고 왔슴에도 한줄만 출력하고 싶다면?

row_title = ws[1]  # 1번째 로우만 갖고오기

# for cell in row_title:
#     print(cell.value)

row_range = ws[2:6]  # 2번째 줄에서 6번째 줄까지 갖고오기
# for rows in row_range:  # 2-6낒; 한줄씩 갖고와서
#     for cell in rows:  # 그 줄에 있는걸 하나의 셀씩 끊어서 값출력
#         print(cell.value, end=" ")
#  #Q. end를 넣는이유는?  end를 넣어 준 이유는 해당 결괏값을 출력할 때 다음줄로 넘기지 않고 그 줄에 계속해서 출력하기 위해서이다.
#     print()

# 줄이많아서

row_ran = ws[2:ws.max_row]  # 2번째부터 마지막 줄 까지
for row in row_ran:
    for cell in row:
        # print(cell.value, end=" ")
        # print(cell.coordinate, end =" ") #CElL의 좌표정보 갖고옴
        xy = coordinate_from_string(cell.coordinate)  # ('A', 2) 튜플을 좌표상태처럼 끊어줌
        # print(xy, end=" ")
       # print(xy[0], end=" ")  # A B C A B C 이렇게 출력
        # print(xy[1], end=" ")  # 222 333 444 555 6 7 8 9 .. 이렇게 출력
        print(xy[0:2], end=" ")  # ('A',) ('B',) ('C',)

    print()

wb.save("sample.xlsx")
