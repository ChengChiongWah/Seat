# -*- coding: utf-8 -*-
import xlrd
import xlwt
import random

data = xlrd.open_workbook("students.xlsx")
table = data.sheets()[0]
nrows = table.nrows
students = []
total_col_numbers = 10
for i in range(1,nrows):
    print(i, table.row_values(i)[1])
    students.append("  ".join([str(int(table.row_values(i)[0])), table.row_values(i)[1]]))    # 把所有的姓名放进列表

random.shuffle(students)   # 打乱列表中的数据
#
# for i in range(len(students)):
#     print(students[i])


wbk = xlwt.Workbook()
sheet = wbk.add_sheet("sheet1")

insert_row_number = 1
col_nuber = 0
sheet.write(0, 10, "讲台")
for s in range(len(students)):
    sheet.write(insert_row_number, col_nuber, students[s])
    if col_nuber + 2 >total_col_numbers*2-1:   #19是班级是10排，中间隔一个excel列故是：2*10-1 = 19.
        col_nuber = 0
        insert_row_number = insert_row_number + 1
    else:
        col_nuber = col_nuber + 2

wbk.save("seat.xls")
