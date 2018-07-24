# from openpyxl import workbook
# from openpyxl import load_workbook
# from openpyxl.writer.excel import ExcelWriter

import xlrd
from xlutils.copy import *
import xlwt
# 1.打开excel文件，读取某一个位置的内容
book = xlrd.open_workbook("./123.xls")
sheet = book.sheet_by_index(0)
# 2.读取道路，案件小类，计量 3列的数据放在3个列表当中
DaoLu = sheet.col_values(3)
AnJianXiaoLei = sheet.col_values(7)
JiLiangShu = sheet.col_values(12)
# 3.读入欲写入表格的内容准备进行比较
book2 = xlrd.open_workbook("./biaoge.xls")
read_sheet = book2.sheet_by_index(0)
# 3.调用xlutils复制要写入的表格的样式
bookwrite = copy(book2)
write_sheet = bookwrite.get_sheet(0)
write_sheet.write(3, 0, 123)
# bookwrite.save("./456.xls")
# mySheet.write(0,0,2) #成功在0，0这个位置写入1
# row = 0
# lie = 0

# myWorkBook.save("./456.xls")
# print(sheet.)
print(DaoLu)
print(AnJianXiaoLei)
print(JiLiangShu)
# print(int(JiLiangShu[1]))
# xiangmuhao = sheet.col_values(1)

# genmulu = 'e:/123'
# os.chdir(genmulu)
# file = open('wang.txt','w')
# file.write('1234567890')
# file.close()

# file2 = open('wang.xlsx','w')
# file2.write('123')
# file2.close()
# for jishu in range(len(xiangmuhao)):
#     ming = xiangmuhao[jishu] + ' ' + xiangmumingcheng[jishu]
#     try:
#         os.makedirs(ming)
#         os.chdir(ming)
#         os.makedirs('需求')
#         os.makedirs('图纸')
#         os.makedirs('程序')
#         os.makedirs('文档')
#         os.chdir(genmulu)
#     except:
#         pass

# path = 'M-400066 湛江港6000HP拖轮'
# try:
#     os.makedirs(path)
# except:
#     pass
#
# os.chdir(genmulu+'/'+path)
# os.makedirs('bo')
