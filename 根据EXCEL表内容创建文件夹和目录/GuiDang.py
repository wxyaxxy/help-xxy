import os
import xlrd

#1.打开excel文件，读取某一个位置的内容
excelFile = xlrd.open_workbook(r'./123.xlsx')
sheet = excelFile.sheet_by_index(0)
#2.读取道路，案件小类，计量 3列的数据放在3个列表当中
DaoLu = sheet.col_values(4)
AnJianXiaoLei = sheet.col_values(7)
JiLiangShu = sheet.col_values(12)
print(sheet.cell(0,0))
print(sheet.)
# print(DaoLu)
# print(AnJianXiaoLei)
# print(JiLiangShu)
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

#path = 'M-400066 湛江港6000HP拖轮'
# try:
#     os.makedirs(path)
# except:
#     pass
#
# os.chdir(genmulu+'/'+path)
# os.makedirs('bo')