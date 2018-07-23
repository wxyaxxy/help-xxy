import os
import xlrd

excelFile = xlrd.open_workbook(r'e:/123/123.xlsx')
sheet = excelFile.sheet_by_index(0)
xiangmumingcheng = sheet.col_values(2)
xiangmuhao = sheet.col_values(1)

genmulu = 'e:/123'
os.chdir(genmulu)
file = open('wang.txt','w')
file.write('1234567890')
file.close()

file2 = open('wang.xlsx','w')
# file2.write('123')
file2.close()
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