#coding=utf-8
#/usr/bin/env python
import xlsxwriter,xlrd
import sys,os.path

from xlrd.timemachine import xrange

fname = 'test.xlsx'
if not os.path.isfile(fname):
    print ('文件路径不存在')
    sys.exit()
data = xlrd.open_workbook(fname)            # 打开fname文件
data.sheet_names()                          # 获取xls文件中所有sheet的名称
table = data.sheet_by_index(0)              # 通过索引获取xls文件第0个sheet
nrows = table.nrows                         # 获取table工作表总行数
ncols = table.ncols                         # 获取table工作表总列数
workbook = xlsxwriter.Workbook('zm6.xls')  #创建一个excel文件
# worksheet = workbook.add_worksheet()        #创建一个工作表对象


print(data.sheet_names())

for a, val in enumerate(data.sheet_names()):
    worksheet = workbook.add_worksheet(val)
    for i in xrange( data.sheet_by_name(val).nrows):
        for j in xrange(data.sheet_by_name(val).ncols):
            cell_value = table.cell_value(i, j, )  # 获取第i行中第j列的值
            worksheet.write(i, j, cell_value)  # 把获取到的值写入文件对应的行列
workbook.close()