from xlrd import open_workbook
import xlwt
import xlsxwriter
from xlutils.copy import copy
import xlrd
import win32com,re
import os,sys,re
#import unicode




def format_str(content):
    content = re.sub("[A-Za-z0-9\!\%\[\]\,\。]", "", content)
    content=content.replace(" ","")
    return content

#读取源文件
filepath_source=r'E:\教务处工作\19年校级质量工程\质量工程学生核对表.xls';
sheet_source_name=u'Sheet1'
#要核对的文件数据
filepath_check=r'E:\教务处工作\19年校级质量工程\result.xls'
sheet_check_name=u'延期材料中提取信息'
#源文件中的索引列
col_source=3
#源文件值填充列
val_source=2

#目标文件索引列
col_check=3
#目标文件值
val_check=2
book_source = xlrd.open_workbook(filepath_source)
write_book_source = copy(book_source)
write_sheet_source = write_book_source.get_sheet(sheet_source_name)

# 通过sheet_by_index()获取的sheet没有write()方法
sheet_source = book_source.sheet_by_name(sheet_source_name)
nrows_source = sheet_source.nrows  # 行数

book_check=xlrd.open_workbook(filepath_check)
write_book_check = copy(book_check)
write_sheet_check = write_book_check.get_sheet(sheet_check_name)
sheet_check=book_check.sheet_by_name(sheet_check_name)
nrows_check = sheet_check.nrows  # 行数


for index_source in range(2, nrows_source):
    value='';
    for index_check in range(0, nrows_check):
        str_source=format_str(sheet_source.cell(index_source,col_source).value)
        str_check=format_str(sheet_check.cell(index_check,col_check).value)
        if sheet_source.cell(index_source,10).value=='否'and str_source==str_check:
            value=value+'*'+sheet_check.cell(index_check, val_check).value
            value.lstrip('*')
            write_sheet_source.write(index_source, val_source,value)
            #sheet_source.cell(index_source, val_source)=sheet_check.cell(index_check, val_check)
write_book_source.save(filepath_source)
write_book_check.save(filepath_check)
m=1;








