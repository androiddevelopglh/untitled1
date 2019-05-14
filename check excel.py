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
str = '8590gd 中ddas 国akfag'
format_str(str)
#读取源文件
filepath_source=r'E:\教务处工作\19年校级质量工程\result.xls';
sheet_source_name=u'文件读取记录表'
#要核对的文件数据
filepath_check=r'E:\教务处工作\19年校级质量工程\result.xls'
sheet_check_name=u'文件读取记录表'
#源文件中的索引列
col_source=5
#目标文件索引列
col_check=8

book_source = xlrd.open_workbook(filepath_source)
# 通过sheet_by_index()获取的sheet没有write()方法
sheet_source = book_source.sheet_by_name(sheet_source_name)
nrows_source = sheet_source.nrows  # 行数
write_book_source = copy(book_source)

book_check=xlrd.open_workbook(filepath_check)
sheet_check=book_check.sheet_by_name(sheet_check_name)
nrows_check = sheet_check.nrows  # 行数


for index_source in range(1, nrows_source):
    for index_check in range(1, nrows_check):
        str_source=format_str(sheet_source.cell(index_source,col_source))
        str_check=format_str(sheet_check.cell(index_check,col_check))
        if sheet_source.cell(index_source,9)=='否'and str_source==str_check:
            m=1;










