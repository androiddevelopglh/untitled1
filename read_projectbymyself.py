import poplib
import email
import os
import ssl
import zip
import shutil
import fun_read_word
from xlutils.copy import copy
import xlrd
import time
from datetime import datetime
import sys

from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr

date = datetime.strftime(datetime.now(),'%Y-%m-%d %H:%M:%S')
path=r'E:\教务处工作\19年校级质量工程\邮件附件\临时文件'
def renamebyself(path,date):#检查所下载附件是否所需附件，如果不是则直接删除临时文件夹；否者将临时文件夹中的文件重命名，并且将临时文件夹修改后的名字返回in，
	for fpathe, dirs, files in os.walk(path):
		for file in files:
			print(os.path.join(fpathe, file))
			if '.doc'in file:
				fun_read_word.read_word(r'E:\教务处工作\19年校级质量工程',os.path.join(fpathe, file),date)
renamebyself(path,date)