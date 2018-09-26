# -*- coding: utf-8 -*-

import xlrd
import xlwt
import os,sys

reload(sys)

sys.setdefaultencoding('utf8')

from datetime import date,datetime

path = "2017-5-24.xls"

def get_excel():
	workbook = xlrd.open_workbook(path)
	print workbook.sheet_names()
	rb = xlwt.Workbook()

	sheet = workbook.sheet_by_name("Sheet1")
	sheet2 = rb.add_sheet("sheet1",cell_overwrite_ok=True)
	print sheet.name,sheet.nrows,sheet.ncols
	lens = sheet.nrows
	count = 0
	for i in range(3,lens):
		rows = sheet.row_values(i)
		if "市" in rows[1] or "区" in rows[1]:
			print rows
			count+=1
			for j  in  range(0,sheet.ncols):
				sheet2.write(count,j,rows[j])
	rb.save("write.xls")

if __name__== "__main__":
	get_excel()
