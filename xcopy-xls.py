# -*- coding: windows-1251 -*-

# pip install xlrd
# pip install xlwt
# pip install xlutils

import xlrd,xlwt
from xlutils.copy import copy
def func(s):
  rb = xlrd.open_workbook("blank.xls",formatting_info=True)
  wb = copy(rb)
  # help(wb)
  font0 = xlwt.Font()
  font0.name = "Times New Roman"
  font0.height = 16*20 # 16 points
  font0.bold = True
  style0 = xlwt.XFStyle()
  style0.font = font0

  ws = wb.get_sheet(0)
  ws.write(0,0,s,style0) # А1
  wb.save("out\\"+s.strip()[:120]+".xls")

filename = "list.txt" # Создать копии файла по списку с подстановкой ячейки А1
f = open(filename)
lines = f.readlines()
for line in lines:
  a = line.split("\t")
  print a
  func(unicode(a,"cp1251"))
f.close()
