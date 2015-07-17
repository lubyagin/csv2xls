# -* coding: cp1251 -*-
# python-csv2xls; License: MIT; Author: Alexander Lubyagin

import sys

if len(sys.argv) < 2: sys.exit()
filename = sys.argv[1]

import csv
import xlwt
font0 = xlwt.Font()
font0.name = "DejaVu Sans"
font0.height = 12*20 # 12 points
style0 = xlwt.XFStyle()
style0.font = font0

wb = xlwt.Workbook()
ws = wb.add_sheet("RESULT")

n = 0
with open(filename,"r") as csvfile:
  w = csv.reader(csvfile, delimiter=';', quotechar='"')
  for row in w:
    k = 0
    for item in row:
      ws.write(n,k,unicode(item,"cp1251"), style0)
      k += 1
    n += 1

wb.save(filename.replace(".csv",".xls"))

sys.exit()
