# -*- coding:utf-8 -*-
import pyExcelerator,OTA

fin = open("galaxy_ota_stat_upperleft_first_group_plan_xiecheng.xiecheng",'r')
mark = 0
data = {}
#读取文件并修正问题
for line in fin:
	line = line.strip()
	if OTA.check_line_item(line) == 10:
		data[mark] = line.split('\t')
	elif OTA.check_line_item(line) == 9:
		data[mark] = OTA.repair_line_9(line)
	else:
		print "%d line has problem,please check"%mark
	mark += 1

workbook = pyExcelerator.Workbook()
sheet1 = workbook.add_sheet('sheet1')
for j in range(mark):
	for k in range(10):
		if k ==6 or k == 7 or k == 8:
			sheet1.write(j,k,int(data[j][k]))
		else:
			sheet1.write(j,k,data[j][k].decode('gbk', "ignore"))

workbook.save("testfor1.xls")