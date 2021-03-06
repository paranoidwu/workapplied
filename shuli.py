# -*- coding:utf-8 -*-
import pyExcelerator,OTA
#定义开始和结束文件名称
startdate = "20160805"
enddate = "20160806"
#写入文件为双重字典，一级key为yyyymmdd字符串，二级key为从0开始的行数
data = {}
for d in OTA.timeseries(startdate,enddate):
	fin = open(d,'r')
	mark = 0
	data[d] = {}
#读取文件并修正问题
	for line in fin:
		line = line.strip()
		if OTA.check_line_item(line) == 10:
			data[d][mark] = line.split('\t')
		elif OTA.check_line_item(line) == 9:
			data[d][mark] = OTA.repair_line_9(line)
		else:
			print "%s,%d line has problem,please check"%(d,mark)
		mark += 1
#输出文件
workbook = pyExcelerator.Workbook()
sheet1 = workbook.add_sheet('sheet1')
myfont = pyExcelerator.Font()
#myfont.name = u'Times New Roman'
myfont.bold = True
mystyle = pyExcelerator.XFStyle()
mystyle.font = myfont
sheet1.write(0,0,'日期'.decode('gbk', "ignore"),mystyle)
sheet1.write(0,1,"账户ID".decode('gbk', "ignore"),mystyle)
sheet1.write(0,2,"计划名称".decode('gbk', "ignore"),mystyle)
sheet1.write(0,3,"组名称".decode('gbk', "ignore"),mystyle)
sheet1.write(0,4,"关键词".decode('gbk', "ignore"),mystyle)
sheet1.write(0,5,"查询词".decode('gbk', "ignore"),mystyle)
sheet1.write(0,6,"展现".decode('gbk', "ignore"),mystyle)
sheet1.write(0,7,"点击".decode('gbk', "ignore"),mystyle)
sheet1.write(0,8,"消耗".decode('gbk', "ignore"),mystyle)
sheet1.write(0,9,"高端样式".decode('gbk', "ignore"),mystyle)
write_line = 1
for d in OTA.timeseries(startdate,enddate):
	lenth = len(data[d])
	for j in range(lenth):
		if data[d][j][9] == "null":
			print "%s,%d is null"%(d,j)
			continue
		elif data[d][j][8]==0 and data[d][j][7]> 0:
			print "%s,%d is wrong between click and cost"%(d,j)
			continue
		else:
			for k in range(10):
				if k ==6 or k == 7 or k == 8 or k == 1:
					sheet1.write(write_line,k,int(data[d][j][k]))
				else:
					sheet1.write(write_line,k,data[d][j][k].decode('gbk', "ignore"))
			write_line += 1
savename = startdate+"-"+enddate+"高端定制样式数据报表"+".xls"
workbook.save(savename)