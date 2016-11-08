# -*- coding:utf-8 -*-
import xlrd,datetime,pyExcelerator
import OTA
#定义开始和结束文件名称，每次只需要修改这个
startdate = "20161031"
enddate = "20161106"
####################################################
ts = OTA.timeseries(startdate,enddate)
dic = {}
EPV = 0
EClick = 0
ECost = 0
CPV = 0
CClick = 0
CCost = 0
#七天数据都写入字典，一级为日期，二级为样式，三级记录实验组对照组PV点击消耗
for d in ts:
	dic[d] = {}
	filename_with_date = "C:\\Users\\wuenyu\\save\\OTA\\"+d+"\\"+d+"\\galaxy_ota_stat_upperleft_indus_style_"+d
	fin = open(filename_with_date,'r')
	for line in fin:
		line = line.strip().split()
		dic[d][line[0]] = [int(line[1]),int(line[2]),float(line[3]),int(line[9]),int(line[10]),float(line[11])]
		EPV += int(line[1])
		EClick += int(line[2])
		ECost += float(line[3])
		CPV += int(line[9])
		CClick += int(line[10])
		CCost += float(line[11])
#计算CTR提升，RPM提升
ECTR = float(EClick)/EPV
CCTR = float(CClick)/CPV
ERPM = ECost/EPV*1000
CRPM = CCost/CPV*1000
CTRchange = ECTR/CCTR - 1
RPMchange = ERPM/CRPM - 1
RPMdelta = ERPM-CRPM

#计算PV3
PV3 = 0
for d in ts:
	finxiecheng= open("C:\\Users\\wuenyu\\save\\OTA\\"+d+"\\"+d+"\\galaxy_ota_stat_upperleft_first_group_plan_xiecheng",'r')
	finyilong= open("C:\\Users\\wuenyu\\save\\OTA\\"+d+"\\"+d+"\\galaxy_ota_stat_upperleft_first_group_plan_yilong",'r')
	fintongcheng= open("C:\\Users\\wuenyu\\save\\OTA\\"+d+"\\"+d+"\\galaxy_ota_stat_upperleft_first_group_plan_tongcheng",'r')
#读取三家数据文件并求和
	for line in finxiecheng:
		mark = 0
		line = line.strip()
		if OTA.check_line_item(line) == 10:
			line = line.split('\t')
		elif OTA.check_line_item(line) == 9:
			line = OTA.repair_line_9(line)
		else:
			print "there is problem in line %r,xiecheng,%r"%(mark,d)
		PV3 += int(line[6])
		mark += 1
	for line in finyilong:
		mark = 0
		line = line.strip()
		if OTA.check_line_item(line) == 10:
			line = line.split('\t')
		elif OTA.check_line_item(line) == 9:
			line = OTA.repair_line_9(line)
		else:
			print "there is problem in line %r,yilong,%r"%(mark,d)
		PV3 += int(line[6])
		mark += 1
	for line in fintongcheng:
		mark = 0
		line = line.strip()
		if OTA.check_line_item(line) == 10:
			line = line.split('\t')
		elif OTA.check_line_item(line) == 9:
			line = OTA.repair_line_9(line)
		else:
			print "there is problem in line %r,tongcheng,%r"%(mark,d)
		PV3 += int(line[6])
		mark += 1

Costdelta = PV3*RPMdelta/1000/len(ts)
#########################################计算完所有的数据，准备写入#######################################
workbook = pyExcelerator.Workbook()
sheet1 = workbook.add_sheet('origin')
sheet2 = workbook.add_sheet('SumAll')
row = 0
sheet1.write(row,1,'Style')
sheet1.write(row,2,'ExpPV')
sheet1.write(row,3,'ExpClick')
sheet1.write(row,4,'ExpCost')
sheet1.write(row,5,'ComPV')
sheet1.write(row,6,'ComClick')
sheet1.write(row,7,'ComCost')
row += 1
for d in ts:
	sheet1.write(row,0,d)
	for key in dic[d]:
		sheet1.write(row,1,key.decode('gbk', "ignore"))
		i = 2
		for keyval in dic[d][key]:
			sheet1.write(row,i,keyval)
			i += 1
		row += 1
##############第一张表完成#############
row = 0
sheet2.write(row,0,'ExpPV')
sheet2.write(row,1,'ExpClick')
sheet2.write(row,2,'ExpCost')
sheet2.write(row,3,'ExpCTR')
sheet2.write(row,4,'ExpRPM')
sheet2.write(row,5,'ComPV')
sheet2.write(row,6,'ComClick')
sheet2.write(row,7,'ComCost')
sheet2.write(row,8,'ComCTR')
sheet2.write(row,9,'ComRPM')
sheet2.write(row,10,'CTRchange')
sheet2.write(row,11,'RPMchange')
sheet2.write(row,12,'PV3')
sheet2.write(row,13,'CostIncreased')
row = 1
sheet2.write(row,0,EPV)
sheet2.write(row,1,EClick)
sheet2.write(row,2,ECost)
sheet2.write(row,3,ECTR)
sheet2.write(row,4,ERPM)
sheet2.write(row,5,CPV)
sheet2.write(row,6,CClick)
sheet2.write(row,7,CCost)
sheet2.write(row,8,CCTR)
sheet2.write(row,9,CRPM)
sheet2.write(row,10,CTRchange)
sheet2.write(row,11,RPMchange)
sheet2.write(row,12,PV3)
sheet2.write(row,13,Costdelta)
savename = startdate+"-"+enddate+"RPM计算"+".xls"
workbook.save(savename)