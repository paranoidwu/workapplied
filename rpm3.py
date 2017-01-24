# -*- coding:utf-8 -*-
import xlrd,datetime,pyExcelerator
import OTA
#定义开始和结束文件名称，每次只需要修改这个
startdate = "20170116"
enddate = "20170122"
####################################################
ts = OTA.timeseries(startdate,enddate)
ts7back = OTA.timeseries7back(startdate,enddate)
print ts
print ts7back
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

#计算PV3,统计当周的数据情况
PV3 = 0
tongcheng_PV = 0
xiecheng_PV = 0
yilong_PV = 0
tongcheng_Click = 0
xiecheng_Click = 0
yilong_Click = 0
tongcheng_Cost = 0
xiecheng_Cost = 0
yilong_Cost = 0
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
		xiecheng_PV += int(line[6])
		xiecheng_Click += int(line[7])
		xiecheng_Cost += int(line[8])
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
		yilong_PV += int(line[6])
		yilong_Click += int(line[7])
		yilong_Cost += int(line[8])
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
		tongcheng_PV += int(line[6])
		tongcheng_Click += int(line[7])
		tongcheng_Cost += int(line[8])
		mark += 1
Costdelta = PV3*RPMdelta/1000/len(ts)

#统计上周的数据情况
tongcheng_PV7back = 0
xiecheng_PV7back = 0
yilong_PV7back = 0
tongcheng_Click7back = 0
xiecheng_Click7back = 0
yilong_Click7back = 0
tongcheng_Cost7back = 0
xiecheng_Cost7back = 0
yilong_Cost7back = 0
for d in ts7back:
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
		xiecheng_PV7back += int(line[6])
		xiecheng_Click7back += int(line[7])
		xiecheng_Cost7back += int(line[8])
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
		yilong_PV7back += int(line[6])
		yilong_Click7back += int(line[7])
		yilong_Cost7back += int(line[8])
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
		tongcheng_PV7back += int(line[6])
		tongcheng_Click7back += int(line[7])
		tongcheng_Cost7back += int(line[8])
		mark += 1
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
#print "customer\tPVlastweek\tPVthisweek\tIncreatment\tClicklastweek\tClickthisweek\tIncreatment\tCostlastweek\tCostthisweek\tIncreatment"
shichang = float(len(ts))
print "tongcheng\t%d\t%d\t%.2f\t%d\t%d\t%.2f\t%d\t%d\t%.2f"%(tongcheng_PV7back/shichang,tongcheng_PV/shichang,float(tongcheng_PV)/tongcheng_PV7back,tongcheng_Click7back/shichang,tongcheng_Click/shichang,float(tongcheng_Click)/tongcheng_Click7back,tongcheng_Cost7back/shichang,tongcheng_Cost/shichang,float(tongcheng_Cost)/tongcheng_Cost7back)
print "yilong\t%d\t%d\t%.2f\t%d\t%d\t%.2f\t%d\t%d\t%.2f"%(yilong_PV7back/shichang,yilong_PV/shichang,float(yilong_PV)/yilong_PV7back,yilong_Click7back/shichang,yilong_Click/shichang,float(yilong_Click)/yilong_Click7back,yilong_Cost7back/shichang,yilong_Cost/shichang,float(yilong_Cost)/yilong_Cost7back)
print "xiecheng\t%d\t%d\t%.2f\t%d\t%d\t%.2f\t%d\t%d\t%.2f"%(xiecheng_PV7back/shichang,xiecheng_PV/shichang,float(xiecheng_PV)/xiecheng_PV7back,xiecheng_Click7back/shichang,xiecheng_Click/shichang,float(xiecheng_Click)/xiecheng_Click7back,xiecheng_Cost7back/shichang,xiecheng_Cost/shichang,float(xiecheng_Cost)/xiecheng_Cost7back)
huizong_PV = tongcheng_PV+yilong_PV+xiecheng_PV
huizong_Click = tongcheng_Click+yilong_Click+xiecheng_Click
huizong_Cost = tongcheng_Cost+yilong_Cost+xiecheng_Cost
huizong_PV7back = tongcheng_PV7back+yilong_PV7back+xiecheng_PV7back
huizong_Click7back = tongcheng_Click7back+yilong_Click7back+xiecheng_Click7back
huizong_Cost7back = tongcheng_Cost7back+yilong_Cost7back+xiecheng_Cost7back
print "huizong\t%d\t%d\t%.2f\t%d\t%d\t%.2f\t%d\t%d\t%.2f"%(huizong_PV7back/shichang,huizong_PV/shichang,float(huizong_PV)/huizong_PV7back,huizong_Click7back/shichang,huizong_Click/shichang,float(huizong_Click)/huizong_Click7back,huizong_Cost7back/shichang,huizong_Cost/shichang,float(huizong_Cost)/huizong_Cost7back)