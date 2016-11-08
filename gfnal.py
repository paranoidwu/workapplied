# -*- coding:utf-8 -*-
import xlrd,datetime,pyExcelerator
import OTA
#定义开始和结束文件名称
startdate = "20161031"
enddate = "20161106"
ts = OTA.timeseries(startdate,enddate)
for d in ts:
	filename_with_date = "C:\\Users\\wuenyu\\save\\OTA\\"+d+"\\"+d+"\\galaxy_ota_stat_upperleft_first_indus_style_"+d
	fin = open(filename_with_date,'r')
	dic = {}
	for line in fin:
		line = line.strip().split()
		if line[0] == "机票-通用词":
			line[0] = "机票通用词"
			print "fixed"
		if float(line[12]) == 0:
			CTchange == 0
		else:
			CTchange = float(line[5])/float(line[12])-1
		dic[line[0]] = [int(line[1]),int(line[2]),float(line[3]),float(line[5]),int(line[9]),int(line[10]),float(line[11]),float(line[12]),CTchange]

	hotelEPV = 0
	hotelECLICK = 0
	hotelECost = 0
	hotelCPV = 0
	hotelCCLICK = 0
	hotelCCost = 0

	flightEPV = 0
	flightECLICK = 0
	flightECost = 0
	flightCPV = 0
	flightCCLICK = 0
	flightCCost = 0

	travelEPV = 0
	travelECLICK = 0
	travelECost = 0
	travelCPV = 0
	travelCCLICK = 0
	travelCCost = 0

	for key in dic:
		if key in ["酒店-单体品牌词","酒店-地域酒店","酒店-通用词","酒店-连锁品牌词"]:
			hotelEPV += dic[key][0]
			hotelECLICK += dic[key][1]
			hotelECost += dic[key][2]
			hotelCPV += dic[key][4]
			hotelCCLICK += dic[key][5]
			hotelCCost += dic[key][6]
		elif key in ["旅游-通用词","旅游-单线路"]:
			travelEPV += dic[key][0]
			travelECLICK += dic[key][1]
			travelECost += dic[key][2]
			travelCPV += dic[key][4]
			travelCCLICK += dic[key][5]
			travelCCost += dic[key][6]
		else:
			flightEPV += dic[key][0]
			flightECLICK += dic[key][1]
			flightECost += dic[key][2]
			flightCPV += dic[key][4]
			flightCCLICK += dic[key][5]
			flightCCost += dic[key][6]
	
	hotelECTR = float(hotelECLICK)/hotelEPV
	hotelCCTR = float(hotelCCLICK)/hotelCPV
	hotelCTR_change = hotelECTR/hotelCCTR - 1
	dic["酒店"]=[hotelEPV,hotelECLICK,hotelECost,hotelECTR,hotelCPV,hotelCCLICK,hotelCCost,hotelCCTR,hotelCTR_change]
	
	flightECTR = float(flightECLICK)/flightEPV
	flightCCTR = float(flightCCLICK)/flightCPV
	flightCTR_change = flightECTR/flightCCTR - 1
	dic["机票"]=[flightEPV,flightECLICK,flightECost,flightECTR,flightCPV,flightCCLICK,flightCCost,flightCCTR,flightCTR_change]
	
	travelECTR = float(travelECLICK)/travelEPV
	travelCCTR = float(travelCCLICK)/travelCPV
	travelCTR_change = travelECTR/travelCCTR - 1
	dic["旅游"]=[travelEPV,travelECLICK,travelECost,travelECTR,travelCPV,travelCCLICK,travelCCost,travelCCTR,travelCTR_change]

	savename = d+".xls"
	OTA.write_on_3(dic,savename)
	print "finish",d
