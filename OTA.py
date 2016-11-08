# -*- coding:utf-8 -*-
import datetime,xlrd,pyExcelerator
def convert2dic(f):
	"""load the file and convert it into dic,key is query word,and the values are pv,style,click,cost"""
	#定义字典，依次存储值为PV，样式，点击，消耗
	dic = {}
	fileError = 0
	for line in f:
		line = line.strip()
		#定义数据项为7、6及其他时的字典操作
		if check_line_item(line) == 7:
			PV_ID,Account_ID,KeyWord,QueryWord,Style_ID,Click,Cost = line.split('\t')
		elif check_line_item(line) == 6:
			PV_ID,Account_ID,KeyWord,QueryWord,Style_ID,Click,Cost = repair_line_6(line)
		else:
			fileError += 1

		if dic.has_key(QueryWord):
			dic[QueryWord][0] += 1
			if dic[QueryWord][1] != Style_ID:
				print QueryWord,'doesn''t have the same styid'
			dic[QueryWord][2] += int(Click)
			dic[QueryWord][3] += int(Cost)
		else:
			dic[QueryWord] = []
			dic[QueryWord].append(1)
			dic[QueryWord].append(Style_ID)
			dic[QueryWord].append(int(Click))
			dic[QueryWord].append(int(Cost))
	print "origin file has %d mistakes with 7 item enough"%fileError
	return dic

def check_line_item(line):
	"""judge the item whether if is 7"""
	return len(line.split('\t'))

def repair_line_6(line):
	PV_ID,Account_ID,X,Y,Click,Cost = line.split('\t')
	if check_number(Y):
		Style_ID = int(Y)
		KeyWord,QueryWord = split_question(X)
		return PV_ID,Account_ID,KeyWord,QueryWord,Style_ID,Click,Cost
	else:
		KeyWord = X
		QueryWord,Style_ID = split_question(Y)
		return PV_ID,Account_ID,KeyWord,QueryWord,Style_ID,Click,Cost

def repair_line_9(line):
	a1,b2,c3,d4,e5,f6,h8,i9,j10 = line.split('\t')
	f6,g7 = split_question(f6)
	return [a1,b2,c3,d4,e5,f6,int(g7),int(h8),int(i9),j10]

		
def check_number(s):
	try:
		return int(s)
	except ValueError:
		return None

def split_question(s):
	'''split the string with the last question mark into 2 parts'''
	mark = 0
	qposition =0
	for ch in s:
		if ch == '?':
			qposition = mark
		mark += 1
	former = s[:qposition]
	latter = s[qposition+1:]
	return former,latter

#求两个字典的交集，实现vlookup的功能
def dic_intersection(a,b):
	c = {}
	for key in a:
		if b.has_key(key):
			#存储数据位实验组PV0,样式1，点击2，消耗3，对照组PV4，样式5，点击6，消耗7
			c[key] = (a[key][0],a[key][1],a[key][2],a[key][3],b[key][0],b[key][1],b[key][2],b[key][3])
	return c

#基于sytle_ID输出总结数据
def calculate_result(dic):
	result = {}
	for key in dic:
		if dic[key][1] == dic[key][5]:
			if result.has_key(dic[key][1]):
				result[dic[key][1]][0] += dic[key][0]
				result[dic[key][1]][1] += dic[key][2]
				result[dic[key][1]][2] += dic[key][3]
				result[dic[key][1]][3] += dic[key][4]
				result[dic[key][1]][4] += dic[key][6]
				result[dic[key][1]][5] += dic[key][7]
			else:
				result[dic[key][1]] = [dic[key][0],dic[key][2],dic[key][3]]
				result[dic[key][1]].append(dic[key][4])
				result[dic[key][1]].append(dic[key][6])
				result[dic[key][1]].append(dic[key][7])
		else:
			print "Wrong,same query word's style doesn't match"
			#打开以下语句，寻找对应问题query词
			#print dic[key],key
	#上侧的结果是样式：实验组PV，点击，消耗，对照组PV，点击，消耗
	
	#总数据
	ExpPV = 0
	ExpClick = 0
	ExpCost = 0
	ComPV = 0
	ComClick = 0
	ComCost = 0
	
	#酒店数据
	Hotel_ExpPV = 0
	Hotel_ExpClick = 0
	Hotel_ExpCost = 0
	Hotel_ComPV = 0
	Hotel_ComClick = 0
	Hotel_ComCost = 0
	
	#旅游数据
	Travel_ExpPV = 0
	Travel_ExpClick = 0
	Travel_ExpCost = 0
	Travel_ComPV = 0
	Travel_ComClick = 0
	Travel_ComCost = 0
	
	#机票数据
	Ticket_ExpPV = 0
	Ticket_ExpClick = 0
	Ticket_ExpCost = 0
	Ticket_ComPV = 0
	Ticket_ComClick = 0
	Ticket_ComCost = 0
	#基于样式汇总
	for key in result:
		ExpPV += result[key][0]
		ExpClick += result[key][1]
		ExpCost += result[key][2]
		ComPV += result[key][3]
		ComClick += result[key][4]
		ComCost += result[key][5]
		
		result[key].append(float(result[key][1])/result[key][0])
		result[key].append(float(result[key][4])/result[key][3])
		################################################################修改
		if result[key][7] == 0:
			result[key].append(0.01)
		else:
			result[key].append(result[key][6]/result[key][7]-1)
		
		#汇总酒店
		if key == "11" or key == "12"or key == "13"or key == "14":
			Hotel_ExpPV += result[key][0]
			Hotel_ExpClick += result[key][1]
			Hotel_ExpCost += result[key][2]
			Hotel_ComPV += result[key][3]
			Hotel_ComClick += result[key][4]
			Hotel_ComCost += result[key][5]
		#汇总旅游
		if key == "21" or key == "22":
			Travel_ExpPV += result[key][0]
			Travel_ExpClick += result[key][1]
			Travel_ExpCost += result[key][2]
			Travel_ComPV += result[key][3]
			Travel_ComClick += result[key][4]
			Travel_ComCost += result[key][5]
		#汇总机票
		if key == "31" or key == "32":
			Ticket_ExpPV += result[key][0]
			Ticket_ExpClick += result[key][1]
			Ticket_ExpCost += result[key][2]
			Ticket_ComPV += result[key][3]
			Ticket_ComClick += result[key][4]
			Ticket_ComCost += result[key][5]

	ExpCTR =float(ExpClick) / ExpPV
	ComCTR =float(ComClick) / ComPV
	if ComCTR == 0:
		CTR_change = 0.01
	else:
		CTR_change =ExpCTR / ComCTR - 1
	result['99'] = [ExpPV,ExpClick,ExpCost,ComPV,ComClick,ComCost,ExpCTR,ComCTR,CTR_change]
	if float(Hotel_ExpClick) == 0 or float(Hotel_ComClick) == 0:
		Hotel_ExpCTR = 0
		Hotel_ComCTR = 0
		Hotel_CTR_change = 0.01
	else:
		Hotel_ExpCTR =float(Hotel_ExpClick) / Hotel_ExpPV
		Hotel_ComCTR =float(Hotel_ComClick) / Hotel_ComPV
		Hotel_CTR_change =Hotel_ExpCTR / Hotel_ComCTR - 1
	result['97'] = [Hotel_ExpPV,Hotel_ExpClick,Hotel_ExpCost,Hotel_ComPV,Hotel_ComClick,Hotel_ComCost,Hotel_ExpCTR,Hotel_ComCTR,Hotel_CTR_change]
	
	Travel_ExpCTR =float(Travel_ExpClick) / Travel_ExpPV
	Travel_ComCTR =float(Travel_ComClick) / Travel_ComPV
	if Travel_ComCTR == 0:
		Travel_CTR_change = 0.01
	else:
		Travel_CTR_change =Travel_ExpCTR / Travel_ComCTR - 1
	result['98'] = [Travel_ExpPV,Travel_ExpClick,Travel_ExpCost,Travel_ComPV,Travel_ComClick,Travel_ComCost,Travel_ExpCTR,Travel_ComCTR,Travel_CTR_change]
	
	if Ticket_ExpPV != 0 and Ticket_ComPV != 0:
		Ticket_ExpCTR =float(Ticket_ExpClick) / Ticket_ExpPV
		Ticket_ComCTR =float(Ticket_ComClick) / Ticket_ComPV
		if Ticket_ComCTR == 0:
			Ticket_CTR_change = 0.01
		else:
			Ticket_CTR_change =Ticket_ExpCTR / Ticket_ComCTR - 1
		result['96'] = [Ticket_ExpPV,Ticket_ExpClick,Ticket_ExpCost,Ticket_ComPV,Ticket_ComClick,Ticket_ComCost,Ticket_ExpCTR,Ticket_ComCTR,Ticket_CTR_change]
	
	return result

def write_origin(fA,fB,sheet,x,y):
	#写实验组表头
	row = x
	sheet.write(row,y,'PV_ID')
	sheet.write(row,y+1,'Account_ID')
	sheet.write(row,y+2,'KeyWord')
	sheet.write(row,y+3,'QueryWord')
	sheet.write(row,y+4,'Style_ID')
	sheet.write(row,y+5,'Click')
	sheet.write(row,y+6,'Cost')
	#写实验组数据
	for line in fA:
		row += 1
		line = line.strip()
		if check_line_item(line) == 7:
			PV_ID,Account_ID,KeyWord,QueryWord,Style_ID,Click,Cost = line.split('\t')
		elif check_line_item(line) == 6:
			PV_ID,Account_ID,KeyWord,QueryWord,Style_ID,Click,Cost = repair_line_6(line)
		sheet.write(row,y,PV_ID)
		sheet.write(row,y+1,Account_ID)
		sheet.write(row,y+2,KeyWord.decode('gbk', "ignore"))
		sheet.write(row,y+3,QueryWord.decode('gbk', "ignore"))
		sheet.write(row,y+4,Style_ID)
		sheet.write(row,y+5,int(Click))
		sheet.write(row,y+6,int(Cost))
	#写对照组表头
	row = x
	sheet.write(row,y+8,'PV_ID')
	sheet.write(row,y+9,'Account_ID')
	sheet.write(row,y+10,'KeyWord')
	sheet.write(row,y+11,'QueryWord')
	sheet.write(row,y+12,'Style_ID')
	sheet.write(row,y+13,'Click')
	sheet.write(row,y+14,'Cost')
	#写对照组数据
	for line in fB:
		row += 1
		line = line.strip()
		if check_line_item(line) == 7:
			PV_ID,Account_ID,KeyWord,QueryWord,Style_ID,Click,Cost = line.split('\t')
		elif check_line_item(line) == 6:
			PV_ID,Account_ID,KeyWord,QueryWord,Style_ID,Click,Cost = repair_line_6(line)
		sheet.write(row,y+8,PV_ID)
		sheet.write(row,y+9,Account_ID)
		sheet.write(row,y+10,KeyWord.decode('gbk', "ignore"))
		sheet.write(row,y+11,QueryWord.decode('gbk', "ignore"))
		sheet.write(row,y+12,Style_ID)
		sheet.write(row,y+13,int(Click))
		sheet.write(row,y+14,int(Cost))
	print "origin data is saved into EXCEL"

def write_QuerySum(a,b,sheet,x,y):
	row = x
	#创建表头
	sheet.write(row,y,'ExpQueryWord')
	sheet.write(row,y+1,'ExpPV')
	sheet.write(row,y+2,'ExpClick')
	sheet.write(row,y+3,'ExpCost')
	sheet.write(row,y+5,'ComQueryWord')
	sheet.write(row,y+6,'ComPV')
	sheet.write(row,y+7,'ComClick')
	sheet.write(row,y+8,'ComCost')
	
	#先写左侧实验组数据
	for key in a:
		row += 1
		sheet.write(row,y,key.decode('gbk', "ignore"))
		sheet.write(row,y+1,a[key][0])
		sheet.write(row,y+2,a[key][2])
		sheet.write(row,y+3,a[key][3])

		#再写右侧对照组数据
	row = x
	for key in b:
		row += 1
		sheet.write(row,y+5,key.decode('gbk', "ignore"))
		sheet.write(row,y+6,b[key][0])
		sheet.write(row,y+7,b[key][2])
		sheet.write(row,y+8,b[key][3])
	
	print "QuerySum data is saved into EXCEL"

def write_ValidData(a,sheet):
	row = 0
	sheet.write(row,0,'QueryWord')
	sheet.write(row,1,'ExpPV')
	sheet.write(row,2,'ExpClick')
	sheet.write(row,3,'ExpCost')
	sheet.write(row,4,'ComPV')
	sheet.write(row,5,'ComClick')
	sheet.write(row,6,'ComCost')
	sheet.write(row,7,'Style_ID')
	for key in a:
		row += 1
		sheet.write(row,0,key.decode('gbk', "ignore"))
		sheet.write(row,1,a[key][0])
		sheet.write(row,2,a[key][2])
		sheet.write(row,3,a[key][3])
		sheet.write(row,4,a[key][4])
		sheet.write(row,5,a[key][6])
		sheet.write(row,6,a[key][7])
		sheet.write(row,7,a[key][1])
	
	print "ValidData data is saved into EXCEL"

def write_result(a,sheet,b,shunxu):
	#写表头
	row = 0
	sheet.write(row,0,'Style')
	sheet.write(row,1,'ExpPV')
	sheet.write(row,2,'ExpClick')
	sheet.write(row,3,'ExpCost')
	sheet.write(row,4,'ExpCTR')
	sheet.write(row,5,'ComPV')
	sheet.write(row,6,'ComClick')
	sheet.write(row,7,'ComCost')
	sheet.write(row,8,'ComCTR')
	sheet.write(row,9,'CTR_change')
	
	for key in shunxu:
		if a.has_key(shunxu[key]):
			row += 1
			sheet.write(row,0,b[shunxu[key]])
			sheet.write(row,1,a[shunxu[key]][0])
			sheet.write(row,2,a[shunxu[key]][1])
			sheet.write(row,3,a[shunxu[key]][2])
			sheet.write(row,4,a[shunxu[key]][6])
			sheet.write(row,5,a[shunxu[key]][3])
			sheet.write(row,6,a[shunxu[key]][4])
			sheet.write(row,7,a[shunxu[key]][5])
			sheet.write(row,8,a[shunxu[key]][7])
			sheet.write(row,9,a[shunxu[key]][8])
		else:
			print "Dic doesn's has key %r"%shunxu[key]
	
	print "ALL is successfully done"

def timeseries(startdate,enddate):
	format = "%Y%m%d"
	sd = datetime.datetime.strptime(startdate,format)
	ed = datetime.datetime.strptime(enddate,format)
	delta = ed-sd
	ts = []
	for i in range(delta.days+1):
		ts.append(datetime.datetime.strftime(sd+datetime.timedelta(days = i),format))
	return ts

def loadandmerge(a,startdate,enddate):
	AllData = {}
	Stylekey = ''
	AllStyle = [u"酒店",u"旅游",u"机票",u"酒店-单体品牌词",u"酒店-通用词",u"酒店-地域酒店",u"酒店-连锁品牌词",u"旅游-通用词",u"旅游-单线路",u"机票通用词",u"机票-交通出行聚合",u"总计"]
	data = xlrd.open_workbook(a)
	table = data.sheets()[1]
	for i in range(table.nrows):
		line = table.row_values(i)[:12]
		if line[0] in AllStyle:
			Stylekey = line[0]
			AllData[Stylekey]={}
			continue
		if line[0] ==u"总计" or line[0] ==u'' or line[0] == '':
			continue
		AllData[Stylekey][line[0]] = line[1:]
	for k in timeseries(startdate,enddate):
		oneday = loadoneday(k)
		for key in oneday:
			if AllData.has_key(key):
				AllData[key][k] = oneday[key]
			else:
				print "doont has key",key
			########################修改
			if AllData[key][k][4] == 0 or AllData[key][k][6] == 0:
				PV_exp_comp = 0.01
				Cost_exp_comp = 0.01
			else:
				PV_exp_comp = float(AllData[key][k][0])/AllData[key][k][4]
				Cost_exp_comp = float(AllData[key][k][2])/AllData[key][k][6]
			AllData[key][k].append(PV_exp_comp)
			AllData[key][k].append(Cost_exp_comp)
	#读取所有数据，开始增加所有样式的分天总计项，不能带“机票”、"旅游"、“酒店”，否则重复
	AllData[u"总计"] = {}
	for k in timeseries("20160622",enddate):
		ExpPV1=0
		ExpClick1 =0
		ExpCost1=0
		ComPV1= 0
		ComClick1= 0
		ComCost1= 0
		#print "print date",k
		for key in [u"酒店-单体品牌词",u"酒店-通用词",u"酒店-地域酒店",u"酒店-连锁品牌词",u"旅游-通用词",u"旅游-单线路",u"机票通用词",u"机票-交通出行聚合"]:
			if AllData[key].has_key(k):
				#print "%s date has  %s"%(k,key)
				ExpPV,ExpClick,ExpCost,ExpCTR,ComPV,ComClick,ComCost,ComCTR,CTR_change,PV_exp_comp,Cost_exp_comp = AllData[key][k]
				ExpPV1 += ExpPV
				ExpClick1 += ExpClick
				ExpCost1 += ExpCost
				ComPV1 += ComPV
				ComClick1 += ComClick
				ComCost1 += ComCost
			else :
				continue
		if ExpPV1 == 0:
			print "%r has no data"%k
			continue
		else:
			ExpCTR1 = float(ExpClick1)/ExpPV1
			if ComClick1 == 0:
				ComCTR1 = 0.01
				CTR_change1 = 0.01
				PV_exp_comp1 = 0.01
				Cost_exp_comp1 = 0.01
			else:
				ComCTR1 = float(ComClick1)/ComPV1
				CTR_change1 = ExpCTR1/ComCTR1-1
				PV_exp_comp1 = float(ExpPV1) / ComPV1
				Cost_exp_comp1 = float(ExpCost1) / ComCost1
		#print "%s data is"%k,ExpPV1,ExpClick1,ExpCost1,ExpCTR1,ComPV1,ComClick1,ComCost1,ComCTR1,CTR_change1,PV_exp_comp1,Cost_exp_comp1
		AllData[u"总计"][k]=[ExpPV1,ExpClick1,ExpCost1,ExpCTR1,ComPV1,ComClick1,ComCost1,ComCTR1,CTR_change1,PV_exp_comp1,Cost_exp_comp1]
	#增加分样式总计：
	AllSumStyle = [u"酒店",u"旅游",u"机票",u"酒店-单体品牌词",u"酒店-通用词",u"酒店-地域酒店",u"酒店-连锁品牌词",u"旅游-通用词",u"旅游-单线路",u"机票通用词",u"机票-交通出行聚合",u"总计"]
	for key in AllSumStyle:
		ExpPV2=0
		ExpClick2 =0
		ExpCost2=0
		ComPV2= 0
		ComClick2= 0
		ComCost2= 0
		for k in timeseries("20160622",enddate):
			if AllData[key].has_key(k):
				ExpPV,ExpClick,ExpCost,ExpCTR,ComPV,ComClick,ComCost,ComCTR,CTR_change,PV_exp_comp,Cost_exp_comp = AllData[key][k]
				ExpPV2 += ExpPV
				ExpClick2 += ExpClick
				ExpCost2 += ExpCost
				ComPV2 += ComPV
				ComClick2 += ComClick
				ComCost2 += ComCost
			else :
				continue
		if ExpPV2 == 0:
			print "%r has no data"%k
			continue
		else:
			ExpCTR2 = float(ExpClick2)/ExpPV2
			if ComClick2 == 0:
				ComCTR2 = 0.01
				CTR_change2 = 0.01
				PV_exp_comp2 = 0.01
				Cost_exp_comp2 = 0.01
			else:
				ComCTR2 = float(ComClick2)/ComPV2
				CTR_change2 = ExpCTR2/ComCTR2-1
				PV_exp_comp2 = float(ExpPV2) / ComPV2
				Cost_exp_comp2 = float(ExpCost2) / ComCost2
		AllData[key][u"总计"]=[ExpPV2,ExpClick2,ExpCost2,ExpCTR2,ComPV2,ComClick2,ComCost2,ComCTR2,CTR_change2,PV_exp_comp2,Cost_exp_comp2]
	return AllData
	
def loadoneday(date):
	s = date+'.xls'
	oneday_data = xlrd.open_workbook(s)
	table = oneday_data.sheets()[3]
	dic = {}
	for j in range(table.nrows):
		line = table.row_values(j)[:10]
		if j == 0:
			continue
		if line[0] ==u"汇总":
			continue
		dic[line[0]] = line[1:]
	return dic
	

def writeall(dic,startdate,enddate,savename):
	#AllStyle = [u"酒店",u"旅游",u"机票",u"酒店-单体品牌词",u"酒店-通用词",u"酒店-地域酒店",u"酒店-连锁品牌词",u"旅游-通用词",u"旅游-单线路",u"机票通用词",u"机票-交通出行聚合"]
	workbook = pyExcelerator.Workbook()
	sheet1 = workbook.add_sheet(u'结论')
	sheet2 = workbook.add_sheet(u'最终结论表')
	
	writefirst(sheet1,dic,startdate,enddate)
	writesecond(sheet2,dic,startdate,enddate)
	
	workbook.save(savename)
	
	print "all is write into the EXCEL"
	
def writefirst(sheet,dic,startdate,enddate):
	myfont = pyExcelerator.Font()
	#myfont.name = u'Times New Roman'
	myfont.bold = True
	pattern1 = pyExcelerator.Pattern() 
	pattern1.pattern = pyExcelerator.Pattern.SOLID_PATTERN 
	pattern1.pattern_fore_colour = 5
	pattern2 = pyExcelerator.Pattern() 
	pattern2.pattern = pyExcelerator.Pattern.SOLID_PATTERN 
	pattern2.pattern_fore_colour = 22
	pattern3 = pyExcelerator.Pattern() 
	pattern3.pattern = pyExcelerator.Pattern.SOLID_PATTERN 
	pattern3.pattern_fore_colour = 44 
	mystyle1 = pyExcelerator.XFStyle()
	mystyle1.font = myfont
	mystyle1.pattern = pattern1
	#加粗标黄 
	mystyle2 = pyExcelerator.XFStyle()
	mystyle2.pattern = pattern1
	#标黄
	mystyle3 = pyExcelerator.XFStyle()
	mystyle3.pattern = pattern2
	#为空标灰
	mystyle4 = pyExcelerator.XFStyle()
	mystyle4.pattern = pattern3
	mystyle4.font = myfont
	#为蓝加粗
	row = 0
	for day in timeseries(startdate,enddate):
		sheet.write(row,0,day)
		firsttitle_write(sheet,row)
		row += 1
		for key in [u"酒店-单体品牌词",u"酒店-地域酒店",u"酒店-通用词",u"酒店-连锁品牌词",u"酒店",u"旅游-通用词",u"旅游-单线路",u"旅游",u"机票通用词",u"机票-交通出行聚合",u"机票",u"总计"]:
			if dic[key].has_key(day):
				if key == u"酒店" or key == u"旅游" or key == u"机票":
					sheet.write(row,0,key,mystyle1)
					ExpPV,ExpClick,ExpCost,ExpCTR,ComPV,ComClick,ComCost,ComCTR,CTR_change,PV_exp_comp,Cost_exp_comp = dic[key][day]
					sheet.write(row,1,ExpPV,mystyle2)
					sheet.write(row,2,ExpClick,mystyle2)
					sheet.write(row,3,ExpCost,mystyle2)
					sheet.write(row,4,ExpCTR,mystyle2)
					sheet.write(row,5,ComPV,mystyle2)
					sheet.write(row,6,ComClick,mystyle2)
					sheet.write(row,7,ComCost,mystyle2)
					sheet.write(row,8,ComCTR,mystyle2)
					sheet.write(row,9,CTR_change,mystyle2)
					sheet.write(row,10,PV_exp_comp,mystyle2)
					sheet.write(row,11,Cost_exp_comp,mystyle2)
				elif key == u"总计":
					sheet.write(row,0,key,mystyle4)
					ExpPV,ExpClick,ExpCost,ExpCTR,ComPV,ComClick,ComCost,ComCTR,CTR_change,PV_exp_comp,Cost_exp_comp = dic[key][day]
					sheet.write(row,1,ExpPV,mystyle4)
					sheet.write(row,2,ExpClick,mystyle4)
					sheet.write(row,3,ExpCost,mystyle4)
					sheet.write(row,4,ExpCTR,mystyle4)
					sheet.write(row,5,ComPV,mystyle4)
					sheet.write(row,6,ComClick,mystyle4)
					sheet.write(row,7,ComCost,mystyle4)
					sheet.write(row,8,ComCTR,mystyle4)
					sheet.write(row,9,CTR_change,mystyle4)
					sheet.write(row,10,PV_exp_comp,mystyle4)
					sheet.write(row,11,Cost_exp_comp,mystyle4)
				else:
					sheet.write(row,0,key)
					ExpPV,ExpClick,ExpCost,ExpCTR,ComPV,ComClick,ComCost,ComCTR,CTR_change,PV_exp_comp,Cost_exp_comp = dic[key][day]
					sheet.write(row,1,ExpPV)
					sheet.write(row,2,ExpClick)
					sheet.write(row,3,ExpCost)
					sheet.write(row,4,ExpCTR)
					sheet.write(row,5,ComPV)
					sheet.write(row,6,ComClick)
					sheet.write(row,7,ComCost)
					sheet.write(row,8,ComCTR)
					sheet.write(row,9,CTR_change)
					sheet.write(row,10,PV_exp_comp)
					sheet.write(row,11,Cost_exp_comp)
			else:
				sheet.write(row,0,key,mystyle3)
				sheet.write(row,1,'',mystyle3)
				sheet.write(row,2,'',mystyle3)
				sheet.write(row,3,'',mystyle3)
				sheet.write(row,4,'',mystyle3)
				sheet.write(row,5,'',mystyle3)
				sheet.write(row,6,'',mystyle3)
				sheet.write(row,7,'',mystyle3)
				sheet.write(row,8,'',mystyle3)
				sheet.write(row,9,'',mystyle3)
				sheet.write(row,10,'',mystyle3)
				sheet.write(row,11,'',mystyle3)
			#else:
				#print "%s don't have data of %s"%(key,day)
#########################################################以下用来增加单天备注
			if key == u"酒店-通用词" and day == "20160622":
				sheet.write(row,12,u"当日badcase较多，抽取日PV大于10的query进行分析")
			if key == u"旅游-单线路" and day == "20160623":
				sheet.write(row,12,u"PV量较少数据不稳定，待观察")
			row += 1
	print "the first sheet is write"
#定义表头函数		
def firsttitle_write(sheet,row):
	myfont = pyExcelerator.Font()
	#myfont.name = u'Times New Roman'
	myfont.bold = True
	pattern = pyExcelerator.Pattern() 
	pattern.pattern = pyExcelerator.Pattern.SOLID_PATTERN 
	pattern.pattern_fore_colour = 52 
	mystyle = pyExcelerator.XFStyle()
	mystyle.font = myfont
	mystyle.pattern = pattern
	sheet.write(row,1,u'实验组总PV',mystyle)
	sheet.write(row,2,u'实验组总点击',mystyle)
	sheet.write(row,3,u'实验组总消耗',mystyle)
	sheet.write(row,4,u'实验组CTR',mystyle)
	sheet.write(row,5,u'对照组总PV',mystyle)
	sheet.write(row,6,u'对照组总点击',mystyle)
	sheet.write(row,7,u'对照组总消耗',mystyle)
	sheet.write(row,8,u'对照组CTR',mystyle)
	sheet.write(row,9,u'CTR变化',mystyle)
	sheet.write(row,10,u'PV对比',mystyle)
	sheet.write(row,11,u'消耗对比',mystyle)
	sheet.write(row,12,u'备注',mystyle)

def secondtitle_write(sheet,row):
	myfont = pyExcelerator.Font()
	#myfont.name = u'Times New Roman'
	myfont.bold = True
	pattern = pyExcelerator.Pattern() 
	pattern.pattern = pyExcelerator.Pattern.SOLID_PATTERN 
	pattern.pattern_fore_colour = 52 
	mystyle = pyExcelerator.XFStyle()
	mystyle.font = myfont
	mystyle.pattern = pattern
	sheet.write(row,1,u'实验组总PV',mystyle)
	sheet.write(row,2,u'实验组总点击',mystyle)
	sheet.write(row,3,u'实验组总消耗',mystyle)
	sheet.write(row,4,u'实验组CTR',mystyle)
	sheet.write(row,5,u'对照组总PV',mystyle)
	sheet.write(row,6,u'对照组总点击',mystyle)
	sheet.write(row,7,u'对照组总消耗',mystyle)
	sheet.write(row,8,u'对照组CTR',mystyle)
	sheet.write(row,9,u'CTR变化',mystyle)
	sheet.write(row,10,u'PV对比',mystyle)
	sheet.write(row,11,u'消耗对比',mystyle)

def writesecond(sheet,dic,startdate,enddate):
	myfont = pyExcelerator.Font()
	#myfont.name = u'Times New Roman'
	myfont.bold = True
	pattern1 = pyExcelerator.Pattern() 
	pattern1.pattern = pyExcelerator.Pattern.SOLID_PATTERN 
	pattern1.pattern_fore_colour = 52 
	mystyle1 = pyExcelerator.XFStyle()
	mystyle1.font = myfont
	mystyle1.pattern = pattern1
	pattern2 = pyExcelerator.Pattern() 
	pattern2.pattern = pyExcelerator.Pattern.SOLID_PATTERN 
	pattern2.pattern_fore_colour = 44 
	mystyle2 = pyExcelerator.XFStyle()
	mystyle2.font = myfont
	mystyle2.pattern = pattern2
	row = 0
	for key in [u"总计",u"酒店",u"旅游",u"机票",u"酒店-单体品牌词",u"酒店-通用词",u"酒店-地域酒店",u"酒店-连锁品牌词",u"旅游-通用词",u"旅游-单线路",u"机票通用词",u"机票-交通出行聚合"]:
		sheet.write(row,0,key,mystyle1)
		secondtitle_write(sheet,row)
		row += 1
		for day in timeseries(startdate,enddate):
			if dic[key].has_key(day):
				sheet.write(row,0,day)
				ExpPV,ExpClick,ExpCost,ExpCTR,ComPV,ComClick,ComCost,ComCTR,CTR_change,PV_exp_comp,Cost_exp_comp = dic[key][day]
				sheet.write(row,1,ExpPV)
				sheet.write(row,2,ExpClick)
				sheet.write(row,3,ExpCost)
				sheet.write(row,4,ExpCTR)
				sheet.write(row,5,ComPV)
				sheet.write(row,6,ComClick)
				sheet.write(row,7,ComCost)
				sheet.write(row,8,ComCTR)
				sheet.write(row,9,CTR_change)
				sheet.write(row,10,PV_exp_comp)
				sheet.write(row,11,Cost_exp_comp)
				row += 1
			if isWeekornot(day) == True:
				last7 = last7days(day)
				WeekExpCTR,WeekComCTR,WeekCTRChange = calcWeekCTR(dic,key,last7)
				if WeekExpCTR != 0:
					sheet.write(row,12,u"上周A路CTR",mystyle2)
					sheet.write(row,13,WeekExpCTR)
					sheet.write(row,14,u"上周对照组CTR",mystyle2)
					sheet.write(row,15,WeekComCTR)
					sheet.write(row,16,u"CTR变化",mystyle2)
					sheet.write(row,17,WeekCTRChange)
			#else:
				#print "%s don't have data of %s"%(key,day)
			##此处可增写备注
		sheet.write(row,0,u'总计',mystyle2)
		ExpPV,ExpClick,ExpCost,ExpCTR,ComPV,ComClick,ComCost,ComCTR,CTR_change,PV_exp_comp,Cost_exp_comp = dic[key][u"总计"]
		sheet.write(row,1,ExpPV,mystyle2)
		sheet.write(row,2,ExpClick,mystyle2)
		sheet.write(row,3,ExpCost,mystyle2)
		sheet.write(row,4,ExpCTR,mystyle2)
		sheet.write(row,5,ComPV,mystyle2)
		sheet.write(row,6,ComClick,mystyle2)
		sheet.write(row,7,ComCost,mystyle2)
		sheet.write(row,8,ComCTR,mystyle2)
		sheet.write(row,9,CTR_change,mystyle2)
		sheet.write(row,10,PV_exp_comp,mystyle2)
		sheet.write(row,11,Cost_exp_comp,mystyle2)
		row += 2
	print "the second sheet is write"

def isWeekornot(weekdate):
	superstartdate = "20160620"
	format = "%Y%m%d"
	sd = datetime.datetime.strptime(superstartdate,format)
	ed = datetime.datetime.strptime(weekdate,format)
	delta = ed-sd
	if (delta.days+1)%7 == 0:
		return True
	else:
		return False

def last7days(weekdate):
	format = "%Y%m%d"
	ed = datetime.datetime.strptime(weekdate,format)
	sd = ed - datetime.timedelta(days = 6)
	sd = datetime.datetime.strftime(sd,format)
	ts = timeseries(sd,weekdate)
	
	return ts

def calcWeekCTR(dic,key,ts):
	ExpPV = 0
	ExpClick = 0
	ComPV = 0
	ComClick = 0
	for day in ts:
		if dic[key].has_key(day):
			ExpPV += dic[key][day][0]
			ExpClick += dic[key][day][1]
			ComPV += dic[key][day][4]
			ComClick += dic[key][day][5]
	if ExpPV == 0:
		return 0,0,0
	else:
		return float(ExpClick)/ExpPV,float(ComClick)/ComPV,(float(ExpClick)/ExpPV)/(float(ComClick)/ComPV)-1
		

def write_on_3(dic,savename):
	workbook = pyExcelerator.Workbook()
	sheet1 = workbook.add_sheet('origin_empty')
	sheet2 = workbook.add_sheet('QuerySum_empty')
	sheet3 = workbook.add_sheet('ValidData_empty')
	sheet4 = workbook.add_sheet('result')
	row = 0
	sheet1.write(row,0,'Just to fill blank,please read sheet4')
	sheet2.write(row,0,'Just to fill blank,please read sheet4')
	sheet3.write(row,0,'Just to fill blank,please read sheet4')
	sheet4.write(row,0,'Style')
	sheet4.write(row,1,'ExpPV')
	sheet4.write(row,2,'ExpClick')
	sheet4.write(row,3,'ExpCost')
	sheet4.write(row,4,'ExpCTR')
	sheet4.write(row,5,'ComPV')
	sheet4.write(row,6,'ComClick')
	sheet4.write(row,7,'ComCost')
	sheet4.write(row,8,'ComCTR')
	sheet4.write(row,9,'CTR_change')
	
	for key in dic:
		row += 1
		sheet4.write(row,0,key.decode('gbk', "ignore"))
		sheet4.write(row,1,dic[key][0])
		sheet4.write(row,2,dic[key][1])
		sheet4.write(row,3,dic[key][2])
		sheet4.write(row,4,dic[key][3])
		sheet4.write(row,5,dic[key][4])
		sheet4.write(row,6,dic[key][5])
		sheet4.write(row,7,dic[key][6])
		sheet4.write(row,8,dic[key][7])
		sheet4.write(row,9,dic[key][8])
	workbook.save(savename)
