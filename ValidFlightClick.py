# -*- coding:utf-8 -*-
#����VR�μ������б�
'''
#21208501	all	PC-��ͨ-�Ҳ��Ʊ����(ȥ�Ķ���ҵ)	�۳�wangfan,dancheng��������Ϊ�㣬��С���ɺ���
#21208401	all	PC-��ͨ-�Ҳ��Ʊ��·(ȥ�Ķ���ҵ)	ͬ�ϣ�ģ��һ��
#20114302	all	PC-��Ʊ-���ں��ߴ�flash(ȥ�Ķ���ҵ��	NULL��submit,flash
#21136201	all	PC-��Ʊ-ʡ��ʡ(ȥ�Ķ���ҵ)	Chaxun = submit(�Ʋ��û���Ʒ��¼�ظ�),title
20006620	all	PC-��Ʊ-����������(ȥ�Ķ���ҵ)	submit,title,more,morelink,tejiapiao1-8,
#21145301	all	PC-��Ʊ-���ʺ��߲���(ȥ�Ķ���ҵ)	NULL,submit
#20000901	all	PC-��Ʊ-���ʺ���(ȥ�Ķ���ҵ)	ͬ��
#21228101	all	PC-��Ʊ-���ں��߲���(ȥ�Ķ���ҵ)	ͬ��
20009504	all	PC-��Ʊ-�������б����ؼ�(ȥ�Ķ���ҵ)	title,submit,link,tejiapiao1-6
#20009503	all	PC-��Ʊ-����������(ȥ�Ķ���ҵ)	title,submit,more
#20009603	all	PC-��Ʊ-���ں��ߴ�����(ȥ�Ķ���ҵ)	NULL,submit
'''
import re,pyExcelerator
#���ݸ�ʽ��title,���У����ڣ�VRID��PV��������Լ�����ַ���
#������ݶ�ȡ�����Ϊ�ֵ�flightdata,��¼һ�׹ؼ���Ϊ���ڣ�����Ϊvrid����¼��ʽΪ�б�PV�������
flightdata = {}
fin = open("a.csv","r+")
for line in fin:
	line = line.strip()
	(date,vrid,pv,click1,detail) = line.split(',',4)
	pv = int(pv)
	if vrid == "21208501" or vrid == "21208401":
		click = int(click1)
	elif vrid == "20114302" or vrid == "21145301" or vrid == "20000901" or vrid == "21228101"or vrid == "20009603":
		click = 0
		#null
		pattern = re.compile('NULL:(\d*)',re.S)
		item = re.findall(pattern,detail)
		if item:
			click += int(item[0])
		#submit
		pattern = re.compile('submit:(\d*)',re.S)
		item = re.findall(pattern,detail)
		if item:
			click += int(item[0])
		#flash
			pattern = re.compile('flash:(\d*)',re.S)
		item = re.findall(pattern,detail)
		if item:
			click += int(item[0])
	elif vrid == "21136201" or vrid == "20009503":
		click = 0
		#title
		pattern = re.compile('title:(\d*)',re.S)
		item = re.findall(pattern,detail)
		if item:
			click += int(item[0])
		#submit
		pattern = re.compile('submit:(\d*)',re.S)
		item = re.findall(pattern,detail)
		if item:
			click += int(item[0])
		#more
		pattern = re.compile('more:(\d*)',re.S)
		item = re.findall(pattern,detail)
		if item:
			click += int(item[0])
	elif vrid == "20006620" or vrid == "20009504":
		click = 0
		#title
		pattern = re.compile('title:(\d*)',re.S)
		item = re.findall(pattern,detail)
		if item:
			click += int(item[0])
		#submit
		pattern = re.compile('submit:(\d*)',re.S)
		item = re.findall(pattern,detail)
		if item:
			click += int(item[0])
		#more
		pattern = re.compile('more:(\d*)',re.S)
		item = re.findall(pattern,detail)
		if item:
			click += int(item[0])
		#morelink
		pattern = re.compile('morelink:(\d*)',re.S)
		item = re.findall(pattern,detail)
		if item:
			click += int(item[0])
		#link
		pattern = re.compile('link:(\d*)',re.S)
		item = re.findall(pattern,detail)
		if item:
			click += int(item[0])
		#tejiapiao1-8
		pattern = re.compile('tejiapiao1:(\d*)',re.S)
		item = re.findall(pattern,detail)
		if item:
			click += int(item[0])
		pattern = re.compile('tejiapiao2:(\d*)',re.S)
		item = re.findall(pattern,detail)
		if item:
			click += int(item[0])
		pattern = re.compile('tejiapiao3:(\d*)',re.S)
		item = re.findall(pattern,detail)
		if item:
			click += int(item[0])
		pattern = re.compile('tejiapiao4:(\d*)',re.S)
		item = re.findall(pattern,detail)
		if item:
			click += int(item[0])
		pattern = re.compile('tejiapiao5:(\d*)',re.S)
		item = re.findall(pattern,detail)
		if item:
			click += int(item[0])
		pattern = re.compile('tejiapiao6:(\d*)',re.S)
		item = re.findall(pattern,detail)
		if item:
			click += int(item[0])
		pattern = re.compile('tejiapiao7:(\d*)',re.S)
		item = re.findall(pattern,detail)
		if item:
			click += int(item[0])
		pattern = re.compile('tejiapiao8:(\d*)',re.S)
		item = re.findall(pattern,detail)
		if item:
			click += int(item[0])
	else:
		continue
	if flightdata.has_key(date):
		flightdata[date][vrid] = [pv,click]
	else:
		flightdata[date] = {}
		flightdata[date][vrid] = [pv,click]

#�����ݼ�¼��excel
workbook = pyExcelerator.Workbook()
sheet1 = workbook.add_sheet('sheet1')
myfont = pyExcelerator.Font()
#myfont.name = u'Times New Roman'
myfont.bold = True
mystyle = pyExcelerator.XFStyle()
mystyle.font = myfont
sheet1.write(0,0,'Date',mystyle)
sheet1.write(0,1,"VRID",mystyle)
sheet1.write(0,2,"PV",mystyle)
sheet1.write(0,3,"Click",mystyle)
write_line = 1
for daydata in flightdata:
	for key in flightdata[daydata]:
		sheet1.write(write_line,0,daydata)
		sheet1.write(write_line,1,key)
		sheet1.write(write_line,2,flightdata[daydata][key][0])
		sheet1.write(write_line,3,flightdata[daydata][key][1])
		write_line += 1
workbook.save("flightdata.xls")