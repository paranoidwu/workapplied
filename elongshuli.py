# -*- coding:utf-8 -*-
import pyExcelerator,OTA
#���忪ʼ�ͽ����ļ�����
startdate = "20161031"
enddate = "20161106"
#д���ļ�Ϊ˫���ֵ䣬һ��keyΪyyyymmdd�ַ���������keyΪ��0��ʼ������
data = {}
ts = OTA.timeseries(startdate,enddate)
for d in ts:
#for d in ["20160808","20160809","20160810","20160811","20160812","20160813","20160814"]:
	fin = open("C:\\Users\\wuenyu\\save\\OTA\\"+d+"\\"+d+"\\galaxy_ota_stat_upperleft_first_group_plan_yilong",'r')
	mark = 0
	data[d] = {}
#��ȡ�ļ�����������
	for line in fin:
		line = line.strip()
		if OTA.check_line_item(line) == 10:
			data[d][mark] = line.split('\t')
			if data[d][mark][9] in ["��Ʊ-��ͨ���оۺ�","��Ʊ-ͨ�ô�"] and int(data[d][mark][1]) not in [8243024,7510850,425560,18732105]:
				print "date %r,%r has wrong accountid"%(d,mark)
			if data[d][mark][9] in ["����-����·","����-ͨ�ô�"] and int(data[d][mark][1]) not in [18669566,18669561,18669559,18371715,18360413,18199240,7414539]:
				print "date %r,%r has wrong accountid"%(d,mark)
			if data[d][mark][9] in ["�Ƶ�-����Ʒ�ƴ�","�Ƶ�-ͨ�ô�","�Ƶ�-����Ƶ�","�Ƶ�-����Ʒ�ƴ�"] and int(data[d][mark][1]) not in [18872878,18872838,18872871,18872833,18872869,18872813,18872867,18872792,18872865,18872863,18872883,18872854,18872881,18872844,18872879,18872840,18872877,18872836,18872870,18872819,18872868,18872800,18872866,18872784,18872864,18872884,18872859,18872882,18872851,18872880,18872842]:
				print "date %r,%r has wrong accountid"%(d,mark)
		elif OTA.check_line_item(line) == 9:
			data[d][mark] = OTA.repair_line_9(line)
			if data[d][mark][9] in ["��Ʊ-��ͨ���оۺ�","��Ʊ-ͨ�ô�"] and int(data[d][mark][1]) not in [8243024,7510850,425560,18732105]:
				print "date %r,%r has wrong accountid"%(d,mark)
			if data[d][mark][9] in ["����-����·","����-ͨ�ô�"] and int(data[d][mark][1]) not in [18669566,18669561,18669559,18371715,18360413,18199240,7414539]:
				print "date %r,%r has wrong accountid"%(d,mark)
			if data[d][mark][9] in ["�Ƶ�-����Ʒ�ƴ�","�Ƶ�-ͨ�ô�","�Ƶ�-����Ƶ�","�Ƶ�-����Ʒ�ƴ�"] and int(data[d][mark][1]) not in [18872878,18872838,18872871,18872833,18872869,18872813,18872867,18872792,18872865,18872863,18872883,18872854,18872881,18872844,18872879,18872840,18872877,18872836,18872870,18872819,18872868,18872800,18872866,18872784,18872864,18872884,18872859,18872882,18872851,18872880,18872842]:
				print "date %r,%r has wrong accountid"%(d,mark)
		else:
			print "%s,%d line has problem,please check"%(d,mark)
		mark += 1
#����ļ�
workbook = pyExcelerator.Workbook()
sheet1 = workbook.add_sheet('sheet1')
myfont = pyExcelerator.Font()
#myfont.name = u'Times New Roman'
myfont.bold = True
mystyle = pyExcelerator.XFStyle()
mystyle.font = myfont
sheet1.write(0,0,'����'.decode('gbk', "ignore"),mystyle)
sheet1.write(0,1,"�˻�ID".decode('gbk', "ignore"),mystyle)
sheet1.write(0,2,"�ƻ�����".decode('gbk', "ignore"),mystyle)
sheet1.write(0,3,"������".decode('gbk', "ignore"),mystyle)
sheet1.write(0,4,"�ؼ���".decode('gbk', "ignore"),mystyle)
sheet1.write(0,5,"��ѯ��".decode('gbk', "ignore"),mystyle)
sheet1.write(0,6,"չ��".decode('gbk', "ignore"),mystyle)
sheet1.write(0,7,"���".decode('gbk', "ignore"),mystyle)
sheet1.write(0,8,"����".decode('gbk', "ignore"),mystyle)
sheet1.write(0,9,"�߶���ʽ".decode('gbk', "ignore"),mystyle)
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
savename = startdate+"-"+enddate+"�߶˶�����ʽ���ݱ���-����"+".xls"
workbook.save(savename)