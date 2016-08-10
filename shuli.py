# -*- coding:utf-8 -*-
import pyExcelerator,OTA
#���忪ʼ�ͽ����ļ�����
startdate = "20160805"
enddate = "20160806"
#д���ļ�Ϊ˫���ֵ䣬һ��keyΪyyyymmdd�ַ���������keyΪ��0��ʼ������
data = {}
for d in OTA.timeseries(startdate,enddate):
	fin = open(d,'r')
	mark = 0
	data[d] = {}
#��ȡ�ļ�����������
	for line in fin:
		line = line.strip()
		if OTA.check_line_item(line) == 10:
			data[d][mark] = line.split('\t')
		elif OTA.check_line_item(line) == 9:
			data[d][mark] = OTA.repair_line_9(line)
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
savename = startdate+"-"+enddate+"�߶˶�����ʽ���ݱ���"+".xls"
workbook.save(savename)