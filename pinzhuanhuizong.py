# -*- coding:utf-8 -*-
#��������PCƷר��ֱͨ��������Ʒר��suggestion������
#Դ�ļ���Ӧ����������������Ϊa,b,c,d����ʽͳһΪcsv

result = "����"+"\t"+"��Ʒ"+"\t"+"�ع�"+"\t"+"���"+"CTR"+"\n"
fout = open("��������.txt",'w')

#pcƷר����
pcbrand = open("a.csv",'r')
title = pcbrand.readline()

for line in pcbrand.readlines():
	line = line.strip()
	(shunxu,riqi,kehu,hetonghao,pv,click_search,click_BD,click_explorer,click_input,click,ctr) = line.split(',')
	if shunxu == '"�ܼ�"':
		continue
	result += riqi + "\t"+"PCƷר" + "\t"+pv + "\t"+click+"\t"+ctr+"\n"
pcbrand.close()

#ֱͨ������
zhitongche = open("b.csv",'r')
title = zhitongche.readline()

for line in zhitongche:
	line = line.strip()
	(hetonghao,kehu,riqi,pv,click,ctr) = line.split(',')
	result += riqi + "\t"+"ֱͨ��" + "\t"+pv + "\t"+click+"\t"+ctr+"\n"
zhitongche.close()

#����Ʒר����
wlbrand = open("c.csv",'r')
title = wlbrand.readline()

for line in wlbrand:
	line = line.strip()
	(shunxu,riqi,kehu,hetonghao,pv,click,ctr,pv_mob,pv_color,click_mob,click_color) = line.split(',')
	if shunxu =='"�ܼ�"':
		break
	result += riqi + "\t"+"����Ʒר" + "\t"+pv + "\t"+click+"\t"+ctr+"\n"
wlbrand.close()

#suggestion����
wlsugg = open("d.csv",'r')
title = wlsugg.readline()

for line in wlsugg:
	line = line.strip()
	(shunxu,riqi,kehu,hetonghao,pv,click,ctr) = line.split(',')
	if shunxu == '"�ܼ�"':
		break
	result += riqi + "\t"+"����suggestion" + "\t"+pv + "\t"+click+"\t"+ctr+"\n"
wlsugg.close()

#print result

fout.write(result)
fout.close()
