# -*- coding:utf-8 -*-
#函数用于PC品专、直通车、无线品专、suggestion的数据
#源文件对应数据排序命名依次为a,b,c,d，格式统一为csv

result = "日期"+"\t"+"产品"+"\t"+"曝光"+"\t"+"点击"+"CTR"+"\n"
fout = open("汇总数据.txt",'w')

#pc品专处理
pcbrand = open("a.csv",'r')
title = pcbrand.readline()

for line in pcbrand.readlines():
	line = line.strip()
	(shunxu,riqi,kehu,hetonghao,pv,click_search,click_BD,click_explorer,click_input,click,ctr) = line.split(',')
	if shunxu == '"总计"':
		continue
	result += riqi + "\t"+"PC品专" + "\t"+pv + "\t"+click+"\t"+ctr+"\n"
pcbrand.close()

#直通车处理
zhitongche = open("b.csv",'r')
title = zhitongche.readline()

for line in zhitongche:
	line = line.strip()
	(hetonghao,kehu,riqi,pv,click,ctr) = line.split(',')
	result += riqi + "\t"+"直通车" + "\t"+pv + "\t"+click+"\t"+ctr+"\n"
zhitongche.close()

#无线品专处理
wlbrand = open("c.csv",'r')
title = wlbrand.readline()

for line in wlbrand:
	line = line.strip()
	(shunxu,riqi,kehu,hetonghao,pv,click,ctr,pv_mob,pv_color,click_mob,click_color) = line.split(',')
	if shunxu =='"总计"':
		break
	result += riqi + "\t"+"无线品专" + "\t"+pv + "\t"+click+"\t"+ctr+"\n"
wlbrand.close()

#suggestion处理
wlsugg = open("d.csv",'r')
title = wlsugg.readline()

for line in wlsugg:
	line = line.strip()
	(shunxu,riqi,kehu,hetonghao,pv,click,ctr) = line.split(',')
	if shunxu == '"总计"':
		break
	result += riqi + "\t"+"无线suggestion" + "\t"+pv + "\t"+click+"\t"+ctr+"\n"
wlsugg.close()

#print result

fout.write(result)
fout.close()
