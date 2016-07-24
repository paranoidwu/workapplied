# -*- coding:utf-8 -*-
import xlrd,datetime,pyExcelerator
import OTA
#输入之前的文件和需要增加的起止日期,将其录入字典，一级key为样式，二级key为日期
AllData = OTA.loadandmerge("result_analysis.xls","20160719","20160719")
#将汇总的字典输出，此次起始时间要求不动，只更改结束时间
OTA.writeall(AllData,"20160622","20160719")
