# -*- coding:utf-8 -*-
import xlrd,datetime,pyExcelerator
import OTA
#����֮ǰ���ļ�����Ҫ���ӵ���ֹ����,����¼���ֵ䣬һ��keyΪ��ʽ������keyΪ����
AllData = OTA.loadandmerge("result_analysis.xls","20160719","20160719")
#�����ܵ��ֵ�������˴���ʼʱ��Ҫ�󲻶���ֻ���Ľ���ʱ��
OTA.writeall(AllData,"20160622","20160719")
