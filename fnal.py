# -*- coding:utf-8 -*-
import xlrd,datetime,pyExcelerator
import OTA
#����֮ǰ���ļ�����Ҫ���ӵ���ֹ����,����¼���ֵ䣬һ��keyΪ��ʽ������keyΪ����
AllData = OTA.loadandmerge("result_1024.xls","20161024","20161106")
#�����ܵ��ֵ�������˴���ʼʱ��Ҫ�󲻶���ֻ���Ľ���ʱ��,���һ������Ϊϣ���γɵ�excel����
OTA.writeall(AllData,"20160622","20161106","result_1107.xls")
