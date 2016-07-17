# -*- coding:utf-8 -*-
import OTA
import pyExcelerator
#每次处理，请修改fileAB name
fileA_name = 'expr_account_key.with_click.20160715'
fileB_name = 'comp_account_key.with_click.20160715'

fileA = open(fileA_name)
fileB = open(fileB_name)

#对文件处理，输出为字典
dicA = OTA.convert2dic(fileA)
print "Data of Experiment is loaded"
dicB = OTA.convert2dic(fileB)
print "Data of Control is loaded"

#求取交集，实现vlookup功能，记录在字典中
dicC = OTA.dic_intersection(dicA,dicB)

#定义样式对照表
dicE = {"11":u"酒店-地域酒店","12":u"酒店-通用词","13":u"酒店-单体品牌词","14":u"酒店-连锁品牌词","21":u"旅游-单线路","22":u"旅游-通用词","31":u"机票通用词","32":u"机票-交通出行聚合","96":u"机票","97":u"酒店","98":u"旅游","99":u"汇总"}

#有机票请用这个
dicF = {1:"13",2:"11",3:"14",4:"21",5:"22",6:"32",7:"31",8:"96",9:"97",10:"98",11:"99"}

'''
#没机票用这个
dicF = {1:"13",2:"11",3:"14",4:"21",5:"22",9:"97",10:"98",11:"99"}
'''

#汇总数据
dicD = OTA.calculate_result(dicC)


workbook = pyExcelerator.Workbook()
sheet1 = workbook.add_sheet('origin')
sheet2 = workbook.add_sheet('QuerySum')
sheet3 = workbook.add_sheet('ValidData')
sheet4 = workbook.add_sheet('result')

#重新打开文件
fileA = open(fileA_name)
fileB = open(fileB_name)

#左侧为实验组，右侧为对照组,写入原始数据
OTA.write_origin(fileA,fileB,sheet1,0,0)

#写入汇总文件
OTA.write_QuerySum(dicA,dicB,sheet2,0,0)

#写入有效vlookup数据
OTA.write_ValidData(dicC,sheet3)

#写入结果
OTA.write_result(dicD,sheet4,dicE,dicF)

#寻找合适的命名
Final_name = fileA_name[-8:]+'.xls'

workbook.save(Final_name)