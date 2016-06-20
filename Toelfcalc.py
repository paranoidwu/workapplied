#-*- coding:utf-8 -*-
from sys import argv

script,filename = argv

fp = open(filename)

title = list()
a1 = list()
a2 = list()

for line in fp.readlines():
	start_position = 0
	mark = 0
	for lett in line:
		mark+=1
		if lett ==":":
			kw = line[start_position:mark-1]
			print "kewwords is %r,it's position is %d to %d"%(kw,start_position,mark-1)
			start_position = mark
		if lett =="," or lett == "\n":
			kwnumber = line[start_position:mark-1]
			start_position = mark
			print "KN is %r "%kwnumber
			if kw =="title":
				title.append(kwnumber)
				print "Title append %r"%kwnumber
			elif kw == "a1":
				a1.append(kwnumber)
				print "a1 append %r"%kwnumber
			else:
				a2.append(kwnumber)
				print "a2 append %r"%kwnumber

f = open("cpyied.txt",'w')
f.write("title\ta1\ta2\n")
for i in range(4):
	f.write(title[i])
	f.write("\t")
	f.write(a1[i])
	f.write("\t")
	f.write(a2[i])
	f.write("\n")
f.close()

			
