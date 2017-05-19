# -*- coding: UTF-8 -*-
import sys
import xlwt
import os
reload(sys)
sys.setdefaultencoding('utf-8')

import sys
try:
	import xml.etree.cElementTree as ET
except:
	import xml.etree.ElementTree as ET


XmlNameList = [] #xml 文件列表
XlsNameList = [] #xls 文件列表
if (len(sys.argv) < 2):
	print "Usage: python test.py xmlfiledir,"
	exit()


#遍历目录下面的xml文件,返回要读取的xml文件名，和要写入的xls名
def ListFile(dir):
	xmlfilelist = []
	xlsfilelist = []
	files = os.listdir(dir)
	for name in files:
        #print name
        #print name.replace('xml','xls')
		#name = name.replace(' ','\ ')
		if (name.find('.xml') == -1):
			continue
		xmlfilelist.append(dir+'/'+name)
		xlsfilelist.append(name.replace('xml','xls'))
	return xmlfilelist,xlsfilelist
	
	
def ReadXmlToXls(xmlfile,xlsfile):
	try:
		tree = ET.parse(xmlfile)
		root = tree.getroot()
	except:
		print "Error:cannot parse file:",xmlfile
		sys.exit(1)
	print root.tag,"---",root.attrib

	workbook = xlwt.Workbook(encoding = 'ascii')
	worksheet = workbook.add_sheet("sheet1")


	i = 0
	j = 0
	for para in root[1].findall('para'):
		worksheet.write(i, 0, 'para');
		worksheet.write(i, 1, para.get('id'));
		i = i + 1
		for sent in para.findall('sent'):
			print sent.get('id'),sent.get('cont')
			worksheet.write(i, 0, sent.get('id'));
			worksheet.write(i, 1, sent.get('cont'));
			i = i + 1
			print '*' * 20
			for word in sent.findall('word'):
				sem_parent = ''
				sem_relate = ''
				for sem in word.findall('sem'):
					sem_parent = sem.get('parent')
					sem_relate = sem.get('relate')
				worksheet.write(i, 0, word.get('cont'));
				worksheet.write(i, 1, word.get('pos'));
				worksheet.write(i, 2, word.get('ne'));
				worksheet.write(i, 3, word.get('parent'));
				worksheet.write(i, 4, word.get('relate'));
				worksheet.write(i, 5, word.get('semparent'));
				worksheet.write(i, 6, word.get('semrelate'));
				worksheet.write(i, 7, sem_parent);
				worksheet.write(i, 8, sem_relate);
				print word.get('cont'), word.get('pos'), word.get('ne'), word.get('parent'), word.get('relate'), word.get('semparent'),word.get('semrelate'),sem_parent,sem_relate
				i = i + 1
	workbook.save(xlsfile)	

	
if (os.path.isfile(sys.argv[1])):
	ReadXmlToXls(sys.argv[1],sys.argv[1].replace('.xml','.xls'))
	exit()
	
XmlNameList = ListFile(sys.argv[1])[0]
XlsNameList = ListFile(sys.argv[1])[1]
i = 0
for i in range(len(XmlNameList)):
	ReadXmlToXls(XmlNameList[i],XlsNameList[i])
