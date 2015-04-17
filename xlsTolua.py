# -*- coding: utf-8 -*-
import codecs
import xdrlib ,sys
import xlrd
import os
import os.path

reload(sys)
sys.setdefaultencoding('utf-8')

def open_excel(file):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception,e:
        print str(e)

def readxls(file, clientdir, serverdir):
	data = open_excel(file)
	namearr = data.sheet_names()
	for shindex in range(0,2):
		namesplit = namearr[shindex].split('.')
		if(namesplit[0]=="c"):
			outputdir = clientdir
		elif(namesplit[0]=="s"):
			outputdir = serverdir
		else:
			return 

		print(outputdir+namesplit[1]+".lua")
		table = data.sheets()[shindex]
		output = codecs.open(outputdir+namesplit[1]+".lua", "w", "utf-8")
		output.write("\nreturn {\n\n")
		for row in range(2, table.nrows):
			tmp = "\t["+str(table.cell(row, 0).value)+"] = {  "
			for col in range(0, table.ncols):
				cchar = str(table.cell(row, col).value)
				tmp += str(table.cell(1, col).value) + " = " + cchar.decode('utf-8') + ", "
			output.write(tmp+"}, \n")
		output.write("\n}\n")
		output.close()

def makeconfig(rootdir, cdir, sdir):
	if (os.path.exists(cdir)==False):
		os.makedirs(cdir)
	if (os.path.exists(sdir)==False):
		os.makedirs(sdir)
	for parent,dirnames,filenames in os.walk(rootdir):     
		for dirname in  dirnames:
			makeconfig(rootdir+dirname, cdir+dirname+"/", sdir+dirname+"/")
		for filename in filenames:                         
			if (filename[0] != ".") and (parent == rootdir):
				# print(os.path.join(parent,filename).split(rootdir)[1])
				readxls(os.path.join(parent,filename), cdir, sdir)

if __name__=="__main__":
	makeconfig("策划数值配置/", "客户端配置表/", "服务端配置表/")


