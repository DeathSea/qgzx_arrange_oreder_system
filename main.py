# -*- coding: cp936 -*-
import xlrd,xlwt
fname = "test1.xls"
bk = xlrd.open_workbook(fname)
shxrange = range(bk.nsheets)
try:
    sh = bk.sheet_by_index(0)
except:
    print "no sheet in %s named Sheet1" % fname
nrows = sh.nrows
ncols = sh.ncols
#print "nrows %d, ncols %d" % (nrows,ncols)
 
#cell_value = sh.cell_value(1,1)
#print cell_value
 
row_list = []
for i in range(0,nrows):
    row_data = sh.row_values(i)
    row_list.append(row_data)
for i in  row_list:
    for j in i:
        print j,
    print 

####################################################33



fname = "test2.xls"
def add_values(excel_name,row_list,belong_name = ""):
    bk = xlrd.open_workbook(excel_name)
    shxrange = range(bk.nsheets)
    try:
        sh = bk.sheet_by_index(0)
    except:
        print "no sheet in %s named Sheet1" % fname
    nrows = sh.nrows
    ncols = sh.ncols
    print "nrows %d, ncols %d" % (nrows,ncols)
    #9行16列
    #cell_value = sh.cell_value(1,1)
    #print cell_value
     
    for i in  range(0,9):
        if sh.row_values(i) == row_list[i]:pass
        else:
            row_va = sh.row_values(i);
            for j in range(0,15):
                print i,j
                if row_list[i][j] == row_va[j]:pass
                elif row_list[i][j] == "":
                    row_list[i][j] == row_va[j]
                elif row_va[j]!="":
                    if row_va[j] == "√".decode("cp936"):
                        row_va[j] = belong_name.decode("cp936")
                    row_list[i][j] = row_list[i][j] + "、".decode("cp936") + row_va[j]
def writeXls(file_name,row_list):
    xls   = xlwt.Workbook()
    sheet = xls.add_sheet("Sheet1")
    for i in range(0,9):
        for j in range(0,15):
            sheet.write(i,j,row_list[i][j])
    xls.save(file_name)
add_values(fname,row_list)
add_values("test3.xls",row_list)
#add_values("test4.xls",row_list)
#add_values("test5.xls",row_list)
#add_values("test3.xls",row_list)
for i in  row_list:
    for j in i:
        print j,
    print 
writeXls("test4.xls",row_list)
