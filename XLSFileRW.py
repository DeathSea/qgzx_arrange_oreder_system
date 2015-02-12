# -*- coding: utf-8 -*-
__author__ = 'DeathSea'
__doc__ = """
读取与写入课表信息
"""
import xlrd,xlwt

color_list = [
    ['aqua', '0x00FFFF'],
    ['black', '0x000000'],
    ['blue', '0x0000FF'],
    ['blue_gray', '0x666699'],
    ['bright_green', '0x00ff00'],
    ['brown', '0x993300'],
    ['coral', '0xFF7F50'],
    ['cyan_ega', '0x00ffff'],
    ['dark_blue', '0x000080'],
    ['dark_blue_ega', '0x000080'],
    ['dark_green', '0x003300'],
    ['dark_green_ega', '0x008000'],
    ['dark_purple', '0x660066'],
    ['dark_red', '0x800000'],
    ['dark_red_ega', '0x800000'],
    ['dark_teal', '0x003366'],
    ['dark_yellow' ,'0x808000'],
    ['gold', '0xFFD700'],
    ['gray_ega', '0x808080'],
    ['gray25', '0xc0c0c0'],
    ['gray40', '0x969696'],
    ['gray50', '0x808080'],
    ['gray80', '0x333333'],
    ['green', '0x008000'],
    ['ice_blue', '0xccccff'],
    ['indigo', '0x4B0082'],
    ['ivory', '0xFFFFF0'],
    ['lavender', '0xE6E6FA'],
    ['light_blue', '0x3366ff'],
    ['light_green', '0xccffcc'],
    ['light_orange', '0xff9900'],
    ['light_turquoise', '0xccffff'],
    ['light_yellow', '0xffff99'],
    ['lime', '0x00FF00'],
    ['magenta_ega', '0xff00ff'],
    ['ocean_blue', '0x0066cc'],
    ['olive_ega', '0x808000'],
    ['olive_green', '0x333300'],
    ['orange', '0xFFA500'],
    ['pale_blue', '0x99ccff'],
    ['periwinkle', '0x9999ff'],
    ['pink', '0xFFC0CB'],
    ['plum', '0xDDA0DD'],
    ['purple_ega', '0x800080'],
    ['red', '0xFF0000'],
    ['rose', '0xff99cc'],
    ['sea_green', '0x339966'],
    ['silver_ega', '0xc0c0c0'],
    ['sky_blue', '0x00ccff'],
    ['tan', '0xD2B48C'],
    ['teal', '0x008080'],
    ['teal_ega', '0x008080'],
    ['turquoise', '0x40E0D0'],
    ['violet', '0xEE82EE'],
    ['white', '0xFFFFFF'],
    ['yellow', '0xFFFF00']
]

def ReadFile(filename):
    return xlrd.open_workbook(filename).sheet_by_index(0)



def WhichDepartmentTheStaffIn(StaffName, DepartmentStaffDict):
    """
        DepartmentStaffDict   部门员工dict
        格式为
        {
        "部门1":[staff1,staff2,staff3,.....staffn],
        "部门2":[staff1,staff2,staff3,.....staffn],
        }
    """
    for key,value in DepartmentStaffDict.items():
        if StaffName in value:
            return key
    return None

def XlsWriter(filename, WriteIntoFileList, ColorDict, DepartmentStaffDict):
    """
    filename          导出文件名
    WriteIntoFileList 写入文件列表
       格式为：
       [
       [col1,col2,col3,......,coln].#row1
       [col1,col2,col3,......,coln].#row2
       ]
    ColorDict         颜色dict
        格式为:
        {
        "部门":颜色编号，
        "部门2":颜色编号
        }
        颜色编号为color_list第i个（从0数起）
    DepartmentStaffDict   部门员工dict
        格式为
        {
        "部门1":[staff1,staff2,staff3,.....staffn],
        "部门2":[staff1,staff2,staff3,.....staffn],
        }
    """
    DepartmentStyleDict = {}
    for key,value in ColorDict.items():
        DepartmentStyleDict[key] = xlwt.easyxf('pattern: pattern solid, fore_colour '+color_list[value][0]+';')
    FileListRowLength = len(WriteIntoFileList)
    FileListColLength = len(WriteIntoFileList[0])

    wb = xlwt.Workbook()
    ws = wb.add_sheet('Sheet1')
    for row in range(0, FileListRowLength):
        for col in range(0, FileListColLength):
            department = WhichDepartmentTheStaffIn(WriteIntoFileList[row][col], DepartmentStaffDict)
            if department:
                xlsstyle = DepartmentStyleDict[department]
            else:
                xlsstyle = xlwt.easyxf("pattern: pattern solid, fore_colour white;")
            ws.write(row, col, WriteIntoFileList[row][col], xlsstyle)
    wb.save(filename)



if __name__ == "__main__":
    print "test start"
    filename = "TotestXlsFileRw.py.xls"
    print "filename:",filename
    ToFileList = [
        ["",           "Sat","Wet","Thi","Wet","Fri","","Sta","Sun"],
        ["8:00-10:20", "who1",  "who2",  "",      "",      "",      "", "",     ""],
    ]
    ColorDict = {
        "dep1":1,
        "dep2":2,
    }
    StaffDict = {
        "dep1":["who1"],
        "dep2":["who2"],
    }
    XlsWriter(filename, ToFileList, ColorDict, StaffDict)
    print "test complete"