# -*- coding: utf-8 -*-
__author__ = 'DeathSea'
__doc__ = """
将课表转化为班表的模块
Translate
函数：
    __init__(xls_sh)
        初始化传入xlrd.open_workbook(FileName).sheet_by_index(0)
    GetRow()
        获取课表的行数
    GetCol()
        获取课表的列数
    JudgeDayLorS(int)
        判断非周末班是否为短版
        传入非周末班的序号（0,1,2,3,4...）
        短班返回真，长版返回假，短板<=150分钟
    JudgeEndLorS(class_index)
        判断周末班是否为短版
        传入周末班的序号（0,1,2,3,4...）
        短班返回真，长版返回假，短板<=150分钟
    SetWhoHeIs(str)
        设定是谁课表
        传入非空str
        生成的班表以该名称标记可以上班的班次
    SetLongOnly(boolean)
        设定是否只上长版，默认不是
        传入Boolean值或int
    SetShortOnly(boolean)
        设定是否只上短班，默认不是
        传入Boolean值或int
    SetClassesClasses(int, int)
        设置非周末时间与周末时间多少个班
        前一个参数非周末时间，后一个参数周末时间
    SetWeekDayClassesTime(list)
        设置非周末时间的上班时间
        list格式为["hour:min-hour:min","hour:min-hour:min",...]
    SetWeekEndClassesTime(list)
        设置周末时间的上班时间
        list格式为["hour:min-hour:min","hour:min-hour:min",...]
    SetWeekDayClassesMustBlank(list)
        设置非周末时间需要从第i节课没课到第j节没课
        以二元组储存(i, j)，几个班就存几组
        如五个班就可以传入为[(1, 2), (3, 4), (5, 6), (6, 10), (10, 13)]
    SetWeekEndClassesMustBlank(list)
        设置周末时间需要从第i节课没课到第j节没课
        以二元组储存(i, j)，几个班就存几组
        如五个班就可以传入为[(1, 2), (3, 4), (5, 6), (6, 10), (10, 13)]
    ReturnResult()
        返回最后生成的班表结果
        二元数组
        [
        [col,col,col...],
        [col,col,col...],
        ...
        ]
        示例：
        [
        ['', u'\u5468\u4e00', u'\u5468\u4e8c', u'\u5468\u4e09', u'\u5468\u56db', u'\u5468\u4e94', '', u'\u5468\u516d', u'\u5468\u65e5'],
        ['8:00-9:50', '', 't', 't', '', '', '', '', ''],
        ['9:50-12:00', '', '', '', '', '', '9:30-12:00', 't', 't'],
        ['12:00-14:10', '', 't', '', 't', '', '12:00-14:30', 't', 't'],
        ['14:10-17:50', '', '', '', '', '', '14:30-18:00', '', ''],
        ['17:50-21:30', '', '', '', '', '', '18:00-21:30', '', '']
        ]
    Reset()
        重置所有设置参数
    InitResultList()
        设置全部完成后初始化返回班表的结果
    CreateResultList()
        生成班表
"""
import xlrd
class Translate(object):
    def __init__(self, xls_sh):
        self.xls_sh = xls_sh
        self.rows = 0
        self.cols = 0
        self.classes = (0, 0)
        #一天有多少个班，非周末时间前，周末时间后
        self.result_list = [
            ["", "周一".decode("u8"), "周二".decode("u8"), "周三".decode("u8"), "周四".decode("u8"), "周五".decode("u8"), "",  "周六".decode("u8"), "周日".decode("u8")],
        ]
        self.WeekDayTime = []
        #非周末上班时间
        self.WeekEndTime = []
        #周末上班时间
        self.WeekDayClassesMustBlank = []
        #非周末时间需要从第i节课没课到第j节没课
        #以二元组储存(i, j)，几个班就存几组
        #如五个班就可以设为[(1, 2), (3, 4), (5, 6), (6, 10), (10, 13)]
        #下同
        self.WeekEndClassesMustBlank = []
        self.GetRow()
        self.GetCol()
        self.OnlyLong = 0
        self.OnlyShort = 0
        #self.InitResultList()

    def GetRow(self):
        self.rows = self.xls_sh.nrows
        return self.xls_sh.nrows

    def GetCol(self):
        self.cols = self.xls_sh.ncols
        return self.xls_sh.ncols

    def JudgeDayLorS(self, class_index):
        #短班返回真
        #长斑返回假
        time = self.WeekDayTime[class_index]
        s, e = time.split("-")
        s = int(s.split(':')[0]) * 60 + int(s.split(":")[1])
        e = int(e.split(":")[0]) * 60 + int(e.split(":")[1])
        if e - s <= 150:
            return True
        return False

    def JudgeEndLorS(self, class_index):
        time = self.WeekEndTime[class_index]
        s, e = time.split("-")
        s = int(s.split(':')[0]) * 60 + int(s.split(":")[1])
        e = int(e.split(":")[0]) * 60 + int(e.split(":")[1])
        if e - s <= 150:
            return True
        return False

    def SetWhoHeIs(self, whoheis):
        if whoheis:
            self.whoheis = whoheis
            return 0
        return "the name is blank"

    def SetLongOnly(self, default = False):
        self.OnlyLong = default
        return 0

    def SetShortOnly(self, default = False):
        self.OnlyShort = default
        return 0

    def SetClassesClasses(self, weekday, weekend):
        self.classes = (weekday, weekend)
        #设置非周末时间与周末时间多少个班
        return 0

    def SetWeekDayClassesTime(self, list):
        self.WeekDayTime = list
        if len(self.WeekDayTime) != self.classes[0]:
            return "Setting the number of Classes Error"
        return 0

    def SetWeekEndClassesTime(self, list):
        self.WeekEndTime = list
        if len(self.WeekEndTime) != self.classes[1]:
            return "Setting the number of Classes Error"
        return 0

    def SetWeekDayClassesMustBlank(self, list):
        self.WeekDayClassesMustBlank = list
        if len(self.WeekDayClassesMustBlank) != self.classes[0]:
            return "class must settng error"
        return 0

    def SetWeekEndClassesMustBlank(self, list):
        self.WeekEndClassesMustBlank = list
        if len(self.WeekEndClassesMustBlank) != self.classes[1]:
            return "class must blank setting error"
        return 0

    def ReturnResult(self):
        if self.result_list:
            return self.result_list
        else:
            return "haven't init and create the result list"

    def Reset(self):
        self.rows = 0
        self.cols = 0
        self.WeekDayTime = []
        self.WeekEndTime = []
        self.classes = (0, 0)
        self.result_list = [
            ["", "周一".decode("u8"), "周二".decode("u8"), "周三".decode("u8"), "周四".decode("u8"), "周五".decode("u8"), "",  "周六".decode("u8"), "周日".decode("u8")],
        ]
        self.WeekDayClassesMustBlank = []
        self.WeekEndClassesMustBlank = []
        self.OnlyLong = 0
        self.OnlyShort = 0

    def InitResultList(self):
        if len(self.WeekEndTime) != 0 and len(self.WeekDayTime) != 0:
            #时间全部设定
            #self.ListCols = len(self.WeekDayTime) > len(self.WeekEndTime) and len(self.WeekDayTime) or len(self.WeekEndTime)
            #每天多少个班统一格式
            #    周一   周二   周三  周四  周五        周六  周日
            #时间1
            #时间2                              周末时间1
            #时间3
            #共self.ListCols+1行
            if len(self.WeekDayTime) != self.classes[0] and len(self.WeekEndTime) != self.classes[1]:
                return "Setting the number of Classes Error"
                #班表数目设置错误

            for i in range(0, len(self.WeekDayTime)):
                self.result_list.append([self.WeekDayTime[i], "", "", "", "", "", "", "", ""])

            for i in range(0, len(self.WeekEndTime)):
                self.result_list[i + 1 + len(self.WeekDayTime) - len(self.WeekEndTime)][6] = self.WeekEndTime[i]
        else:
            return  "haven't set the week day time and week end time"

    def CreateResultList(self):
        if self.xls_sh.cell_value(3, 0) == "周一".decode("u8"):
            start = (3, 1)
        elif self.xls_sh.cell_value(2, 0) == "周一".decode("u8"):
            start = (2, 1)
        else:
            return "the TimeTable is no export from jwc"

        Day_Sub_End = len(self.WeekDayTime) - len(self.WeekEndTime)
        for row in range(start[0], self.rows):
            #for col in range(start[1], self.cols):
            if row <= start[0] + 4:
                #非周末时间
                for class_index in range(0, self.classes[0]):
                    col = 0
                    col_start, col_end = self.WeekDayClassesMustBlank[class_index]
                    for col in range(col_start, col_end + 1):
                        if self.xls_sh.cell_value(row, col) != "":
                            break
                    if col == col_end and self.xls_sh.cell_value(row, col) == "":
                        if not self.OnlyShort and not self.OnlyLong:
                            self.result_list[class_index + 1][row - start[0] + 1] = self.whoheis
                            continue
                        if self.OnlyShort:
                            if self.JudgeDayLorS(class_index):
                                self.result_list[class_index + 1][row - start[0] + 1] = self.whoheis
                                continue
                            else:
                                pass
                        if self.OnlyLong:
                            if not self.JudgeDayLorS(class_index):
                                self.result_list[class_index + 1][row - start[0] + 1] = self.whoheis
                                continue
                            else:
                                pass
            else:
                #周末时间
                for class_index in range(0, self.classes[1]):
                    col = 0
                    col_start, col_end = self.WeekEndClassesMustBlank[class_index]
                    for col in range(col_start, col_end + 1):
                        if self.xls_sh.cell_value(row, col) != "":
                            break
                    if col == col_end and self.xls_sh.cell_value(row, col) == "":
                        if not self.OnlyShort and not self.OnlyLong:
                            self.result_list[class_index + Day_Sub_End + 1][row - start[0] + 2] = self.whoheis
                            continue
                        if self.OnlyShort:
                            if self.JudgeEndLorS(class_index):
                                self.result_list[class_index + Day_Sub_End + 1][row - start[0] + 2] = self.whoheis
                                continue
                            else:
                                pass
                        if self.OnlyLong:
                            if not self.JudgeEndLorS(class_index):
                                self.result_list[class_index + Day_Sub_End + 1][row - start[0] + 2] = self.whoheis
                                continue
                            else:
                                pass

if __name__ == "__main__":
    wob = xlrd.open_workbook("test.xls")
    sh  = wob.sheet_by_index(0)
    TT = Translate(sh)
    TT.SetClassesClasses(5, 4)
    TT.SetWeekDayClassesTime(["8:00-9:50", "9:50-12:00", "12:00-14:10", "14:10-17:50", "17:50-21:30"])
    TT.SetWeekEndClassesTime(["9:30-12:00", "12:00-14:30", "14:30-18:00", "18:00-21:30"])
    #TT.SetClassesClasses(3, 3)
    #TT.SetWeekDayClassesTime(["10:20-12:30", "12:30-14:40", "14:50-17:50"])
    #TT.SetWeekEndClassesTime(["10:20-12:30", "12:30-14:40", "14:50-17:50"])
    TT.SetWeekDayClassesMustBlank([(1, 2), (3, 4), (5, 6), (6, 10), (10, 13)])
    TT.SetWeekEndClassesMustBlank([(3, 4), (5, 6), (6, 10), (10, 13)])
    TT.SetWhoHeIs("t")
    TT.SetShortOnly(1)
    TT.InitResultList()
    TT.CreateResultList()
    print TT.ReturnResult()