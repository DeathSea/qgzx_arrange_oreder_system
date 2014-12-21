# -*- coding: utf-8 -*-
import wx
import wx.lib.filebrowsebutton
import wx.lib.dialogs
import wx.grid
import xlrd
import os
import xlwt

class SouthCampus(object):
    def __init__(self , xls_sh):
        self.xls_sh = xls_sh
        self.GetRows()
        self.GetCols()
        self.return_list = [
            ["", "周一".decode("u8"), "周二".decode("u8"), "周三".decode("u8"), "周四".decode("u8"), "周五".decode("u8"), "周六".decode("u8"), "周日".decode("u8")],
            ["10:20-12:30", ],
            ["12:30-14:40", ],
            ["14:50-17:50", ]]

    def GetRows(self):
        self.rows = self.xls_sh.nrows
        return self.rows

    def GetCols(self):
        self.cols = self.xls_sh.ncols
        return self.cols

    def do(self , whoheis):
        #row行
        #col列
        #print self.cols,self.rows
        for i in range(3 , self.rows):
            a1 , a2 , a3 = 0 , 0 , 0
            for j in range(1 , self.cols):
                if 3<=j<=9:
                    if self.xls_sh.cell_value(i , j) == "":
                         if j == 3 or j == 4:
                             a1 += 1
                         if j == 5 or j == 6:
                             a2 += 1
                         if 6 <= j <= 9:
                             a3 += 1
            if a1 == 2:#第一个班可以上
                self.return_list[1].append(whoheis)
            if a2 == 2:#第二个班可以上
                self.return_list[2].append(whoheis)
            if a3 == 4:#第三个班可以上
                self.return_list[3].append(whoheis)
            if a1 != 2:#第一个班不能上
                self.return_list[1].append("")
            if a2 != 2:#第二个班不能上
                self.return_list[2].append("")
            if a3 != 4:#第三个班不能上
                self.return_list[3].append("")

    def clearup(self):
        self.rows = 0
        self.cols = 0
        self.return_list = [
            ["" , "周一".decode("u8") , "周二".decode("u8") , "周三".decode("u8") , "周四".decode("u8") , "周五".decode("u8") , "周六".decode("u8") , "周日".decode("u8")],
            ["10:20-12:30"  , ],
            ["12:30-14:40"  , ],
            ["14:50-17:50"  , ]]

    def set_xls_sh(self , xls_sh):
        self.xls_sh = xls_sh
        self.clearup()
        self.GetRows()
        self.GetCols()
        return

    def return_lists(self):
        return self.return_list

class Integrated(object):#本部融合
    def __init__(self):
        self.return_list = []

    def Main(self, li):
        Ish = self.ReadFile(li[0])
        del li[0]
        self.FirstOneToAdd(Ish)
        for i in li:
            Ish = self.ReadFile(i)
            whoheis = self.WhatTheManName(i)
            self.AddToReturnList(Ish,whoheis)

    def FirstOneToAdd(self, Tsh):
        self.rows = Tsh.nrows
        self.cols = Tsh.ncols
        for i in range(0, self.rows):
            self.return_list.append(Tsh.row_values(i))

    def AddToReturnList(self, Tsh, whoheis):
        for i in range(0, self.rows):
            try:
                if Tsh.row_values(i) == self.return_list[i]:pass
                else:
                    row_va = Tsh.row_values(i)
                    for j in range(0, self.cols):
                        try:
                            if self.return_list[i][j] == row_va[j]:pass
                            elif self.return_list[i][j] == "":
                                self.return_list[i][j] = row_va[j]
                            elif row_va[j] != "":
                                if row_va[j] == "√".decode("u8"):
                                    row_va[j] = whoheis.decode("u8")
                                self.return_list[i][j] = self.return_list[i][j] + "、".decode("u8") + row_va[j]
                        except IndexError, p:
                            wx.MessageBox(
                                message = "出了点错误，请检查名字是%s的班表是否有多余的单元".decode("u8")%whoheis,
                                caption = "警告".decode("u8"),
                                style = wx.OK,
                            )
            except IndexError, p:
                wx.MessageBox(
                    message = "出了点错误，请检查名字是%s的班表是否有多余的单元".decode("u8")%whoheis,
                    caption = "警告".decode("u8"),
                    style = wx.OK,
                )

    def ReadFile(self, filename):
        wob = xlrd.open_workbook(filename)
        sh  = wob.sheet_by_index(0)
        return sh

    def WriteFile(self, filename):
        xs = xlwt.Workbook()
        st = xs.add_sheet("Sheet1")
        for i in range(0, self.rows):
            for j in range(0, self.cols):
                st.write(i, j, self.return_list[i][j])
        xs.save(filename)

    def WhatTheManName(self, filename):
        return os.path.split(filename)[1].split(".")[0]

    def ReturnList(self):
        return self.return_list

    def ReturnRC(self):
        return self.rows, self.cols

class Integrated_V2(object):#南校区融合
    def __init__(self):
        self.return_list = []
        self.rows = 4
        self.cols = 8

    def Main(self, li):
        Th = self.ReadFile(li[0])
        self.NSC = SouthCampus(Th)
        self.NSC.do(self.WhatTheManName(li[0]))
        thisList = self.NSC.return_lists()
        self.return_list = thisList
        del li[0]
        for i in li:
            Th = self.ReadFile(i)
            self.AddToReturnList(Th, self.WhatTheManName(i))

    def AddToReturnList(self, Tsh, whoheis):
        self.NSC.set_xls_sh(Tsh)
        self.NSC.do(whoheis)
        thisList = self.NSC.return_lists()
        print thisList
        for i in range(0, self.rows):
            if thisList[i] == self.return_list[i]:
                continue
            for j in range(0, self.cols):
                print i, j
                if thisList[i][j] != "" and self.return_list[i][j] == thisList[i][j]:
                    continue
                elif self.return_list[i][j] == "":
                    self.return_list[i][j] = thisList[i][j]
                elif self.return_list[i][j] != "" and thisList[i][j] != "":
                    self.return_list[i][j] = self.return_list[i][j] + "、".decode("u8") + thisList[i][j]
        # wx.MessageBox(
        #     message = "出了点错误，请检查名字是%s的班表是否有多余的单元".decode("u8")%whoheis,
        #     caption = "警告".decode("u8"),
        #     style = wx.OK,
        # )

    def ReadFile(self, filename):
        wob = xlrd.open_workbook(filename)
        sh  = wob.sheet_by_index(0)
        return sh

    def WriteFile(self, filename):
        xs = xlwt.Workbook()
        st = xs.add_sheet("Sheet1")
        for i in range(0, self.rows):
            for j in range(0, self.cols):
                st.write(i, j, self.return_list[i][j])
        xs.save(filename)

    def WhatTheManName(self, filename):
        return os.path.split(filename)[1].split(".")[0]

    def ReturnList(self):
        return self.return_list

    def ReturnRC(self):
        return self.rows, self.cols

class MyFrame(wx.Frame):
    def __init__(self, parent, title):
        wx.Frame.__init__(self, parent, title=title, size=(800, 545))
        self.integrated = None
        self.panel = wx.Panel(self)
        self.FileBrowser = wx.lib.filebrowsebutton.FileBrowseButton(
            parent = self.panel,
            id = -1,
            pos = (0, 0),
            size = (500, 35),
            style = wx.TAB_TRAVERSAL,
            labelText = "选择班表路径:".decode("u8"),
            buttonText = "浏览文件".decode("u8"),
            toolTip = "键入路径+文件名或者按浏览文件按钮浏览文件".decode("u8"),
            dialogTitle = "选择文件".decode("u8"),
            startDirectory = ".",
            initialValue = "",
            fileMask = "*.xls;*.xlsx",
            fileMode = wx.FD_OPEN,
            changeCallback = lambda x: x,
            labelWidth = 0,
            name = "fileBrowseButton"
        )
        self.addButton = wx.Button(
            parent = self.panel,
            id = -1,
            label = "添加".decode("u8"),
            pos = (500, 3),
            size = (80, 29),
            style = 0,
            validator = wx.DefaultValidator,
            name = "addButton"
        )
        wx.StaticLine(
            parent = self.panel,
            id = -1,
            pos = (54, 43),
            size = (526, 2),
            style = wx.LI_HORIZONTAL,
            name = wx.StaticLineNameStr
        )
        wx.StaticText(
            parent = self.panel,
            id = -1,
            label = "已选班表".decode("u8"),
            pos = (4,35),
            style = 0,
            name = wx.StaticTextNameStr
        )
        self.xlslist = wx.ListBox(
            parent = self.panel,
            id = -1,
            pos = (3, 59),
            size = (575, 200),
        )
        wx.StaticLine(
            parent = self.panel,
            id = -1,
            pos = (54,277),
            size = (526,2)
        )
        wx.StaticText(
            parent = self.panel,
            id = -1,
            label = "班表预览".decode("u8"),
            pos = (4,267),
            style = 0,
            name = wx.StaticTextNameStr
        )
        self.grid = wx.grid.Grid(
            parent = self.panel,
            id = -1,
            pos = (3, 292),
            size = (495, 200),
        )
        self.grid.CreateGrid(0, 0)
        self.grid.EnableEditing(False)
        self.grid.AppendRows(1)
        self.grid.AppendCols(1)
        self.PreviewButton = wx.Button(
            parent = self.panel,
            id = -1,
            label = "预览".decode("u8"),
            pos = (500, 360),
            size = (80, 29),
            style = 0,
            validator = wx.DefaultValidator,
            name = "PreViewButton"
        )
        self.PreviewSButton = wx.Button(
            parent = self.panel,
            id = -1,
            label = "南校区预览".decode("u8"),
            pos = (500, 330),
            size = (80, 29),
            style = 0,
            validator = wx.DefaultValidator,
            name = "PreViewSButton"
        )
        self.ExportButton = wx.Button(
            parent = self.panel,
            id = -1,
            label = "导出".decode("u8"),
            pos = (500, 410),
            size = (80, 29),
            style = 0,
            validator = wx.DefaultValidator,
            name = "ExportButton"
        )
        sb = wx.StaticBox(
            parent = self.panel,
            id = -1,
            label = "说明".decode("u8"),
            pos = (585, 0),
            size = (196, 500),
            style = 0,
            name = wx.StaticBoxNameStr
        )
        wx.StaticText(
            parent = sb,
            id = -1,
            label = "先按浏览文件，再点击添加。\n导出时，先点击预览，确认没有\n错误之后再点击导出。\n文件最好以人名命名，以方便在\n整合的时候处理那些打钩的表格.".decode("u8"),
            pos = (10, 20),
        )
        self.addButton.Bind(wx.EVT_BUTTON, self.ClickAddButton, self.addButton)
        self.PreviewButton.Bind(wx.EVT_BUTTON, self.ClickPreviewButton, self.PreviewButton)
        self.PreviewSButton.Bind(wx.EVT_BUTTON, self.ClickSPreviewButton, self.PreviewSButton)
        self.ExportButton.Bind(wx.EVT_BUTTON, self.ClickExportButton, self.ExportButton)

        self.Show(True)

    def ClickAddButton(self, event):
        filename = self.FileBrowser.GetValue()
        if(not filename):  #文件名非空
            wx.MessageBox(
                message = "文件名不能为空".decode("u8"),
                caption = "警告".decode("u8"),
                style = wx.OK,
            )
            return
        self.FileBrowser.SetValue("")
        if(self.xlslist.FindString(filename) == wx.NOT_FOUND):  ###插入时注意没有存在xlslist中
            self.xlslist.InsertItems(items = [filename], pos = 0)
        else:
            wx.MessageBox(
                message = "所选班表已经存在列表中".decode("u8"),
                caption = "警告".decode("u8"),
                style = wx.OK,
            )

    def ClickPreviewButton(self, event):

        self.grid.DeleteCols(0, self.grid.GetNumberCols())
        self.grid.DeleteRows(0, self.grid.GetNumberRows())
        self.integrated = Integrated()
        self.filelist = []
        for i in range(0, self.xlslist.GetCount()):
            filename = self.xlslist.GetString(i)
            if filename != "":
                self.filelist.append(filename)
        if len(self.filelist) == 0:
            return
        self.integrated.Main(self.filelist)
        r, c = self.integrated.ReturnRC()
        self.grid.AppendRows(r)
        self.grid.AppendCols(c)
        for i in range(0, self.xlslist.GetCount()):
            self.xlslist.SetString(i, "")
        result_list = self.integrated.ReturnList()
        for i in range(0, len(result_list)):
            for j in range(0, len(result_list[i])):
                self.grid.SetCellValue(row = i, col = j, s = result_list[i][j])

    def ClickSPreviewButton(self, event):
        self.grid.DeleteCols(0, self.grid.GetNumberCols())
        self.grid.DeleteRows(0, self.grid.GetNumberRows())
        self.integrated = Integrated_V2()
        self.filelist = []
        for i in range(0, self.xlslist.GetCount()):
            filename = self.xlslist.GetString(i)
            if filename != "":
                self.filelist.append(filename)
        if len(self.filelist) == 0:
            return
        self.integrated.Main(self.filelist)
        r, c = self.integrated.ReturnRC()
        self.grid.AppendRows(r)
        self.grid.AppendCols(c)
        for i in range(0, self.xlslist.GetCount()):
            self.xlslist.SetString(i, "")
        result_list = self.integrated.ReturnList()
        for i in range(0, len(result_list)):
            for j in range(0, len(result_list[i])):
                self.grid.SetCellValue(row = i, col = j, s = result_list[i][j])

    def ClickExportButton(self, event):

        filedialog = wx.FileDialog(
            parent = self.panel,
            message = "save the file...",
            defaultDir = "",
            defaultFile = ".xls",
            wildcard = "*.xls;*.xlsx",
            style = wx.FD_SAVE,
            pos = wx.DefaultPosition,
            size = wx.DefaultSize,
        )
        flag = filedialog.ShowModal()
        if flag == wx.ID_CANCEL:
            wx.MessageBox(
                message = "取消保存".decode("u8"),
                caption = "提示".decode("u8"),
                style = wx.OK,
            )
        elif flag == wx.ID_OK:
            filename = filedialog.GetFilename()
            if self.integrated == None:
                wx.MessageBox(
                    message = "班表未生成请先预览".decode("u8"),
                    caption = "警告".decode("u8"),
                    style = wx.OK,
                )
                return
            elif self.integrated.ReturnList() == []:
                wx.MessageBox(
                    message = "班表未生成，请先预览".decode("u8"),
                    caption = "警告".decode("u8"),
                    style = wx.OK,
                )
                return
            elif len(filename.split(".")) != 2:
                wx.MessageBox(
                    message = "文件名不能为空".decode("u8"),
                    caption = "警告".decode("u8"),
                    style = wx.OK,
                )
                return
            self.integrated.WriteFile(filedialog.GetPath())
            wx.MessageBox(
                message = "保存成功".decode("u8"),
                caption = "提示".decode("u8"),
                style = wx.OK,
            )

app = wx.App(False)
frame = MyFrame(None, "test")
app.MainLoop()
# newint = Integrated()
# li = ["test1.xls", "test2.xls","test3.xls","test4.xls","test5.xls"]
# newint.Main(li)
# newint.WriteFile("test7.xls")