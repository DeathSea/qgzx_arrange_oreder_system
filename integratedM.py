# -*- coding: utf-8 -*-
__author__ = 'DeathSea'
__doc__ = """

"""
import os
#import XLSFileRW
from translateM import Translate

class Integrate(object):
    def __init__(self):
        #self.selflist = personal_list
        self.resultList      = []
        self.rows            = 0
        self.cols            = 0
        self.OnlyLongList    = []#只上长班
        self.OnlyShoertList  = []#只上短班
        self.TuiGuangCaiBian = []#推广采编
        self.TuiGuangOther   = []#推广其他
        self.RenLi           = []#人力
        self.CaiWu           = []#财务
        self.ShiChang        = []#市场
        self.ZongJiang       = []#总监团队
        self.LianLuo         = []#联络部
        self.Color           = {
            "TuiGuang":54,
            "RenLi":54,
            "CaiWu":54,
            "ShiChang":54,
            "LianLuo":54,
            "ZongJiang":54,
        }#背景色设置

    def SetColor(self, TuiGuang = 54, RenLi = 54, CaiWu = 54, ShiChang = 54, LianLuo = 54, ZongJiang = 54):
        self.Color["TuiGuang"] = TuiGuang
        self.Color["ShiChang"] = ShiChang
        self.Color["RenLi"] = RenLi
        self.Color["CaiWu"] = CaiWu
        self.Color["LianLuo"] = LianLuo
        self.Color["ZongJiang"] = ZongJiang

    def Main(self, li):
        pass

    def WhatTheManName(self, filename):
        return os.path.split(filename)[1].split(".")[0]

    def ReturnRC(self):
        return self.rows, self.cols

    def SetTuiguangCaibianList(self, list):
        """设定推广部采编名单
        """
        if list:
            self.TuiGuangCaiBian = list
            return
        return "please give a list at last one person"

    def SetTuiguangOtherList(self, list):
        """设定推广部其他人的名单"""
        if list:
            self.TuiGuangOther = list
            return
        return "please give a list at last one person"

    def SetRenliList(self, list):
        """设定人力名单"""
        if list:
            self.RenLi = list
            return
        return "please give a list at last one person"

    def SetCaiwuList(self, list):
        """设定财务名单"""
        if list:
            self.CaiWu = list
            return
        return "please give a list at last one person"

    def SetShichangList(self, list):
        """设定市场名单"""
        if list:
            self.ShiChang = list
            return
        return "please give a list at last one person"

    def SetLianluoList(self, list):
        """设定联络部名单"""
        if list:
            self.LianLuo = list
            return
        return "please give a list at last one person"

    def SetZongjiangList(self, list):
        """设定总监团队名单"""
        if list:
            self.ZongJiang = list
            return
        return "please give a list at last one person"

    def Reset(self):
        self.resultList = []
        self.rows = 0
        self.cols = 0

    def ReturnResult(self):
        return self.resultList