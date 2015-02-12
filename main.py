# -*- coding: utf-8 -*-
__author__ = 'DeathSea'
__version__ = "V3.0"
__doc__ = """
this file edited by DeathSea 2014.11.6
main file in run
"""
import xlwt
ll = [['', u'\u5468\u4e00', u'\u5468\u4e8c', u'\u5468\u4e09', u'\u5468\u56db', u'\u5468\u4e94', '', u'\u5468\u516d', u'\u5468\u65e5'], ['8:00-9:50', '', 't', 't', '', '', '', '', ''], ['9:50-12:00', '', '', '', '', '', '9:30-12:00', 't', 't'], ['12:00-14:10', '', 't', '', 't', '', '12:00-14:30', 't', 't'], ['14:10-17:50', '', '', '', '', '', '14:30-18:00', '', ''], ['17:50-21:30', '', '', '', '', '', '18:00-21:30', '', '']]
def WriteFile(filename):
    xs = xlwt.Workbook()
    st = xs.add_sheet("Sheet1")
    for i in range(0, len(ll)):
        for j in range(0, len(ll[0])):
            st.write(i, j, ll[i][j])
    xs.save(filename)
WriteFile("output.xls")