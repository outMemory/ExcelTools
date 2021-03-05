# -*- coding: utf-8 -*-
# -*- coding: mbcs -*-

import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open('D:/test/待合并表格/qq.xls')
# FileFormat = 51 is for .xlsx extension
try:
    wb.SaveAs('D:\\test1/q1.xlsx', FileFormat=51)
except Exception as e:
    print(e)
    pass
wb.Close()  # FileFormat = 56 is for .xls extension
excel.Application.Quit()
