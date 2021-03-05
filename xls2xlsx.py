# -*- coding: utf-8 -*-
import os
import sys

import win32com.client as win32
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QRadioButton, QButtonGroup


def xls_2_xlsx(xlspath):
    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(xlspath)
        # FileFormat = 51 is for .xlsx extension
        wb.SaveAs(xlspath + "x", FileFormat=51)
        wb.Close()  # FileFormat = 56 is for .xls extension
        excel.Application.Quit()
    except Exception as e:
        Ui_MainWindow.message = str(e)


def xlsx_2_xls(xlsxpath):
    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(xlsxpath)
        wb.SaveAs(xlsxpath[:-4] + 'xls', FileFormat=56)
        wb.Close()
        excel.Application.Quit()
    except Exception as e:
        Ui_MainWindow.message = str(e)


class Ui_MainWindow(object):
    filepath = 'D:/test/待合并表格'
    path_save = 'D:/test/合并结果'
    message = ''
    panduan = 1  # 判断转换方式，1为xls转xlsx，默认1

    def setupUi(self, MainWindow):
        # 主窗口参数设置
        MainWindow.setObjectName("MainWindow")
        MainWindow.setWindowTitle("拆分表格")
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        # 实例化
        MainWindow.setCentralWidget(self.centralwidget)
        MainWindow.setGeometry(QtCore.QRect(550, 250, 500, 500))

        # 设置读取文件夹的窗口，一般是待处理文件所在文件夹
        self.openfile = QtWidgets.QPushButton(self.centralwidget)
        self.openfile.setGeometry(QtCore.QRect(50, 100, 333, 35))
        self.openfile.setText("选择要拆分的Excel所在的文件夹")

        # 设置保存文件的窗口
        self.savepath = QtWidgets.QPushButton(self.centralwidget)
        self.savepath.setGeometry(QtCore.QRect(50, 160, 333, 35))
        self.savepath.setText("选择保存位置")

        # 设置选框，确定xls转xlsx或反向操作，转换方式
        self.xls_ = QRadioButton('xls转xlsx', self.centralwidget)
        self.xls_.move(50, 240)
        self.xls_.setChecked(True)
        self.xlsx_ = QRadioButton('xlsx转xls', self.centralwidget)
        self.xlsx_.move(200, 240)
        self.bt = QButtonGroup(self.centralwidget)
        self.bt.addButton(self.xls_, 1)
        self.bt.addButton(self.xlsx_, 2)

        self.bt.setId(self.xls_, 1)
        self.bt.setId(self.xlsx_, 2)  # 设置默认值，表示选择了该项之后返回的CheckId值

        self.bt.buttonClicked.connect(self.pd)

        # 设置开始运行按钮
        self.startrun = QtWidgets.QPushButton(self.centralwidget)
        self.startrun.setGeometry(QtCore.QRect(50, 280, 100, 35))
        self.startrun.setText("开始运行")

        # 显示运行信息
        self.yxxx = QtWidgets.QLabel(self.centralwidget)
        self.yxxx.setGeometry(QtCore.QRect(50, 340, 400, 400))
        self.yxxx.setStyleSheet("color:red;font-size:20px")
        self.yxxx.setText("")

        ################button按钮点击事件回调函数################
        self.savepath.clicked.connect(self.save_Excel)
        self.openfile.clicked.connect(self.reader_excel)
        self.startrun.clicked.connect(self.start)

    @staticmethod
    def reader_excel():
        m = QtWidgets.QFileDialog.getExistingDirectory(
            None, "选取文件夹", "D:/test/待合并表格/")  # 起始路径
        Ui_MainWindow.filepath = m
        print(m)

    @staticmethod
    def save_Excel():
        m = QtWidgets.QFileDialog.getExistingDirectory(
            None, "选取文件夹", "D:/test/合并结果/")  # 起始路径
        Ui_MainWindow.path_save = m
        print(m)

    def start(self):
        try:
            print(self.panduan)
            for file in os.listdir(self.filepath):
                self.message_put(file)
                if self.panduan == 1:
                    if file.endswith('.xls'):
                        xls_2_xlsx(self.filepath + '/' + file)
                    else:
                        self.message_put('该文件不是指定格式的文件' + file)
                elif self.panduan == 2:
                    if file.endswith('.xlsx'):
                        xlsx_2_xls(self.filepath + '/' + file)
                    else:
                        self.message_put('该文件不是指定格式的文件' + file)
                else:
                    self.message_put('未选择转换方式')
                    pass
        except Exception as e:
            print(e)
            self.message = str(e)
        self.message_put('运行结束')

    def message_put(self, aa: 'str'):
        QtWidgets.QApplication.processEvents()
        self.yxxx.setText(self.message + aa)

    def pd(self):
        print(self.bt.checkedId())
        if self.bt.checkedId() == 2:
            self.panduan = 2
        else:
            self.panduan = 1


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    mainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(mainWindow)
    mainWindow.show()
    sys.exit(app.exec_())
