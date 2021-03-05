# -*- coding: utf-8 -*-
import os
import sys
import time

from PyQt5 import QtCore, QtWidgets
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, get_column_interval


class Ui_MainWindow(object):
    png_ = ''
    list_pdf_1 = []
    list_pdf_2 = []
    list_pdf_1_ = []
    list_pdf_2_ = []
    filepath_excel = 'D:/test/待合并表格'
    path_save = 'D:/test/合并结果'
    biaotou = 1
    message = ''

    def setupUi(self, MainWindow):
        # 主窗口参数设置
        MainWindow.setObjectName("MainWindow")
        MainWindow.setWindowTitle("拆分表格")
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        # 实例化
        MainWindow.setCentralWidget(self.centralwidget)
        MainWindow.setGeometry(QtCore.QRect(550, 250, 500, 500))

        self.fileExcel = QtWidgets.QPushButton(self.centralwidget)
        self.fileExcel.setGeometry(QtCore.QRect(50, 100, 333, 35))
        self.fileExcel.setText("选择要拆分的Excel所在的文件夹")

        self.saveExcel = QtWidgets.QPushButton(self.centralwidget)
        self.saveExcel.setGeometry(QtCore.QRect(50, 160, 333, 35))
        self.saveExcel.setText("选择保存位置")

        # 设置分类依据
        self.exceltitle_text = QtWidgets.QLabel(self.centralwidget)
        self.exceltitle_text.setGeometry(QtCore.QRect(50, 220, 290, 35))
        self.exceltitle_text.setText("根据该列内容进行拆分")
        self.exceltitle = QtWidgets.QLineEdit(self.centralwidget)
        self.exceltitle.setGeometry(QtCore.QRect(350, 220, 33, 35))
        self.exceltitle.setText("1")

        # 设置开始运行按钮
        self.startrun = QtWidgets.QPushButton(self.centralwidget)
        self.startrun.setGeometry(QtCore.QRect(50, 280, 100, 35))
        self.startrun.setText("开始运行")

        # 显示运行信息
        self.yxxx = QtWidgets.QLabel(self.centralwidget)
        self.yxxx.setGeometry(QtCore.QRect(50, 340, 400, 135))
        self.yxxx.setStyleSheet("color:red;font-size:20px")
        self.yxxx.setText("kaishi")

        ################button按钮点击事件回调函数################
        self.saveExcel.clicked.connect(self.save_Excel)
        self.fileExcel.clicked.connect(self.reader_excel)
        self.startrun.clicked.connect(self.start)

    def start(self):
        self.yxxx.setText('kaishi')
        for i in range(0, 10):
            time_(i)
            QtWidgets.QApplication.processEvents()
            self.yxxx.setText(str(i))
            print(self.message)

    def reader_excel(self):
        m = QtWidgets.QFileDialog.getExistingDirectory(
            None, "选取文件夹", "D:/test/待合并表格/")  # 起始路径
        Ui_MainWindow.filepath_excel = m
        print(m)

    def save_Excel(self):
        m = QtWidgets.QFileDialog.getExistingDirectory(
            None, "选取文件夹", "D:/test/合并结果/")  # 起始路径
        Ui_MainWindow.path_save = m
        print(m)


def time_(i):
    time.sleep(4)
    Ui_MainWindow.message = str(i)


if __name__ == '__main__':
    # excel('C:/Users/Gis04/Documents/ArcGIS/scratch/工作簿2.xlsx', 'D:/test/拆分结果/qq2', 1)
    app = QtWidgets.QApplication(sys.argv)
    mainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(mainWindow)
    mainWindow.show()
    sys.exit(app.exec_())
