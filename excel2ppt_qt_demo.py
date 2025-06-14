# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'excel2ppt_qt.ui'
#
# Created by: PyQt5 UI code generator 5.15.9
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from readExcelClass import ReadExcelClass
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLineEdit, QVBoxLayout, QWidget, QFileDialog, QHeaderView
from PyQt5.QtGui import QColor, QBrush
from exception_handler import exception_signal_instance
import pandas as pd
from searchClass import SearchClass
from bindClass import BindClass


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setEnabled(True)
        MainWindow.resize(1000, 857)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(MainWindow.sizePolicy().hasHeightForWidth())
        MainWindow.setSizePolicy(sizePolicy)
        MainWindow.setMinimumSize(QtCore.QSize(400, 400))
        MainWindow.setMaximumSize(QtCore.QSize(1000, 1000))
        font = QtGui.QFont()
        font.setStyleStrategy(QtGui.QFont.PreferAntialias)
        MainWindow.setFont(font)
        MainWindow.setAcceptDrops(False)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/icon/10692.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)
        MainWindow.setAutoFillBackground(False)
        MainWindow.setStyleSheet("background-image: url(:/background/white-concrete-wall.jpg);\n"
"background-color: rgba(0,0,0,0);")
        MainWindow.setAnimated(False)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.centralwidget.sizePolicy().hasHeightForWidth())
        self.centralwidget.setSizePolicy(sizePolicy)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout_7 = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout_7.setObjectName("gridLayout_7")
        self.tabWidget = QtWidgets.QTabWidget(self.centralwidget)
        self.tabWidget.setTabPosition(QtWidgets.QTabWidget.North)
        self.tabWidget.setObjectName("tabWidget")
        self.tab_2 = QtWidgets.QWidget()
        self.tab_2.setObjectName("tab_2")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.tab_2)
        self.verticalLayout.setObjectName("verticalLayout")
        spacerItem = QtWidgets.QSpacerItem(168, 16, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout.addItem(spacerItem)
        self.gridLayout_3 = QtWidgets.QGridLayout()
        self.gridLayout_3.setObjectName("gridLayout_3")
        self.lineEdit_3 = QtWidgets.QLineEdit(self.tab_2)
        font = QtGui.QFont()
        font.setFamily("Audi Type")
        font.setBold(True)
        font.setWeight(75)
        self.lineEdit_3.setFont(font)
        self.lineEdit_3.setStyleSheet("QLineEdit {\n"
"    border: 2px solid #808080;\n"
"    border-radius: 5px;\n"
"    padding: 5px;\n"
"    background-color: #f0f0f0;\n"
"    selection-color: white;\n"
"    selection-background-color:#5CACEE;\n"
"}")
        self.lineEdit_3.setReadOnly(True)
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.gridLayout_3.addWidget(self.lineEdit_3, 0, 0, 1, 1)
        self.pushButton_4 = QtWidgets.QPushButton(self.tab_2)
        font = QtGui.QFont()
        font.setFamily("Audi Type")
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_4.setFont(font)
        self.pushButton_4.setAutoFillBackground(False)
        self.pushButton_4.setStyleSheet("\n"
"QPushButton {\n"
"    background-color: #FF0000 !important;\n"
"    border: 2px solid #FF0000;\n"
"    color: black;\n"
"    padding: 10px 20px;\n"
"    border-radius: 5px;\n"
"    font-weight: bold;\n"
"}\n"
"\n"
"QPushButton:hover {\n"
"    background-color: #FF3333 !important;\n"
"    border-color: #FF3333;\n"
"}\n"
"\n"
"QPushButton:pressed {\n"
"    background-color: #CC0000 !important;\n"
"    border-color: #CC0000;\n"
"}")
        self.pushButton_4.setAutoDefault(False)
        self.pushButton_4.setDefault(False)
        self.pushButton_4.setFlat(False)
        self.pushButton_4.setObjectName("pushButton_4")
        self.gridLayout_3.addWidget(self.pushButton_4, 0, 1, 1, 1)
        self.verticalLayout.addLayout(self.gridLayout_3)
        spacerItem1 = QtWidgets.QSpacerItem(168, 17, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout.addItem(spacerItem1)
        self.gridLayout_4 = QtWidgets.QGridLayout()
        self.gridLayout_4.setObjectName("gridLayout_4")
        self.lineEdit_4 = QtWidgets.QLineEdit(self.tab_2)
        font = QtGui.QFont()
        font.setFamily("Audi Type")
        font.setBold(True)
        font.setWeight(75)
        self.lineEdit_4.setFont(font)
        self.lineEdit_4.setStyleSheet("QLineEdit {\n"
"    border: 2px solid #808080;\n"
"    border-radius: 5px;\n"
"    padding: 5px;\n"
"    background-color: #f0f0f0;\n"
"    selection-color: white;\n"
"    selection-background-color:#5CACEE;\n"
"}")
        self.lineEdit_4.setReadOnly(True)
        self.lineEdit_4.setObjectName("lineEdit_4")
        self.gridLayout_4.addWidget(self.lineEdit_4, 0, 0, 1, 1)
        self.pushButton_5 = QtWidgets.QPushButton(self.tab_2)
        font = QtGui.QFont()
        font.setFamily("Audi Type")
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_5.setFont(font)
        self.pushButton_5.setAutoFillBackground(False)
        self.pushButton_5.setStyleSheet("\n"
"QPushButton {\n"
"    background-color: #FF0000 !important;\n"
"    border: 2px solid #FF0000;\n"
"    color: black;\n"
"    padding: 10px 20px;\n"
"    border-radius: 5px;\n"
"    font-weight: bold;\n"
"}\n"
"\n"
"QPushButton:hover {\n"
"    background-color: #FF3333 !important;\n"
"    border-color: #FF3333;\n"
"}\n"
"\n"
"QPushButton:pressed {\n"
"    background-color: #CC0000 !important;\n"
"    border-color: #CC0000;\n"
"}")
        self.pushButton_5.setAutoDefault(False)
        self.pushButton_5.setDefault(False)
        self.pushButton_5.setFlat(False)
        self.pushButton_5.setObjectName("pushButton_5")
        self.gridLayout_4.addWidget(self.pushButton_5, 0, 1, 1, 1)
        self.verticalLayout.addLayout(self.gridLayout_4)
        spacerItem2 = QtWidgets.QSpacerItem(168, 16, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout.addItem(spacerItem2)
        self.gridLayout_5 = QtWidgets.QGridLayout()
        self.gridLayout_5.setObjectName("gridLayout_5")
        self.lineEdit_5 = QtWidgets.QLineEdit(self.tab_2)
        font = QtGui.QFont()
        font.setFamily("Audi Type")
        font.setBold(True)
        font.setWeight(75)
        self.lineEdit_5.setFont(font)
        self.lineEdit_5.setStyleSheet("QLineEdit {\n"
"    border: 2px solid #808080;\n"
"    border-radius: 5px;\n"
"    padding: 5px;\n"
"    background-color: #f0f0f0;\n"
"    selection-color: white;\n"
"    selection-background-color:#5CACEE;\n"
"}")
        self.lineEdit_5.setReadOnly(True)
        self.lineEdit_5.setObjectName("lineEdit_5")
        self.gridLayout_5.addWidget(self.lineEdit_5, 0, 0, 1, 1)
        self.pushButton_6 = QtWidgets.QPushButton(self.tab_2)
        font = QtGui.QFont()
        font.setFamily("Audi Type")
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_6.setFont(font)
        self.pushButton_6.setAutoFillBackground(False)
        self.pushButton_6.setStyleSheet("\n"
"QPushButton {\n"
"    background-color: #FF0000 !important;\n"
"    border: 2px solid #FF0000;\n"
"    color: black;\n"
"    padding: 10px 20px;\n"
"    border-radius: 5px;\n"
"    font-weight: bold;\n"
"}\n"
"\n"
"QPushButton:hover {\n"
"    background-color: #FF3333 !important;\n"
"    border-color: #FF3333;\n"
"}\n"
"\n"
"QPushButton:pressed {\n"
"    background-color: #CC0000 !important;\n"
"    border-color: #CC0000;\n"
"}")
        self.pushButton_6.setAutoDefault(False)
        self.pushButton_6.setDefault(False)
        self.pushButton_6.setFlat(False)
        self.pushButton_6.setObjectName("pushButton_6")
        self.gridLayout_5.addWidget(self.pushButton_6, 0, 1, 1, 1)
        self.verticalLayout.addLayout(self.gridLayout_5)
        spacerItem3 = QtWidgets.QSpacerItem(168, 17, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout.addItem(spacerItem3)
        self.gridLayout_6 = QtWidgets.QGridLayout()
        self.gridLayout_6.setObjectName("gridLayout_6")
        self.lineEdit_6 = QtWidgets.QLineEdit(self.tab_2)
        font = QtGui.QFont()
        font.setFamily("Audi Type")
        font.setBold(True)
        font.setWeight(75)
        self.lineEdit_6.setFont(font)
        self.lineEdit_6.setStyleSheet("QLineEdit {\n"
"    border: 2px solid #808080;\n"
"    border-radius: 5px;\n"
"    padding: 5px;\n"
"    background-color: #f0f0f0;\n"
"    selection-color: white;\n"
"    selection-background-color:#5CACEE;\n"
"}")
        self.lineEdit_6.setReadOnly(True)
        self.lineEdit_6.setObjectName("lineEdit_6")
        self.gridLayout_6.addWidget(self.lineEdit_6, 0, 0, 1, 1)
        self.pushButton_8 = QtWidgets.QPushButton(self.tab_2)
        font = QtGui.QFont()
        font.setFamily("Audi Type")
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_8.setFont(font)
        self.pushButton_8.setAutoFillBackground(False)
        self.pushButton_8.setStyleSheet("\n"
"QPushButton {\n"
"    background-color: #FF0000 !important;\n"
"    border: 2px solid #FF0000;\n"
"    color: black;\n"
"    padding: 10px 20px;\n"
"    border-radius: 5px;\n"
"    font-weight: bold;\n"
"}\n"
"\n"
"QPushButton:hover {\n"
"    background-color: #FF3333 !important;\n"
"    border-color: #FF3333;\n"
"}\n"
"\n"
"QPushButton:pressed {\n"
"    background-color: #CC0000 !important;\n"
"    border-color: #CC0000;\n"
"}")
        self.pushButton_8.setAutoDefault(False)
        self.pushButton_8.setDefault(False)
        self.pushButton_8.setFlat(False)
        self.pushButton_8.setObjectName("pushButton_8")
        self.gridLayout_6.addWidget(self.pushButton_8, 0, 1, 1, 1)
        self.verticalLayout.addLayout(self.gridLayout_6)
        spacerItem4 = QtWidgets.QSpacerItem(168, 44, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout.addItem(spacerItem4)
        self.gridLayout_8 = QtWidgets.QGridLayout()
        self.gridLayout_8.setObjectName("gridLayout_8")
        self.lineEdit_7 = QtWidgets.QLineEdit(self.tab_2)
        font = QtGui.QFont()
        font.setFamily("Audi Type")
        font.setBold(True)
        font.setWeight(75)
        self.lineEdit_7.setFont(font)
        self.lineEdit_7.setStyleSheet("QLineEdit {\n"
"    border: 2px solid #808080;\n"
"    border-radius: 5px;\n"
"    padding: 5px;\n"
"    background-color: #f0f0f0;\n"
"    selection-color: white;\n"
"    selection-background-color:#5CACEE;\n"
"}")
        self.lineEdit_7.setReadOnly(True)
        self.lineEdit_7.setObjectName("lineEdit_7")
        self.gridLayout_8.addWidget(self.lineEdit_7, 0, 0, 1, 1)
        self.pushButton_9 = QtWidgets.QPushButton(self.tab_2)
        font = QtGui.QFont()
        font.setFamily("Audi Type")
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_9.setFont(font)
        self.pushButton_9.setAutoFillBackground(False)
        self.pushButton_9.setStyleSheet("\n"
"QPushButton {\n"
"    background-color: #FF0000 !important;\n"
"    border: 2px solid #FF0000;\n"
"    color: black;\n"
"    padding: 10px 20px;\n"
"    border-radius: 5px;\n"
"    font-weight: bold;\n"
"}\n"
"\n"
"QPushButton:hover {\n"
"    background-color: #FF3333 !important;\n"
"    border-color: #FF3333;\n"
"}\n"
"\n"
"QPushButton:pressed {\n"
"    background-color: #CC0000 !important;\n"
"    border-color: #CC0000;\n"
"}")
        self.pushButton_9.setAutoDefault(False)
        self.pushButton_9.setDefault(False)
        self.pushButton_9.setFlat(False)
        self.pushButton_9.setObjectName("pushButton_9")
        self.gridLayout_8.addWidget(self.pushButton_9, 0, 1, 1, 1)
        self.verticalLayout.addLayout(self.gridLayout_8)
        spacerItem5 = QtWidgets.QSpacerItem(168, 44, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout.addItem(spacerItem5)
        self.gridLayout_9 = QtWidgets.QGridLayout()
        self.gridLayout_9.setObjectName("gridLayout_9")
        self.lineEdit_8 = QtWidgets.QLineEdit(self.tab_2)
        font = QtGui.QFont()
        font.setFamily("Audi Type")
        font.setBold(True)
        font.setWeight(75)
        self.lineEdit_8.setFont(font)
        self.lineEdit_8.setStyleSheet("QLineEdit {\n"
"    border: 2px solid #808080;\n"
"    border-radius: 5px;\n"
"    padding: 5px;\n"
"    background-color: #f0f0f0;\n"
"    selection-color: white;\n"
"    selection-background-color:#5CACEE;\n"
"}")
        self.lineEdit_8.setReadOnly(True)
        self.lineEdit_8.setObjectName("lineEdit_8")
        self.gridLayout_9.addWidget(self.lineEdit_8, 0, 0, 1, 1)
        self.pushButton_10 = QtWidgets.QPushButton(self.tab_2)
        font = QtGui.QFont()
        font.setFamily("Audi Type")
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_10.setFont(font)
        self.pushButton_10.setAutoFillBackground(False)
        self.pushButton_10.setStyleSheet("\n"
"QPushButton {\n"
"    background-color: #FF0000 !important;\n"
"    border: 2px solid #FF0000;\n"
"    color: black;\n"
"    padding: 10px 20px;\n"
"    border-radius: 5px;\n"
"    font-weight: bold;\n"
"}\n"
"\n"
"QPushButton:hover {\n"
"    background-color: #FF3333 !important;\n"
"    border-color: #FF3333;\n"
"}\n"
"\n"
"QPushButton:pressed {\n"
"    background-color: #CC0000 !important;\n"
"    border-color: #CC0000;\n"
"}")
        self.pushButton_10.setAutoDefault(False)
        self.pushButton_10.setDefault(False)
        self.pushButton_10.setFlat(False)
        self.pushButton_10.setObjectName("pushButton_10")
        self.gridLayout_9.addWidget(self.pushButton_10, 0, 1, 1, 1)
        self.verticalLayout.addLayout(self.gridLayout_9)
        spacerItem6 = QtWidgets.QSpacerItem(168, 44, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout.addItem(spacerItem6)
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        spacerItem7 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_4.addItem(spacerItem7)
        self.pushButton_7 = QtWidgets.QPushButton(self.tab_2)
        font = QtGui.QFont()
        font.setFamily("Audi Type")
        font.setPointSize(11)
        font.setBold(True)
        font.setItalic(False)
        font.setUnderline(False)
        font.setWeight(75)
        font.setStrikeOut(False)
        font.setKerning(True)
        font.setStyleStrategy(QtGui.QFont.PreferDefault)
        self.pushButton_7.setFont(font)
        self.pushButton_7.setStyleSheet("\n"
"QPushButton {\n"
"    background-color: #FFA500;\n"
"    border: 2px solid #FFA500;\n"
"    color: black;\n"
"    padding: 10px 20px;\n"
"    border-radius: 5px;\n"
"}\n"
"\n"
"QPushButton:hover {\n"
"    background-color: #FFC04D;\n"
"    border-color: #FFC04D;\n"
"}\n"
"\n"
"QPushButton:pressed {\n"
"    background-color: #FF8500;\n"
"    border-color: #FF8500;\n"
"}")
        self.pushButton_7.setObjectName("pushButton_7")
        self.horizontalLayout_4.addWidget(self.pushButton_7)
        spacerItem8 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_4.addItem(spacerItem8)
        self.verticalLayout.addLayout(self.horizontalLayout_4)
        self.tabWidget.addTab(self.tab_2, "")
        self.tab = QtWidgets.QWidget()
        self.tab.setObjectName("tab")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.tab)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.gridLayout = QtWidgets.QGridLayout()
        self.gridLayout.setObjectName("gridLayout")
        self.pushButton_2 = QtWidgets.QPushButton(self.tab)
        font = QtGui.QFont()
        font.setFamily("Audi Type")
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setAutoFillBackground(False)
        self.pushButton_2.setStyleSheet("\n"
"QPushButton {\n"
"    background-color: #FF0000 !important;\n"
"    border: 2px solid #FF0000;\n"
"    color: black;\n"
"    padding: 10px 20px;\n"
"    border-radius: 5px;\n"
"    font-weight: bold;\n"
"}\n"
"\n"
"QPushButton:hover {\n"
"    background-color: #FF3333 !important;\n"
"    border-color: #FF3333;\n"
"}\n"
"\n"
"QPushButton:pressed {\n"
"    background-color: #CC0000 !important;\n"
"    border-color: #CC0000;\n"
"}")
        self.pushButton_2.setAutoDefault(False)
        self.pushButton_2.setDefault(False)
        self.pushButton_2.setFlat(False)
        self.pushButton_2.setObjectName("pushButton_2")
        self.gridLayout.addWidget(self.pushButton_2, 0, 2, 1, 1)
        self.lineEdit_2 = QtWidgets.QLineEdit(self.tab)
        font = QtGui.QFont()
        font.setFamily("Audi Type")
        font.setBold(True)
        font.setWeight(75)
        self.lineEdit_2.setFont(font)
        self.lineEdit_2.setStyleSheet("QLineEdit {\n"
"    border: 2px solid #808080;\n"
"    border-radius: 5px;\n"
"    padding: 5px;\n"
"    background-color: #f0f0f0;\n"
"    selection-color: white;\n"
"    selection-background-color:#5CACEE;\n"
"}")
        self.lineEdit_2.setReadOnly(True)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.gridLayout.addWidget(self.lineEdit_2, 0, 1, 1, 1)
        self.verticalLayout_2.addLayout(self.gridLayout)
        self.gridLayout_10 = QtWidgets.QGridLayout()
        self.gridLayout_10.setObjectName("gridLayout_10")
        self.pushButton_11 = QtWidgets.QPushButton(self.tab)
        font = QtGui.QFont()
        font.setFamily("Audi Type")
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_11.setFont(font)
        self.pushButton_11.setAutoFillBackground(False)
        self.pushButton_11.setStyleSheet("\n"
"QPushButton {\n"
"    background-color: #FF0000 !important;\n"
"    border: 2px solid #FF0000;\n"
"    color: black;\n"
"    padding: 10px 20px;\n"
"    border-radius: 5px;\n"
"    font-weight: bold;\n"
"}\n"
"\n"
"QPushButton:hover {\n"
"    background-color: #FF3333 !important;\n"
"    border-color: #FF3333;\n"
"}\n"
"\n"
"QPushButton:pressed {\n"
"    background-color: #CC0000 !important;\n"
"    border-color: #CC0000;\n"
"}")
        self.pushButton_11.setAutoDefault(False)
        self.pushButton_11.setDefault(False)
        self.pushButton_11.setFlat(False)
        self.pushButton_11.setObjectName("pushButton_11")
        self.gridLayout_10.addWidget(self.pushButton_11, 0, 2, 1, 1)
        self.lineEdit_10 = QtWidgets.QLineEdit(self.tab)
        font = QtGui.QFont()
        font.setFamily("Audi Type")
        font.setBold(True)
        font.setWeight(75)
        self.lineEdit_10.setFont(font)
        self.lineEdit_10.setStyleSheet("QLineEdit {\n"
"    border: 2px solid #808080;\n"
"    border-radius: 5px;\n"
"    padding: 5px;\n"
"    background-color: #f0f0f0;\n"
"    selection-color: white;\n"
"    selection-background-color:#5CACEE;\n"
"}")
        self.lineEdit_10.setReadOnly(False)
        self.lineEdit_10.setObjectName("lineEdit_10")
        self.gridLayout_10.addWidget(self.lineEdit_10, 0, 1, 1, 1)
        self.verticalLayout_2.addLayout(self.gridLayout_10)
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.tableWidget = QtWidgets.QTableWidget(self.tab)
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(6)
        self.tableWidget.setRowCount(31)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(4, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(5, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(6, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(7, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(8, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(9, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(10, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(11, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(12, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(13, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(14, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(15, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(16, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(17, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(18, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(19, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(20, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(21, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(22, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(23, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(24, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(25, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(26, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(27, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(28, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(29, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(30, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(4, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(5, item)
        self.horizontalLayout_3.addWidget(self.tableWidget)
        self.verticalLayout_2.addLayout(self.horizontalLayout_3)
        self.tabWidget.addTab(self.tab, "")
        self.gridLayout_7.addWidget(self.tabWidget, 1, 1, 1, 1)
        self.label = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Audi Type")
        font.setPointSize(15)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setObjectName("label")
        self.gridLayout_7.addWidget(self.label, 0, 1, 1, 1)
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setEnabled(False)
        font = QtGui.QFont()
        font.setFamily("Audi Type")
        font.setPointSize(10)
        font.setBold(True)
        font.setItalic(False)
        font.setWeight(75)
        self.statusbar.setFont(font)
        self.statusbar.setMouseTracking(False)
        self.statusbar.setTabletTracking(False)
        self.statusbar.setAcceptDrops(True)
        self.statusbar.setStyleSheet("color: red;")
        self.statusbar.setSizeGripEnabled(True)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        self.tabWidget.setCurrentIndex(1)
        self.pushButton_4.clicked.connect(self.log_record_clicked) # type: ignore
        self.pushButton_5.clicked.connect(self.employee_list_clicked) # type: ignore
        self.pushButton_6.clicked.connect(self.leave_list_clicked) # type: ignore
        self.pushButton_8.clicked.connect(self.business_trip_list_clicked) # type: ignore
        self.pushButton_9.clicked.connect(self.tax2gti_on_button_clicked) # type: ignore
        self.pushButton_10.clicked.connect(self.card_problem_list_clicked) # type: ignore
        self.pushButton_2.clicked.connect(self.gti_file_button_clicked) # type: ignore
        self.pushButton_11.clicked.connect(self.enter_clicked) # type: ignore
        exception_signal_instance.signal.connect(self.update_status_bar)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Time Attendance Management Tool"))
        self.pushButton_4.setText(_translate("MainWindow", "Choose Log record File"))
        self.pushButton_5.setText(_translate("MainWindow", "Choose Employee list"))
        self.pushButton_6.setText(_translate("MainWindow", "Choose leave list"))
        self.pushButton_8.setText(_translate("MainWindow", "Choose Business trip list"))
        self.pushButton_9.setText(_translate("MainWindow", "Choose Out of office list"))
        self.pushButton_10.setText(_translate("MainWindow", "Choose Card Problem list"))
        self.pushButton_7.setText(_translate("MainWindow", "Generate"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_2), _translate("MainWindow", "Choose Files to generate result"))
        self.pushButton_2.setText(_translate("MainWindow", "Choose result file"))
        self.pushButton_11.setText(_translate("MainWindow", "Enter"))
        item = self.tableWidget.verticalHeaderItem(0)
        item.setText(_translate("MainWindow", "1"))
        item = self.tableWidget.verticalHeaderItem(1)
        item.setText(_translate("MainWindow", "2"))
        item = self.tableWidget.verticalHeaderItem(2)
        item.setText(_translate("MainWindow", "3"))
        item = self.tableWidget.verticalHeaderItem(3)
        item.setText(_translate("MainWindow", "4"))
        item = self.tableWidget.verticalHeaderItem(4)
        item.setText(_translate("MainWindow", "5"))
        item = self.tableWidget.verticalHeaderItem(5)
        item.setText(_translate("MainWindow", "6"))
        item = self.tableWidget.verticalHeaderItem(6)
        item.setText(_translate("MainWindow", "7"))
        item = self.tableWidget.verticalHeaderItem(7)
        item.setText(_translate("MainWindow", "8"))
        item = self.tableWidget.verticalHeaderItem(8)
        item.setText(_translate("MainWindow", "9"))
        item = self.tableWidget.verticalHeaderItem(9)
        item.setText(_translate("MainWindow", "10"))
        item = self.tableWidget.verticalHeaderItem(10)
        item.setText(_translate("MainWindow", "11"))
        item = self.tableWidget.verticalHeaderItem(11)
        item.setText(_translate("MainWindow", "12"))
        item = self.tableWidget.verticalHeaderItem(12)
        item.setText(_translate("MainWindow", "13"))
        item = self.tableWidget.verticalHeaderItem(13)
        item.setText(_translate("MainWindow", "14"))
        item = self.tableWidget.verticalHeaderItem(14)
        item.setText(_translate("MainWindow", "15"))
        item = self.tableWidget.verticalHeaderItem(15)
        item.setText(_translate("MainWindow", "16"))
        item = self.tableWidget.verticalHeaderItem(16)
        item.setText(_translate("MainWindow", "17"))
        item = self.tableWidget.verticalHeaderItem(17)
        item.setText(_translate("MainWindow", "18"))
        item = self.tableWidget.verticalHeaderItem(18)
        item.setText(_translate("MainWindow", "19"))
        item = self.tableWidget.verticalHeaderItem(19)
        item.setText(_translate("MainWindow", "20"))
        item = self.tableWidget.verticalHeaderItem(20)
        item.setText(_translate("MainWindow", "21"))
        item = self.tableWidget.verticalHeaderItem(21)
        item.setText(_translate("MainWindow", "22"))
        item = self.tableWidget.verticalHeaderItem(22)
        item.setText(_translate("MainWindow", "23"))
        item = self.tableWidget.verticalHeaderItem(23)
        item.setText(_translate("MainWindow", "24"))
        item = self.tableWidget.verticalHeaderItem(24)
        item.setText(_translate("MainWindow", "25"))
        item = self.tableWidget.verticalHeaderItem(25)
        item.setText(_translate("MainWindow", "26"))
        item = self.tableWidget.verticalHeaderItem(26)
        item.setText(_translate("MainWindow", "27"))
        item = self.tableWidget.verticalHeaderItem(27)
        item.setText(_translate("MainWindow", "28"))
        item = self.tableWidget.verticalHeaderItem(28)
        item.setText(_translate("MainWindow", "29"))
        item = self.tableWidget.verticalHeaderItem(29)
        item.setText(_translate("MainWindow", "30"))
        item = self.tableWidget.verticalHeaderItem(30)
        item.setText(_translate("MainWindow", "31"))
        item = self.tableWidget.horizontalHeaderItem(0)
        item.setText(_translate("MainWindow", "Staff ID"))
        item = self.tableWidget.horizontalHeaderItem(1)
        item.setText(_translate("MainWindow", "Name"))
        item = self.tableWidget.horizontalHeaderItem(2)
        item.setText(_translate("MainWindow", "Clock in"))
        item = self.tableWidget.horizontalHeaderItem(3)
        item.setText(_translate("MainWindow", "Clock out"))
        item = self.tableWidget.horizontalHeaderItem(4)
        item.setText(_translate("MainWindow", "Working Hour"))
        item = self.tableWidget.horizontalHeaderItem(5)
        item.setText(_translate("MainWindow", "Status"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), _translate("MainWindow", "employee filter"))
        self.label.setText(_translate("MainWindow", "Time Attendance Management Tool"))
        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tableWidget.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.tableWidget.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)

    def enter_clicked(self):
        staff_id = self.lineEdit_10.text()
        search_class = SearchClass()
        search_class.path = self.lineEdit_2.text().replace('/', '\\')
        search_class.find_search_result(staff_id)
        for i in range(31):
            self.tableWidget.setItem(i, 0, QtWidgets.QTableWidgetItem(f'{search_class.df.iloc[i, 0]}'))
            self.tableWidget.setItem(i, 1, QtWidgets.QTableWidgetItem(f'{search_class.df.iloc[i, 1]}'))
            self.tableWidget.setItem(i, 2, QtWidgets.QTableWidgetItem(f'{search_class.df.iloc[i, 2]}'))
            self.tableWidget.setItem(i, 3, QtWidgets.QTableWidgetItem(f'{search_class.df.iloc[i, 3]}'))
            self.tableWidget.setItem(i, 4, QtWidgets.QTableWidgetItem(f'{search_class.df.iloc[i, 4]}'))
            self.tableWidget.setItem(i, 5, QtWidgets.QTableWidgetItem(f'{search_class.df.iloc[i, 5]}'))
            if str(search_class.df.iloc[i, 5]) == "abnormal" or str(search_class.df.iloc[i, 5]) == "nan" or str(search_class.df.iloc[i, 5]) == "come late" or str(search_class.df.iloc[i, 5]) == "leave early" or str(search_class.df.iloc[i, 5]) == "clock in missing" or str(search_class.df.iloc[i, 5]) == "clock out missing":
                self.tableWidget.item(i, 5).setBackground(QBrush(QColor(255, 0, 0)))
            # else:
            #     self.tableWidget.setItem(i, 0, QtWidgets.QTableWidgetItem(''))
            #     self.tableWidget.setItem(i, 1, QtWidgets.QTableWidgetItem(''))
            #     self.tableWidget.setItem(i, 2, QtWidgets.QTableWidgetItem(''))
            #     self.tableWidget.setItem(i, 3, QtWidgets.QTableWidgetItem(''))
            #     self.tableWidget.setItem(i, 4, QtWidgets.QTableWidgetItem(''))
            #     self.tableWidget.setItem(i, 5, QtWidgets.QTableWidgetItem(''))


    def gti_file_button_clicked(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(QMainWindow(), "Select File", "",
                                                   "All Files (*);;Text Files (*.txt)",
                                                   options=options)
        if file_name:
            self.lineEdit_2.setText(file_name)
            self.statusbar.showMessage("GTI excel file path: " + file_name, 5000)

    def log_record_clicked(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(QMainWindow(), "Select File", "",
                                                   "All Files (*);;Text Files (*.txt)",
                                                   options=options)
        if file_name:
            self.lineEdit_3.setText(file_name)
            self.statusbar.showMessage("log record file path: " + file_name, 5000)

    def employee_list_clicked(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(QMainWindow(), "Select File", "",
                                                   "All Files (*);;Text Files (*.txt)",
                                                   options=options)
        if file_name:
            self.lineEdit_4.setText(file_name)
            self.statusbar.showMessage("employee list file path: " + file_name, 5000)

    def leave_list_clicked(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(QMainWindow(), "Select File", "",
                                                   "All Files (*);;Text Files (*.txt)",
                                                   options=options)
        if file_name:
            self.lineEdit_5.setText(file_name)
            self.statusbar.showMessage("leave list file path: " + file_name, 5000)

    def business_trip_list_clicked(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(QMainWindow(), "Select File", "",
                                                   "All Files (*);;Text Files (*.txt)",
                                                   options=options)
        if file_name:
            self.lineEdit_6.setText(file_name)
            self.statusbar.showMessage("business trip file path: " + file_name, 5000)

    def tax2gti_on_button_clicked(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(QMainWindow(), "Select File", "",
                                                   "All Files (*);;Text Files (*.txt)",
                                                   options=options)
        if file_name:
            self.lineEdit_7.setText(file_name)
            self.statusbar.showMessage("out of office list file path: " + file_name, 5000)

    def card_problem_list_clicked(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(QMainWindow(), "Select File", "",
                                                   "All Files (*);;Text Files (*.txt)",
                                                   options=options)
        if file_name:
            self.lineEdit_8.setText(file_name)
            self.statusbar.showMessage("card problem list file path: " + file_name, 5000)

    def update_status_bar(self, error_message):
        self.statusbar.showMessage(f'Error: {error_message}', 5000)
import pictures_rc
