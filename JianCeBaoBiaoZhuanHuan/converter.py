# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'converter.ui'
#
# Created by: PyQt5 UI code generator 5.8.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(469, 550)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(MainWindow.sizePolicy().hasHeightForWidth())
        MainWindow.setSizePolicy(sizePolicy)
        MainWindow.setMinimumSize(QtCore.QSize(450, 550))
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout.setObjectName("verticalLayout")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(False)
        font.setWeight(50)
        self.label_3.setFont(font)
        self.label_3.setAlignment(QtCore.Qt.AlignCenter)
        self.label_3.setObjectName("label_3")
        self.verticalLayout.addWidget(self.label_3)
        self.line_3 = QtWidgets.QFrame(self.centralwidget)
        self.line_3.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_3.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_3.setObjectName("line_3")
        self.verticalLayout.addWidget(self.line_3)
        self.horizontalLayout_5 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_5.setObjectName("horizontalLayout_5")
        self.radioButton_single = QtWidgets.QRadioButton(self.centralwidget)
        self.radioButton_single.setChecked(True)
        self.radioButton_single.setObjectName("radioButton_single")
        self.horizontalLayout_5.addWidget(self.radioButton_single, 0, QtCore.Qt.AlignHCenter)
        self.radioButton_multi = QtWidgets.QRadioButton(self.centralwidget)
        self.radioButton_multi.setObjectName("radioButton_multi")
        self.horizontalLayout_5.addWidget(self.radioButton_multi, 0, QtCore.Qt.AlignHCenter)
        self.verticalLayout.addLayout(self.horizontalLayout_5)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        self.lineEdit_excel = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_excel.setObjectName("lineEdit_excel")
        self.horizontalLayout.addWidget(self.lineEdit_excel)
        self.pushButton_excel = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_excel.setObjectName("pushButton_excel")
        self.horizontalLayout.addWidget(self.pushButton_excel)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setObjectName("label_2")
        self.horizontalLayout_2.addWidget(self.label_2)
        self.lineEdit_excel_out = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_excel_out.setObjectName("lineEdit_excel_out")
        self.horizontalLayout_2.addWidget(self.lineEdit_excel_out)
        self.pushButton_excel_out = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_excel_out.setObjectName("pushButton_excel_out")
        self.horizontalLayout_2.addWidget(self.pushButton_excel_out)
        self.verticalLayout.addLayout(self.horizontalLayout_2)
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setObjectName("label_4")
        self.horizontalLayout_4.addWidget(self.label_4)
        self.lineEdit_conf = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_conf.setObjectName("lineEdit_conf")
        self.horizontalLayout_4.addWidget(self.lineEdit_conf)
        self.pushButton_conf = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_conf.setObjectName("pushButton_conf")
        self.horizontalLayout_4.addWidget(self.pushButton_conf)
        self.verticalLayout.addLayout(self.horizontalLayout_4)
        self.line = QtWidgets.QFrame(self.centralwidget)
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.verticalLayout.addWidget(self.line)
        self.textBrowser_out = QtWidgets.QTextBrowser(self.centralwidget)
        self.textBrowser_out.setLineWrapMode(QtWidgets.QTextEdit.NoWrap)
        self.textBrowser_out.setObjectName("textBrowser_out")
        self.verticalLayout.addWidget(self.textBrowser_out)
        self.label_out_info = QtWidgets.QLabel(self.centralwidget)
        self.label_out_info.setText("")
        self.label_out_info.setAlignment(QtCore.Qt.AlignCenter)
        self.label_out_info.setObjectName("label_out_info")
        self.verticalLayout.addWidget(self.label_out_info)
        self.line_2 = QtWidgets.QFrame(self.centralwidget)
        self.line_2.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")
        self.verticalLayout.addWidget(self.line_2)
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        spacerItem = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_3.addItem(spacerItem)
        self.pushButton_converte = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_converte.setObjectName("pushButton_converte")
        self.horizontalLayout_3.addWidget(self.pushButton_converte)
        spacerItem1 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_3.addItem(spacerItem1)
        self.verticalLayout.addLayout(self.horizontalLayout_3)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 469, 23))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.label.setBuddy(self.lineEdit_excel)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "报表转换器"))
        self.label_3.setText(_translate("MainWindow", "<html><head/><body><p>报表转化</p><p>--- design by liu sc</p></body></html>"))
        self.radioButton_single.setText(_translate("MainWindow", "单一转换"))
        self.radioButton_multi.setText(_translate("MainWindow", "批量转换"))
        self.label.setText(_translate("MainWindow", "报表路径："))
        self.pushButton_excel.setText(_translate("MainWindow", "选择"))
        self.label_2.setText(_translate("MainWindow", "输出路径："))
        self.pushButton_excel_out.setText(_translate("MainWindow", "选择"))
        self.label_4.setText(_translate("MainWindow", "配置路径："))
        self.pushButton_conf.setText(_translate("MainWindow", "选择"))
        self.pushButton_converte.setText(_translate("MainWindow", "转换"))

