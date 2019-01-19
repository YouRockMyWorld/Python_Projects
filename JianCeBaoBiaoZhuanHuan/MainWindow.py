from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from PyQt5.QtGui import QTextCursor
import sys, os
from converter import *
import json
import MyThread


class MyWindow(QMainWindow, Ui_MainWindow):

    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.__excel_path = ''
        self.__excel_out_path = ''
        self.__conf_path = ''

        self.initUI()
        self.initConf()

    def initUI(self):
        #self.setFixedSize(500, 500)
        self.pushButton_excel.clicked.connect(self.pushButton_excel_Clicked)
        self.pushButton_excel_out.clicked.connect(self.pushButton_excel_out_Clicked)
        self.pushButton_converte.clicked.connect(self.pushButton_converte_Clicked)
        self.pushButton_conf.clicked.connect(self.pushButton_conf_Clicked)
        self.lineEdit_excel_out.editingFinished.connect(self.lineEdit_excel_out_editingFinished)
        self.show()


    def initConf(self):
        if os.path.exists('./conf.sc') and os.path.isfile('./conf.sc'):
            with open('./conf.sc') as f:
                content = f.read()
            jsonfile = json.loads(content)
            if os.path.exists(jsonfile['conf_path']) and os.path.isfile(jsonfile['conf_path']):
                self.__conf_path = jsonfile['conf_path']
                self.lineEdit_conf.setText(str(self.__conf_path))
                self.lineEdit_conf.setDisabled(True)


    def pushButton_excel_Clicked(self):
        path = QFileDialog.getOpenFileName(self, '打开报表', '.', 'Excel files (*.xlsx)')
        if path[0] != '':
            self.lineEdit_excel.setText(str(path[0]))
            self.__excel_path = path[0]
            self.lineEdit_excel.setDisabled(True)


    def pushButton_conf_Clicked(self):
        path = QFileDialog.getOpenFileName(self, '选择配置文件', '.', '配置文件 (*.json)')
        if path[0] != '':
            self.lineEdit_conf.setText(str(path[0]))
            self.__conf_path = path[0]
            self.lineEdit_conf.setDisabled(True)


    def pushButton_excel_out_Clicked(self):
        path = QFileDialog.getSaveFileName(self, '输出', '.', 'Excel files (*.xlsx);;Excel 97-03 files (*.xls)')
        if path[0] != '':
            self.lineEdit_excel_out.setText(path[0])
            self.__excel_out_path = path[0]
            # self.lineEdit_excel_out.setDisabled(True)


    # def pushButton_converte_Clicked(self):
    #     try:
    #         if self.__excel_path and self.__excel_out_path:
    #             with open(r'E:\Projects\Python_Projects\JianCePingTai\conf_map_platform_new.json', encoding='utf-8') as f:
    #                 jsonstr = f.read()
    #             dataconf = json.loads(jsonstr)
    #             self.label_out_info.setText('开始转换...')
    #             self.pushButton_converte.setDisabled(True)
    #             info = ''
    #             if self.__excel_out_path[-1] == 'x':
    #                 info = ProcessXlsx(dataconf, self.__excel_path, self.__excel_out_path)
    #             elif self.__excel_out_path[-1] == 's':
    #                 info = ProcessXls(dataconf, self.__excel_path, self.__excel_out_path)
    #             self.textBrowser_out.insertPlainText(info)
    #             self.label_out_info.setText('转换完成！')
    #             self.pushButton_converte.setDisabled(False)
    #     except Exception as e:
    #         self.textBrowser_out.insertPlainText(str(e))
    #         self.label_out_info.setText('转换失败！')
    #         self.pushButton_converte.setDisabled(False)

    def pushButton_converte_Clicked(self):
        try:
            if self.__excel_path and self.__excel_out_path and self.__conf_path:
                with open(self.__conf_path, encoding='utf-8') as f:
                    jsonstr = f.read()
                dataconf = json.loads(jsonstr)

                self.thread = MyThread.ExcelThread()
                self.thread.signalOut.connect(self.setInfoText)
                self.thread.started.connect(self.convert_start)
                self.thread.finished.connect(self.convert_finished)

                self.thread.readPath = self.__excel_path
                self.thread.writePath = self.__excel_out_path
                self.thread.conf = dataconf

                self.thread.start()

            else:
                QMessageBox.warning(self, '警告', '未能正确配置！', QMessageBox.Ok)

        except Exception as e:
            self.textBrowser_out.insertPlainText(str(e))
            self.label_out_info.setText('转换失败！')
            self.pushButton_converte.setDisabled(False)


    def setInfoText(self, text):
        self.textBrowser_out.insertPlainText(text)
        textCursor = self.textBrowser_out.textCursor()
        textCursor.movePosition(QTextCursor.End)
        self.textBrowser_out.setTextCursor(textCursor)


    def convert_start(self):
        self.label_out_info.setText('开始转换...')
        self.pushButton_converte.setDisabled(True)


    def convert_finished(self):
        self.label_out_info.setText('转换完成！')
        self.pushButton_converte.setDisabled(False)


    def lineEdit_excel_out_editingFinished(self):
        path = self.lineEdit_excel_out.text()
        print(path)
        if (path.endswith('.xls') and len(path) > 4) or (path.endswith('.xlsx') and len(path) > 5):
            self.__excel_out_path = path


    def closeEvent(self, event):
        # reply = QtWidgets.QMessageBox.question(self,
        #                                        '本程序',
        #                                        "是否要退出程序？",
        #                                        QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
        #                                        QtWidgets.QMessageBox.No)
        # if reply == QtWidgets.QMessageBox.Yes:
        #     event.accept()
        # else:
        #     event.ignore()
        if self.__conf_path != '':
            conf = {}
            conf['conf_path'] = self.__conf_path
            with open('./conf.sc', 'w') as f:
                f.write(json.dumps(conf))


if __name__ == '__main__':
    app = QApplication(sys.argv)
    m_window = MyWindow()
    sys.exit(app.exec_())

