# -*- coding: utf-8 -*-

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(500, 450)

        font = QtGui.QFont()
        font.setPointSize(8)

        #源文件地址
        self.sourceButton = QtWidgets.QPushButton(Form)
        self.sourceButton.setGeometry(QtCore.QRect(380, 20, 75, 23))
        self.sourceButton.setFont(font)
        self.sourceButton.setObjectName("pushButton")

        self.sourceValue = QtWidgets.QLineEdit(Form)
        self.sourceValue.setGeometry(QtCore.QRect(100, 20, 250, 20))
        self.sourceValue.setObjectName("sourceValue")

        self.sourceText = QtWidgets.QLabel(Form)
        self.sourceText.setFont(font)
        self.sourceText.setGeometry(QtCore.QRect(20, 20, 80, 23))

        # sheet
        self.sheetNameText = QtWidgets.QLabel(Form)
        self.sheetNameText.setFont(font)
        self.sheetNameText.setGeometry(QtCore.QRect(20, 50, 80, 23))

        self.sheetNameValue = QtWidgets.QLineEdit(Form)
        self.sheetNameValue.setGeometry(QtCore.QRect(100, 50, 250, 20))
        self.sheetNameValue.setObjectName("sheetNameValue")

        #起始行
        self.startLineText = QtWidgets.QLabel(Form)
        self.startLineText.setFont(font)
        self.startLineText.setGeometry(QtCore.QRect(20, 80, 80, 23))

        self.startLineValue = QtWidgets.QLineEdit(Form)
        self.startLineValue.setGeometry(QtCore.QRect(100, 80, 250, 20))
        self.startLineValue.setObjectName("startLineValue")

        # 起始列
        self.startColumnText = QtWidgets.QLabel(Form)
        self.startColumnText.setFont(font)
        self.startColumnText.setGeometry(QtCore.QRect(20, 110, 80, 23))

        self.startColumnValue = QtWidgets.QLineEdit(Form)
        self.startColumnValue.setGeometry(QtCore.QRect(100, 110, 250, 20))
        self.startColumnValue.setObjectName("startColumnValue")

        # 终止行
        self.endLineText = QtWidgets.QLabel(Form)
        self.endLineText.setFont(font)
        self.endLineText.setGeometry(QtCore.QRect(20, 140, 80, 23))

        self.endLineValue = QtWidgets.QLineEdit(Form)
        self.endLineValue.setGeometry(QtCore.QRect(100, 140, 250, 20))
        self.endLineValue.setObjectName("endLineValue")
        # 终止列
        self.endColumnText = QtWidgets.QLabel(Form)
        self.endColumnText.setFont(font)
        self.endColumnText.setGeometry(QtCore.QRect(20, 170, 80, 23))

        self.endColumnValue = QtWidgets.QLineEdit(Form)
        self.endColumnValue.setGeometry(QtCore.QRect(100, 170, 250, 20))
        self.endColumnValue.setObjectName("endColumnValue")
        # 文件生成地址
        self.resultFilePathText = QtWidgets.QLabel(Form)
        self.resultFilePathText.setFont(font)
        self.resultFilePathText.setGeometry(QtCore.QRect(20, 200, 80, 23))

        self.resultFilePathValue = QtWidgets.QLineEdit(Form)
        self.resultFilePathValue.setGeometry(QtCore.QRect(100, 200, 250, 20))
        self.resultFilePathValue.setObjectName("resultFilePathValue")

        self.resultFileButton = QtWidgets.QPushButton(Form)
        self.resultFileButton.setGeometry(QtCore.QRect(380, 200, 75, 23))
        self.resultFileButton.setFont(font)

        # 执行按钮
        self.runButton = QtWidgets.QPushButton(Form)
        self.runButton.setGeometry(QtCore.QRect(180, 235, 100, 40))

        #正常信息文本框
        self.resultText = QtWidgets.QTextBrowser(Form)
        self.resultText.setGeometry(QtCore.QRect(20, 280, 220, 150))

        #异常信息文本框
        self.exceptionText = QtWidgets.QTextBrowser(Form)
        self.exceptionText.setGeometry(QtCore.QRect(245, 280, 220, 150))

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.sourceButton.setText(_translate("Form", "选择文件夹"))
        self.sourceText.setText(_translate("Form", "源文件地址"))
        self.sheetNameText.setText(_translate("Form", "sheetName"))
        self.startLineText.setText(_translate("Form", "起始行"))
        self.startColumnText.setText(_translate("Form", "起始列"))
        self.endLineText.setText(_translate("Form", "终止行"))
        self.endColumnText.setText(_translate("Form", "终止列"))
        self.resultFilePathText.setText(_translate("Form", "生成文件地址"))
        self.resultFileButton.setText(_translate("Form", "选择文件夹"))
        self.runButton.setText(_translate("Form", "执行"))

        #事件
        self.sourceButton.clicked.connect(self.source_bt_event)
        self.resultFileButton.clicked.connect(self.result_bt_event)
        self.runButton.clicked.connect(self.deal_with)

    def source_bt_event(self):
        _translate = QtCore.QCoreApplication.translate
        directory1 = QFileDialog.getExistingDirectory(None, "选择文件夹", "./")
        self.sourceValue.setText(_translate("Form", directory1))

    def result_bt_event(self):
        _translate = QtCore.QCoreApplication.translate
        directory1 = QFileDialog.getExistingDirectory(None, "选择文件夹", "./")
        # print(directory1)  # 打印文件夹路径
        self.resultFilePathValue.setText(_translate("Form", directory1))

    def rest_text(self,text):
        try:
            cursor = self.resultText.textCursor()
            cursor.movePosition(QtGui.QTextCursor.End)
            cursor.insertText(str(text)+"\n")
            self.resultText.setTextCursor(cursor)
            self.resultText.ensureCursorVisible()
        except Exception as e:
            print('except:', e)

    def exception_text(self,text):
        try:
            cursor = self.exceptionText.textCursor()
            cursor.movePosition(QtGui.QTextCursor.End)
            cursor.insertText(str(text)+"\n")
            self.exceptionText.setTextCursor(cursor)
            self.exceptionText.ensureCursorVisible()
        except Exception as e:
            print('except:', e)