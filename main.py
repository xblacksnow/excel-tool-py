#import easygui as g
#coding=utf-8
import sys
import os
import xlrd
import numpy as np
from PyQt5.QtWidgets import QDialog
from PyQt5 import QtWidgets
from decimal import Decimal
from xlwt import Workbook
from form import Ui_Form

# def easygui():
#     fieldNames = ["文件路径","sheet名称", "起始行", "起始列", "终止行", "终止列","文件生成地址"]
#     fieldValues = g.multenterbox("表格统计", "excel处理工具-xxh", fieldNames)
#     while True:
#         if fieldValues == None:
#             break
#         errmsg = ""
#         for i in range(len(fieldNames)):
#             option = fieldNames[i].strip()
#             if fieldValues[i].strip() == "" and option[0] == "*":
#                 errmsg += ("【%s】为必填项   " % fieldNames[i])
#         if errmsg == "":
#             break
#         fieldValues = g.multenterbox(errmsg, title, fieldNames, fieldValues)
#     print("获取信息如下:%s" %str(fieldValues))
#     excel(fieldValues[0],fieldValues[1],fieldValues[2],fieldValues[3],fieldValues[4],fieldValues[5],fieldValues[6])

class LoginDlg(QDialog,Ui_Form):
    def __init__(self, parent=None):
        super(LoginDlg, self).__init__(parent)
        self.setupUi(self)
        self.setWindowTitle("表单处理")
        # self.resize(1800, 1000)


    def deal_with(self):
        try:
            self.excel()
        except Exception as e:
            self.exception_text(e)

    def excel(self):
        sourceFilePath = self.sourceValue.text()
        sheetName = self.sheetNameValue.text()
        startLineNo = self.startLineValue.text()
        startColumnNo = self.startColumnValue.text()
        endLineNo = self.endLineValue.text()
        endColumnNo = self.endColumnValue.text()
        resFilePath = self.resultFilePathValue.text()
        res = np.zeros((int(endLineNo)-int(startLineNo)+1, int(endColumnNo)-int(startColumnNo)+1), dtype=Decimal)
        #正常文件数
        normalNum = 0
        #异常文件数
        errNums = 0
        self.rest_text("-------正常文件-------")
        self.exception_text("-------异常文件-------")
        for filepath, dirnames, filenames in os.walk(sourceFilePath):
            for filename in filenames:
                if filename.lower().find(".xlsx") == -1 and filename.lower().find(".xls") == -1:
                    # 不是 .xlsx | .xls文件跳过
                    errNums = errNums +1
                    self.exception_text(filename)
                    continue
                d_file_name = (os.path.join(filepath, filename))
                self.rest_text(filename)
                xl = xlrd.open_workbook(d_file_name)
                if sheetName in xl.sheet_names():
                    normalNum = normalNum + 1
                    # 存在 内审统03表 sheet 进行汇总
                    sheet = xl.sheet_by_name(sheetName)
                    for r in range(sheet.nrows):
                        if r >= int(startLineNo) and r <= int(endLineNo):
                            for l in range(int(startColumnNo), int(endColumnNo)+1):
                                # l 会被重新赋值 ??
                                l_copy = l
                                sheet_copy = l
                                if sheet.cell(r, sheet_copy).ctype == 2:
                                    res[r - int(startLineNo)][l_copy - int(startColumnNo)] += Decimal(
                                        sheet.cell_value(r, sheet_copy)).quantize(Decimal("0.00"))
                            # print(sheet.cell(r,4))
                else:
                    # 打印未处理文件
                    errNums += 1
                    self.exception_text(filename)
        #print(res.astype(np.str_))
        self.rest_text("=======end======")
        self.rest_text("共执行:"+ str(normalNum) +"文件")
        self.exception_text("=======end======")
        self.exception_text("异常文件:"+ str(errNums) +"文件")
        file = Workbook(encoding='utf-8')
        table = file.add_sheet('data')
        for i in range(len(res)):
            for j in range(len(res[i])):
                table.write(i, j, label=res[i][j])
        resultFilePath = resFilePath + "/result.xlsx"
        file.save(resultFilePath)

if __name__ == '__main__':
    #easygui();
    app = QtWidgets.QApplication(sys.argv)
    window = LoginDlg()
    window.show()
    sys.exit(app.exec_())