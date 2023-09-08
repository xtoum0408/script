from PySide2.QtWidgets import QFileDialog, QApplication, QMessageBox, QButtonGroup
from PySide2.QtCore import QDir
from PySide2.QtUiTools import QUiLoader
import openpyxl as xl


class Stats:
    def __init__(self):
        self.ui = QUiLoader().load('./excelTool/tool.ui')
        # self.ui.jobChoose.clicked.connect(self.perpare)
        self.ui.shigong.clicked.connect(self.perpare)
        self.ui.tiaoshi.clicked.connect(self.perpare)
        self.ui.huiyi.clicked.connect(self.perpare)
        self.ui.qita.clicked.connect(self.perpare)
        self.ui.addButton.clicked.connect(self.addData)
        self.ui.delButton.clicked.connect(self.deleteData)
        self.ui.clearButton.clicked.connect(self.clearData)
        self.ui.saveData.clicked.connect(self.saveDataToFile)
        self.ui.saveFile.clicked.connect(self.saveFile)


    def perpare(self):
        if self.ui.shigong.isChecked():
            job = '施工'
            print('施工')
        elif self.ui.tiaoshi.isChecked():
            job = '系统调试'
            print('调试')
        elif self.ui.huiyi.isChecked():
            job = '会议'
            print('会议')
        else:
            job = '系统调试'
            print('其他')

        # 先清除候选框内的内容
        allRow = self.ui.houxuanList.count()
        for r in range(0, allRow):
            self.ui.houxuanList.takeItem(0)

        # 再从表格中将内容添加到候选
        personalData = xl.load_workbook('./excelTool/excelModel/测试名单.xlsx')
        sheet = personalData[job]
        for row in range(2, 6):
            name = sheet[f'A{row}'].value
            self.ui.houxuanList.addItem(name)

    # 将人员名单加入已选
    def addData(self):
        chose = self.ui.houxuanList.currentItem().text()
        rowInList = self.ui.houxuanList.currentRow()
        self.ui.houxuanList.takeItem(rowInList)
        self.ui.yixuanList.addItem(chose)

    # 将人员移出已选
    def deleteData(self):
        chose = self.ui.yixuanList.currentItem().text()
        rowInList = self.ui.yixuanList.currentRow()
        self.ui.yixuanList.takeItem(rowInList)
        self.ui.houxuanList.addItem(chose)

    # 将全部人员移出已选
    def clearData(self):
        allRow = self.ui.yixuanList.count()
        for r in range(0, allRow):
            self.ui.houxuanList.addItem(self.ui.yixuanList.item(0).text())
            self.ui.yixuanList.takeItem(0)
    # 将已选写入excel
    # 要做两件事：1.将选择的人的信息写入申请单和方法声明 2.将信息保存进一个dic
    def saveDataToFile(self):
        # 将信息写入一个dic
        
        allRow = self.ui.yixuanList.count()
        for index in range(0, allRow):
            value = self.ui.houxuanList.item(index).text()  # 读取信息



    # 保存文件
    def saveFile(self):
        print('看！我能保存文件！')
        # 此处操作是为了获取保存文件的路径和文件名
        dialog = QFileDialog()
        dialog.setFilter(dialog.filter() | QDir.Hidden)
        # 设置保存的默认类型
        dialog.setDefaultSuffix('xlsx')
        # 设置为保存文件
        dialog.setAcceptMode(QFileDialog.AcceptSave)
        # 设置保存的文件类型（显示在保存界面上的文件类型）
        dialog.setNameFilters(['Excel工作簿 (*.xlsx)'])
        if dialog.exec_():
            filePath = dialog.selectedFiles()[0]
            # 要插入一句book.save
            print('保存的路径为：', filePath)
            QMessageBox.information(
                self.ui,
                '保存',
                '保存成功！')
        else:
            print('Cancelled')


app = QApplication([])
stats = Stats()
stats.ui.show()
app.exec_()
