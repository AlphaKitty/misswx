from PySide2.QtGui import QIcon
from PySide2.QtWidgets import QApplication, QMessageBox, QFileDialog
from PySide2.QtUiTools import QUiLoader
from twilio.rest import Client
from SalesContractMng.contract import SheetWizard

from winreg import *


def send_sms(msg='兄弟', my_number='(+86) 17600313722'):
    # 从官网获得以下信息
    account_sid = 'ACda14f1382d4eaabc2dba04d42b534512'
    auth_token = '0c0d008f66f0fc0c5a5fcace5424f38f'
    # twilio_number = '(469) 530-3834'
    twilio_number = '+14695303834'

    client = Client(account_sid, auth_token)
    try:
        client.messages.create(to=my_number, from_=twilio_number, body=msg)
        print('短信已经发送！')
    except ConnectionError as e:
        print('发送失败，请检查你的账号是否有效或网络是否良好！')
        return e


class Stats:

    def __init__(self):
        self.regpath = "SOFTWARE\\zylTemp\\"
        self.ui = QUiLoader().load('ExcelPicker.ui')
        self.ui.operatorBox.addItems(['不是', '是', '包含', '大于', '大于等于', '小于', '小于等于', '等于', '不等于'])
        self.ui.loadButton.clicked.connect(self.loadExcelEvent)
        self.ui.anlzButton.clicked.connect(self.doExcel)
        self.ui.titleButton.clicked.connect(self.tips)

    def loadExcelEvent(self):
        # 初始化fileDialog
        self.fileDialog = QFileDialog()
        self.fileDialog.setNameFilter('Excel|*.xls')
        self.fileDialog.selectNameFilter('Images|*.xls')
        # self.fileDialog.getOpenFileName(self, '打开Excel', 'C:', 'Excel(*.xls)')
        name_ = self.fileDialog.getOpenFileName()[0]
        if name_ != '':
            self.fileDialog.fileSelected.connect(self.getExcelFullPath(name_))
        else:
            return
        # 包装工作簿
        try:
            self.wizard = SheetWizard(self.ui.pathLabel.text())
        except Exception:
            return
        self.firstRowList = self.wizard.get_row_by_index(0)
        self.ui.conditionBox.addItems(self.firstRowList)
        # for node in self.firstRowList:
        #     checkBox = QCheckBox(self.verticalLayoutWidget)
        #     checkBox.setObjectName(node)
        #     checkBox.setText(node)
        #     checkBox.toggled.connect(lambda: self.checkBoxClicked(node))
        #     self.verticalLayout_2.addWidget(checkBox)
        # QMessageBox.about(self.window, 'Hey', '我猜这不是一个Excel')
        # self.comboBox.

    def getExcelFullPath(self, fileName):
        self.ui.pathLabel.setText(fileName)

    def doExcel(self):
        # 获得主工作表
        main_sheet = self.wizard.get_main_sheet()
        # 根据条件获取符合条件的行号
        try:
            rows = self.wizard.get_rows_by_condition(main_sheet, self.ui.conditionBox.currentText(),
                                                     self.ui.operatorBox.currentText(), self.ui.valueEdit.text())
        except TypeError:
            QMessageBox.about(self.ui, 'warning', '你输入的条件貌似不合适')
            return
        textList = self.ui.plainTextEdit.toPlainText()[:-1].split()
        if len(textList) == 0:
            QMessageBox.about(self.ui, 'warning', '得有生成依据才行~')
            return
        # 根据需要的列名获取列索引
        cols = self.wizard.get_cols_by_col_names(textList)
        # 根据行和列获取全部值
        values = self.wizard.get_values_by_coordinate(rows, cols)
        # 写入Excel
        book = self.wizard.write_excel(values, '新的应收工作表')
        # 保存
        try:
            self.wizard.save_book(book)
        except PermissionError:
            QMessageBox.about(self.ui, 'tips', '你需要先关闭生成的工作簿')
            return
        print(rows)
        print(cols)
        print(values)
        QMessageBox.about(self.ui, 'success!', '成功啦!')

    def tips(self):
        reg = CreateKey(HKEY_CURRENT_USER, self.regpath)
        value = '0'
        try:
            value = QueryValue(reg, 'index')
            SetValue(reg, 'index', REG_SZ, str(int(value) + 1))
        except Exception:
            SetValue(reg, 'index', REG_SZ, '1')
        if '' != value and int(value) == 100:
            self.ui.titleButton.setIcon(QIcon('sources/ico/smile.ico'))
            SetValue(reg, 'index', REG_SZ, str(100))
            return
        elif '' == value or int(value) % 10 == 0:
            QMessageBox.about(self.ui, 'tips', '选择"加载Excel" 选中一个XLS工作表')
        # elif int(value) == 13:
        #     question = QMessageBox.question(self.ui, 'ask', 'do you have a boyfriend?')
        #     if question == QMessageBox.Yes:
        #         SetValue(reg, 'index', REG_SZ, str(100))
        #         QMessageBox.about(self.ui, 'tips', '嗯... 当然...')
        #         send_sms('yes i have.. sorry..')
        #     else:
        #         SetValue(reg, 'index', REG_SZ, str(100))
        #         QMessageBox.about(self.ui, 'tips', '(๑•̀ㅂ•́)و✧')
        #         send_sms('of course no! hahaha')
        elif int(value) % 10 == 1:
            QMessageBox.about(self.ui, 'tips', '在第一个下拉框中选择过滤条件 填入阈值')
        elif int(value) % 10 == 2:
            QMessageBox.about(self.ui, 'tips', '在第二个下拉框中选择操作类型\n注意: 除了"是"和"不是"之外 其他选项都是针对数字的')
        elif int(value) % 10 == 3:
            QMessageBox.about(self.ui, 'tips', '在文本框中填入数字或文本\nagain: [是不是]对应文本\n[等不等于]对应数字')
        elif int(value) % 10 == 4:
            QMessageBox.about(self.ui, 'tips', '如果一切顺利 新的工作簿会被保存在桌面 名字就叫"新的工作簿"')
        elif int(value) % 10 == 5:
            QMessageBox.about(self.ui, 'tips', '时间仓促 bug好多 出现崩溃? 重开试试!')
        elif int(value) % 10 == 6:
            QMessageBox.about(self.ui, 'tips', '不要吐槽本作品的UI 我很丑但我很温柔')
        elif int(value) % 10 == 7:
            QMessageBox.about(self.ui, 'tips', '加载完Excel后复制主Sheet中你想要的汇总的标题行')
        elif int(value) % 10 == 8:
            QMessageBox.about(self.ui, 'tips', '最后一步 点击解析Excel 会生成你想要的结果')
        elif int(value) % 10 == 9:
            QMessageBox.about(self.ui, 'tips', '不要吐槽本作品的UI 我很丑但我很温柔')


app = QApplication([])
stats = Stats()
stats.ui.show()
app.exec_()
