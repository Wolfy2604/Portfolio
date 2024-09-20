import json, datetime

from PyQt6 import QtCore, QtGui, QtWidgets
import sys, shutil, traceback
import main


class OpenCVLabel(QtWidgets.QLabel):
    def __init__(self, *args, **kwargs):
        super(OpenCVLabel, self).__init__(*args, **kwargs)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.accept()
        else:
            e.ignore()

    def dropEvent(self, e):
        if e.mimeData().hasUrls():
            e.accept()
            for url in e.mimeData().urls():
                self.copyFile(url.toLocalFile())
        else:
            e.ignore()

    def copyFile(self, filename):
        ex.short = filename.split('/').pop()
        shutil.copyfile(filename, f'input/{ex.short}')
        ex.alert.append('Протокол загружен!')
        ex.push_alert()
        ex.alert = []


class Example(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.alert = []
        self.data = {}
        self.short = ''
        self.initUI()

    def initUI(self):
        label1 = OpenCVLabel()
        Form = QtWidgets.QWidget()
        self.listView_2 = QtWidgets.QListView(parent=Form)
        self.listView_2.setGeometry(QtCore.QRect(10, 140, 281, 381))
        self.listView_2.setObjectName("listView_2")
        self.setCentralWidget(Form)
        lay = QtWidgets.QVBoxLayout(Form)
        lay.addWidget(label1)
        self.resize(609, 612)
        Form.setObjectName("Form")
        self.listView = QtWidgets.QListView(parent=Form)
        self.listView.setGeometry(QtCore.QRect(315, 340, 281, 261))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.listView.setObjectName("listView")
        self.calendarWidget = QtWidgets.QCalendarWidget(parent=Form)
        self.calendarWidget.setGeometry(QtCore.QRect(313, 90, 281, 183))
        self.calendarWidget.setObjectName("calendarWidget")
        self.label = QtWidgets.QLabel(parent=Form)
        self.label.setGeometry(QtCore.QRect(326, 300, 281, 31))
        font = QtGui.QFont()
        font.setPointSize(16)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.label_6 = QtWidgets.QLabel(parent=Form)
        self.label_6.setGeometry(QtCore.QRect(326, 10, 91, 31))
        font = QtGui.QFont()
        font.setPointSize(16)
        self.label_6.setFont(font)
        self.label_6.setObjectName("label_6")
        self.label_7 = QtWidgets.QLabel(parent=Form)
        self.label_7.setGeometry(QtCore.QRect(326, 40, 91, 31))
        font = QtGui.QFont()
        font.setPointSize(16)
        self.label_7.setFont(font)
        self.label_7.setObjectName("label_7")
        self.lineEdit_4 = QtWidgets.QLineEdit(parent=Form)
        self.lineEdit_4.setGeometry(QtCore.QRect(410, 10, 181, 31))
        font = QtGui.QFont()
        font.setPointSize(16)
        with open('bd.json', encoding='utf-8') as bd:
            exp = json.load(bd)["expert"]
        self.lineEdit_4.setFont(font)
        self.lineEdit_4.setPlaceholderText(exp)
        self.lineEdit_4.setObjectName("lineEdit")
        self.line = QtWidgets.QFrame(parent=Form)
        self.line.setGeometry(QtCore.QRect(289, 10, 31, 581))
        self.line.setFrameShape(QtWidgets.QFrame.Shape.VLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Shadow.Sunken)
        self.line.setObjectName("line")
        self.pushButton = QtWidgets.QPushButton(parent=Form)
        self.pushButton.setGeometry(QtCore.QRect(30, 540, 241, 51))
        font = QtGui.QFont()
        font.setPointSize(16)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.start)
        self.lineEdit = QtWidgets.QLineEdit(parent=Form)
        self.lineEdit.setGeometry(QtCore.QRect(10, 10, 131, 31))
        font = QtGui.QFont()
        font.setPointSize(16)
        self.lineEdit.setFont(font)
        self.lineEdit.setText("")
        self.lineEdit.setObjectName("lineEdit")
        self.dateEdit = QtWidgets.QDateEdit(parent=Form)
        self.dateEdit.setGeometry(QtCore.QRect(10, 50, 131, 31))
        self.dateEdit.setObjectName("dateEdit")
        self.dateEdit.setDate(QtCore.QDate.fromString('01/01/2024', "dd/MM/yyyy"))
        self.lineEdit_3 = QtWidgets.QLineEdit(parent=Form)
        self.lineEdit_3.setGeometry(QtCore.QRect(10, 89, 131, 31))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.dateEdit.setFont(font)
        self.dateEdit.setObjectName("dateEdit")
        self.lineEdit_3.setFont(font)
        self.lineEdit_3.setText("")
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.label_2 = QtWidgets.QLabel(parent=Form)
        self.label_2.setGeometry(QtCore.QRect(160, 10, 121, 21))
        font = QtGui.QFont()
        font.setPointSize(16)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(parent=Form)
        self.label_3.setGeometry(QtCore.QRect(160, 50, 121, 21))
        font = QtGui.QFont()
        font.setPointSize(16)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.label_4 = QtWidgets.QLabel(parent=Form)
        self.label_4.setGeometry(QtCore.QRect(160, 90, 121, 21))
        font = QtGui.QFont()
        font.setPointSize(16)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(parent=Form)
        self.label_5.setGeometry(QtCore.QRect(50, 490, 221, 21))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.label_5.setFont(font)
        self.label_5.setObjectName("label_5")
        self.listView.raise_()
        self.label.raise_()
        self.line.raise_()
        self.pushButton.raise_()
        self.lineEdit.raise_()
        self.dateEdit.raise_()
        self.lineEdit_3.raise_()
        self.label_2.raise_()
        self.label_3.raise_()
        self.label_4.raise_()
        self.label_5.raise_()
        self.label_6.raise_()
        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def start(self):
        try:
            self.data["num_request"] = int(self.lineEdit.text())
            self.data["date_request"] = self.dateEdit.text()
            self.data["num_ez"] = int(self.lineEdit_3.text())
            self.data["expert"] = self.lineEdit_4.text()
        except ValueError:
            self.alert.append('Некорректное заполнение полей слева!')
            self.push_alert()
            return
        if not self.data["num_request"] or not self.data["num_ez"]:
            self.alert.append('Не заполнено одно из полей слева!')
            self.push_alert()
            return
        self.data["num_request"] = str(self.data["num_request"])
        self.data["num_ez"] = str(self.data["num_ez"])
        if len(self.data["num_ez"]) > 5:
            self.alert.append('Неправильный номер ЭЗ!')
            self.push_alert()
            return
        if len(self.data["num_request"]) > 2:
            self.alert.append('Неправильный номер заявки!')
            self.push_alert()
            return
        if not self.short:
            self.alert.append('Не загружен протокол!')
            self.push_alert()
            return
        if len(self.data["num_ez"]) < 4:
            zeros = (4 - len(self.data["num_ez"])) * '0'
            self.data["num_ez"] = zeros + self.data["num_ez"]
        if len(self.data["num_request"]) < 2:
            zeros = (2 - len(self.data["num_request"])) * '0'
            self.data["num_request"] = zeros + self.data["num_request"]
        a = self.calendarWidget.selectedDate().toPyDate()
        self.data["ez_date"] = datetime.datetime.strftime(a, '%d.%m.%Y')
        self.alert = self.alert + main.start(self.short, self.data)
        self.push_alert()

    def push_alert(self):
        print('PUSH')
        if not self.alert:
            self.alert.append('Нет ошибок')
        self.model_1 = QtCore.QStringListModel(self)
        self.model_1.setStringList(self.alert)
        self.listView.setModel(self.model_1)
        self.alert = []


    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.label.setText(_translate("Form", "Предупреждения:"))
        self.pushButton.setText(_translate("Form", "Сформировать ЭЗ ГО"))
        self.label_2.setText(_translate("Form", "№ заявки"))
        self.label_3.setText(_translate("Form", "Дата заявки"))
        self.label_4.setText(_translate("Form", "№ ЭЗ"))
        self.label_5.setText(_translate("Form", "Сюда бросить протокол"))
        self.label_6.setText(_translate("Form", "Эксперт:"))
        self.label_7.setText(_translate("Form", "Дата ЭЗ:"))


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    ex = Example()
    ex.show()
    sys.exit(app.exec())