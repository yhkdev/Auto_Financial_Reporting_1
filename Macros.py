# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Macros.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QDialog, QMessageBox, QLineEdit

class Ui_Macro_Dialog(QDialog):
    def __init__(self):
        super(Ui_Macro_Dialog, self).__init__()
        self.initUI(Macro_Dialog)

    def initUI(self, Macro_Dialog):
        Macro_Dialog.setObjectName("Macro_Dialog")
        Macro_Dialog.resize(600, 675)

        self.lineEdit_macro_title = QtWidgets.QLineEdit(Macro_Dialog)
        self.lineEdit_macro_title.setGeometry(QtCore.QRect(100, 20, 400, 30))
        self.lineEdit_macro_title.setMaxLength(75)
        self.lineEdit_macro_title.setAlignment(QtCore.Qt.AlignCenter)
        self.lineEdit_macro_title.setObjectName("lineEdit_macro_title")

        self.textEdit_macro_description = QtWidgets.QTextEdit(Macro_Dialog)
        self.textEdit_macro_description.setGeometry(QtCore.QRect(20, 70, 560, 100))
        self.textEdit_macro_description.setObjectName("textEdit_macro_description")

        self.Label_sheets = QtWidgets.QLabel(Macro_Dialog)
        self.Label_sheets.setGeometry(QtCore.QRect(45, 210, 210, 20))
        self.Label_sheets.setAlignment(QtCore.Qt.AlignCenter)
        self.Label_sheets.setObjectName("Label_sheets")

        self.Label_cell_entries = QtWidgets.QLabel(Macro_Dialog)
        self.Label_cell_entries.setGeometry(QtCore.QRect(345, 210, 210, 20))
        self.Label_cell_entries.setAlignment(QtCore.Qt.AlignCenter)
        self.Label_cell_entries.setObjectName("Label_cell_entries")

        self.tableWidget_sheets = QtWidgets.QTableWidget(Macro_Dialog)
        self.tableWidget_sheets.setGeometry(QtCore.QRect(45, 240, 218, 320))
        self.tableWidget_sheets.setDragEnabled(True)
        self.tableWidget_sheets.setDragDropMode(QtWidgets.QAbstractItemView.DragDrop)
        self.tableWidget_sheets.setColumnCount(2)
        self.tableWidget_sheets.setObjectName("tableWidget_sheets")
        self.tableWidget_sheets.setRowCount(0)  # << Set to 0 myself. Use commented code later if needed
        # item = QtWidgets.QTableWidgetItem()
        # self.tableWidget_sheets.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget_sheets.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget_sheets.setHorizontalHeaderItem(1, item)
        # item = QtWidgets.QTableWidgetItem()
        # self.tableWidget_sheets.setItem(0, 0, item)
        # item = QtWidgets.QTableWidgetItem()
        # self.tableWidget_sheets.setItem(0, 1, item)

        self.tableWidget_cell_entries = QtWidgets.QTableWidget(Macro_Dialog)
        self.tableWidget_cell_entries.setGeometry(QtCore.QRect(345, 240, 218, 320))
        self.tableWidget_cell_entries.setDragEnabled(True)
        self.tableWidget_cell_entries.setDragDropMode(QtWidgets.QAbstractItemView.DragDrop)
        self.tableWidget_cell_entries.setObjectName("tableWidget_cell_entries")
        self.tableWidget_cell_entries.setColumnCount(2)
        self.tableWidget_cell_entries.setRowCount(0)  # << Set to 0 myself. Use commented code later if needed
        # item = QtWidgets.QTableWidgetItem()
        # self.tableWidget_cell_entries.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget_cell_entries.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget_cell_entries.setHorizontalHeaderItem(1, item)
        # item = QtWidgets.QTableWidgetItem()
        # self.tableWidget_cell_entries.setItem(0, 0, item)
        # item = QtWidgets.QTableWidgetItem()
        # self.tableWidget_cell_entries.setItem(0, 1, item)

        self.Button_del_sheets = QtWidgets.QPushButton(Macro_Dialog)
        self.Button_del_sheets.setGeometry(QtCore.QRect(90, 570, 50, 30))
        self.Button_del_sheets.setAutoDefault(False)
        self.Button_del_sheets.setObjectName("Button_del_sheets")

        self.Button_add_sheets = QtWidgets.QPushButton(Macro_Dialog)
        self.Button_add_sheets.setGeometry(QtCore.QRect(170, 570, 50, 30))
        self.Button_add_sheets.setAutoDefault(False)
        self.Button_add_sheets.setObjectName("Button_add_sheets")

        self.Button_del_cell_entries = QtWidgets.QPushButton(Macro_Dialog)
        self.Button_del_cell_entries.setGeometry(QtCore.QRect(390, 570, 50, 30))
        self.Button_del_cell_entries.setAutoDefault(False)
        self.Button_del_cell_entries.setObjectName("Button_del_cell_entries")

        self.Button_add_cell_entries = QtWidgets.QPushButton(Macro_Dialog)
        self.Button_add_cell_entries.setGeometry(QtCore.QRect(470, 570, 50, 30))
        self.Button_add_cell_entries.setAutoDefault(False)
        self.Button_add_cell_entries.setObjectName("Button_add_cell_entries")

        self.Button_cancel = QtWidgets.QPushButton(Macro_Dialog)
        self.Button_cancel.setGeometry(QtCore.QRect(190, 630, 100, 30))
        self.Button_cancel.setAutoDefault(False)
        self.Button_cancel.setObjectName("Button_cancel")

        self.Button_save = QtWidgets.QPushButton(Macro_Dialog)
        self.Button_save.setGeometry(QtCore.QRect(310, 630, 100, 30))
        self.Button_save.setObjectName("Button_save")

        self.retranslateUi(Macro_Dialog)
        QtCore.QMetaObject.connectSlotsByName(Macro_Dialog)

    def retranslateUi(self, Macro_Object):
        _translate = QtCore.QCoreApplication.translate
        Macro_Object.setWindowTitle(_translate("Macro_Object", "Dialog"))
        self.Label_sheets.setText(_translate("Macro_Object", "Sheets"))
        self.Label_cell_entries.setText(_translate("Macro_Object", "Cell Entries"))
        self.lineEdit_macro_title.setPlaceholderText(_translate("Macro_Object", "Enter Macro Title"))
        self.Button_add_sheets.setText(_translate("Macro_Object", "+"))
        self.Button_del_sheets.setText(_translate("Macro_Object", "-"))
        self.Button_del_cell_entries.setText(_translate("Macro_Object", "-"))
        self.Button_add_cell_entries.setText(_translate("Macro_Object", "+"))
        self.Button_save.setText(_translate("Macro_Object", "Save"))
        self.Button_cancel.setText(_translate("Macro_Object", "Cancel"))
        self.textEdit_macro_description.setPlaceholderText(_translate("Macro_Object", "Enter Macro Descriptions"))

        # Sheets Table
        item = self.tableWidget_sheets.horizontalHeaderItem(0)
        item.setText(_translate("Macro_Object", "Copy from:"))
        item = self.tableWidget_sheets.horizontalHeaderItem(1)
        item.setText(_translate("Macro_Object", "Paste to:"))
        __sortingEnabled = self.tableWidget_sheets.isSortingEnabled()
        self.tableWidget_sheets.setSortingEnabled(False)
        self.tableWidget_sheets.setSortingEnabled(__sortingEnabled)

        # Cell Entries Table
        item = self.tableWidget_cell_entries.horizontalHeaderItem(0)
        item.setText(_translate("Macro_Object", "Copy from:"))
        item = self.tableWidget_cell_entries.horizontalHeaderItem(1)
        item.setText(_translate("Macro_Object", "Paste to:"))
        __sortingEnabled = self.tableWidget_cell_entries.isSortingEnabled()
        self.tableWidget_cell_entries.setSortingEnabled(False)
        self.tableWidget_cell_entries.setSortingEnabled(__sortingEnabled)



if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Macro_Dialog = QtWidgets.QDialog()
    ui = Ui_Macro_Dialog()
    ui.initUI(Macro_Dialog)
    Macro_Dialog.show()
    sys.exit(app.exec_())

