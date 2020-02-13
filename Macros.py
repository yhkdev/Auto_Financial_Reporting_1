# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Macros.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QDialog, QMessageBox, QLineEdit

from SqliteHelper import SqliteHelper

class Ui_Macro_Dialog(QDialog):
    def __init__(self):
        super(Ui_Macro_Dialog, self).__init__()
        self.initUI(self)

    def initUI(self, Macro_Dialog):
        Macro_Dialog.setObjectName("Macro_Dialog")
        Macro_Dialog.resize(600, 675)

        self.macro_title = QtWidgets.QLineEdit(Macro_Dialog)
        self.macro_title.setGeometry(QtCore.QRect(100, 20, 400, 30))
        self.macro_title.setMaxLength(75)
        self.macro_title.setAlignment(QtCore.Qt.AlignCenter)
        self.macro_title.setObjectName("lineEdit_macro_title")

        self.macro_description = QtWidgets.QTextEdit(Macro_Dialog)
        self.macro_description.setGeometry(QtCore.QRect(20, 70, 560, 100))
        self.macro_description.setObjectName("textEdit_macro_description")

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
        # self.tableWidget_sheets.setDragEnabled(True)
        # self.tableWidget_sheets.setDragDropMode(QtWidgets.QAbstractItemView.DragDrop)
        self.tableWidget_sheets.setEditTriggers(QtWidgets.QAbstractItemView.DoubleClicked | QtWidgets.QAbstractItemView.AnyKeyPressed)
        self.tableWidget_sheets.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
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

        self.tableWidget_cells = QtWidgets.QTableWidget(Macro_Dialog)
        self.tableWidget_cells.setGeometry(QtCore.QRect(345, 240, 218, 320))
        # self.tableWidget_cells.setDragEnabled(True)
        # self.tableWidget_cells.setDragDropMode(QtWidgets.QAbstractItemView.DragDrop)
        self.tableWidget_cells.setEditTriggers(QtWidgets.QAbstractItemView.DoubleClicked | QtWidgets.QAbstractItemView.AnyKeyPressed)
        self.tableWidget_cells.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.tableWidget_cells.setObjectName("tableWidget_cell_entries")
        self.tableWidget_cells.setColumnCount(2)
        self.tableWidget_cells.setRowCount(0)  # << Set to 0 myself. Use commented code later if needed
        # item = QtWidgets.QTableWidgetItem()
        # self.tableWidget_cell_entries.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget_cells.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget_cells.setHorizontalHeaderItem(1, item)
        # item = QtWidgets.QTableWidgetItem()
        # self.tableWidget_cell_entries.setItem(0, 0, item)
        # item = QtWidgets.QTableWidgetItem()
        # self.tableWidget_cell_entries.setItem(0, 1, item)

        self.Button_del_sheets = QtWidgets.QPushButton(Macro_Dialog)
        self.Button_del_sheets.setGeometry(QtCore.QRect(90, 570, 50, 30))
        self.Button_del_sheets.setAutoDefault(False)
        self.Button_del_sheets.setObjectName("Button_del_sheets")
        self.Button_del_sheets.clicked.connect(lambda: self.remove_row(self.tableWidget_sheets))  # Del Sheet entry

        self.Button_add_sheets = QtWidgets.QPushButton(Macro_Dialog)
        self.Button_add_sheets.setGeometry(QtCore.QRect(170, 570, 50, 30))
        self.Button_add_sheets.setAutoDefault(False)
        self.Button_add_sheets.setObjectName("Button_add_sheets")
        self.Button_add_sheets.clicked.connect(lambda: self.add_row(self.tableWidget_sheets))  # Add entry to sheet table

        self.Button_del_cell_entries = QtWidgets.QPushButton(Macro_Dialog)
        self.Button_del_cell_entries.setGeometry(QtCore.QRect(390, 570, 50, 30))
        self.Button_del_cell_entries.setAutoDefault(False)
        self.Button_del_cell_entries.setObjectName("Button_del_cell_entries")
        self.Button_del_cell_entries.clicked.connect(lambda: self.remove_row(self.tableWidget_cells)) # Del cell entry

        self.Button_add_cell_entries = QtWidgets.QPushButton(Macro_Dialog)
        self.Button_add_cell_entries.setGeometry(QtCore.QRect(470, 570, 50, 30))
        self.Button_add_cell_entries.setAutoDefault(False)
        self.Button_add_cell_entries.setObjectName("Button_add_cell_entries")
        self.Button_add_cell_entries.clicked.connect(lambda: self.add_row(self.tableWidget_cells))  # Add entry to cell entry table

        self.Button_cancel = QtWidgets.QPushButton(Macro_Dialog)
        self.Button_cancel.setGeometry(QtCore.QRect(190, 630, 100, 30))
        self.Button_cancel.setAutoDefault(False)
        self.Button_cancel.setObjectName("Button_cancel")
        self.Button_cancel.clicked.connect(lambda: Macro_Dialog.close())  # << Exit Dialog

        self.Button_save = QtWidgets.QPushButton(Macro_Dialog)
        self.Button_save.setGeometry(QtCore.QRect(310, 630, 100, 30))
        self.Button_save.setObjectName("Button_save")
        self.Button_save.clicked.connect(self.save_macro)  # << Save/Update Dialog to sqlite


        self.retranslateUi(Macro_Dialog)
        QtCore.QMetaObject.connectSlotsByName(Macro_Dialog)

    def retranslateUi(self, Macro_Object):
        _translate = QtCore.QCoreApplication.translate
        Macro_Object.setWindowTitle(_translate("Macro_Object", "Dialog"))
        self.Label_sheets.setText(_translate("Macro_Object", "Sheets"))
        self.Label_cell_entries.setText(_translate("Macro_Object", "Cell Entries"))
        self.macro_title.setPlaceholderText(_translate("Macro_Object", "Enter Macro Title"))
        self.Button_add_sheets.setText(_translate("Macro_Object", "+"))
        self.Button_del_sheets.setText(_translate("Macro_Object", "-"))
        self.Button_del_cell_entries.setText(_translate("Macro_Object", "-"))
        self.Button_add_cell_entries.setText(_translate("Macro_Object", "+"))
        self.Button_save.setText(_translate("Macro_Object", "Save"))
        self.Button_cancel.setText(_translate("Macro_Object", "Cancel"))
        self.macro_description.setPlaceholderText(_translate("Macro_Object", "Enter Macro Descriptions"))

        # Sheets Table
        item = self.tableWidget_sheets.horizontalHeaderItem(0)
        item.setText(_translate("Macro_Object", "Copy from:"))
        item = self.tableWidget_sheets.horizontalHeaderItem(1)
        item.setText(_translate("Macro_Object", "Paste to:"))
        __sortingEnabled = self.tableWidget_sheets.isSortingEnabled()
        self.tableWidget_sheets.setSortingEnabled(False)
        self.tableWidget_sheets.setSortingEnabled(__sortingEnabled)

        # Cell Entries Table
        item = self.tableWidget_cells.horizontalHeaderItem(0)
        item.setText(_translate("Macro_Object", "Copy from:"))
        item = self.tableWidget_cells.horizontalHeaderItem(1)
        item.setText(_translate("Macro_Object", "Paste to:"))
        __sortingEnabled = self.tableWidget_cells.isSortingEnabled()
        self.tableWidget_cells.setSortingEnabled(False)
        self.tableWidget_cells.setSortingEnabled(__sortingEnabled)

    def add_row(self, TableWidget):
        rowPosition = TableWidget.rowCount()
        TableWidget.insertRow(rowPosition)

    def remove_row(self, TableWidget):
        # Note: Edit, if possible, to allow row deletion by clicking on a cell, not the entire row

        if TableWidget.selectionModel().selectedRows():  # TableWidget.selectionModel().hasSelection() << use when editting to allow row deletion by clicking on a cell, not entire row

            reply = QMessageBox.warning(self, "Remove Entry", "Remove selected rows?", QMessageBox.Yes | QMessageBox.No)

            if reply == QMessageBox.Yes:

                index_list = []
                for model_index in TableWidget.selectionModel().selectedRows():
                    index = QtCore.QPersistentModelIndex(model_index)
                    index_list.append(index)

                for index in index_list:
                    TableWidget.removeRow(index.row())


        #
        # if item:
        #     reply = QMessageBox.warning(self, "Remove Entry",
        #                                  "Remove Entry '{0}'?".format(str(item.text())),  # <<<
        #                                  QMessageBox.Yes | QMessageBox.No)
        #
        #     if reply == QMessageBox.Yes:
        #         item = TableWidget.takeItem(row) # <<<
        #         del item

    def save_macro(self):  # << Add func to save to sqlite !!!!!!
        title = self.macro_title.text()
        description = self.macro_description.toPlainText()
        sheets = [self.tableWidget_sheets.items()]

        add_macros_sql = ''' INSERT INTO Macros (title,description)
                      VALUES(?,?) '''  # The '?'s will be filled with 'data_to_insert'

        # add_sheets_sql = ''' INSERT INTO Sheets (macroid, copy_sheet, paste_sheet)
        #           VALUES(?,?,?,?) '''

        # add_cells_sql = ''' INSERT INTO Cells (sheetid, copy_cell, paste_cell)
        #           VALUES(?,?,?) '''  # The '?'s will be filled with 'data_to_insert'

        sqlhelper = SqliteHelper("Macros_db")

        if sqlhelper is not None:  # If db is connected:
            macro_data = (title, description)
            sqlhelper.insert(add_macros_sql, macro_data)
            # macro_id = sqlhelper.insert(add_macros_sql, macro_data)

        self.close()

    def new_dialog(self, conn):
        """ To be used when 'save' button is pressed """
        sqlhelper = SqliteHelper("Macros_db")

        if conn is not None:
            macro_data = (self.macro_title, self.macro_description)
            macro_id = sqlhelper.insert(add_macros_sql, macro_data)

            # sheet_data = ()
            # sheet_id = _add_item(conn, sheet_data, add_sheets_sql)

            # cell_data = ()
            # _add_item(conn, cell_data, add_cells_sql)

        else:
            print("Error! can't create db connection")



if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Macro_Dialog = QtWidgets.QDialog()
    ui = Ui_Macro_Dialog()
    ui.initUI(Macro_Dialog)
    Macro_Dialog.show()
    sys.exit(app.exec_())


