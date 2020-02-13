# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Main_Window.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMainWindow, QFileDialog, QDialog, QInputDialog, QMessageBox, QLineEdit
from PyQt5.QtGui import QIcon

import sqlite3

from CopyPasteExcel import copy_paste
from Macros import Ui_Macro_Dialog
from SqliteHelper import SqliteHelper

import os


class FileEdit(QLineEdit):
    """ Custom Subclass of QLineEdit tp allow users to drag & drop for file selection (instead of browsing) """
    def __init__(self, parent):
        super(FileEdit, self).__init__(parent)
        self.setDragEnabled(True)

    def dragEnterEvent(self, event):
        data = event.mimeData()
        urls = data.urls()
        if urls and urls[0].scheme() == 'file':
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        data = event.mimeData()
        urls = data.urls()
        if urls and urls[0].scheme() == 'file':
            event.acceptProposedAction()

    def dropEvent(self, event):
        data = event.mimeData()
        urls = data.urls()
        if urls and urls[0].scheme() == 'file':
            filepath = str(urls[0].path())[1:]
            # any file type here
            self.setText(filepath)


            # BELOW IS CODE FROM: https://www.reddit.com/r/learnpython/comments/97z5dq/pyqt5_drag_and_drop_file_option/e4cv39x/
            # WHEN YOU HAVE TIME, REFACTOR ABOVE CODE WITH REGEX TO ONLY OPEN EXCEL FILES (anything ending with .xl...)

            # if filepath[-4:].upper() in [".txt", ".x"]:
            #     self.setText(filepath)
            # else:
            #     dialog = QMessageBox()
            #     dialog.setWindowTitle("Error: Invalid File")
            #     dialog.setText("Only Excel files are accepted")
            #     dialog.setIcon(QMessageBox.Warning)
            #     dialog.exec_()


class Ui_MainWindow(QMainWindow):
    def __init__(self):
        super(Ui_MainWindow, self).__init__()
        self.initUI(MainWindow)


    def initUI(self, MainWindow):
        sqlhelper = SqliteHelper("Macros_db")  # Create db/tables if it doesn't exist yet
        sqlhelper.create_table()

        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(550, 500)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.label_select_excel_files = QtWidgets.QLabel(self.centralwidget)
        self.label_select_excel_files.setGeometry(QtCore.QRect(210, 70, 131, 21))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.label_select_excel_files.setFont(font)
        self.label_select_excel_files.setAlignment(QtCore.Qt.AlignCenter)
        self.label_select_excel_files.setObjectName("label_select_excel_files")

        self.Frame_fileimport = QtWidgets.QFrame(self.centralwidget)
        self.Frame_fileimport.setGeometry(QtCore.QRect(20, 100, 511, 100))
        self.Frame_fileimport.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.Frame_fileimport.setFrameShadow(QtWidgets.QFrame.Raised)
        self.Frame_fileimport.setObjectName("frame_fileimport")

        self.Label_copyfrom = QtWidgets.QLabel(self.Frame_fileimport)
        self.Label_copyfrom.setGeometry(QtCore.QRect(10, 10, 79, 35))
        self.Label_copyfrom.setObjectName("Label_copyfrom")

        self.Label_destination = QtWidgets.QLabel(self.Frame_fileimport)
        self.Label_destination.setGeometry(QtCore.QRect(10, 50, 71, 31))
        self.Label_destination.setObjectName("Label_destination")

        self.textEdit_copyfrom = FileEdit(self.Frame_fileimport)
        self.textEdit_copyfrom.setGeometry(QtCore.QRect(90, 20, 319, 21))
        self.textEdit_copyfrom.setObjectName("textEdit_copyfrom")

        self.textEdit_destination = FileEdit(self.Frame_fileimport)
        self.textEdit_destination.setGeometry(QtCore.QRect(90, 60, 319, 21))
        self.textEdit_destination.setObjectName("textEdit_destination")

        self.Button_browse_copyfrom = QtWidgets.QPushButton(self.Frame_fileimport)
        self.Button_browse_copyfrom.setGeometry(QtCore.QRect(410, 10, 91, 41))
        self.Button_browse_copyfrom.setObjectName("Button_browse_copyfrom")
        self.Button_browse_copyfrom.clicked.connect(lambda: self.open_excel_file(self.textEdit_copyfrom))  # Added by me (browse func for copy file)

        self.Button_browse_destination = QtWidgets.QPushButton(self.Frame_fileimport)
        self.Button_browse_destination.setGeometry(QtCore.QRect(410, 50, 91, 41))
        self.Button_browse_destination.setObjectName("Button_browse_destination")
        self.Button_browse_destination.clicked.connect(lambda: self.open_excel_file(self.textEdit_destination))  # Added by me (browse func for destination file)

        self.line = QtWidgets.QFrame(self.centralwidget)
        self.line.setGeometry(QtCore.QRect(0, 220, 550, 5))
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")

        self.verticalLayoutWidget = QtWidgets.QWidget(self.centralwidget)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(380, 250, 151, 151))
        self.verticalLayoutWidget.setObjectName("verticalLayoutWidget")

        self.verticalLayout = QtWidgets.QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")

        self.Label_macros = QtWidgets.QLabel(self.centralwidget)
        self.Label_macros.setGeometry(QtCore.QRect(135, 230, 120, 16))
        self.Label_macros.setAlignment(QtCore.Qt.AlignCenter)
        self.Label_macros.setObjectName("Label_macros")

        self.listWidget_macros = QtWidgets.QListWidget(self.centralwidget)
        self.listWidget_macros.setGeometry(QtCore.QRect(20, 250, 350, 150))
        self.listWidget_macros.setObjectName("listWidget_macros")


        # ADD REFRESH LIST HERE?
        self.refresh()  # << Fill listtable with macro title data from db


        self.Button_new_macro = QtWidgets.QPushButton(self.verticalLayoutWidget)
        self.Button_new_macro.setObjectName("Button_new_macro")
        self.verticalLayout.addWidget(self.Button_new_macro)
        self.Button_new_macro.clicked.connect(self.new_macro)    # Create new entry in macro list

        self.Button_edit_macro = QtWidgets.QPushButton(self.verticalLayoutWidget)
        self.Button_edit_macro.setObjectName("Button_edit_macro")
        self.verticalLayout.addWidget(self.Button_edit_macro)
        self.Button_edit_macro.clicked.connect(self.edit_macro)    # Edit selected entry in macro list

        self.Button_remove_macro = QtWidgets.QPushButton(self.verticalLayoutWidget)
        self.Button_remove_macro.setObjectName("Button_remove_macro")
        self.verticalLayout.addWidget(self.Button_remove_macro)
        self.Button_remove_macro.clicked.connect(self.remove_macro)    # Remove selected entry in macro list

        self.Button_Run = QtWidgets.QPushButton(self.centralwidget)
        self.Button_Run.setGeometry(QtCore.QRect(120, 410, 311, 51))
        self.Button_Run.setObjectName("Button_Run")
        self.Button_Run.clicked.connect(self.refresh)

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 550, 22))
        self.menubar.setObjectName("menubar")

        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")

        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.Button_new_macro.setText(_translate("MainWindow", "New"))
        self.Button_edit_macro.setText(_translate("MainWindow", "Edit"))
        self.Button_remove_macro.setText(_translate("MainWindow", "Remove"))
        self.listWidget_macros.setSortingEnabled(False)
        self.Label_macros.setText(_translate("MainWindow", "Macros"))
        self.Button_Run.setText(_translate("MainWindow", "RUN"))
        self.Label_copyfrom.setText(_translate("MainWindow", "Copy From:"))
        self.Label_destination.setText(_translate("MainWindow", "Destination:"))
        self.Button_browse_destination.setText(_translate("MainWindow", "Browse"))
        self.Button_browse_copyfrom.setText(_translate("MainWindow", "Browse"))
        self.label_select_excel_files.setText(_translate("MainWindow", "Select  Excel  Files"))

    def refresh(self):
        """ Refresh Macros """
        # If db is not None:
        sqlhelper = SqliteHelper("Macros_db")

        if sqlhelper:
            self.listWidget_macros.clear()

            data = sqlhelper.load_table("Macros")  # == (id, title, description)

            for macro in data:
                print(macro)
                self.listWidget_macros.addItem(macro[1])
                # self.listWidget_macros.setCurrentRow(0)

    def _add_table(self, columns):
        pass
        # row_pos = self.listWidget_macros.count()
        # last_row = self.listWidget_macros.
        # self.listWidget_macros.insertItem()
        #
        # for i, col in enumerate(columns):
        #     self.listWidget_macros.setitem

    def open_excel_file(self, textEdit):
        """ open file browser and get path to designated copy or destination file """
        fname = QFileDialog.getOpenFileName(self, "Open file")

        if fname[0]:
            file = open(fname[0], 'r')

            with file:
                text = file.name  # << Saves file PATH to textEdit next to it
                textEdit.setText(text)

    def new_macro(self):
        macros_dialog = Ui_Macro_Dialog()
        macros_dialog.show()

        # # self.listWidget_macros.addItem("Added new Macro")
        # row = self.listWidget_macros.currentRow()
        # title, label = "Add Macro", "Macro Name:"
        # text, ok = QInputDialog.getText(self, title, label)  # << Conventional syntax for QInputDialog.getText()
        # # 'QInputDialog.getText' -> (str, bool)
        # # 'text' is text I/O. 'ok' == True if user pressed "ok" button in the popup input dialog
        #
        # if ok and text is not None:  # If user entered input and pressed ok:
        #     self.listWidget_macros.insertItem(row, text)  # Add new item to table

    def edit_macro(self):
        row = self.listWidget_macros.currentRow()  # get currently selected row idx
        item = self.listWidget_macros.item(row)  # Get row item at currently selected row

        if item is not None:
            title = "Edit Macro: '{0}'".format(str(item.text()))
            text, ok = QInputDialog.getText(self, title, title,
                                              QLineEdit.Normal, item.text())
            if ok and text is not None:
                item.setText(text)  # Edit

    def remove_macro(self):
        row = self.listWidget_macros.currentRow()
        item = self.listWidget_macros.item(row)

        if item:

            reply = QMessageBox.warning(self, "Remove Macro?",
                                         "Remove Macro '{0}'?".format(str(item.text())),
                                         QMessageBox.Yes | QMessageBox.No)

            if reply == QMessageBox.Yes:
                item = self.listWidget_macros.takeItem(row)
                del item

    def sort(self):
        self.listWidget_macros.sortItems()

    def run(self):
        """  """
        copy_wb_path = self.textEdit_copyfrom.text()
        destination_wb_path = self.textEdit_destination
        rule = [["B2", "B2"], ["C2:C4", "D2:D4"]]  # TEST rule
        pass



if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.initUI(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

# ----------- TEST ----------

# open_excel_file()

