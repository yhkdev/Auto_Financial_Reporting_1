# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Main_Window.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets, QtSql
from PyQt5.QtWidgets import QMainWindow, QFileDialog, QDialog, QInputDialog, QMessageBox, QLineEdit
from PyQt5.QtGui import QIcon

import sqlite3

from CopyPasteExcel import copy_paste
from Macros import Ui_Macro_Dialog as Macro_Dialog

import os

def create_connection(database):
    db = QtSql.QSqlDatabase.addDatabase("QSQLITE")
    db.setDatabaseName(database)
    if not db.open():
        print("Cannot open database")
        print(
            "Unable to establish a database connection.\n"
            "This example needs SQLite support. Please read "
            "the Qt SQL driver documentation for information "
            "how to build it.\n\n"
            "Click Cancel to exit."
        )
        return False

    query = QtSql.QSqlQuery()
    if not query.exec_(
        """CREATE TABLE IF NOT EXISTS Macros (
    "id" INTEGER PRIMARY KEY AUTOINCREMENT,
    "title" TEXT NOT NULL,
    "description" TEXT)"""
    ):
        print(query.lastError().text())
        return False
    return True



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

    def __init__(self, parent=None):
        super(Ui_MainWindow, self).__init__(parent)
        self.initUI(MainWindow)


    def initUI(self, MainWindow):

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

        self._model = QtSql.QSqlTableModel(MainWindow)  # Added SQL Table model
        self.model.setTable("Macros")
        self.model.select()

        self.listView_macros = QtWidgets.QListView(self.centralwidget)
        self.listView_macros.setGeometry(QtCore.QRect(20, 250, 350, 150))
        self.listView_macros.setObjectName("tableWidget_macros")
        self.listView_macros.setModel(self.model)
        self.listView_macros.setModelColumn(self.model.record().indexOf("title"))

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
        # self.Button_Run.clicked.connect(self.refresh)

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
        # self.listView_macros.setSortingEnabled(False)
        self.Label_macros.setText(_translate("MainWindow", "Macros"))
        self.Button_Run.setText(_translate("MainWindow", "RUN"))
        self.Label_copyfrom.setText(_translate("MainWindow", "Copy From:"))
        self.Label_destination.setText(_translate("MainWindow", "Destination:"))
        self.Button_browse_destination.setText(_translate("MainWindow", "Browse"))
        self.Button_browse_copyfrom.setText(_translate("MainWindow", "Browse"))
        self.label_select_excel_files.setText(_translate("MainWindow", "Select  Excel  Files"))

    @property
    def model(self):
        return self._model

    @QtCore.pyqtSlot()
    def new_macro(self):

        d = Macro_Dialog()
        if d.exec_() == QtWidgets.QDialog.Accepted:
            r = self.model.record()
            r.setValue("title", d.title)
            r.setValue("description", d.description)
            if self.model.insertRecord(self.model.rowCount(), r):
                self.model.select()

    @QtCore.pyqtSlot()
    def edit_macro(self):
        ixs = self.listView_macros.selectionModel().selectedIndexes()
        if ixs:
            d = Macro_Dialog(self.model, ixs[0].row())
            d.exec_()

    @QtCore.pyqtSlot()
    def remove_macro(self):
        ixs = self.listView_macros.selectionModel().selectedIndexes()
        if ixs:
            reply = QMessageBox.warning(self, "Remove Macro?",
                                        "Remove Macro?",
                                        QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.model.removeRow(ixs[0].row())
                self.model.select()



    def _add_table(self, columns):
        pass
        # row_pos = self.tableWidget_macros.count()
        # last_row = self.tableWidget_macros.
        # self.tableWidget_macros.insertItem()
        #
        # for i, col in enumerate(columns):
        #     self.tableWidget_macros.setitem



    def open_excel_file(self, textEdit):
        """ open file browser and get path to designated copy or destination file """
        fname = QFileDialog.getOpenFileName(self, "Open file")

        if fname[0]:
            file = open(fname[0], 'r')

            with file:
                text = file.name  # << Saves file PATH to textEdit next to it
                textEdit.setText(text)

    def run(self):
        """  """
        copy_wb_path = self.textEdit_copyfrom.text()
        destination_wb_path = self.textEdit_destination
        rule = [["B2", "B2"], ["C2:C4", "D2:D4"]]  # TEST rule
        pass






if __name__ == "__main__":
    import sys
    database_name = "Macros_db"  # ":memory:"
    app = QtWidgets.QApplication(sys.argv)
    if not create_connection(database_name):
        sys.exit(app.exec_())
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.initUI(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

# ----------- TEST ----------

# open_excel_file()

