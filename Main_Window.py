# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Main_Window.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMainWindow, QMessageBox, QLineEdit
from PyQt5.QtGui import QIcon

from CopyPasteExcel import copy_paste

import os


excel_folder_path = "/local_test"

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
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(550, 500)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.verticalLayoutWidget = QtWidgets.QWidget(self.centralwidget)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(380, 250, 151, 151))
        self.verticalLayoutWidget.setObjectName("verticalLayoutWidget")

        self.verticalLayout = QtWidgets.QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")

        self.Button_new_rule = QtWidgets.QPushButton(self.verticalLayoutWidget)
        self.Button_new_rule.setObjectName("Button_new_rule")
        self.verticalLayout.addWidget(self.Button_new_rule)
        self.Button_edit_rule = QtWidgets.QPushButton(self.verticalLayoutWidget)
        self.Button_edit_rule.setObjectName("Button_edit_rule")

        self.verticalLayout.addWidget(self.Button_edit_rule)
        self.Button_remove_rule = QtWidgets.QPushButton(self.verticalLayoutWidget)
        self.Button_remove_rule.setObjectName("Button_remove_rule")

        self.verticalLayout.addWidget(self.Button_remove_rule)
        self.listWidget_rules = QtWidgets.QListWidget(self.centralwidget)
        self.listWidget_rules.setGeometry(QtCore.QRect(20, 250, 350, 150))
        self.listWidget_rules.setObjectName("listWidget_rules")

        self.line = QtWidgets.QFrame(self.centralwidget)
        self.line.setGeometry(QtCore.QRect(0, 220, 550, 5))
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")

        self.Label_rules = QtWidgets.QLabel(self.centralwidget)
        self.Label_rules.setGeometry(QtCore.QRect(135, 230, 120, 16))
        self.Label_rules.setAlignment(QtCore.Qt.AlignCenter)
        self.Label_rules.setObjectName("Label_rules")

        self.Button_Run = QtWidgets.QPushButton(self.centralwidget)
        self.Button_Run.setGeometry(QtCore.QRect(120, 410, 311, 51))
        self.Button_Run.setObjectName("Button_Run")
        self.Button_Run.clicked.connect(self.run)

        self.file_import_frame = QtWidgets.QFrame(self.centralwidget)
        self.file_import_frame.setGeometry(QtCore.QRect(20, 100, 511, 100))
        self.file_import_frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.file_import_frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.file_import_frame.setObjectName("file_import_frame")

        self.Label_copyfrom = QtWidgets.QLabel(self.file_import_frame)
        self.Label_copyfrom.setGeometry(QtCore.QRect(10, 10, 79, 35))
        self.Label_copyfrom.setObjectName("Label_copyfrom")

        self.Label_destination = QtWidgets.QLabel(self.file_import_frame)
        self.Label_destination.setGeometry(QtCore.QRect(10, 50, 71, 31))
        self.Label_destination.setObjectName("Label_destination")

        self.Button_browse_copyfrom = QtWidgets.QPushButton(self.file_import_frame)
        self.Button_browse_copyfrom.setGeometry(QtCore.QRect(410, 10, 91, 41))
        self.Button_browse_copyfrom.setObjectName("Button_browse_copyfrom")
        self.Button_browse_copyfrom.clicked.connect(lambda: self.open_excel_file(self.textEdit_copyfrom))  # Added by me (browse func for copy file)

        self.Button_browse_destination = QtWidgets.QPushButton(self.file_import_frame)
        self.Button_browse_destination.setGeometry(QtCore.QRect(410, 50, 91, 41))
        self.Button_browse_destination.setObjectName("Button_browse_destination")
        self.Button_browse_destination.clicked.connect(lambda: self.open_excel_file(self.textEdit_destination))  # Added by me (browse func for destination file)

        self.textEdit_copyfrom = FileEdit(self.file_import_frame)
        self.textEdit_copyfrom.setGeometry(QtCore.QRect(90, 20, 319, 21))
        self.textEdit_copyfrom.setObjectName("textEdit_copyfrom")

        self.textEdit_destination = FileEdit(self.file_import_frame)
        self.textEdit_destination.setGeometry(QtCore.QRect(90, 60, 319, 21))
        self.textEdit_destination.setObjectName("textEdit_destination")

        self.label_select_excel_files = QtWidgets.QLabel(self.centralwidget)
        self.label_select_excel_files.setGeometry(QtCore.QRect(210, 70, 131, 21))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.label_select_excel_files.setFont(font)
        self.label_select_excel_files.setAlignment(QtCore.Qt.AlignCenter)
        self.label_select_excel_files.setObjectName("label_select_excel_files")

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
        self.Button_new_rule.setText(_translate("MainWindow", "New"))
        self.Button_edit_rule.setText(_translate("MainWindow", "Edit"))
        self.Button_remove_rule.setText(_translate("MainWindow", "Remove"))
        self.listWidget_rules.setSortingEnabled(False)
        self.Label_rules.setText(_translate("MainWindow", "Copy/Paste Macros"))
        self.Button_Run.setText(_translate("MainWindow", "RUN"))
        self.Label_copyfrom.setText(_translate("MainWindow", "Copy From:"))
        self.Label_destination.setText(_translate("MainWindow", "Destination:"))
        self.Button_browse_destination.setText(_translate("MainWindow", "Browse"))
        self.Button_browse_copyfrom.setText(_translate("MainWindow", "Browse"))
        self.label_select_excel_files.setText(_translate("MainWindow", "Select  Excel  Files"))


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
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.initUI(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

