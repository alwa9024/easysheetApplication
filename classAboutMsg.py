#!/usr/bin/env python3

# import all the necessary libraries for the construction of the application
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *

########################################################################################################################

# dialog box (message window for the 'help' menu about button. An instance is created when the about action is triggered
class AboutMsg(QDialog):
    def __init__(self):
        QDialog.__init__(self)

        # format the window
        self.setGeometry(650, 550, 350, 170)
        self.setWindowTitle("About the Application")
        self.setWindowIcon(QIcon('ES_icon.png'))

        # set the layout of the dialog and add text to it; make it so the user can't write in the text edit
        self.dialogLayout = QVBoxLayout(self)
        self.textWidget = QTextEdit("This application is a simple spreadsheet developer. Easy Sheet allows users to create customized spreadheets capable of performing basic mathematical expressions as well as text. Sheets developed with Easy Sheet can be saved, reopened, and edited. A user may also choose to open any .csv file to edit using Easy Sheet.")
        self.textWidget.setReadOnly(True)
        self.textWidget.setAlignment(Qt.AlignJustify)
        self.dialogLayout.addWidget(self.textWidget)

        # make and add an ok button that will close the dialog box when clicked
        ok_button = QPushButton("OK")
        self.dialogLayout.addWidget(ok_button, alignment = Qt.AlignJustify)
        ok_button.clicked.connect(self.close)


        # show the dialog
        self.exec()