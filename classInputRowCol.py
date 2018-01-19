#!/usr/bin/env python3

# import all the necessary libraries for the construction of the application
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *

########################################################################################################################

# this class takes user input for a number of rows and columns and kicks off the meat of the app with those inputs to
# initialize a spreadsheet

class inputRowCol(QDialog):

    # initialize
    def __init__(self):
        super().__init__()

        # initialize the number of rows and columns as 0
        self.numRows = 0
        self.numCols = 0

        # set the window title and geometry
        self.setWindowTitle("Dimensions For Spreadsheet")
        self.setWindowIcon(QIcon('ES_icon.png'))

        # create the labels
        self.labelNumRow = QLabel("Number of Rows: ")
        self.labelNumCol = QLabel("Number of Columns: ")

        # create the line edits for user input and place a validator on them
        self.inputNumRow = QLineEdit(str(20))
        self.inputNumRow.setValidator(QIntValidator(1, 100))
        self.inputNumCol = QLineEdit(str(10))
        self.inputNumCol.setValidator(QIntValidator(1, 100))

        # create a submit button and add an icon to it
        self.subRowColButton = QPushButton("Submit")
        self.subRowColButton.setIcon(QIcon('sub_img.png'))

        # create a cancel button and add an icon to it
        self.cancelNewButton = QPushButton("Cancel")
        self.cancelNewButton.setIcon(QIcon('close_img.png'))

        # adjust the size of the button and line edits to be prettier
        width = max(self.labelNumRow.sizeHint().width(), self.labelNumCol.sizeHint().width())
        self.labelNumRow.setFixedWidth(width)
        self.labelNumCol.setFixedWidth(width)

        # add layouts for the rows/columns input box
        mainLayout = QVBoxLayout()
        labelEditLayout1 = QHBoxLayout()
        labelEditLayout2 = QHBoxLayout()
        buttonLayout = QHBoxLayout()
        buttonLayout.addStretch(0)

        # add the widgets to their respective layouts
        labelEditLayout1.addWidget(self.labelNumRow)
        labelEditLayout1.addWidget(self.inputNumRow)
        labelEditLayout2.addWidget(self.labelNumCol)
        labelEditLayout2.addWidget(self.inputNumCol)
        buttonLayout.addWidget(self.cancelNewButton)
        buttonLayout.addWidget(self.subRowColButton)


        # add sub-layouts onto main layout
        self.setLayout(mainLayout)
        mainLayout.addLayout(labelEditLayout1)
        mainLayout.addLayout(labelEditLayout2)
        mainLayout.addLayout(buttonLayout)
        #mainLayout.addWidget(self.subRowColButton, alignment=Qt.AlignRight)

        # leave the submit button disabled unless both line edits have values
        self.subRowColButton.setDisabled(False)
        self.inputNumCol.textChanged.connect(self.disableSubButton)
        self.inputNumRow.textChanged.connect(self.disableSubButton)

        # if the values entered are positive, non-zero integers, go to function
        self.subRowColButton.clicked.connect(self.clickSubmitRowsCols)

        # when the cancel button is clicked, close out of self without changing the table
        self.cancelNewButton.clicked.connect(self.cancelFunction)

    # function to check if the inputted values are valid each time the user changes them. If the line edits are empty,
    # the submit button is submitted. If they have values in them, the submit button works
    def disableSubButton(self):
        if len(self.inputNumCol.text()) == 0 or len(self.inputNumRow.text()) == 0:
            self.subRowColButton.setDisabled(True)
        else:
            self.subRowColButton.setDisabled(False)

    # function to take the values the user has typed upon clicking submit, and assign them to the row and column values
    def clickSubmitRowsCols(self):

        # get the number of desired rows and columns from the user
        self.numRows = self.inputNumRow.text()
        self.numCols = self.inputNumCol.text()

        # close the dialog box and initialize the sheet with the inputted number of columns and rows
        self.close()

    # function to cancel out of the dialog box without altering what the table currently looks like
    def cancelFunction(self):
        # assign the number of rows and columns to be 0. This doesn't actually make them 0, but is used in the table
        # class to detect when this screen was cancelled to do its thing
        self.numRows = 0
        self.numCols = 0

        # close out of the dialog box
        self.close()

