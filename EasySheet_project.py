#!/usr/bin/env python3

# import all the necessary libraries for the construction of the application
import sys, os, csv, time, re
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
########################################################################################################################

# Create the spreadsheet table and its functions

# ***** ADD: MATH FUNCTIONS, RIGHT CLICK FUNCTIONS *****
class MyTable(QTableWidget):

    # initialize self.value as an empty string for use in the paste/ copy cell functions
    value = ''

##### INITIALIZE THE TABLE #####
    def __init__(self, rows, columns):
        super().__init__(rows,columns)

        # set to only single cell selection
        self.setSelectionMode(QAbstractItemView.SingleSelection)

        # initialize each table widget item as an empty string
        for row_inc in range(rows):
            for col_inc in range(columns):
                self.setItem(row_inc,col_inc, QTableWidgetItem(""))

        # FORMATTING OF TABLE
        # style the table headers light cyan
        style_lightcyan = "::section {""background-color: lightcyan; }"
        self.horizontalHeader().setStyleSheet(style_lightcyan)
        self.verticalHeader().setStyleSheet(style_lightcyan)
        # stretch the table to take up the entire whitespace of the window
        self.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)

    ##### MATH SECTION #####

        # have a regular expression to check if the contents of the cell are in a math string format
        self.regex = re.compile('^[ *[(?:\-?\d *) *[\-+/*]*[(?:\-?\d)][\-+*/]?]*|[\-?\d[\-+/*]*[\-?\d][\-+/*]?]*$')

        # when a cell is changed, connect to the math function where the cell will be evaluated.
        self.itemChanged.connect(self.math_stuff)

    # This function evaluates a math expression within a cell.
    def math_stuff(self, item):
        if self.regex.fullmatch(item.text()):
            # if it's a valid mathematical statement, set the text to the result
            try:
                math_evald = eval(item.text())
                item.setText(str(math_evald))
            # if it's not valid but satisfies the regular expression, set the item to an empty string and print an error
            except:
                error_message = QMessageBox()
                error_message.setIcon(QMessageBox.Critical)
                error_message.setWindowTitle("Woopsie-Daisy")
                error_message.setText("Warning: invalid mathematical statement.")
                error_message.exec()
                item.setText("")

    ##### END MATH SECTION #####


##### CONTEXT MENU #####

    # CREATE A CONTEXT MENU FOR THE TABLE WIDGET
    def contextMenuEvent(self, event):
        # create a menu
        menu = QMenu(self)

        # define the actions for the menu
        insert_row_act = menu.addAction("Insert Row")
        insert_col_act = menu.addAction("Insert Column")
        remove_specific_row_act = menu.addAction("Remove Row")
        remove_specific_col_act = menu.addAction("Remove Column")
        upper_act = menu.addAction("Uppercase cell")
        lower_act = menu.addAction("Lowercase cell")

        # execute the menu when a cell is right clicked and get the position
        action = menu.exec_(self.mapToGlobal(event.pos()))

        # functionality for clicking add a specific row action
        if action == insert_row_act:
            # get the specific row number, then insert a row in the location directly below it
            curRow = self.currentRow()
            self.insertRow(curRow)
            # make each of the cells in this new row filled by an empty string
            for col_inc in range(self.columnCount()):
                self.setItem(curRow, col_inc, QTableWidgetItem(""))

        # functionality for clicking remove a specific row action
        if action == remove_specific_row_act:
            if self.rowCount() > 1:
                # get the specific row number
                curRow = self.currentRow()
                # delete the row from the table
                self.removeRow(curRow)
                self.horizontalHeader().setStretchLastSection(True)
                self.verticalHeader().setStretchLastSection(True)
            else:
                # remove all the contents of the remaining columns
                for col_inc in range(self.columnCount()):
                    self.item(0, col_inc).setText("")

        # functionality for clicking add a specific column action
        if action == insert_col_act:
            # get the current column number, then insert a column in the location directly to the right of it
            curCol = self.currentColumn()
            self.insertColumn(curCol)
            # make each of the cells in this new column filled by an empty string
            for row_inc in range(self.rowCount()):
                self.setItem(row_inc, curCol, QTableWidgetItem(""))

        # functionality for clicking remove a specific column action
        if action == remove_specific_col_act:
            if self.columnCount() > 1:
                # get the specific row number
                curCol = self.currentColumn()
                # delete the row from the table
                self.removeColumn(curCol)
                self.horizontalHeader().setStretchLastSection(True)
                self.verticalHeader().setStretchLastSection(True)
            else:
                # remove the contents of the remaining columns
                for row_inc in range(self.rowCount()):
                    self.item(row_inc, 0).setText("")

        # functionality for capitalizing the contents of the cell selected
        if action == upper_act:
            # get the current row and column of the selected cell
            curRow = self.currentRow()
            curCol = self.currentColumn()
            # get the text in that cell and save it to a variable to be capitalized
            uneditedText = self.item(curRow,curCol).text()
            editedText = uneditedText.upper()
            self.item(curRow,curCol).setText(editedText)

        # functionality for lower-casing the contents of the cell selected
        if action == lower_act:
            # get the current row and column of the selected cell
            curRow = self.currentRow()
            curCol = self.currentColumn()
            # get the text in that cell and save it to a variable to be lower-cased
            uneditedText = self.item(curRow,curCol).text()
            editedText = uneditedText.lower()
            self.item(curRow,curCol).setText(editedText)


##### MENU ITEM FUNCTIONS #####

    # FUNCTION TO OPEN A CSV FILE INTO A TABLE
    def open_sheet(self):
        # declare the path name as the one the user chooses from a file dialog (filter for only csv files)
        path = QFileDialog.getOpenFileName(self, 'Open CSV', os.getenv('HOME'), 'CSV(*.csv)')
        # if the path is real (not empty: i.e. an actual file), open the csv
        if path[0] != '':
            with open(path[0], newline='') as csv_file:

                # use csv pkg to assign the contents of the file to my_file
                my_file = csv.reader(csv_file, delimiter=',', quotechar='|')

                # assign the data from the file into a tuple
                my_data = list(my_file)

                # assign row and column counts as the length of the tuple (number of lists) and the length of each list
                rowCount = len(my_data)
                colCount = len(my_data[0])

                # if the csv file is empty, open up a 1x1 "table"
                if(rowCount <= 0 or colCount <= 0):
                    self.setColumnCount(1)
                    self.setRowCount(1)

                else:

                    # set the default row/column count of the table to be filled in with following for loop
                    self.setRowCount(rowCount)
                    self.setColumnCount(colCount)

                    # based on the data in the file, fill in the table cell by cell
                    for row in range(rowCount):
                        for column in range(colCount):
                            item = QTableWidgetItem(my_data[row][column])
                            self.setItem(row, column, item)
        # note that the user is not able to select an empty path within the file dialog, so don't worry about an else:


    # FUNCTION TO SAVE THE SHEET
    def save_sheet(self):
        path = QFileDialog.getSaveFileName(self, 'Save CSV', os.getenv('HOME'), 'CSV(*.csv)')

        # if the user includes .csv in the file name, save file as specified name
        if '.csv' in str(path[0]):
            if path[0] != '':
                with open(path[0], 'w') as csv_file:
                    writer = csv.writer(csv_file, dialect = 'excel')
                    for row in range(self.rowCount()):
                        row_data = []
                        for column in range(self.columnCount()):
                            item = self.item(row,column)
                            row_data.append(item.text())
                        writer.writerow(row_data)

        # set up a case for if the user doesn't put .csv in the file name; add .csv in
        else:
            if path[0] != '':
                path_withCSV = str(path[0])+'.csv'
                with open(path_withCSV, 'w') as csv_file:
                    writer = csv.writer(csv_file, dialect = 'excel')
                    for row in range(self.rowCount()):
                        row_data = []
                        for column in range(self.columnCount()):
                            item = self.item(row,column)
                            row_data.append(item.text())
                        writer.writerow(row_data)
        # if the path is empty, it won't let the user save, so don't worry 'bout it


    # FUNCTION TO ADD A ROW TO TABLE
    def add_row(self):
        # add the row to the end of the table
        currentNumRows = self.rowCount()
        self.insertRow(currentNumRows)

        # loop through each QTableWidget and add an item with an empty string to each of the new cells
        currentNumCols = self.columnCount()
        for row_of_cells in range(currentNumCols):
            self.setItem(currentNumRows, row_of_cells, QTableWidgetItem(""))


    # FUNCTION TO ADD A COLUMN TO TABLE
    def add_column(self):
        # add a column to the end of the table
        currentNumCols = self.columnCount()
        self.insertColumn(currentNumCols)

        # loop through each new QTableWidget and add an item with an empty string to each of the new cells
        currentNumRows = self.rowCount()
        for col_of_cells in range(currentNumRows):
            self.setItem(col_of_cells, currentNumCols, QTableWidgetItem(""))


    # FUNCTION TO COPY THE CONTENTS OF A CELL
    def copy_cell(self):
        # get the highlighted cell
        curRow = self.currentRow()
        curCol = self.currentColumn()

        self.value = self.item(curRow, curCol).text()


    # FUNCTION TO PASTE THE CONTENTS OF A COPIED CELL INTO A NEW CELL
    def paste_cell(self):
        # get the highlighted cell
        curRow = self.currentRow()
        curCol = self.currentColumn()

        # if there is a value assigned to self.value, use it. If not, do nothing
        if self.value:
            self.item(curRow, curCol).setText(self.value)


    # FUNCTION TO CLEAR OUT A TABLE AND REINITIALIZE IT
    def new_table(self):

        # display the dialog that prompts the user for a number of rows and columns
        self.new_prompt = inputRowCol()
        self.new_prompt.exec()


        # set the number of rows & columns to the values from the input row and column dialog if they were given
        if int(self.new_prompt.numRows) > 0:
            print("made it")
            self.numRows = self.new_prompt.numRows
            self.setRowCount(int(self.numRows))
            self.numCols = self.new_prompt.numCols
            self.setColumnCount(int(self.numCols))

            # initialize each new table widget item as an empty string
            for row_inc in range(int(self.numRows)):
                for col_inc in range(int(self.numCols)):
                    self.setItem(row_inc, col_inc, QTableWidgetItem(""))

            # stretch the table to fit the entire window
            self.horizontalHeader().setStretchLastSection(True)
            self.verticalHeader().setStretchLastSection(True)





########################################################################################################################

# the main window creates an instance of a table called self.form_widget and contains the menu bar with file, edit, and
# help functions

class Sheet(QMainWindow):

    # create a custom signal to be caught by MyTable and used to initialize a new table
    new_table_sig = pyqtSignal()

    # initialize the sheet
    def __init__(self, rows, columns):
        super().__init__()

        # the r and c values passed in are strings. Make them into integers
        rows = int(rows)
        columns = int(columns)

        # create the spreadsheet with initial dimensions specified by user
        self.form_widget = MyTable(rows, columns)

        # set the geometry of the window and do some formatting
        self.setWindowTitle("Easy Sheet Application")
        self.setWindowIcon(QIcon('ES_icon.png'))
        self.setGeometry(400, 250, 900, 700)

        # set the central widget as the newly created instance of form_widget
        self.setCentralWidget(self.form_widget)

        # set up a menu in main window
        bar = self.menuBar()
        file = bar.addMenu('File')
        edit = bar.addMenu('Edit')
        help = bar.addMenu('Help')


        ##### BEGIN MENU ACTIONS ######

        # FILE MENU ACTIONS
        # create a new action for file menu and link to new sheet function
        new_action = QAction('New', self)
        new_action.setShortcut('Ctrl+N')
        new_action.setIcon(QIcon('new_img.png'))
        file.addAction(new_action)
        new_action.triggered.connect(self.form_widget.new_table)

        # create a save action for file menu and link to the save function
        save_action = QAction('Save', self)
        save_action.setShortcut('Ctrl+S')
        save_action.setIcon(QIcon('save_img.png'))
        file.addAction(save_action)
        save_action.triggered.connect(self.form_widget.save_sheet)

        # create an open action for file menu and link to the open function
        open_action = QAction('Open', self)
        open_action.setShortcut('Ctrl+O')
        open_action.setIcon(QIcon('open_img.png'))
        file.addAction(open_action)
        open_action.triggered.connect(self.form_widget.open_sheet)

        # create a quit action for file menu and link to the quit function
        quit_action = QAction('Quit', self)
        quit_action.setShortcut('Ctrl+X')
        quit_action.setIcon(QIcon('close_img.png'))
        file.addAction(quit_action)
        quit_action.triggered.connect(self.quit_app)


        # EDIT MENU ACTIONS
        # create an add row action for the edit menu and link to add_row function
        add_row_action = QAction('Append Row', self)
        add_row_action.setShortcut('Ctrl+R')
        add_row_action.setIcon(QIcon('add_row_img.png'))
        edit.addAction(add_row_action)
        add_row_action.triggered.connect(self.form_widget.add_row)

        # create and add column action for the edit menu and link to add_col function
        add_col_action = QAction('Append Column', self)
        add_col_action.setShortcut('Ctrl+I')
        add_col_action.setIcon(QIcon('add_col_img.png'))
        edit.addAction(add_col_action)
        add_col_action.triggered.connect(self.form_widget.add_column)

        # create a copy cell action for the edit menu and link to copy_cell function
        copy_cell_action = QAction('Copy Cell', self)
        copy_cell_action.setShortcut('Ctrl+C')
        copy_cell_action.setIcon(QIcon('copy_img.png'))
        edit.addAction(copy_cell_action)
        copy_cell_action.triggered.connect(self.form_widget.copy_cell)

        # create a paste cell action for the edit menu and link to paste_cell function
        paste_cell_action = QAction('Paste Cell', self)
        paste_cell_action.setShortcut('Ctrl+V')
        paste_cell_action.setIcon(QIcon('paste_img.png'))
        edit.addAction(paste_cell_action)
        paste_cell_action.triggered.connect(self.form_widget.paste_cell)


        # HELP MENU ACTIONS
        # create an about the app action for the help menu and link to about app function
        disp_about_action = QAction('About', self)
        disp_about_action.setShortcut('Ctrl+A')
        disp_about_action.setIcon(QIcon('about_img.png'))
        help.addAction(disp_about_action)
        disp_about_action.triggered.connect(self.about_app)

        ##### END MENU ACTIONS #####

    # FUNCTIONS FOR MENU BAR

    # define function to exit out of the application
    def quit_app(self):
        qApp.quit()

    # define function to create instance of AboutMsg class (QDialog)
    def about_app(self):
        self.msg = AboutMsg()


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
        self.dialogLayout.addWidget(self.textWidget)

        # make and add an ok button that will close the dialog box when clicked
        ok_button = QPushButton("OK")
        self.dialogLayout.addWidget(ok_button, alignment = Qt.AlignJustify)
        ok_button.clicked.connect(self.close)


        # show the dialog
        self.exec()

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

########################################################################################################################

# run the application
def main():

    # define the application
    app = QApplication(sys.argv)

    # assign an image of watermelons for the splash screen
    splash_pixmap = QPixmap('Spreadsheet_img.png')

    # create a splash screen, called splash
    splash = QSplashScreen(splash_pixmap)

    # show the splash screen
    splash.show()
    #splash.showMessage("<h1><font color='white'>Spreadsheet Application</font></h1>", Qt.AlignCenter)

    # declare how long the splash screen will run
    t = time.time()
    while time.time() < t + 5:
        app.processEvents()
        # once the splash screen is no longer visible, stop processing this way
        if(not splash.isVisible()):
            break

    # close the splash page and open the window that prompts the user for a number of columns and rows
    form = inputRowCol()
    splash.finish(form)
    form.exec()

    # kick off the spreadsheet with the number of rows and columns
    startItUp = Sheet(form.numRows, form.numCols)
    startItUp.show()

    sys.exit(app.exec_())

########################################################################################################################

if(__name__ == "__main__"):
   main()
