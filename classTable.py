#!/usr/bin/env python3

# import all the necessary libraries for the construction of the application
import os, csv, re
import classInputRowCol
from PyQt5.QtWidgets import *

########################################################################################################################

# Create the spreadsheet table and its functions

# ***** ADD: MATH FUNCTIONS, RIGHT CLICK FUNCTIONS *****
class MyTable(QTableWidget):

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
                try:
                    rowCount = len(my_data)
                    colCount = len(my_data[0])

                    # set the default row/column count of the table to be filled in with following for loop
                    self.setRowCount(rowCount)
                    self.setColumnCount(colCount)

                    # based on the data in the file, fill in the table cell by cell
                    for row in range(rowCount):
                        for column in range(colCount):
                            item = QTableWidgetItem(my_data[row][column])
                            self.setItem(row, column, item)

                # if there isn't any data in the csv, set a blank 1 x 1 table
                except:
                    rowCount = 1
                    colCount = 1

                    # set the default row/column count of the table to be filled in with following for loop
                    self.setRowCount(rowCount)
                    self.setColumnCount(colCount)
                    # stretch out the table to fit the size of the window
                    self.horizontalHeader().setStretchLastSection(True)
                    self.verticalHeader().setStretchLastSection(True)

                # # if the csv file is empty, open up a 1x1 "table"
                # if(rowCount <= 0 or colCount <= 0):
                #     self.setColumnCount(1)
                #     self.setRowCount(1)

                #else:

                # # set the default row/column count of the table to be filled in with following for loop
                # self.setRowCount(rowCount)
                # self.setColumnCount(colCount)
                #
                # # based on the data in the file, fill in the table cell by cell
                # for row in range(rowCount):
                #     for column in range(colCount):
                #         item = QTableWidgetItem(my_data[row][column])
                #         self.setItem(row, column, item)
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
        self.new_prompt = classInputRowCol.inputRowCol()
        self.new_prompt.exec()


        # set the number of rows & columns to the values from the input row and column dialog if they were given
        if int(self.new_prompt.numRows) > 0:
            self.numRows = self.new_prompt.numRows
            self.setRowCount(int(self.numRows))
            self.numCols = self.new_prompt.numCols
            self.setColumnCount(int(self.numCols))

            # initialize each new table widget item as an empty string
            for row_inc in range(int(self.numRows)):
                for col_inc in range(int(self.numCols)):
                    self.setItem(row_inc, col_inc, QTableWidgetItem(""))

            # stretch out the table to fit the size of the window
            self.horizontalHeader().setStretchLastSection(True)
            self.verticalHeader().setStretchLastSection(True)
