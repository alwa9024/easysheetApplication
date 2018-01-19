#!/usr/bin/env python3

# import all the necessary libraries for the construction of the application
import sys
import classTable, classAboutMsg
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *

########################################################################################################################

# the main window creates an instance of a table called self.form_widget and contains the menu bar with file, edit, and
# help functions

class Sheet(QMainWindow):

    # initialize the sheet
    def __init__(self, rows, columns):
        super().__init__()

        # the r and c values passed in are strings. Make them into integers
        rows = int(rows)
        columns = int(columns)

        # create the spreadsheet with initial dimensions specified by user
        self.form_widget = classTable.MyTable(rows, columns)

        # set the geometry of the window and do some formatting
        self.setWindowTitle("Easy Sheet Application")
        self.setWindowIcon(QIcon('ES_icon.png'))
        self.setGeometry(400, 250, 900, 700)

        # set the central widget as the newly created instance of form_widget
        self.setCentralWidget(self.form_widget)

        # if the number of rows or columns given to the table was 0, close the App
        if str(rows) == '0' or str(columns) == '0':
            self.quit_app()

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
        sys.exit()

    # define function to create instance of AboutMsg class (QDialog)
    def about_app(self):
        self.msg = classAboutMsg.AboutMsg()

########################################################################################################################