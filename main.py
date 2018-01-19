#!/usr/bin/env python3

# import all the necessary libraries for the construction of the application
import sys, time
import classSheet, classInputRowCol
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *

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
    form = classInputRowCol.inputRowCol()
    splash.finish(form)
    form.exec()

    # kick off the spreadsheet with the number of rows and columns
    startItUp = classSheet.Sheet(form.numRows, form.numCols)
    startItUp.show()

    sys.exit(app.exec_())

########################################################################################################################

if(__name__ == "__main__"):
   main()