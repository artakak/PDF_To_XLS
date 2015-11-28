#!-*-coding:utf-8-*-
import sys
# import PyQt4 QtCore and QtGui modules
from PyQt4.QtCore import *
from PyQt4.QtGui import *
# from PyQt4 import uic
from ui_test import Ui_MainWindow

s
#( Ui_MainWindow, QMainWindow ) = uic.loadUiType( 'ui_test.ui' )

class MainWindow(QMainWindow):
    """MainWindow inherits QMainWindow"""

    def __init__(self, parent=None):
        QMainWindow.__init__(self, parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

    def __del__(self):
        self.ui = None


#-----------------------------------------------------#
if __name__ == '__main__':
    # create application
    app = QApplication(sys.argv)
    app.setApplicationName('PDF_To_XLS')

    # create widget
    w = MainWindow()
    w.setWindowTitle('PDF_To_XLS')
    w.show()

    # connection
    QObject.connect(app, SIGNAL('lastWindowClosed()'), app, SLOT('quit()'))

    # execute application
    sys.exit(app.exec_())