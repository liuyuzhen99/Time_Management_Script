import sys
from PyQt5 import QtCore, QtGui, QtWidgets
import excel2ppt_qt
from exception_handler import exception_signal_instance


def except_hook(typo, value, traceback):
    exception_signal_instance.signal.emit(f'{typo.__name__} - {value}')


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    sys.excepthook = except_hook
    MainWindow = QtWidgets.QMainWindow()
    Ui = excel2ppt_qt.Ui_MainWindow()
    Ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
