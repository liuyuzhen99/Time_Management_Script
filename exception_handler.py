from PyQt5.QtCore import QObject, pyqtSignal


class ExceptionSignal(QObject):
    signal = pyqtSignal(str)

exception_signal_instance = ExceptionSignal()