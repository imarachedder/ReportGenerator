import sys

from PyQt6.QtCore import QObject, pyqtSlot,pyqtSignal, QRunnable
from PyQt6.QtWidgets import QMessageBox
from icecream import ic
import traceback

def log_uncaught_exceptions(ex_cls, ex, tb):
    text = '{}: {}:\n'.format(ex_cls.__name__, ex)

    text += ''.join(traceback.format_tb(tb))

    print(text)
    QMessageBox.critical(None, 'Error', text)
    quit()
sys.excepthook = log_uncaught_exceptions
class WorkerSignals(QObject):
    started = pyqtSignal()
    finish = pyqtSignal()
    error = pyqtSignal(tuple)
    result = pyqtSignal(dict)
    progress = pyqtSignal(int)


class Worker(QRunnable):
    '''
    Worker thread

    Inherits from QRunnable to handler worker thread setup, signals and wrap-up.

    :param callback: The function callback to run on this worker thread. Supplied args and
                     kwargs will be passed through to the runner.
    :type callback: function
    :param args: Arguments to pass to the callback function
    :param kwargs: Keywords to pass to the callback function

    '''

    def __init__(self, fn, *args, **kwargs):
        super(Worker, self).__init__()
        # Store constructor arguments (re-used for processing)
        self.result = None
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()
        kwargs['progress_callback'] = self.signals.progress

    @pyqtSlot()
    def run(self):
        '''
        Initialise the runner function with passed args, kwargs.
        '''
        try:
            self.signals.started.emit()
            self.result = self.fn(*self.args)
        except Exception as e:
           ic(e)
        else:
            self.signals.result.emit(self.result)  # Return the result of the processing
            #ic(self.result)
        finally:
            self.signals.finish.emit()  # Done
            #return self.result
            #ic(self.result)
