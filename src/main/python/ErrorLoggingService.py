import logging
import os
import AppInfo

logFileDir = os.getenv('APPDATA') + '\\' + AppInfo.Name
if not os.path.exists(logFileDir):
    os.makedirs(logFileDir)
logFilePath = logFileDir + '\\pbi_theme_generator.log'

logging.basicConfig(
    filename=logFilePath,
    filemode='w',
    level=logging.DEBUG
)


def LogException(exception):
    logging.exception(exception)
    return ReadException()


def ReadException():
    try:
        if os.path.isfile(logFilePath):
            with open(logFilePath, 'r') as logFile:
                return logFile.read()
        else:
            print('No file')
    except Exception as e:
        logging.exception('Error reading log file')


try:
    from PyQt5.QtWidgets import QMessageBox
    from PyQt5.QtCore import Qt

    def ShowErrorDialog(errorMessage):

        messageBoxError = QMessageBox()
        messageBoxError.setWindowTitle('Error in Application')
        messageBoxError.setIcon(QMessageBox.Critical)
        messageBoxError.setTextFormat(Qt.RichText)
        messageBoxErrorText = 'Please click on show detail button and copy text from text box and create new ' \
                              'issue on below link.<br/>'
        messageBoxErrorText += '<a href="' + AppInfo.GitHubRepoIssuesURL + '/new">Click here to create issue</a><br/>'
        messageBoxErrorText += 'Below error message is also stored in ' + logFilePath
        messageBoxError.setText(messageBoxErrorText)
        messageBoxError.setStandardButtons(QMessageBox.Ok)
        messageBoxError.setDetailedText(errorMessage)

        messageBoxError.exec_()

except ImportError:
    logging.exception('No PyQt installed.')
except Exception as e:
    logging.exception(e)
