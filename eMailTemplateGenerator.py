import sys
import os
import re
import datetime
import time
import pandas as pd
import win32com.client
from threading import Thread
from PyQt5.QtWidgets import QWidget, QApplication, QMainWindow, QDialog, QFileDialog, QMessageBox, QErrorMessage
from PyQt5 import QtCore, QtGui, uic
from PyQt5.QtGui import QIntValidator
from concurrent.futures import Future

txtEmailTemplatePath = ""
txtDataListPath = ""
txtDestinationPath = ""


def call_with_future(fn, future, args, kwargs):
    try:
        result = fn(*args, **kwargs)
        future.set_result(result)
    except Exception as exc:
        future.set_exception(exc)

def threaded(fn):
    def wrapper(*args, **kwargs):
        future = Future()
        Thread(target=call_with_future, args=(fn, future, args, kwargs)).start()
        return future
    return wrapper

class UI(QMainWindow):
    def __init__(self):
        super(UI, self).__init__()
        uic.loadUi(os.path.dirname(__file__) + '/resource/main.ui', self)
        self.setWindowIcon(QtGui.QIcon(os.path.dirname(__file__) +  '/resource/logo.png'))
        self.btnBrowseTemplate.clicked.connect(self.choose_source_data)
        self.btnBrowseDataList.clicked.connect(self.choose_data_list)
        self.btnBrowseDestination.clicked.connect(self.choose_destination_folder)

        self.buttonConfirm.accepted.connect(self.do_process)
        self.buttonConfirm.rejected.connect(self.close)
        self.onloaded()
        self.show()

    def onloaded(self):
        self.progressBar.hide()
        self.progressBar.value = 0
        self.txtEmailTemplatePath.setText("")
        self.txtDataListPath.setText("")
        self.txtDestinationPath.setText("")
        self.txtEmailTemplatePath.setReadOnly(True)
        self.txtDataListPath.setReadOnly(True)
        self.txtDestinationPath.setReadOnly(True)

    def choose_source_data(self):
        global txtEmailTemplatePath
        fileName, _ = QFileDialog.getOpenFileName(self,"QFileDialog.getOpenFileName()", "","HTML Template (*.html *htm *msg)")
        if fileName:
            self.txtEmailTemplatePath.setText(fileName)
            txtEmailTemplatePath = fileName

    def choose_data_list(self):
        global txtDataListPath
        fileName, _ = QFileDialog.getOpenFileName(self,"QFileDialog.getOpenFileName()", "","Excel File (*.xlsx *.xls)")
        if fileName:
            self.txtDataListPath.setText(fileName)
            txtDataListPath = fileName
            
    def choose_destination_folder(self):
        global txtDestinationPath
        directory_path = str(QFileDialog.getExistingDirectory(self, "Select Directory"))
        if directory_path:
            self.txtDestinationPath.setText(directory_path)
            txtDestinationPath = directory_path

    def do_process(self):
        global txtEmailTemplatePath, txtDestinationPath, txtDataListPath

        if not txtEmailTemplatePath or not txtDataListPath:
            msg = QMessageBox.about(self, "Warning", "Please select your eMail template and data source before proceeds.")
            return

        self.progressBar.show()
        try:
            future_result = self.generate_mail()
            result = future_result.result()
            if not os.path.exists(txtDestinationPath):
                os.mkdir(txtDestinationPath)
            time.sleep(1)

            msg = QMessageBox.about(self, "Completed", "Your request was proceesed completly.")

        except Exception as e:
            msg = QMessageBox.about(self, "Something is wrong retry again. ", str(e))
            
        time.sleep(1)
        self.onloaded()

    def close(self):
        sys.exit(app.exec_())

    @threaded
    def generate_mail(self):
        global txtEmailTemplatePath, txtDestinationPath, txtDataListPath
        try:
            if os.path.exists(txtEmailTemplatePath) and os.path.exists(txtDataListPath):
                data_list = pd.read_excel(txtDataListPath)
                progress_value_step = 100/data_list.shape[0]

                for index, item in data_list.iterrows():
                    header_key = item.keys()

                    newMail_body = None
                    if txtEmailTemplatePath.endswith((".html", ".htm")):
                        with open(txtEmailTemplatePath, encoding="utf-8") as f:
                            newMail_body = f.read()
                    elif txtEmailTemplatePath.endswith(".msg"):
                        obj = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
                        msg = obj.OpenSharedItem(txtEmailTemplatePath)
                        newMail_body = str(msg.HTMLBody)
                        del obj, msg


                    for key in header_key:
                        newMail_body = re.sub(r"(\["+key+r"\])", str(data_list.loc[index,key]), newMail_body)
                    mail_subject = f"BROKER TURN OVER FOR {data_list.loc[index,'COMPANY_NAME']}, Waybill No"

                    olMailItem = 0x0
                    obj = win32com.client.Dispatch("Outlook.Application")
                    newMail = obj.CreateItem(olMailItem)
                    newMail.Subject = mail_subject[:-4]
                    newMail.BodyFormat = 2
                    newMail.HTMLBody = newMail_body
                    newMail.To = ';'.join(data_list.loc[index,"EMAIL_TO"].split(','))
                    newMail.CC = ';'.join(data_list.loc[index,"EMAIL_CC"].split(','))
                    newMail.SaveAs(os.path.join(txtDestinationPath, str(f"{mail_subject}.msg")))
                    self.progressBar.setValue(int(progress_value_step * (index + 1)))
            
            return True
        except Exception as e:
            raise e


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = UI()
    sys.exit(app.exec_())