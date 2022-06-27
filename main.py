import sys
import os
import shutil
import pandas as pd
from datetime import date
from dataclasses import dataclass

import win32com.client
from PyQt5 import QtWidgets, uic
from PyQt5.QtGui import QStandardItem, QStandardItemModel
from PyQt5.QtWidgets import QFileDialog, QMessageBox
from PyQt5.QtCore import QFileInfo

import excelreader
from fedex_api import FedexAPI

# region Constants
SITE_DETAILS = pd.read_excel("SiteList.xlsx")
SITE_DETAILS.set_index("Station#", inplace=True)
# endregion

class MainWindow(QtWidgets.QMainWindow):

    def __init__(self):
        super(MainWindow, self).__init__()
        uic.loadUi('./ui/main.ui', self)
        self.show()

        self.snowModel = QStandardItemModel()
        self.snowModel.setHorizontalHeaderLabels(["Project/Name", "Type", "Date"])
        self.snowTreeView.setModel(self.snowModel)
        self.snowTreeView.header().setDefaultSectionSize(200)

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec())
