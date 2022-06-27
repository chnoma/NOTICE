import sys
import os
import shutil
import shelve

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


@dataclass
class shipmentNotification:
    """Class detailing information about a given shipment notification/missing dn entry"""
    project: str
    title: str
    is_shipment_notification: bool
    email_generated: bool
    site_code: int
    date_added: date
    date_sent: date
    uid: int


# region Constants
SITE_DETAILS = pd.read_excel("./settings/SiteList.xlsx")
SITE_DETAILS.set_index("Station#", inplace=True)
# endregion


class MainWindow(QtWidgets.QMainWindow):

    def __init__(self):
        super(MainWindow, self).__init__()
        uic.loadUi('./ui/main.ui', self)
        self.show()
        self.shelve = shelve.open("./settings/registry.db")
        self.snowModel = QStandardItemModel()
        self.snowModel.setHorizontalHeaderLabels(["Project/Name", "Type", "Date"])
        root = self.snowModel.invisibleRootItem()
        root.appendRow([QStandardItem("Supporting Technologies")])
        root.appendRow([QStandardItem("PVaaS")])
        root.appendRow([QStandardItem("Specialized Devices")])
        root.appendRow([QStandardItem("Other")])

        self.snowTreeView.setModel(self.snowModel)
        self.snowTreeView.header().setDefaultSectionSize(200)

    def close(self) -> bool:
        self.shelve.close()


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec())
