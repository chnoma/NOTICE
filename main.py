import sys
import os
import shutil
import shelve

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
TYPE_SHIPMENT = 0
TYPE_DN_REQUEST = 1
PROJECT_SUPPORTING_TECH = 0
PROJECT_PVAAS = 1
PROJECT_SPECIALIZED_DEVICES = 3
PROJECT_OTHER = 4
# endregion


@dataclass
class Shipment:
    project: str
    location_code: str
    location_name: str
    ncs_so: str
    ncs_inv: str
    tracking_no: str
    carrier: str
    ship_date: str
    delivery_date: str


@dataclass
class DataEntry:
    """Class detailing information about a given shipment notification/missing dn entry"""
    project: str
    title: str
    type: int
    email_generated: bool
    site_code: int
    date_added: date
    date_sent: date
    data: list  # shipment notification: 0 - XLSX, 1 - PDF | DN Request: array of shipment dataclasses


class MainWindow(QtWidgets.QMainWindow):

    def __init__(self):
        super(MainWindow, self).__init__()
        uic.loadUi('./ui/main.ui', self)
        self.show()

        self.snowModel = QStandardItemModel()
        self.snowModel.setHorizontalHeaderLabels(["Project/Name", "Type", "Date"])

        root = self.snowModel.invisibleRootItem()
        root.appendRow([QStandardItem("Supporting Technologies")])
        root.appendRow([QStandardItem("PVaaS")])
        root.appendRow([QStandardItem("Specialized Devices")])
        root.appendRow([QStandardItem("Other")])

        self.snowTreeView.setModel(self.snowModel)
        self.snowTreeView.header().setDefaultSectionSize(200)

        self.shelve = shelve.open("./settings/registry")
        if "items" not in self.shelve.keys():
            self.shelve["items"] = []

        self.purchaseOrderBrowseButton.pressed.connect(self.browse_po)
        self.purchaseOrderOpenButton.pressed.connect(self.open_po)
        self.shipmentBrowseButton.pressed.connect(self.browse_xlsx)
        self.shipmentOpenButton.pressed.connect(self.open_xlsx)
        self.purchaseOrderLineEdit.textChanged.connect(self.validate_files)
        self.shipmentLineEdit.textChanged.connect(self.validate_files)

    def close(self):
        self.shelve.close()

    def validate_files(self):
        self.saveButton.setEnabled(
            os.path.exists(self.purchaseOrderLineEdit.text()) and os.path.exists(self.shipmentLineEdit.text()))

    def browse_po(self):
        file = QFileDialog.getOpenFileName(self, "QFileDialog.getOpenFileName()", "", "Purchase Orders (*.pdf)")[0]
        if file == "":
            return
        self.purchaseOrderLineEdit.setText(file)

    def open_po(self):
        if not os.path.exists(self.purchaseOrderLineEdit.text()):
            return
        else:
            os.system('"' + self.purchaseOrderLineEdit.text() + '"')

    def browse_xlsx(self):
        file = QFileDialog.getOpenFileName(self, "QFileDialog.getOpenFileName()", "", "Excel Files (*.xlsx)")[0]
        if file == "":
            return
        self.shipmentLineEdit.setText(file)

    def open_xlsx(self):
        if not os.path.exists(self.shipmentLineEdit.text()):
            return
        else:
            os.system('"' + self.shipmentLineEdit.text() + '"')


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec())
