import datetime
import sys
import os
import shutil
import re
import shelve
import typing

from datetime import date
from dataclasses import dataclass

from PyQt5 import QtWidgets, uic
from PyQt5.QtGui import QStandardItem, QStandardItemModel
from PyQt5.QtWidgets import QFileDialog, QMessageBox
from PyQt5.QtCore import QFileInfo

import excelreader

# region Constants
SHELVE_FILENAME = "./settings/registry"
PROJECT_SUPPORTING_TECH = 0
PROJECT_PVAAS = 1
PROJECT_SPECIALIZED_DEVICES = 3
PROJECT_OTHER = 4
PROJECT_NAMES = ["Supporting Technologies",
                 "PVaaS",
                 "Specialized Devices",
                 "Other"]
# These variables are purely for code readability
ENTRY_TYPE_SHIPMENT = 0
ENTRY_TYPE_DN_REQUEST = 1
NODE_TYPES = ["Notification", "Request"]


# endregion


@dataclass
class Shipment:
    """Dataclass providing info about a given shipment within a Delivery Notification Request"""
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
    """Dataclass detailing information about a given shipment notification/missing dn entry"""
    project: int
    excel_file: str
    pdf_file: str
    title: str
    type: int
    email_generated: bool
    date_added: date
    data: list  # shipment notification: ShipmentNotification instance | DN Request: array of shipment dataclasses
    alive: bool = True
    date_sent: date = None

    def group_node_index(self):
        if self.alive:
            return 0
        return 1


class MainWindow(QtWidgets.QMainWindow):

    def __init__(self):
        super(MainWindow, self).__init__()
        uic.loadUi('./ui/main.ui', self)
        self.show()

        with shelve.open(SHELVE_FILENAME) as db:
            if "data_entries" not in db:
                db["data_entries"] = []

        self.snowModel = QStandardItemModel()
        self.roots = []
        self.listed_data_entries = {}
        self.snowTreeView.setModel(self.snowModel)
        self.snowTreeView.header().setDefaultSectionSize(120)
        self.reload_items()

        self.snowTreeView.selectionModel().selectionChanged.connect(self.taskSelectionChanged)
        self.purchaseOrderBrowseButton.pressed.connect(self.browse_po)
        self.purchaseOrderOpenButton.pressed.connect(self.open_po)
        self.shipmentBrowseButton.pressed.connect(self.browse_xlsx)
        self.shipmentOpenButton.pressed.connect(self.open_xlsx)
        self.purchaseOrderLineEdit.textChanged.connect(self.validate_files)
        self.shipmentLineEdit.textChanged.connect(self.validate_files)
        self.shipment_save.pressed.connect(self.save_shipment)

    def reload_items(self):
        self.snowModel.clear()
        self.snowModel.setHorizontalHeaderLabels(["Project/Name", "Type", "Date", "Status"])
        root = self.snowModel.invisibleRootItem()
        self.roots = {}

        for project in PROJECT_NAMES:
            root.appendRow([QStandardItem(project), QStandardItem("Project")])
            self.roots[project] = root.child(root.rowCount() - 1)
        for key in self.roots:
            root = self.roots[key]
            root.appendRow([QStandardItem("Active"), QStandardItem("Group")])
            root.appendRow([QStandardItem("Inactive"), QStandardItem("Group")])

        with shelve.open(SHELVE_FILENAME) as db:
            for entry in db["data_entries"]:
                if entry.type == ENTRY_TYPE_SHIPMENT:
                    root = self.roots[PROJECT_NAMES[entry.project]].child(entry.group_node_index())
                    root.appendRow([QStandardItem(entry.data.order_number),
                                    QStandardItem(NODE_TYPES[entry.type]),
                                    QStandardItem(date.strftime(entry.date_added, "%m/%d/%Y"))])
                    hash(hash(str(root.child(root.rowCount() - 1))))
                    self.listed_data_entries[hash(str(root.child(root.rowCount() - 1)))] = entry

    def save_shipment(self):
        excel_filename = self.shipmentLineEdit.text()
        pdf_filename = self.purchaseOrderLineEdit.text()

        shipment = excelreader.parse_shipment_notification(excel_filename)

        entry = DataEntry(
            project=self.shipment_project.currentIndex(),
            excel_file=QFileInfo(excel_filename).fileName(),
            pdf_file=QFileInfo(pdf_filename).fileName(),
            title=shipment.order_number,
            type=ENTRY_TYPE_SHIPMENT,
            email_generated=False,
            date_added=datetime.datetime.now(),
            data=shipment
        )

        if PROJECT_NAMES[entry.project] == PROJECT_NAMES[PROJECT_PVAAS] \
                and re.search("SCTASK", entry.data.order_number) is None:
            message = QMessageBox()
            message.setIcon(QMessageBox.Question)
            message.setWindowTitle("No SCTASK Found For PVaaS")
            message.setText("""Your 'Project' setting is set to PVaaS, but
            no SCTASK was found in the shipping notification.\n
            Are you sure you would like to continue adding this to PVaaS?""")  # FIX FORMATTING HERE
            message.exec()

        new_item_folder = f"./files/{PROJECT_NAMES[entry.project]}/{entry.data.order_number}/"
        try:
            os.makedirs(new_item_folder)
        except Exception as e:
            message = QMessageBox()
            message.setIcon(QMessageBox.Warning)
            message.setWindowTitle("Exception")
            message.setText(getattr(e, 'message', repr(e)))
            message.exec()
            return

        shutil.copy(pdf_filename, new_item_folder + entry.pdf_file)
        shutil.copy(excel_filename, f"{new_item_folder}{entry.data.order_number}.xlsx")
        wb = excelreader.generateSerialList(
            f"{new_item_folder}{entry.data.order_number}.xlsx")  # Generate one-page serial report
        wb.save(f"{new_item_folder}{entry.data.order_number}_SN.xlsx")  # Save the report to disk

        self.purchaseOrderLineEdit.setText("")
        self.shipmentLineEdit.setText("")
        with shelve.open(SHELVE_FILENAME, writeback=True) as db:
            db["data_entries"].append(entry)
        self.reload_items()

        message = QMessageBox()
        message.setIcon(QMessageBox.Information)
        message.setWindowTitle("Item Created")
        message.setText(f'New item "{entry.data.order_number}" created in {PROJECT_NAMES[entry.project]}.')
        message.exec()

    def taskSelectionChanged(self, selected):
        if len(selected.indexes()) <= 1:
            return
        found = False
        index = selected.indexes()[0]
        while index.parent().isValid():
            selection = index.model().itemFromIndex(index)
            if hash(str(selection)) in self.listed_data_entries:
                found = True
                break
            child = index
            index = index.parent()
        if not found:
            return
        self.load_data_entry(self.listed_data_entries[hash(str(selection))])

    # def load_data_entry(self, data_entry):
    #     task = self.tasks[selection.row() + child.row()]
    #     if task.station in self.sites.keys():
    #         site = self.sites[task.station]
    #     else:
    #         site = excelreader.undefinedSite()
    #     self.selectedTask = task
    #     self.stationLineEdit.setText(task.station)
    #     self.addressLineEdit.setPlainText(task.facility+"\n"+task.address+"\n"+task.city+", "
    #                                       +task.state+" "+task.zip)
    #     self.facilityNameLineEdit.setText(site["facility_name"])
    #     self.oitTextEdit.setPlainText(site["OIT_emails"])
    #     self.emailTextEdit.setPlainText(site["logistics_emails"])
    #     self.procurementLineEdit.setText(task.purchaseOrder)
    #     self.purchaseOrderLineEdit.setText(task.purchaseOrderFile)
    #     self.shipmentLineEdit.setText(task.shipmentFile)

    def validate_files(self):
        self.shipment_save.setEnabled(
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
