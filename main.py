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
from helper_functions import code_to_site

# region Constants
SHELVE_FILENAME = "./settings/registry"
IFCAP_PO_FOLDER = "./settings/IFCAP PO/"
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
STATE_NEW_TASK = 0
STATE_NEW_DN_REQUEST = 1
STATE_SHIPPING_NOTIFICATION_LOADED = 2
STATE_DN_REQUEST_LOADED = 3
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

        # Load available IFCAP POs
        self.po_projects = []
        for i in range(len(PROJECT_NAMES)):
            self.po_projects.append([])
        os.makedirs(IFCAP_PO_FOLDER, exist_ok=True)
        for file in os.listdir(IFCAP_PO_FOLDER):
            fileinfo = QFileInfo(file)
            info = fileinfo.fileName().split(";")
            if fileinfo.isDir() or len(info) < 4:  # File is directory or improperly formatted name
                continue
            try:
                self.po_projects[int(info[0])].append((info[1], info[2], fileinfo.absoluteFilePath()))
            except (ValueError, KeyError):
                continue

        self.selected_entry = None
        with shelve.open(SHELVE_FILENAME) as db:
            if "data_entries" not in db:
                db["data_entries"] = []

        self.snowModel = QStandardItemModel()
        self.roots = []
        self.listed_data_entries = {}
        self.snowTreeView.setModel(self.snowModel)
        self.snowTreeView.header().setDefaultSectionSize(120)
        self.reload_items()

        self.snowTreeView.selectionModel().selectionChanged.connect(self.tree_view_selection_changed)
        self.shipmentBrowseButton.pressed.connect(self.browse_xlsx)
        self.shipmentOpenButton.pressed.connect(self.open_xlsx)
        self.shipmentLineEdit.textChanged.connect(self.validate_files)
        self.save_cancel_button.pressed.connect(self.save_shipment)
        self.shipment_project.activated.connect(self.project_selected)

        self.set_application_state(STATE_NEW_TASK)

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

    def project_selected(self):
        self.po_combobox.clear()
        for purchase_order in self.po_projects[self.shipment_project.currentIndex()]:
            self.po_combobox.addItem(f"{purchase_order[0]}: {purchase_order[1]}")

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

    def tree_view_selection_changed(self, selected):
        if len(selected.indexes()) <= 1:
            return
        found = False
        index = selected.indexes()[0]
        selection = index.model().itemFromIndex(index)
        while index.parent().isValid():
            if hash(str(selection)) in self.listed_data_entries:
                found = True
                break
            child = index
            index = index.parent()
            selection = index.model().itemFromIndex(index)
        if not found:
            return
        self.load_data_entry(self.listed_data_entries[hash(str(selection))])

    def load_data_entry(self, data_entry):
        self.selected_entry = data_entry
        self.shipment_order_number.setText(data_entry.data.order_number)
        if data_entry.type == ENTRY_TYPE_SHIPMENT:
            try:
                site = code_to_site(data_entry.data.station_number, data_entry.data.va_facility)
            except KeyError:
                pass  # Implement proper handling here
            self.stationLineEdit.setText(data_entry.data.station_number)
            self.addressLineEdit.setPlainText(site["Area"] + "\n"
                                              + site["Shipping Address"] + "\n"
                                              + f"{site['Shipping City']}, {site['Shipping State']} {site['Shipping Zip Code']}")
            self.facilityNameLineEdit.setText(site["Area"])
            self.oitTextEdit.setPlainText(site["E-mail Distribution List for OIT"])
            self.emailTextEdit.setPlainText(site["E-mail Distribution List for Logistics"])

            file_folder = f"./files/{PROJECT_NAMES[data_entry.project]}/{data_entry.data.order_number}/"
            excel_file = QFileInfo(file_folder+data_entry.excel_file).absoluteFilePath()
            pdf_file = QFileInfo(file_folder+data_entry.pdf_file).absoluteFilePath()

            self.shipmentLineEdit.setText(excel_file)
            self.shipment_project.setCurrentIndex(data_entry.project)
            self.po_projects.clear()
            self.po_combobox.addItem(QFileInfo(pdf_file).fileName())

            self.set_application_state(STATE_SHIPPING_NOTIFICATION_LOADED)

        else:
            pass

    def set_application_state(self, state):
        edit_mode = False
        if state == STATE_NEW_TASK:
            self.save_cancel_button.setText("Save As New Entry")
        elif state == STATE_SHIPPING_NOTIFICATION_LOADED:
            self.save_cancel_button.setText("Close Shipping Notification")
            self.tab_view.setCurrentWidget(self.shipment_tab)
            self.tab_view.setTabEnabled(1, False)
            edit_mode = True
        elif state == STATE_DN_REQUEST_LOADED:
            self.save_cancel_button.setText("Close Shipping Notification Request")
            self.tab_view.setCurrentWidget(self.shipment_tab)
            self.tab_view.setTabEnabled(1, False)
            edit_mode = True
        self.shipment_project.setEnabled(not edit_mode)
        self.shipmentBrowseButton.setEnabled(not edit_mode)
        self.shipmentOpenButton.setEnabled(not edit_mode)
        self.shipmentLineEdit.setEnabled(not edit_mode)
        self.po_combobox.setEnabled(not edit_mode)
        self.save_cancel_button.setEnabled(edit_mode)
        self.addressLineEdit.setEnabled(edit_mode)
        self.facilityNameLineEdit.setEnabled(edit_mode)
        self.oitTextEdit.setEnabled(edit_mode)
        self.emailTextEdit.setEnabled(edit_mode)
        self.shipment_generate_email.setEnabled(edit_mode)
        self.manufacturerLineEdit.setEnabled(edit_mode)
        self.procurementLineEdit.setEnabled(edit_mode)

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
        self.save_cancel_button.setEnabled(
            os.path.exists(self.shipmentLineEdit.text()) and
            os.path.exists(self.po_projects[self.shipment_project.currentIndex()][self.po_combobox.currentIndex()][3]))

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
