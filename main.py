import datetime
import sys
import os
import shutil
import re
import shelve
import typing

from datetime import date
from dataclasses import dataclass

import fedex_api
from PyQt5 import QtWidgets, uic
from PyQt5.QtGui import QStandardItem, QStandardItemModel
from PyQt5.QtWidgets import QFileDialog, QMessageBox
from PyQt5.QtCore import QFileInfo

import win32com.client

import excelreader
import helper_functions
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
        self.state = None

        self.fedex = fedex_api.FedexAPI(helper_functions.API_KEY, helper_functions.SECRET_KEY)

        # Load available IFCAP POs
        self.po_projects = []
        self.signals = {}
        for i in range(len(PROJECT_NAMES)):
            self.po_projects.append([])
        os.makedirs(IFCAP_PO_FOLDER, exist_ok=True)
        for file in os.listdir(IFCAP_PO_FOLDER):
            fileinfo = QFileInfo(file)
            info = fileinfo.fileName().split(";")
            if fileinfo.isDir() or len(info) < 4:  # File is directory or improperly formatted name
                continue
            try:
                self.po_projects[int(info[0])].append((info[1], info[2], f"{IFCAP_PO_FOLDER}{fileinfo.fileName()}"))
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
        self.po_combobox.activated.connect(self.validate_files)
        self.shipment_project.activated.connect(self.project_selected)
        self.save_cancel_button.pressed.connect(self.save_cancel_pressed)

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
        try:
            for purchase_order in self.po_projects[self.shipment_project.currentIndex()]:
                self.po_combobox.addItem(f"{purchase_order[0]}: {purchase_order[1]}")
        except IndexError:
            pass

    def save_cancel_pressed(self):
        if self.state == STATE_NEW_TASK:
            self.save_notification()
        if self.state == STATE_SHIPPING_NOTIFICATION_LOADED or self.state == STATE_DN_REQUEST_LOADED:
            self.set_application_state(STATE_NEW_TASK)

    def save_notification(self):
        excel_filename = self.shipmentLineEdit.text()
        pdf_filename = self.po_projects[self.shipment_project.currentIndex()][self.po_combobox.currentIndex()][2]

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

        with shelve.open(SHELVE_FILENAME, writeback=True) as db:
            db["data_entries"].append(entry)
        self.reload_items()

        message = QMessageBox()
        message.setIcon(QMessageBox.Information)
        message.setWindowTitle("Item Created")
        message.setText(f'New item "{entry.data.order_number}" created in {PROJECT_NAMES[entry.project]}.')
        message.exec()

        self.set_application_state(STATE_NEW_TASK)
        print("Saved?")

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
            self.shipment_address_text_edit.setPlainText(site["Area"] + "\n"
                                                         + site["Shipping Address"] + "\n"
                                                         + f"{site['Shipping City']}, {site['Shipping State']} {site['Shipping Zip Code']}")
            self.facilityNameLineEdit.setText(site["Area"])
            self.oitTextEdit.setPlainText(site["E-mail Distribution List for OIT"])
            self.emailTextEdit.setPlainText(site["E-mail Distribution List for Logistics"])

            file_folder = f"./files/{PROJECT_NAMES[data_entry.project]}/{data_entry.data.order_number}/"
            excel_file = QFileInfo(file_folder + data_entry.excel_file).absoluteFilePath()
            pdf_file = QFileInfo(file_folder + data_entry.pdf_file).absoluteFilePath()

            self.shipmentLineEdit.setText(excel_file)
            self.shipment_project.setCurrentIndex(data_entry.project)
            self.po_combobox.clear()
            self.po_combobox.addItem(QFileInfo(pdf_file).fileName())

            self.set_application_state(STATE_SHIPPING_NOTIFICATION_LOADED)

        else:
            pass

    def set_application_state(self, state):
        edit_mode = False
        self.state = state

        if state == STATE_NEW_TASK:
            self.save_cancel_button.setText("Save As New Entry")
            self.tab_view.setTabEnabled(0, True)
            self.tab_view.setTabEnabled(1, True)
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

        if not edit_mode:
            self.po_combobox.clear()
            self.shipment_project.setCurrentIndex(0)
            self.shipmentLineEdit.setText("")
            self.facilityNameLineEdit.setText("")
            self.shipment_address_text_edit.setPlainText("")
            self.oitTextEdit.setPlainText("")
            self.emailTextEdit.setPlainText("")
            self.procurementLineEdit.setText("")
            self.manufacturerLineEdit.setText("HP")

        self.shipment_project.setEnabled(not edit_mode)
        self.shipmentBrowseButton.setEnabled(not edit_mode)
        self.shipmentOpenButton.setEnabled(not edit_mode)
        self.shipmentLineEdit.setEnabled(not edit_mode)
        self.po_combobox.setEnabled(not edit_mode)
        self.save_cancel_button.setEnabled(edit_mode)
        self.shipment_address_text_edit.setEnabled(edit_mode)
        self.facilityNameLineEdit.setEnabled(edit_mode)
        self.oitTextEdit.setEnabled(edit_mode)
        self.emailTextEdit.setEnabled(edit_mode)
        self.shipment_generate_email.setEnabled(edit_mode)
        self.manufacturerLineEdit.setEnabled(edit_mode)
        self.procurementLineEdit.setEnabled(edit_mode)

    def validate_files(self):
        print("AAAAAAAAA")
        try:
            print(self.po_projects[self.shipment_project.currentIndex()][self.po_combobox.currentIndex()][2])
            print(os.path.exists(self.shipmentLineEdit.text()), os.path.exists(self.po_projects[self.shipment_project.currentIndex()][self.po_combobox.currentIndex()][2]))
            self.save_cancel_button.setEnabled(
                os.path.exists(self.shipmentLineEdit.text()) and
                os.path.exists(self.po_projects[self.shipment_project.currentIndex()][self.po_combobox.currentIndex()][2]))
        except IndexError:
            pass

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

    def generateEmail(self, task):
        itemQuantities = {}
        for shipment in self.selectedTask.items:
            for item in self.selectedTask.items[shipment]:
                includedItems = self.getItem(item["clin"])["included"];
                if includedItems != None:
                    for subitem in self.getItem(item["clin"])["included"].split(";"):
                        if self.getItem(subitem) != excelreader.undefinedItem():
                            if subitem not in itemQuantities.keys():
                                itemQuantities[subitem] = 0
                            itemQuantities[subitem] += 1
                if item["clin"] not in itemQuantities.keys():
                    itemQuantities[item["clin"] ] = 0
                itemQuantities[item["clin"] ] += 1
        with open("config/emailTemplate.htm") as fin:
            body = fin.read()
        body = body.replace("%FACILITY%", self.facilityNameLineEdit.text())
        body = body.replace("%SUBJECT%", self.procurementLineEdit.text())
        body = body.replace("%PO%", QFileInfo(self.selectedTask.purchaseOrderFile).fileName())
        body = body.replace("%OIT%", self.oitTextEdit.toPlainText())
        body = body.replace("%ADDRESS%", self.addressLineEdit.toPlainText().replace("\n", "<br/>"))
        body = body.replace("%TABLECAPTION%", "IFCAP " + self.selectedTask.purchaseOrder
                            + " - Manufacturer: " + self.manufacturerLineEdit.text() + ", "
                            + "Vendor: Colossal")
        itemTableIndex = body.find("<!-- RECORD ITEMS -->")
        for key in itemQuantities.keys():
            item = self.getItem(key)
            if item["model"] == "Undefined Item" or not item["record"]:
                continue
            rowString =f"""
            <tr>
            <td>{item["model"]}</td>
            <td>{item["csn"]}</td>
            <td>{item["manufacturer_name"]}</td>
            <td>{item["equipment_category"]}</td>
            <td>{item["cost"]}</td>
            <td>{item["warranty"]}</td>
            <td>{str(itemQuantities[key])}</td>
            </tr>"""
            body = body[:itemTableIndex]+rowString+body[itemTableIndex:]
            itemTableIndex+=len(rowString)
        itemTableIndex = body.find("<!-- NON RECORD ITEMS -->")
        foundNonRecordItems = False
        for key in itemQuantities.keys():
            item = self.getItem(key)
            if item["clin"] == "Undefined Item" or item["record"]:
                continue
            foundNonRecordItems = True
            rowString ="<h3>[%s]   -   %s</h3>"%(str(itemQuantities[key]), item["description"])
            body = body[:itemTableIndex]+rowString+body[itemTableIndex:]
            itemTableIndex+=len(rowString)
        outlook = win32com.client.Dispatch("outlook.application")
        itemTableIndex = body.find("<!-- TRACKING NUMBERS -->")
        trackDate = date(1970, 1, 1)
        invalidTracking = False
        packages = {}
        for key in self.selectedTask.items.keys():
            body = body[:itemTableIndex]+str(key)+"<br/>"+body[itemTableIndex:]
            trackingResult = self.fedexApi.trackbynumber(key)
            if not trackingResult.isValid:
                invalidTracking = True
                continue
            if trackingResult.deliveryDate > trackDate:
                trackDate = trackingResult.deliveryDate
            if trackingResult.packageType not in packages.keys():
                packages[trackingResult.packageType] = 0
            packages[trackingResult.packageType] += trackingResult.quantity
            itemTableIndex+=len(key)
        itemTableIndex = body.find("<!-- PACKAGES -->")
        for key in packages.keys():
            rowString ="""
            <tr>
                <td style="background-color:#d9e1f2;"><b><u>%s</u><b></td>
                <td>%s</td>
            </tr>
            """%(key, str(packages[key]))
            body = body[:itemTableIndex]+rowString+body[itemTableIndex:]
            itemTableIndex+=len(rowString)
        if invalidTracking:
            body = '<h3 style="color:red; background-color:yellow;">Some tracking information could not be automatically obtained - please manually enter/verify.</h3><br/>\n'+body
        body = body.replace("%DELIVERDATE%", trackDate.strftime("%A, %B %d, %Y"))
        if not foundNonRecordItems:
            body = body[:body.find("<!-- NON RECORD TEXT -->")]
        outlook = win32com.client.Dispatch("outlook.application")
        mail = outlook.CreateItem(0)
        mail.To = self.emailTextEdit.toPlainText()
        mail.Subject = "Shipment Notification: " + self.selectedTask.purchaseOrder + " ["+self.selectedTask.shipments["name"]+ "] to " + self.facilityNameLineEdit.text()
        mail.HtmlBody = body
        mail.Attachments.Add(QFileInfo(self.selectedTask.purchaseOrderFile).absoluteFilePath())
        mail.Attachments.Add(QFileInfo(self.selectedTask.folderName).absoluteFilePath()+"/SERIAL/PVaaS_SN_"+task.shipments["name"]+".xlsx")
        mail.Save()
        mail.Display(False)


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec())
