import datetime
import os
import re
import shelve
import shutil
import sys
from dataclasses import dataclass
from datetime import date
from pytz import UTC as utc

import fedex_api
import win32com.client
from PyQt5 import QtWidgets, uic
from PyQt5.QtCore import QFileInfo
from PyQt5.QtGui import QStandardItem, QStandardItemModel
from PyQt5.QtWidgets import QFileDialog, QMessageBox

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
    status = "Email not Generated"

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
        self.shipment_generate_email.pressed.connect(self.generate_email)

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
                                    QStandardItem(date.strftime(entry.date_added, "%m/%d/%Y")),
                                    QStandardItem(entry.status)])
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
            excel_file=f"{shipment.order_number}.xlsx",
            pdf_file=QFileInfo(pdf_filename).fileName(),
            title=shipment.order_number,
            type=ENTRY_TYPE_SHIPMENT,
            email_generated=False,
            date_added=datetime.datetime.now(),
            data=shipment
        )

        if PROJECT_NAMES[entry.project] == PROJECT_NAMES[PROJECT_PVAAS] \
                and re.search("SCTASK", entry.data.order_number) is None:
            cont = QMessageBox.question(self, 'No SCTask Found',
                                        'You have selected PVaaS, but no SCTask was located.\nWould you like to continue?',
                                        QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if cont == QMessageBox.No:
                return

        new_item_folder = f"./files/{PROJECT_NAMES[entry.project]}/{entry.data.order_number}/"
        try:
            os.makedirs(new_item_folder)
        except Exception as e:
            QMessageBox.warning(self, 'Exception', getattr(e, 'message', repr(e)))
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
            mod = self.selected_entry.pdf_file.split(';')[1]
            ifcap_po = self.selected_entry.pdf_file.split(';')[2]
            self.stationLineEdit.setText(data_entry.data.station_number)
            self.shipment_address_text_edit.setPlainText(site["Area"] + "\n"
                                                         + site["Shipping Address"] + "\n"
                                                         + f"{site['Shipping City']}, {site['Shipping State']} {site['Shipping Zip Code']}")
            self.procurementLineEdit.setText(
                f"{PROJECT_NAMES[self.selected_entry.project]} {mod} IFCAP PO# {ifcap_po} Shipment to {site['Area']}")
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
        try:
            self.save_cancel_button.setEnabled(
                os.path.exists(self.shipmentLineEdit.text()) and
                os.path.exists(
                    self.po_projects[self.shipment_project.currentIndex()][self.po_combobox.currentIndex()][2]))
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

    def generate_email(self):
        files_folder = f"./files/{PROJECT_NAMES[self.selected_entry.project]}/{self.selected_entry.data.order_number}/"
        pdf_filename = QFileInfo(self.selected_entry.pdf_file).fileName()
        oit_emails = self.oitTextEdit.toPlainText()
        logistics_emails = self.emailTextEdit.toPlainText()
        if re.search("included", oit_emails.lower()) is not None:
            oit_emails = logistics_emails
        if re.search("included", logistics_emails.lower()) is not None:
            logistics_emails = oit_emails
        mod = pdf_filename.split(';')[1]
        ifcap_po = pdf_filename.split(';')[2]
        attachment_name = f"{PROJECT_NAMES[self.selected_entry.project]} {mod} {ifcap_po}.pdf"

        item_quantities = {}
        for shipment in self.selected_entry.data.shipments:
            if shipment.description not in item_quantities.keys():
                item_quantities[shipment.description] = 0
            item_quantities[shipment.description] += shipment.qty

        with open(f"./settings/emailTemplate.htm") as fin:
            body = fin.read()
        body = body.replace("%FACILITY%", self.facilityNameLineEdit.text())
        body = body.replace("%SUBJECT%", self.procurementLineEdit.text())
        body = body.replace("%PO%", attachment_name)
        body = body.replace("%OIT%", oit_emails)
        body = body.replace("%ADDRESS%", self.shipment_address_text_edit.toPlainText().replace("\n", "<br/>"))
        body = body.replace("%TABLECAPTION%", "IFCAP " + ifcap_po
                            + " - Manufacturer: " + self.manufacturerLineEdit.text() + ", "
                            + "Vendor: Colossal")

        found_non_record_items = False
        found_record_items = False
        item_table_index = body.find("<!-- RECORD ITEMS -->") + len("<!-- RECORD ITEMS -->")
        include_all_unknown_items = False
        included_unknown_items = []
        for key in item_quantities.keys():
            print(key)
            ignore_key = False
            for desc in excelreader.IGNORE_LIST[
                "Description"]:  # if key in excelreader.IGNORE_LIST["Description"] was not working.
                if key == desc:  # I do not understand why.
                    ignore_key = True
                    break
            if ignore_key:
                continue
            try:
                item = excelreader.ITEMS[key]
            except KeyError:
                if not include_all_unknown_items:
                    cont = QMessageBox.question(self, 'Item not found.',
                                                f"Item '{key}' not found in the item or ignore spreadsheets."
                                                "\n\nYou may add this item to the item list or ignore list to correct this issue."
                                                "\nWould you like to include this as a non-record item?",
                                                QMessageBox.YesAll | QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel, QMessageBox.No)
                    if cont == QMessageBox.Yes:
                        pass
                    elif cont == QMessageBox.YesAll:
                        include_all_unknown_items = True
                    elif cont == QMessageBox.No:
                        continue
                    elif cont == QMessageBox.Cancel:
                        return
                included_unknown_items.append(key)
                item = excelreader.Item(key, "", "", "", "", "", False)

            if not item.record:
                found_non_record_items = True
                continue
            found_record_items = True
            row_string = f"""
            <tr>
            <td>{item.model}</td>
            <td>{item.csn}</td>
            <td>{item.manufacturer_name}</td>
            <td>{item.equipment_category}</td>
            <td>{item.cost}</td>
            <td>{item.warranty}</td>
            <td>{str(item_quantities[key])}</td>
            </tr>"""
            body = body[:item_table_index] + row_string + body[item_table_index:]
            item_table_index += len(row_string)

        item_table_index = body.find("<!-- NON RECORD ITEMS -->") + len("<!-- NON RECORD ITEMS -->")

        if found_non_record_items:
            for key in item_quantities.keys():
                try:
                    item = excelreader.ITEMS[key]
                    if item.record:
                        continue
                except KeyError:
                    if key not in included_unknown_items:
                        continue
                row_string = f'<h3 style="color:red; background-color:yellow;">[{str(item_quantities[key])}]   -   {key}</h3>'
                body = body[:item_table_index] + row_string + body[item_table_index:]
                item_table_index += len(row_string)

        item_table_index = body.find("<!-- TRACKING NUMBERS -->")
        track_date = utc.localize(datetime.datetime(1970, 1, 1))
        invalid_tracking = False
        packages = {}
        processed_tracking_numbers = []
        for shipment in self.selected_entry.data.shipments:
            ignore_key = False
            for desc in excelreader.IGNORE_LIST["Description"]:  # if key in excelreader.IGNORE_LIST["Description"] was not working.
                if shipment.description == desc:  # I do not understand why.
                    continue
            if shipment.description not in excelreader.ITEMS and shipment.description not in included_unknown_items:
                continue
            tracking_number = shipment.tracking_number
            if tracking_number in processed_tracking_numbers:
                continue
            body = body[:item_table_index] + str(tracking_number) + "<br/>" + body[item_table_index:]
            tracking_result = self.fedex.track_by_number(tracking_number)
            if not tracking_result.is_valid:
                invalid_tracking = True
                continue
            if tracking_result.date_delivery > track_date:
                track_date = tracking_result.date_delivery
            if tracking_result.package is not None:
                if tracking_result.package.type not in packages.keys():
                    packages[tracking_result.package.type] = 0
                packages[tracking_result.package.type] += tracking_result.package.count
            processed_tracking_numbers.append(tracking_number)
            item_table_index += len(tracking_number + "<br/>")
        item_table_index = body.find("<!-- PACKAGES -->")
        for key in packages.keys():
            row_string = """
            <tr>
                <td style="background-color:#d9e1f2;"><b><u>%s</u><b></td>
                <td>%s</td>
            </tr>
            """ % (key, str(packages[key]))
            body = body[:item_table_index] + row_string + body[item_table_index:]
            item_table_index += len(row_string)
        if invalid_tracking:
            body = '<h3 style="color:red; background-color:yellow;">Some tracking information could not be automatically obtained - please manually enter/verify.</h3><br/>\n' + body
        body = body.replace("%DELIVERDATE%", track_date.strftime("%A, %B %d, %Y"))

        if not found_record_items:  # Remove record table if no record items
            body_start = body[:body.find("<!-- RECORD TABLE START -->")]
            body_end = body[body.find("<!-- NON RECORD TEXT -->"):]
            body = body_start + body_end

        if not found_non_record_items:  # Remove non-record section if no non-record items
            body = body[:body.find("<!-- NON RECORD TEXT -->")]

        outlook = win32com.client.Dispatch("outlook.application")
        mail = outlook.CreateItem(0)
        mail.To = logistics_emails
        mail.Subject = self.procurementLineEdit.text()
        mail.HtmlBody = body
        shutil.copy(files_folder + self.selected_entry.pdf_file,
                    files_folder + attachment_name)  # This seems a bit unnecessary, but this is the only way to rename an attachment
        mail.Attachments.Add(QFileInfo(files_folder + attachment_name).absoluteFilePath())
        os.remove(files_folder + attachment_name)  # See two lines above
        mail.Attachments.Add(
            QFileInfo(files_folder + self.selected_entry.excel_file[:-5] + "_SN.xlsx").absoluteFilePath())
        mail.Save()
        mail.Display(True)


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec())
