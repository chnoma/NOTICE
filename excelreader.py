# -*- coding: utf-8 -*-
"""
Created on Sun Feb 13 21:28:21 2022

@author: Adrian
"""

import shutil
import os
import re
from dataclasses import dataclass

import openpyxl
import pandas as pd
from PyQt5.QtCore import QFileInfo

from helper_functions import jam, jam_int


@dataclass
class ShipmentLine:
    district: str
    d_t: str
    location_code: str
    station_number: str
    shipping_address: str
    city: str
    state: str
    va_facility: str
    zip_code: str
    tracking_number: str
    sku: str
    description: str
    clin: str
    qty: int
    service_tag: str
    purchase_order: str
    order_number: str


@dataclass
class ShipmentNotification:
    order_number: str
    shipments: list


def parse_shipment_notification(file_name):
    """
    Parses the given shipment notification file
    Returns a ShipmentNotification instance containing
        all necessary information about the shipment notification
    """
    df = pd.read_excel(file_name)

    order_number = ""
    match = re.search(r"SCTASK(\d+)", df["Unnamed: 0"][0])
    if match is not None:
        order_number = match.group(0)

    alt_order_number = ""
    shipments = []
    for i in range(5, len(df["Unnamed: 0"])):
        new_order_number = jam(df["Unnamed: 16"][i])
        if alt_order_number == "" and new_order_number != "":  # if no SCTASK is found, use the first non-blank order #
            alt_order_number = new_order_number
        new_shipment = ShipmentLine(
            district=jam(df["Unnamed: 0"][i]),
            d_t=jam(df["Unnamed: 1"][i]),
            location_code=jam(df["Unnamed: 2"][i]),
            station_number=jam(df["Unnamed: 3"][i]),
            shipping_address=jam(df["Unnamed: 4"][i]),
            city=jam(df["Unnamed: 5"][i]),
            state=jam(df["Unnamed: 6"][i]),
            va_facility=jam(df["Unnamed: 7"][i]),
            zip_code=jam(df["Unnamed: 8"][i]),
            tracking_number=jam(df["Unnamed: 9"][i]),
            sku=jam(df["Unnamed: 10"][i]),
            description=jam(df["Unnamed: 11"][i]),
            clin=jam(df["Unnamed: 12"][i]),
            qty=jam_int(df["Unnamed: 13"][i], 1),
            service_tag=jam(df["Unnamed: 14"][i]),
            purchase_order=jam(df["Unnamed: 15"][i]),
            order_number=new_order_number)
        shipments.append(new_shipment)

    if order_number == "":
        order_number = alt_order_number

    return ShipmentNotification(order_number, shipments)


# def readShipmentFile(file):
#     shipmentInfo = {}
#     workbook = openpyxl.load_workbook(file)
#     shipments = []
#     sheet_obj = workbook["Shipment"]
#     offset = 6
#     entryCount = 0
#     while (True):
#         if sheet_obj.cell(entryCount + offset, 4).value == None:
#             break
#         entryCount += 1
#     for i in range(0, entryCount):
#         shipment = {}
#         shipment["district"] = sheet_obj.cell(i + offset, 1).value
#         shipment["d-t"] = sheet_obj.cell(i + offset, 2).value
#         shipment["location"] = sheet_obj.cell(i + offset, 3).value
#         shipment["station"] = sheet_obj.cell(i + offset, 4).value
#         shipment["address"] = sheet_obj.cell(i + offset, 5).value
#         shipment["city"] = sheet_obj.cell(i + offset, 6).value
#         shipment["state"] = sheet_obj.cell(i + offset, 7).value
#         shipment["va_facility"] = sheet_obj.cell(i + offset, 8).value
#         shipment["zip"] = sheet_obj.cell(i + offset, 9).value
#         shipment["trackingNumber"] = sheet_obj.cell(i + offset, 10).value
#         shipment["sku"] = sheet_obj.cell(i + offset, 11).value
#         shipment["description"] = sheet_obj.cell(i + offset, 12).value
#         shipment["clin"] = sheet_obj.cell(i + offset, 13).value
#         shipment["qty"] = sheet_obj.cell(i + offset, 14).value
#         shipment["service_tag"] = sheet_obj.cell(i + offset, 15).value
#         shipment["po"] = sheet_obj.cell(i + offset, 16).value
#         shipment["order"] = sheet_obj.cell(i + offset, 17).value
#         shipment["ship_date"] = sheet_obj.cell(i + offset, 18).value
#         shipments.append(shipment)
#     shipmentInfo["shipments"] = shipments
#     shipmentInfo["name"] = sheet_obj.cell(2, 1).value[7:]
#     shipmentInfo["date"] = sheet_obj.cell(2, 1).value[-10:]
#     print(shipmentInfo["date"])
#     shipmentInfo["name"] = shipmentInfo["name"][:shipmentInfo["name"].find(" -")]
#     return shipmentInfo

#
# def readItemList(file):
#     workbook = openpyxl.load_workbook(file)
#     items = []
#     sheet_obj = workbook.active
#     offset = 2
#     entryCount = 0
#     while (True):
#         if sheet_obj.cell(entryCount + offset, 1).value == None:
#             break
#         entryCount += 1
#     for i in range(0, entryCount):
#         item = {}
#         item["contract"] = sheet_obj.cell(i + offset, 1).value
#         item["clin"] = sheet_obj.cell(i + offset, 2).value
#         item["description"] = sheet_obj.cell(i + offset, 3).value
#         item["model"] = sheet_obj.cell(i + offset, 4).value
#         item["csn"] = sheet_obj.cell(i + offset, 5).value
#         item["manufacturer_name"] = sheet_obj.cell(i + offset, 6).value
#         item["equipment_category"] = sheet_obj.cell(i + offset, 7).value
#         item["cost"] = sheet_obj.cell(i + offset, 8).value
#         item["warranty"] = sheet_obj.cell(i + offset, 9).value
#         item["included"] = sheet_obj.cell(i + offset, 10).value
#         item["record"] = sheet_obj.cell(i + offset, 11).value == "Yes"
#         items.append(item)
#     return items
#
#
# def readSiteList(file):
#     workbook = openpyxl.load_workbook(file)
#     sheet_obj = workbook.active
#     sites = {}
#     offset = 1
#     entryCount = 0
#     while (True):
#         if sheet_obj.cell(entryCount + offset, 1).value == None:
#             break
#         entryCount += 1
#     for i in range(0, entryCount):
#         site = {}
#         site["station"] = sheet_obj.cell(i + offset, 1).value
#         site["facility_name"] = sheet_obj.cell(i + offset, 2).value
#         site["logistics_emails"] = sheet_obj.cell(i + offset, 3).value
#         site["OIT_emails"] = sheet_obj.cell(i + offset, 4).value
#         sites[site["station"]] = site
#     return sites
#
#
def generateSerialList(file):
    qfile = QFileInfo(file)
    workbook = openpyxl.load_workbook(file)
    del workbook["Shipment"]
    return workbook
#
#
# def undefinedSite():
#     undefText = "Undefined - please add this site to site_list.xlsx in the config folder"
#     site = {}
#     site["station"] = undefText
#     site["facility_name"] = undefText
#     site["logistics_emails"] = undefText
#     site["OIT_emails"] = undefText
#     return site
#
#
# def undefinedItem():
#     undefText = "Undefined Item"
#     item = {}
#     item["contract"] = undefText
#     item["clin"] = undefText
#     item["description"] = undefText
#     item["model"] = undefText
#     item["csn"] = undefText
#     item["manufacturer_name"] = undefText
#     item["equipment_category"] = undefText
#     item["record"] = False
#     item["cost"] = undefText
#     item["warranty"] = undefText
#     item["included"] = ""
#     return item
