# -*- coding: utf-8 -*-
"""
Created on Fri May 20 11:26:53 2022

@author: Adrian
"""

import os
import re
from datetime import datetime

import extract_msg

EMAILS_FOLDER = "D:/Work/Programming/NVOICE - v2/emails/"


def extractTrackingNumbers(file):
    output = []
    msg = extract_msg.openMsg(file)
    trackingNumbers = re.finditer(r"\b\d{12}|\bKOL-NT\d{2}-\d{4}|\bIADD\d{6}", msg.body)
    for match in trackingNumbers:
        if match.group() not in output:
            output.append(str(match.group()))
    return output


def findFileByNumber(trackingNumber):
    files = os.listdir(EMAILS_FOLDER)
    for file in files:
        if "Supporting Technologies Contract" in file:
            msg = extract_msg.openMsg(EMAILS_FOLDER + file)
            if trackingNumber in msg.body.replace(" ", ""):
                return file
    return "404"


def findFilesByNumbers(trackingNumbers):
    files = []
    files = os.listdir(EMAILS_FOLDER)
    notifications = []
    for k, v in enumerate(trackingNumbers):
        trackingNumbers[k] = v.replace(" ", "")
    for file in files:
        msg = extract_msg.openMsg(EMAILS_FOLDER + file)
        body = msg.body.replace(" ", "")
        cleanup = []
        for k, v in enumerate(trackingNumbers):
            if re.search(f"(?<=:){v}", body) is not None:
                notifications.append(
                    (v, datetime.strptime(msg.date, "%a, %d %b %Y %H:%M:%S %z").strftime("%m/%d/%Y"), file))
                cleanup.append(k)
        for index in cleanup:
            trackingNumbers.pop(index)
        if len(trackingNumbers) == 0:
            break
    for number in trackingNumbers:
        notifications.append((number, "NOT FOUND", "NOT FOUND"))
    return notifications


def parseEmails(emailList=None):
    if emailList == None:
        emailList = {}
    files = os.listdir(EMAILS_FOLDER)
    i = 0
    for file in files:
        if file not in list(emailList.keys()):
            emailList[file] = extractTrackingNumbers(EMAILS_FOLDER + file)
            i += 1
    print(f"Added {i} new emails to index file")
    return emailList


def findEmail(trackingNumber, emailList):
    for key in emailList.keys():
        if str(trackingNumber) in emailList[key]:
            return key
    return None
