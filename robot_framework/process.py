"""This module contains the main process of the robot."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
import random 
import string
import subprocess
import pyautogui
import win32com.client
#import new_password
from openpyxl import load_workbook, workbook
import re
import json
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto import mouse
import win32clipboard

import win32gui
import win32con
import win32api

from collections import Counter

n = 0

# pylint: disable-next=unused-argument
def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:
    """Do the primary process of the robot."""
    orchestrator_connection.log_trace("Running process...")
    
    def get_client():
        #orchestrator_connection.log_trace("get_client started...")
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        if not type(sap_gui_auto) == win32com.client.CDispatch:
            return

        application = sap_gui_auto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            sap_gui_auto = None
            return

        for conn in range(application.Children.Count):
            # Loop through the application and get the connection
            connection = application.Children(conn)

            for sess in range(connection.Children.Count):
                # Loop through each connection and return sessions that are on the main screen 'SESSION_MANAGER'
                session = connection.Children(sess)
                #print(session.Info.Transaction)
                if session.Info.Transaction == 'SESSION_MANAGER':
                    return session
                else:
                    if session.Info.Transaction == 'SBWP':
                        return session
                    else:
                        # Return None and break
                        return

    specific_content = json.loads(queue_element.data)
    # Assign variables from SpecificContent
    invoiceNo = specific_content.get("invoiceNo", None)
    title = specific_content.get("title", None)
    eanNr = specific_content.get("eanNr", None)
    fakturabeløb = specific_content.get("fakturabeløb", None)
    leverandør = specific_content.get("leverandør", None)
    
    orchestrator_connection.log_trace("New: "+title)
    print("New: "+title)
    
    obj_sess = get_client()
    time.sleep(5)
    grid = obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell")
    nr = 0
    nr2 = -1
    data = []
    for r in range(grid.RowCount):
        a = grid.GetCellValue(r,"WI_TEXT")
        data.append({
            "no": nr,
            "title": a
        })
        nr += 1
    id = next((p for p in data if p["title"].lower() == title.lower()), None)
    nr2 = id["no"]
    print("Nr i liste = "+str(nr2))
    orchestrator_connection.log_trace("Nr i liste = "+str(nr2))
    time.sleep(2)
    obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").currentCellColumn = "WI_TEXT"
    obj_sess.findById("wnd[0]/mbar/menu[3]/menu[6]").select() #Opdater siden...
    obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").selectedRows = nr2
    obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").selectionChanged
    obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").pressToolbarButton("APRO") #for 'Haandter afvist' åbnes WebViev
    
    #tree = obj_sess.findById("wnd[0]/usr/cntlSWU20300CONTAINER/shellcont/shell")
    #print("Type:", tree.Type)
    #print("Id:", tree.Id)
    #print("IdSub:", tree.SubType)
    #html = tree.GetSource()
    #print(html)
    #html = tree.GetHtmlSource()
    #print(html)
    #print(dir(tree))

    if queue_element.queue_name=="Bogholderbakke_NulBeløb":
        obj_sess.findById("wnd[0]/usr/cntlSWU20300CONTAINER/shellcont/shell").sapEvent("", "", "SAPEVENT:DECI:0001")
        obj_sess.findById("wnd[0]/usr/txtRBKPV-BELNR").setFocus
        obj_sess.findById("wnd[0]/usr/txtRBKPV-BELNR").caretPosition = 10
        obj_sess.findById("wnd[0]/mbar/menu[0]/menu[6]").select() #Her slettes bilaget
        sbar = obj_sess.findById("wnd[0]/sbar")
        print("invoiceNo: "+invoiceNo+" - Type: "+sbar.MessageType+" - "+sbar.Text)
        orchestrator_connection.log_trace("invoiceNo: "+invoiceNo+" - Type: "+sbar.MessageType+" - "+sbar.Text)
        time.sleep(2)
        obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").pressToolbarButton("EREF")
        #time.sleep(2)
    
    if queue_element.queue_name=="Bogholderbakke_XML":
        obj_sess.findById("wnd[0]/usr/cntlSWU20300CONTAINER/shellcont/shell").sapEvent("","","SAPEVENT:DECI:0002")
        obj_sess.findById("wnd[0]/mbar/menu[0]/menu[3]").select()
        i = 1
        while i < 9:
            sbar = obj_sess.findById("wnd[0]/sbar")
            print("Type: "+sbar.MessageType+" - Text: "+sbar.Text)
            pyautogui.press('enter')
            time.sleep(2)
            if i == 9:
                break
            i += 1
        
        #tree = obj_sess.findById("wnd[0]/usr/txtRBKPV-BELNR")
        #print("Type:", tree.Type) #Type: GuiTextField
        #print("Id:", tree.Id)
    
    if queue_element.queue_name=="Bogholderbakke_DobbeltFaktura":
        obj_sess = get_client()
        obj_sess.findById("wnd[0]/usr/cntlSWU20300CONTAINER/shellcont/shell").sapEvent("","","SAPEVENT:DECI:0002")
        
        #Træk direkte fra siden
        grid = obj_sess.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell")
        """
        cols = grid.ColumnOrder  # GuiCollection af kolonnenavne (tekniske)
        for i in range(cols.Length):
            col = cols.ElementAt(i)
            print(i, col)
        """
        
        tmp = []
        for r in range(grid.RowCount):
            FakNo = grid.GetCellValue(r,"BELNR")
            Reference = grid.GetCellValue(r,"XBLNR")
            FakturaBruttoBelob = grid.GetCellValue(r,"RMWWR")
            Fakturaudsteder = grid.GetCellValue(r,"LIFNR")
            Regnskabsaar = grid.GetCellValue(r,"GJAHR")
            Bilagsdato = grid.GetCellValue(r,"BLDAT")
            EAN = grid.GetCellValue(r,"BKTXT")
            tmp.append({
                'FakNo':FakNo,
                'Reference':Reference, 
                'FakturaBruttoBelob':FakturaBruttoBelob, 
                'Fakturaudsteder':Fakturaudsteder, 
                'Regnskabsaar':Regnskabsaar, 
                'Bilagsdato':Bilagsdato, 
                'EAN':EAN})
          
        resultat = Counter(d["FakNo"] for d in tmp)
        #print(resultat)
        noOfRowsFakturaNr = resultat.most_common(1)[0][1]
        #print(noOfRowsFakturaNr)
        
        noOfRowsTotal = len(tmp)
        noOfRowsFakturaNr = (Counter(d["FakNo"] for d in tmp)).most_common(1)[0][1]
        noOfRowsReference = (Counter(d["Reference"] for d in tmp)).most_common(1)[0][1]
        noOfRowsFakturabeloeb = (Counter(d["FakturaBruttoBelob"] for d in tmp)).most_common(1)[0][1]
        noOfRowsFakturaudsteder = (Counter(d["Fakturaudsteder"] for d in tmp)).most_common(1)[0][1]
        noOfRowsAar = (Counter(d["Regnskabsaar"] for d in tmp)).most_common(1)[0][1]
        noOfRowsBilagsdato = (Counter(d["Bilagsdato"] for d in tmp)).most_common(1)[0][1]
        noOfRowsEAN = (Counter(d["EAN"] for d in tmp)).most_common(1)[0][1]
        
        if noOfRowsFakturaNr == 1: noOfRowsFakturaNr = noOfRowsTotal
        if noOfRowsTotal-noOfRowsReference == 0: noOfRowsReference = 1
        if noOfRowsTotal-noOfRowsFakturabeloeb == 0: noOfRowsFakturabeloeb = 1
        if noOfRowsTotal-noOfRowsFakturaudsteder == 0: noOfRowsFakturaudsteder = 1
        if noOfRowsTotal-noOfRowsAar == 0: noOfRowsAar = 1
        if noOfRowsTotal-noOfRowsBilagsdato == 0: noOfRowsBilagsdato = 1
        if noOfRowsTotal-noOfRowsEAN == 0: noOfRowsEAN = 1
        
        rule = 0
        if (noOfRowsTotal == noOfRowsFakturaNr and noOfRowsReference == 1 and noOfRowsFakturabeloeb == 1 and noOfRowsFakturaudsteder == 1 and noOfRowsAar == 1 and noOfRowsTotal > 1):
            #print("Kontrol af faktura - rule 1")
            rule = 1
        if noOfRowsTotal == 1:
            #print("Kun 1 faktura - rule 2")
            rule = 2
        if (noOfRowsTotal == noOfRowsFakturaNr and noOfRowsReference == 1 and noOfRowsFakturabeloeb == 1 and noOfRowsFakturaudsteder == 1 and noOfRowsAar > 1 and noOfRowsTotal > 1):    
            #print("Aarstal ikke ens - rule 3")
            rule = 3
        if (rule == 0):
            #print("Ingen rule valgt endnu...")
            rule = 4
        
                
        time.sleep(3)
        orchestrator_connection.log_trace("Rule: "+str(rule))
        #print("stop her")
        obj_sess.findById("wnd[0]/tbar[0]/btn[3]").press()
        obj_sess.findById("wnd[0]/tbar[0]/btn[12]").press()
        
    if queue_element.queue_name=="Bogholderbakke_HåndterAfvist":
        print(queue_element.queue_name) 
        time.sleep(20)
        #app = Application(backend="uia").connect(title_re="Bilagsvisning")
        #window = app.window(title_re="Bilagsvisning")
        #window.set_focus()
        
        TITLE_RE = "Bilagsvisning"

        app = Application(backend="uia").connect(title_re=TITLE_RE)
        win = app.window(title_re=TITLE_RE)
        win.set_focus()
        time.sleep(0.3)

        # Find alle børn og print dem (så du kan spotte WebView2-host)
        children = win.descendants()
        for c in children:
            try:
                name = c.window_text()
            except:
                name = ""
            print(c.friendly_class_name(), c.element_info.control_type, name)

        win32clipboard.OpenClipboard()
        try:
            data_WebView = win32clipboard.GetClipboardData()
        finally:
            win32clipboard.CloseClipboard()

        print("Hele tekstindholdet:\n", data_WebView)   
    
    orchestrator_connection.log_trace("Running process - end")