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

n = 0

# pylint: disable-next=unused-argument
def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:
    """Do the primary process of the robot."""
    orchestrator_connection.log_trace("Running process.")
    print(queue_element.queue_name)
    
        
    
    
    
    def get_client():
        orchestrator_connection.log_trace("get_client started...")
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
                print(session.Info.Transaction)
                if session.Info.Transaction == 'SESSION_MANAGER':
                    return session
                else:
                    if session.Info.Transaction == 'SBWP':
                        return session
                    else:
                        # Return None and break
                        return
   


    def data_in_workbook(bakken, aktueltype):
        """
        orchestrator_connection.log_trace("data_in_workbook started...")
        match bakken:
            case "Fakturabeslut.03: Kontroller dob fakt":
                eanNoKolonne = "K"
                titleKolonne = "B"
                invoiceNoKolonne = "B"
            case "Fakturabeslut.07: Inkonsistent XML":
                eanNoKolonne = "A"
                titleKolonne = "A"
                invoiceNoKolonne = "A"
            case "Fakturabeslut.04: Nul beløb i faktura":
                eanNoKolonne = "C"
                titleKolonne = "B"
                invoiceNoKolonne = "B"
            case "Fakturabeslut.08: Håndter afvist faktura":
                eanNoKolonne = "A"
                titleKolonne = "D"
                invoiceNoKolonne = "D"
                fakturabeløbKolonne = "B"
                leverandørKolonne = "D"  
            case "Fakturahandl.07: Ændre faktura":
                eanNoKolonne = "P" #EAN nr findes ikke, så for at finde et tomt felt vælges P
                titleKolonne = "B"
                invoiceNoKolonne = "B"
            case "Kombit Fakturaer":
                eanNoKolonne = "P" #EAN nr findes ikke, så for at finde et tomt felt vælges P
                titleKolonne = "B"
                invoiceNoKolonne = "B"  
            
        wb = load_workbook(filename="C:\\tmp\\EXPORT.XLSX")
        ark1 = wb["Sheet1"]
        ark1 = wb.active
        row_count = ark1.max_row
        
        if (bakken == "Fakturabeslut.08: Håndter afvist faktura" and aktueltype == "KY"):
            bakken = "KY"
        
        match bakken:
            case "Fakturabeslut.08: Håndter afvist faktura":
                list_tmp2 = []
                list_tmp3 = []
                
                obj_sess = get_client()
                obj_sess.findById("wnd[0]/tbar[0]/btn[3]").press()
                
                container = obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell")
                for child in container.Children:
                    print(child.Id, "-", child.Type)
                            
                for row_idx in range(2, row_count+1):
                    title = ark1[f"{titleKolonne}{row_idx}"].value
                    invoiceNo = re.search("[0-9]{10}",ark1[f"{invoiceNoKolonne}{row_idx}"].value).group()
                    eanNr = ark1[f"{eanNoKolonne}{row_idx}"].value
                    fakturabeløb = ark1[f"{fakturabeløbKolonne}{row_idx}"].value.strip()
                    leverandør = ark1[f"{leverandørKolonne}{row_idx}"].value[re.search("leverandør",ark1[f"{leverandørKolonne}{row_idx}"].value).span()[0]:]
                    bilag = ark1[f"{leverandørKolonne}{row_idx}"].value.split()[0]
                    
                    x1 = 0
                    tmp2 = ['1']+[bilag]+[eanNr]+[fakturabeløb]+[leverandør]+[invoiceNo]+[title]
                    findesIkke = True
                    for indhold in list_tmp2:
                        x1 = x1 + 1
                        if indhold[1:5] == tmp2[1:5]:
                            print(indhold[1:5])
                            findesIkke = False
                            print(list_tmp2[x1-1])
                            break
                    
                    if findesIkke:
                        if tmp2[1] == 'Faktura' or tmp2[1] == 'Kreditnota':
                            list_tmp2.append(tmp2)
                            list_tmp3.append(tmp2[1:7])            
                    else:
                        list_tmp2[x1-1][0] = str(int(list_tmp2[x1-1][0])+1)
                        print(list_tmp2[x1-1])
                        list_tmp3.append(tmp2[1:7])
                
                nr3 = 0
                for i in list_tmp3:
                    bilag = i[0]
                    eanNr = i[1]
                    fakturabeløb = i[2]
                    leverandør = i[3]
                    invoiceNo = i[4]
                    if bilag == 'Faktura': andetBilag = 'Kreditnota'
                    if bilag == 'Kreditnota': andetBilag = 'Faktura'
                    
                    index = next((j for j, sublist in enumerate(list_tmp2) if (sublist[1] == bilag and sublist[2] == eanNr and sublist[3] == fakturabeløb and sublist[4] == leverandør)), None)
                        
                    if not index == None:   
                        NumberOfOcurrence = list_tmp2[index][0]
                        list_tmp3[nr3] = list_tmp3[nr3] + [NumberOfOcurrence]
                    
                    nr3 = nr3+1  
                    
                for i in list_tmp3:
                    bilag = i[0]
                    eanNr = i[1]
                    fakturabeløb = i[2]
                    leverandør = i[3]
                    invoiceNo = i[4]
                    title = i[5]
                    NumberOfOcurrence = i[6]
                    if bilag == 'Faktura': andetBilag = 'Kreditnota'
                    if bilag == 'Kreditnota': andetBilag = 'Faktura'
                    
                    index = next((j for j, sublist in enumerate(list_tmp2) if (sublist[0] == NumberOfOcurrence and sublist[1] == andetBilag and sublist[2] == eanNr and sublist[3] == fakturabeløb and sublist[4] == leverandør)), None)
                        
                    if not index == None:
                        #print(NumberOfOcurrence+"  "+invoiceNo+"   "+bilag+" - "+ fakturabeløb+" - "+ eanNr)
                        create_queue_item(ark1, "", aktuel_bogholderbakke, title, invoiceNo, eanNr, fakturabeløb, leverandør)
                    
            case "KY":
                obj_sess = get_client()
                obj_sess.findById("wnd[0]").maximize()
                obj_sess.findById("wnd[0]/tbar[0]/btn[3]").press()
                grid = obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell")
                row_count = grid.RowCount
                
                for row_idx in range(2, row_count+1):
                    eanNr = ark1[f"{eanNoKolonne}{row_idx}"].value
                    bilag = ark1[f"{leverandørKolonne}{row_idx}"].value.split()[0]
                    if eanNr == "5798005775447":
                        title = ark1[f"{titleKolonne}{row_idx}"].value
                        
                        column_names = grid.ColumnOrder
                        for i in range(row_count):
                            nr2 = i
                            title_txt = grid.GetCellValue(i, column_names[3])
                            if title == title_txt:
                                break
                        
                        print(eanNr+" - Nr i listen er "+str(nr2))
                        
                        obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").selectedRows = nr2
                        obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").pressToolbarButton("APRO")
                        
                        time.sleep(20)
                        app = Application(backend="uia").connect(title_re="Bilagsvisning")
                        window = app.window(title_re="Bilagsvisning")
                        window.set_focus()
                        time.sleep(5)
                        send_keys('^a^c')
                        time.sleep(0.5)

                        win32clipboard.OpenClipboard()
                        try:
                            data = win32clipboard.GetClipboardData()
                        finally:
                            win32clipboard.CloseClipboard()

                        print("Hele tekstindholdet:\n", data)
                        
                        if ("KY" or "ky") in data:
                            print("KY fundet...")
                            create_queue_item(ark1, row_idx, bakken, titleKolonne, invoiceNoKolonne, eanNoKolonne, fakturabeløbKolonne, leverandørKolonne)
                        else:
                            print("KY, ikke fundet...")
                            
                        window.close()  
                        
                        obj_sess.findById("wnd[0]/tbar[0]/btn[12]").press()
                        
                    
            case _:
                for row_idx in range(2, row_count+1):
                    create_queue_item(ark1, row_idx, aktuel_bogholderbakke, titleKolonne, invoiceNoKolonne, eanNoKolonne, "", "")
        """

    def create_queue_item(ark1, rowidx, aktuel_bogholderB, titleK, invoiceNoK, eanNoK, fakturabeløbK, leverandørK):
        if not aktuel_bogholderB == "Fakturabeslut.08: Håndter afvist faktura":
            row_data = {
            "title": ark1[f"{titleK}{rowidx}"].value,
            "invoiceNo": re.search("[0-9]{10}",ark1[f"{invoiceNoK}{rowidx}"].value).group(),
            "eanNr": ark1[f"{eanNoK}{rowidx}"].value
            }
            if aktuel_bogholderB == "KY":
                row_data["fakturabeløb"] = ark1[f"{fakturabeløbK}{rowidx}"].value.strip()
                row_data["leverandør"] = ark1[f"{leverandørK}{rowidx}"].value[re.search("leverandør",ark1[f"{leverandørK}{rowidx}"].value).span()[0]:]
        else:
            row_data = {
                "title": titleK,
                "invoiceNo": invoiceNoK,
                "eanNr": eanNoK,
                "fakturabeløb": fakturabeløbK,
                "leverandør": leverandørK
            }
            print("   " + invoiceNoK + "   " + titleK[0:8] + " - " + fakturabeløbK + " - " + eanNoK)
            
        queue_items.append({
            "SpecificContent": row_data,
            "Reference": row_data["invoiceNo"]
        })
        
        orchestrator_connection.log_trace("Queue item no "+str(counter())+": "+row_data["title"])

    def add_queue_item():
        # Prepare references and data for the bulk creation function
        references = tuple(item["Reference"] for item in queue_items)  # Extract references as a tuple
        data = tuple(json.dumps(item["SpecificContent"], ensure_ascii=False) for item in queue_items)  # Convert SpecificContent to JSON strings

        # Bulk add queue items to OpenOrchestrator
        match aktuel_bogholderbakke:
            case "Fakturahandl.07: Ændre faktura":
                queue_name = "Bogholderbakke_ÆndreFaktura"
            case "Fakturabeslut.07: Inkonsistent XML":
                queue_name = "Bogholderbakke_XML"
            case "Fakturabeslut.04: Nul beløb i faktura":
                queue_name = "Bogholderbakke_NulBeløb"
            case "Fakturabeslut.03: Kontroller dob fakt":
                queue_name = "DobbeltFaktura"
            case "Fakturabeslut.08: Håndter afvist faktura":
                queue_name = "Bogholderbakke_HåndterAfvist"
            case "Kombit Fakturaer":
                queue_name = "Bogholderbakke_KombitFaktura"    

        try:
            orchestrator_connection.bulk_create_queue_elements(queue_name, references, data, created_by="Bogholderbakker")
            orchestrator_connection.log_info(f"Successfully added {len(queue_items)} items to the queue.")
        except Exception as e:
            print(f"An error occurred while adding items to the queue: {str(e)}")
        
    def counter():
        global n
        n += 1
        return n

    # Assign variables from SpecificContent
    specific_content = json.loads(queue_element.data)
    invoiceNo = specific_content.get("invoiceNo", None)
    title = specific_content.get("title", None)
    eanNr = specific_content.get("eanNr", None)
    fakturabeløb = specific_content.get("fakturabeløb", None)
    leverandør = specific_content.get("leverandør", None)
    
    orchestrator_connection.log_trace("New: "+title)
    
    obj_sess = get_client()
    
    grid = obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell")
    nr = 0
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
   
    obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").currentCellColumn = "WI_TEXT"
    obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").selectedRows = nr2
    obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").selectionChanged
    obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").pressToolbarButton("APRO") #for Haandter afvist åbnes WebViev
    
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
        print("Title: "+title+" - Type: "+sbar.MessageType+" - "+sbar.Text)
        orchestrator_connection.log_trace("Title: "+title+" - Type: "+sbar.MessageType+" - "+sbar.Text)
        time.sleep(2)
        obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").pressToolbarButton("EREF")
        time.sleep(2)
    if queue_element.queue_name=="Bogholderbakke_XML":
        obj_sess.findById("wnd[0]/usr/cntlSWU20300CONTAINER/shellcont/shell").sapEvent("","","SAPEVENT:DECI:0002")
        #tree = obj_sess.findById("wnd[0]/usr/txtRBKPV-BELNR")
        #print("Type:", tree.Type) #Type: GuiTextField
        #print("Id:", tree.Id)
    
    
    if queue_element.queue_name=="Bogholderbakke_DobbeltFaktura":
        obj_sess.findById("wnd[0]/usr/cntlSWU20300CONTAINER/shellcont/shell").sapEvent("","","SAPEVENT:DECI:0002")
        obj_sess.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        obj_sess.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem("&XXL")
        obj_sess.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\tmp\\"
        obj_sess.findById("wnd[1]/usr/ctxtDY_FILENAME").text = invoiceNo+"_"+queue_element.queue_name+".xlsx"
        obj_sess.findById("wnd[1]/usr/ctxtDY_PATH").setFocus
        obj_sess.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 7
        obj_sess.findById("wnd[1]/tbar[0]/btn[0]").press()
        
        #subprocess.call("taskkill /F /IM excel.exe /T", stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, shell=True)
        wb = load_workbook(filename="C:\\tmp\\"+invoiceNo+"_"+queue_element.queue_name+".XLSX")
        ark1 = wb["Sheet1"]
        ark1 = wb.active
        row_count = ark1.max_row
        tmp = []
        for row_idx in range(2, row_count+1):
            FakNo = ark1[f"{"D"}{row_idx}"].value
            Reference = ark1[f"{"J"}{row_idx}"].value
            FakturaBruttoBelob = ark1[f"{"K"}{row_idx}"].value
            Fakturaudsteder = ark1[f"{"M"}{row_idx}"].value
            Regnskabsaar = ark1[f"{"E"}{row_idx}"].value
            Bilagsdato = ark1[f"{"G"}{row_idx}"].value
            EAN = ark1[f"{"L"}{row_idx}"].value
            tmp.append({
                'FakNo':FakNo,
                'Reference':Reference, 
                'FakturaBruttoBelob':FakturaBruttoBelob, 
                'Fakturaudsteder':Fakturaudsteder, 
                'Regnskabsaar':Regnskabsaar, 
                'Bilagsdato':Bilagsdato, 
                'EAN':EAN})
            
        print("stop her")
        
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
    
    print("End process")    