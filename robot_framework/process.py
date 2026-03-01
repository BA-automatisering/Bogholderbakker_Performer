"""This module contains the main process of the robot."""

import os
import time
import random 
import string
import subprocess
import re
import json
from collections import Counter
from robot_framework import globals

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
from OpenOrchestrator.database.queues import QueueStatus
from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
import pyautogui
import win32com.client
#import new_password
from openpyxl import load_workbook, workbook
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto import mouse
import win32clipboard
import win32gui
import win32con
import win32api

from robot_framework.exceptions import BusinessError

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

    def add_queue_items_to_queue(target, source):
        # Prepare references and data for the bulk creation function
        references = tuple(item["Reference"] for item in queue_items)  # Extract references as a tuple
        data = tuple(json.dumps(item["SpecificContent"], ensure_ascii=False) for item in queue_items)  # Convert SpecificContent to JSON strings

        # Bulk add queue items to OpenOrchestrator
       
        try:
            orchestrator_connection.bulk_create_queue_elements(target, references, data, created_by=source)
            orchestrator_connection.log_info(f"Successfully added {len(queue_items)} items to the queue.")
        except Exception as e:
            print(f"An error occurred while adding items to the queue: {str(e)}")
        
    
    specific_content = json.loads(queue_element.data)
    # Assign variables from SpecificContent
    invoiceNo = specific_content.get("invoiceNo", None)
    title = specific_content.get("title", None).strip()
    eanNr = specific_content.get("eanNr", None)
    fakturabeløb = specific_content.get("fakturabeløb", None)
    leverandør = specific_content.get("leverandør", None)
    
    orchestrator_connection.log_trace("NEW: "+title)
    print("NEW: "+title)
    time.sleep(1)
    
    obj_sess = get_client()
    
    if not globals.aktuel_bogholderbakke == "FakturaKontrolCenter":
        time.sleep(1)
        grid = obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell")
        
        obj_sess.findById("wnd[0]/mbar/menu[3]/menu[6]").select() #Opdater siden...
        time.sleep(1)
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
        
        if not id==None:
            nr2 = id["no"]
            #print("Nr i liste = "+str(nr2))
            #orchestrator_connection.log_trace("Nr i liste = "+str(nr2))
            time.sleep(2)
            obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").currentCellColumn = "WI_TEXT"
            obj_sess.findById("wnd[0]/mbar/menu[3]/menu[6]").select() #Opdater siden...
            obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").selectedRows = nr2
            time.sleep(2)
            obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").selectionChanged
            time.sleep(2)
            obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").pressToolbarButton("APRO") #for 'Haandter afvist' åbnes WebViev
            
            if queue_element.queue_name=="Bogholderbakke_NulBeløb":
                obj_sess.findById("wnd[0]/usr/cntlSWU20300CONTAINER/shellcont/shell").sapEvent("", "", "SAPEVENT:DECI:0001")
                obj_sess.findById("wnd[0]/usr/txtRBKPV-BELNR").setFocus
                obj_sess.findById("wnd[0]/usr/txtRBKPV-BELNR").caretPosition = 10

                invoiceNo_txt = obj_sess.findById("wnd[0]/usr/txtRBKPV-BELNR").Text
                if invoiceNo == invoiceNo_txt:
                    print("invoiceNo_txt "+invoiceNo_txt+" Korrekt er åbnet")
                    obj_sess.findById("wnd[0]/mbar/menu[0]/menu[6]").select() #Her slettes bilaget
                    sbar = obj_sess.findById("wnd[0]/sbar")
                    print("invoiceNo: "+invoiceNo+" - Type: "+sbar.MessageType+" - "+sbar.Text)
                    orchestrator_connection.log_trace("invoiceNo: "+invoiceNo+" - Type: "+sbar.MessageType+" - "+sbar.Text)
                    time.sleep(2)
                    #obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").pressToolbarButton("EREF") #Hvad sker her?
                #time.sleep(2)
            
            if queue_element.queue_name=="Bogholderbakke_XML":
                obj_sess.findById("wnd[0]/usr/cntlSWU20300CONTAINER/shellcont/shell").sapEvent("","","SAPEVENT:DECI:0002")
                obj_sess.findById("wnd[0]/mbar/menu[0]/menu[3]").select() #Gem forudregistreret bilag
                i = 1
                while i < 10:
                    sbar = obj_sess.findById("wnd[0]/sbar")
                    print("Type: "+sbar.MessageType+" - Text: "+sbar.Text)
                    if i == 10 or (not sbar.MessageType == "E" and not sbar.MessageType == "W") :
                        break
                    #pyautogui.press('enter')
                    obj_sess.findById("wnd[0]/tbar[0]/btn[15]").press() #Afslut - gul knap
                    obj_sess.findById("wnd[1]/usr/btnSPOP-OPTION1").press() #Ja
                    time.sleep(1)
                    i += 1
                obj_sess.findById("wnd[0]/tbar[0]/btn[12]").press() #Afbryd - rød knap
                obj_sess.findById("wnd[1]/usr/btnSPOP-OPTION1").press() #Ja
                obj_sess.findById("wnd[1]/usr/btnSPOP-OPTION1").press() #Fortsæt

                
                
                sbar = obj_sess.findById("wnd[0]/sbar")
                print("invoiceNo: "+invoiceNo+" - Type: "+sbar.MessageType+" - "+sbar.Text)
                orchestrator_connection.log_trace("invoiceNo: "+invoiceNo+" - Type: "+sbar.MessageType+" - "+sbar.Text)
                time.sleep(1)

                #pyautogui.press('enter')

                #obj_sess.findById("wnd[1]/usr/btnSPOP-OPTION2").press() #Dette er nej
                #Klik Ja derefter Enter hvis der er en besked i bunden
                #obj_sess.findById("wnd[1]/usr/btnSPOP-OPTION1").press()

            if queue_element.queue_name=="Bogholderbakke_DobbeltFaktura":
                obj_sess = get_client()
                obj_sess.findById("wnd[0]/usr/cntlSWU20300CONTAINER/shellcont/shell").sapEvent("","","SAPEVENT:DECI:0002")
                
                #Træk direkte fra siden
                grid = obj_sess.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell")
                
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
                correct = invoiceNo in Counter(d["FakNo"] for d in tmp)
                if not correct:
                    print("Forkert er valgt...")
                    orchestrator_connection.log_trace("Forkert er valgt... her skal laves ERROR")
                #print("FakNo")
                #print(resultat)
                
                
                noOfRowsTotal = len(tmp)
                noOfRowsFakturaNr = len(Counter(d["FakNo"] for d in tmp))
                noOfRowsReference = len(Counter(d["Reference"] for d in tmp))
                noOfRowsFakturabeloeb = len(Counter(d["FakturaBruttoBelob"] for d in tmp))
                noOfRowsFakturaudsteder = len(Counter(d["Fakturaudsteder"] for d in tmp))
                noOfRowsAar = len(Counter(d["Regnskabsaar"] for d in tmp))
                noOfRowsBilagsdato = len(Counter(d["Bilagsdato"] for d in tmp))
                noOfRowsEAN = len(Counter(d["EAN"] for d in tmp))
                
                rule = 0
                if (noOfRowsTotal == noOfRowsFakturaNr and noOfRowsReference == 1 and noOfRowsFakturabeloeb == 1 and noOfRowsFakturaudsteder == 1 and noOfRowsAar == 1 and noOfRowsEAN == 1 and noOfRowsTotal > 1):
                    rule = 1
                if noOfRowsTotal == 1:
                    rule = 2
                if (noOfRowsTotal == noOfRowsFakturaNr and noOfRowsReference == 1 and noOfRowsFakturabeloeb == 1 and noOfRowsFakturaudsteder == 1 and noOfRowsAar > 1 and noOfRowsTotal > 1):    
                    rule = 3
                if (rule == 0):
                    rule = 4
                
                match rule:
                    case 1:
                        print("Kontrol af faktura - rule 1 - slettes...")
                        orchestrator_connection.log_trace("Kontrol af faktura - rule 1 - slettes...")
                        
                        #Bilag slettes her...
                        obj_sess.findById("wnd[0]/tbar[0]/btn[3]").press()
                        obj_sess.findById("wnd[0]/usr/cntlSWU20300CONTAINER/shellcont/shell").sapEvent("","","SAPEVENT:DECI:0001")
                        invoiceNo_txt = obj_sess.findById("wnd[0]/usr/txtRBKPV-BELNR").Text
                        print("invoiceNo_txt "+invoiceNo_txt)
                
                        if invoiceNo == invoiceNo_txt:
                            #print("Korrekt er åbnet...")
                            obj_sess.findById("wnd[0]/mbar/menu[0]/menu[6]").select()
                            sbar = obj_sess.findById("wnd[0]/sbar")
                            print("invoiceNo: "+invoiceNo+" - Type: "+sbar.MessageType+" - "+sbar.Text)
                            orchestrator_connection.log_trace("invoiceNo: "+invoiceNo+" - Type: "+sbar.MessageType+" - "+sbar.Text)
                        else:
                            print("Korrekt faktura IKKE åbnet...")
                        
                    case 2:
                        print("Kun 1 faktura - rule 2 - kø-element til Faktura Kontrol Center")
                        orchestrator_connection.log_trace("Kun 1 faktura - rule 2 - kø-element til Faktura Kontrol Center")
                        #Til manuel liste
                        globals.manuelliste.append({
                            "Område": "Fakturabeslut 03 - Kontroller dob fakt",
                            "Fakturanr": invoiceNo,
                            "Beskrivelse": "Kun en faktura; derfor slettes den ikke... frigives til Bruger senere i dag"
                        })
                        #Kø-element til Faktura Kontrol Center
                        row_data = {
                        "invoiceNo": invoiceNo,
                        "title": title
                        }
                        queue_items =[]
                        queue_items.append({
                            "SpecificContent": row_data,
                            "Reference": row_data["invoiceNo"]
                        })
                        add_queue_items_to_queue("Bogholderbakke_FakturaKontrolCenter","DobbeltFaktura")
                        
                        obj_sess.findById("wnd[0]/tbar[0]/btn[3]").press()
                        obj_sess.findById("wnd[0]/usr/cntlSWU20300CONTAINER/shellcont/shell").sapEvent("","","SAPEVENT:DECI:0000")
                        
                    case 3:
                        print("Aarstal ikke ens - rule 3 - manuelliste - frigives ikke til Bruger... Antal bilagsdato "+str(noOfRowsBilagsdato))
                        orchestrator_connection.log_trace("Aarstal ikke ens - rule 3 - manuelliste - frigives ikke til Bruger...")
                        globals.manuelliste.append({
                            "Område": "Fakturabeslut 03 - Kontroller dob fakt",
                            "Fakturanr": invoiceNo,
                            "Beskrivelse": "Fakturaer er fra forskellige regnskabsår - ingen er slettet. Vil IKKE blive frigivet til Bruger senere i dag..."
                        })
                        obj_sess.findById("wnd[0]/tbar[0]/btn[3]").press()
                        obj_sess.findById("wnd[0]/usr/cntlSWU20300CONTAINER/shellcont/shell").sapEvent("","","SAPEVENT:DECI:0000")

                    case 4:
                        print("Ingen rule valgt endnu... - rule 4")
                        orchestrator_connection.log_trace("Ingen rule valgt endnu... - rule 4")
                        globals.manuelliste.append({
                            "Område": "Fakturabeslut 03 - Kontroller dob fakt",
                            "Fakturanr": invoiceNo,
                            "Beskrivelse": "Robotten har ingen regler for denne..."
                        })
                        obj_sess.findById("wnd[0]/tbar[0]/btn[3]").press()
                        obj_sess.findById("wnd[0]/usr/cntlSWU20300CONTAINER/shellcont/shell").sapEvent("","","SAPEVENT:DECI:0000")
                        
                    case _:
                        print("Alt andet...")
                        orchestrator_connection.log_trace("Alt andet...")
                        obj_sess.findById("wnd[0]/tbar[0]/btn[3]").press()
                        obj_sess.findById("wnd[0]/usr/cntlSWU20300CONTAINER/shellcont/shell").sapEvent("","","SAPEVENT:DECI:0000")
                    
                                        
                        
                time.sleep(3)
                orchestrator_connection.log_trace("Rule: "+str(rule))
                    
            if queue_element.queue_name=="Bogholderbakke_HåndterAfvist":
                obj_sess = get_client()
                obj_sess.findById("wnd[0]/usr/cntlSWU20300CONTAINER/shellcont/shell").sapEvent("","","SAPEVENT:DECI:0002")
            
                #Tjek om den korrekte er åbnet
                container = obj_sess.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subVENDOR_DATA:SAPLMR1M:6510")
                children = container.Children
                
                for i in range(children.Count):
                    c = children.ElementAt(i)
                    """
                    try:
                        print(i, c.Id, c.Type, getattr(c, "Text", None), getattr(c, "Tooltip", None))
                    except Exception as e:
                        print(i, c.Id, c.Type, "err:", e)
                    """
                
                leverndør_txt = (children.ElementAt(1)).Text
                if not leverandør == ("leverandør "+leverndør_txt):
                    leverndør_txt = (children.ElementAt(3)).Text
                
                invoiceNo_txt = obj_sess.findById("wnd[0]/usr/txtRBKPV-BELNR").Text
                print("invoiceNo_txt "+invoiceNo_txt)
                
                if leverandør == ("leverandør "+leverndør_txt) and invoiceNo == invoiceNo_txt:
                    print("Korrekt er åbnet...")
                    KreditorLinje = (children.ElementAt(0)).Text
                    KreditorLinje = KreditorLinje.split(" ")
                    x = KreditorLinje[1]
                    while x[:1] == "0":
                        x = x[1:]
                    
                    if len(x) < 5:
                        print("Internt bilag - slettes ikke...")
                        orchestrator_connection.log_trace("Internt bilag - slettes ikke...")
                        #Manuelliste...
                        globals.manuelliste.append({
                            "Område": "Fakturabeslut 08 - Haandter afvist faktura",
                            "Fakturanr": invoiceNo,
                            "Beskrivelse": "Internt bilag - slettes ikke"
                        })
                    else:
                        obj_sess.findById("wnd[0]/mbar/menu[0]/menu[6]").select() #Klik Slet
                        time.sleep(2)
                        sbar = obj_sess.findById("wnd[0]/sbar")
                        print("invoiceNo: "+invoiceNo+" - Type: "+sbar.MessageType+" - "+sbar.Text)
                        orchestrator_connection.log_trace("invoiceNo: "+invoiceNo+" - Type: "+sbar.MessageType+" - "+sbar.Text)    
                else:    
                    print("Korrekt faktura IKKE åbnet...")
                    orchestrator_connection.log_trace("Korrekt faktura IKKE åbnet... laves et nyt køelement")
                    
                    queue_items =[]
                    queue_items.append({
                        "SpecificContent": row_data,
                        "Reference": row_data["invoiceNo"]
                    })
                    add_queue_items_to_queue("Bogholderbakke_HåndterAfvist_igen","HaandterafvistFaktura")
            
            if queue_element.queue_name=="Bogholderbakke_ÆndreFaktura":
                def Bogføringsperiode_Moms():
                    info = ""
                    try:
                        obj_sess.findById("wnd[1]/tbar[0]/btn[0]").press() #Klik 'eneste knap' til mulig bogføringsperiode
                        info = info + " Bogføring-vises"
                    except:
                        info = info + " Bogføring-vises-ikke"   
                    try:
                        obj_sess.findById("wnd[1]/usr/btnBUTTON_2").press() #NEJ, til 'ikke momsbærende ændres til momsbærende...'
                        info = info + " MOMS-vises"
                    except:
                        info = info + " MOMS-vises-ikke"
                    print(info)
                    orchestrator_connection.log_trace(info)
                        
                invoiceNo_txt = obj_sess.findById("wnd[0]/usr/txtRBKPV-BELNR").Text
                if invoiceNo == invoiceNo_txt:
                    print("Korrekt åbnet...")
                    obj_sess.findById("wnd[0]/mbar/menu[0]/menu[3]").select() #Gem forudregistreret bilag
                    Bogføringsperiode_Moms()
                    
                    i = 1
                    time.sleep(1)
                    while i < 6:
                        sbar = obj_sess.findById("wnd[0]/sbar")
                        print("Type: "+sbar.MessageType+" - Text: "+sbar.Text)
                        orchestrator_connection.log_trace(str(i)+" Type: "+sbar.MessageType+" - Text: "+sbar.Text)
                        if i == 5 or (not sbar.MessageType == "E" and not sbar.MessageType == "W") :
                            break
                        obj_sess.findById("wnd[0]/tbar[0]/btn[11]").press() #Gem forudregistreret bilag - knap
                        Bogføringsperiode_Moms()
                        time.sleep(1)
                        i += 1
                    time.sleep(1)
                    
                    try:
                        invoiceNo_txt = obj_sess.findById("wnd[0]/usr/txtRBKPV-BELNR").Text
                        if invoiceNo == invoiceNo_txt:
                            obj_sess.findById("wnd[0]/tbar[0]/btn[12]").press() #Afbryd - rød knap
                            obj_sess.findById("wnd[1]/usr/btnSPOP-OPTION1").press() #Ja
                            obj_sess.findById("wnd[1]/usr/btnSPOP-OPTION1").press() #Fortsæt
                    except:
                        print("Er tilbage ved listen...")
                        orchestrator_connection.log_trace("Er tilbage ved listen...")
                    
                
   
        else:
            orchestrator_connection.log_trace("Title '"+title+ "' Opslaget gav intet resultat...")
            #Der skal laves en error her    
               
    else:
        obj_sess.findById("wnd[0]/tbar[0]/okcd").text = "ZMIR6"
        obj_sess.findById("wnd[0]/tbar[0]/btn[0]").press()
        obj_sess.findById("wnd[0]/usr/txtS_BILAG-LOW").text = invoiceNo
        obj_sess.findById("wnd[0]/usr/txtS_BILAG-HIGH").text = ""
        obj_sess.findById("wnd[0]/usr/txtS_GJAHR-LOW").text = "2026"
        obj_sess.findById("wnd[0]/usr/ctxtS_BLART-LOW").text = "RE"
        obj_sess.findById("wnd[0]/usr/ctxtS_CPUDT-LOW").text = ""
        obj_sess.findById("wnd[0]/usr/ctxtS_CPUDT-LOW").setFocus
        obj_sess.findById("wnd[0]/usr/ctxtS_CPUDT-LOW").caretPosition = 0
        obj_sess.findById("wnd[0]/tbar[1]/btn[8]").press()
        obj_sess.findById("wnd[0]/mbar/menu[0]/menu[6]").select()
        sbar = obj_sess.findById("wnd[0]/sbar")
        print("invoiceNo: "+invoiceNo+" - Type: "+sbar.MessageType+" - "+sbar.Text)
        orchestrator_connection.log_trace("invoiceNo: "+invoiceNo+" - Type: "+sbar.MessageType+" - "+sbar.Text)
        if sbar.Text == "Venligst kør program i baggrund, hvis start dato er ældre end 2 månede":
            print("Klik Enter")
        obj_sess.findById("wnd[0]").resizeWorkingPane(139,26,False)
        obj_sess.findById("wnd[0]/usr/cntlCUSTOM_CONTROL/shellcont/shell").setCurrentCell(-1,"")
        obj_sess.findById("wnd[0]/usr/cntlCUSTOM_CONTROL/shellcont/shell").selectAll()
        obj_sess.findById("wnd[0]/usr/cntlCUSTOM_CONTROL/shellcont/shell").pressToolbarButton("EXECUTE")
        obj_sess.findById("wnd[1]/usr/btnBUTTON_1").press()
        obj_sess.findById("wnd[0]/usr/cntlCUSTOM_CONTROL/shellcont/shell").pressToolbarButton("REFRESH")
        time.sleep(1)
        obj_sess.findById("wnd[0]/tbar[0]/btn[12]").press()
        time.sleep(1)
        obj_sess.findById("wnd[0]/tbar[0]/btn[12]").press()
    
        
        orchestrator_connection.log_trace("Workflow er genstartet og opdateret...")
        print("Workflow er genstartet og opdateret...")

    print("Running process - end")
    #orchestrator_connection.log_trace("Running process - end")
