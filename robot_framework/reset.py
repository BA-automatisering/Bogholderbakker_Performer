"""This module handles resetting the state of the computer so the robot can work with a clean slate."""

import subprocess
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from robot_framework import globals

import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
import pyautogui
import win32com.client
import json


def reset(orchestrator_connection: OrchestratorConnection) -> None:
    """Clean up, close/kill all programs and start them again. """
    orchestrator_connection.log_trace("Resetting.")
    clean_up(orchestrator_connection)
    close_all(orchestrator_connection)
    kill_all(orchestrator_connection)
    open_all(orchestrator_connection)


def clean_up(orchestrator_connection: OrchestratorConnection) -> None:
    """Do any cleanup needed to leave a blank slate."""
    orchestrator_connection.log_trace("Doing cleanup.")


def close_all(orchestrator_connection: OrchestratorConnection) -> None:
    """Gracefully close all applications used by the robot."""
    orchestrator_connection.log_trace("Closing all applications.")


def kill_all(orchestrator_connection: OrchestratorConnection) -> None:
    """Forcefully close all applications used by the robot."""
    orchestrator_connection.log_trace("Killing all applications.")

    subprocess.call("taskkill /F /IM msedge.exe /T", stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, shell=True)
    subprocess.call("taskkill /F /IM chrome.exe /T", stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, shell=True)
    subprocess.call("taskkill /F /IM chromedriver.exe /T", stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, shell=True)
    subprocess.call("taskkill /F /IM excel.exe /T", stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, shell=True)
    subprocess.call("taskkill /F /IM saplogon.exe /T", stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, shell=True)
    subprocess.call("taskkill /F /IM sapgui.exe /T", stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, shell=True)

def open_all(orchestrator_connection: OrchestratorConnection) -> None:
    """Open all programs used by the robot."""
    orchestrator_connection.log_trace("Opening all applications.")
    
    opusbruger_navn="OpusBruger_Bog" 
    OpusLogin = orchestrator_connection.get_credential(opusbruger_navn)
    OpusUser = OpusLogin.username
    OpusPassword = OpusLogin.password

    downloads_folder="C:\\tmp"
    bakken = ''
    queue_items = []
    globals.manuelliste = []
    
    file = "C:\\tmp\\EXPORT.XLSX"
    if os.path.exists(file):
        os.remove(file)
        print("C:\\tmp\\EXPORT.XLSX er slettet før start...")

    # Configure Chrome options
    chrome_options = ChromeOptions()
    chrome_options.add_argument('--remote-debugging-pipe')
    #chrome_options.add_argument("--headless=new")  # More stable headless mode
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": downloads_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        })
    chrome_options.add_argument("--incognito")

    chrome_service = ChromeService()
    driver = webdriver.Chrome(service=chrome_service, options=chrome_options)
    
    #Ved PROD bruges denne linje
    globals.aktuel_bogholderbakke = json.loads(orchestrator_connection.process_arguments)['aktuel_bogholderbakke']

    #Ved TEST lokalt bruges nedenstående parametre...
    #globals.aktuel_bogholderbakke = "Fakturahandl.07: Ændre faktura"
    #globals.aktuel_bogholderbakke = "Fakturabeslut.07: Inkonsistent XML"
    #aktuel_bogholderbakke = "Kombit Fakturaer"
    #globals.aktuel_bogholderbakke = "Fakturabeslut.03: Kontroller dob fakt"
    #globals.aktuel_bogholderbakke = "Fakturabeslut.04: Nul beløb i faktura"
    #globals.aktuel_bogholderbakke = "Fakturabeslut.08: Håndter afvist faktura"
    #globals.aktuel_bogholderbakke = "FakturaKontrolCenter"
        
    orchestrator_connection.log_trace("Running: "+globals.aktuel_bogholderbakke)
    print("Running: "+globals.aktuel_bogholderbakke)
    
    def open_RI(driver):
        #orchestrator_connection.log_trace("open_RI started...")
        driver.get("https://portal.kmd.dk/irj/portal")
        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, "logonuidfield")))
        driver.maximize_window()
        driver.find_element(By.ID, "logonuidfield").send_keys(OpusUser)
        driver.find_element(By.ID, "logonpassfield").send_keys(OpusPassword)
        driver.find_element(By.ID, "buttonLogon").click()
        try:
            time.sleep(3)
            driver.find_element(By.CLASS_NAME, "button_inner")
            orchestrator_connection.log_trace("Logged in okay")
        except Exception as e:
            orchestrator_connection.log_trace("Password skal skiftes...")
            #new_password.newpass(driver,opusbruger_navn,OpusUser,OpusPassword)

    
    def open_SAP(driver):
        #orchestrator_connection.log_trace("open_SAP started...")
        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//div[@class="TabText_SmallTabs" and text()="Mine Genveje"]')))
        driver.find_element(By.XPATH, '//div[@class="TabText_SmallTabs" and text()="Mine Genveje"]').click()
                
        time.sleep(2)
        pyautogui.press('enter')
        time.sleep(2)
        
        try:
            path = "C:\\Users\\"+os.getenv('TEMP').split("\\")[2]+"\\Overførsler\\tx.sap"
            os.startfile(path)
        except Exception:
            path = "C:\\Users\\"+os.getenv('TEMP').split("\\")[2]+"\\Downloads\\tx.sap"
            os.startfile(path)
    
        time.sleep(3)
        orchestrator_connection.log_trace("SAP is open")


    def goto_bogholderbakker_i_SAP():
        #orchestrator_connection.log_trace("goto_bogholderbakker_i_SAP started...")
        time.sleep(3)
        obj_sess = get_client()
        obj_sess.findById("wnd[0]").maximize()
        obj_sess.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").expandNode("0000001203")
        obj_sess.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "0000001204"
        obj_sess.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").topNode = "0000000354"
        obj_sess.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode("0000001204")
        
        btn = more_than_200(obj_sess, "wnd[1]/usr/btnDY_VAROPTION3")
        if btn:
            btn.press()
        else:
            print("Not more than 200 in inboks")
            
        obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[0]/shell").expandNode("          5")
        obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[0]/shell").topNode = "          1"
        obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[0]/shell").expandNode("         24")
        obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[0]/shell").topNode = "          1"


    def go_to_specific_bakke():
        orchestrator_connection.log_trace("go_to_specifik_bakkke started...")
        obj_sess = get_client()
        
        tree = obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[0]/shell")
        node_keys = tree.GetAllNodeKeys()
        selectedNodeList = []
        for key in node_keys:
            #print(key, "-", tree.GetNodeTextByKey(key)) #Viser sammenhæng mellem navn og nummer
            selectedNodeList.append({"key":key,"name":tree.GetNodeTextByKey(key)})
            
        id = next((p for p in selectedNodeList if p["name"].lower() == globals.aktuel_bogholderbakke.lower()), None)

        if not id == None:
            for x in id.items():
                if x[0] == "key":
                    nr = x[1]

            obj_sess.findById("wnd[0]").maximize()
            obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "         24"
            obj_sess.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = nr  #Her er nummer som passer med navnet

        else:
            print(f"Bogholderbakken '{globals.aktuel_bogholderbakke}' er ikke aktuel lige nu...")


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
   

    def more_than_200(session, id_str):
        try:
            return session.findById(id_str)
        except Exception:
            return None



    open_RI(driver)
    time.sleep(3)
    open_SAP(driver)
    if not globals.aktuel_bogholderbakke == "FakturaKontrolCenter":    
        goto_bogholderbakker_i_SAP()
        go_to_specific_bakke()
    orchestrator_connection.log_trace("Opening all applications - end")
    
