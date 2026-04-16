import os
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import random 
import string
import time


def newpass(orchestrator_connection,driver,bruger_navn: str,OpusUser: str,old_password: str):
    print('Trying to find change button')
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, "changeButton")))
    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.ID, "changeButton")))
                
    lower = string.ascii_lowercase
    upper = string.ascii_uppercase
    digits = string.digits
    special = "!@#&%"

    password_chars = []
    password_chars += random.choices(lower, k=4)
    password_chars += random.choices(upper, k=4)
    password_chars += random.choices(digits, k=4)
    password_chars += random.choices(special, k=2)

    random.shuffle(password_chars)
    password = ''.join(password_chars)

    driver.find_element(By.ID, "inputUsername").send_keys(old_password)
    driver.find_element(By.NAME, "j_sap_password").send_keys(password)
    driver.find_element(By.NAME, "j_sap_again").send_keys(password)
    driver.find_element(By.ID, "changeButton").click()
    
    orchestrator_connection.update_credential(bruger_navn, OpusUser, password)
    orchestrator_connection.log_info('Password changed and credential updated for '+OpusUser)
    time.sleep(2)