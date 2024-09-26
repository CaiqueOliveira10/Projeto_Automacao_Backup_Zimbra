from selenium import webdriver
import pyautogui 
from selenium.webdriver.common.keys import Keys
from time import sleep 
from selenium.webdriver.common.by import By 
from selenium.webdriver.common.action_chains import ActionChains
import click
import os
import time
import pandas as pd

nome_do_arquivo = "email.xlsx"
df = pd.read_excel(nome_do_arquivo)

for index,row in df.iterrows():
    navegador=webdriver.Firefox()
    navegador.get("https://email.ferreiraechagas.com.br:7071/zimbraAdmin/")
    navegador.maximize_window()
    
    
    navegador.find_element(By.ID, 'ZLoginUserName').send_keys('admin') 
    navegador.find_element(By.ID, 'ZLoginPassword').send_keys('@FcR!f3rr31r4@112K24!2024') 
    navegador.find_element(By.ID, 'ZLoginButton').click() 
    
    sleep(5)
    
    busca = navegador.find_element(By.ID, '_XForm_query_display')
    sleep(3)
    busca.send_keys(row["Email"])
    busca.send_keys(Keys.RETURN) 
    
    sleep(5) 
    
    action = ActionChains(navegador)
    email = navegador.find_element(By.ID, 'zl__DWT84__rows') 
    
    sleep(5)
    
    webdriver.ActionChains(navegador).double_click(on_element=email).perform() 
    
    sleep(5) 

    navegador.find_element(By.ID, 'ztabv__ACCT_EDIT_zimbraAccountStatus_2_display').click()

    sleep(5)

    navegador.find_element(By.ID, 'ztabv__ACCT_EDIT_zimbraAccountStatus_2_choice_0').click()
    
    sleep(5)

    navegador.find_element(By.ID, 'ztabv__ACCT_EDIT_password').send_keys('Mudar@123') 
    navegador.find_element(By.ID, 'ztabv__ACCT_EDIT_confirmPassword').send_keys('Mudar@123') 
    
    sleep(5) 
    
    navegador.find_element(By.ID, 'zb__ZaCurrentAppBar__SAVE_title').click()

    sleep(5) 
    
    navegador.get('https://email.ferreiraechagas.com.br/')

    sleep(5)  

    busca = navegador.find_element(By.ID, 'username')
    busca.send_keys(row["Email"])
    busca.send_keys(Keys.RETURN)

    sleep(5) 
    
    navegador.find_element(By.ID, 'password').send_keys('Mudar@123') 

    sleep(5) 

    navegador.find_element(By.CLASS_NAME, 'ZLoginButton').click()

    sleep(5) 

    navegador.find_element(By. ID, 'zb__App__Options_title').click()
    
    sleep(5)

    navegador.find_element(By.ID, 'zti__main_Options__PREF_PAGE_IMPORT_EXPORT_textCell').click()

    sleep(5)

    navegador.find_element(By.ID, 'EXPORT_BUTTON_title').click()
    
    navegador.quit
   