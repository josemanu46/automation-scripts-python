from selenium import webdriver
import selenium
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
import os
import pdb
import time
from time import sleep
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException



class main:
    
    def __init__(self):
        self.driver = str(os.path.dirname(os.path.realpath(__file__)))+ r"\driver\chromedriver.exe" # RUTA DE EL DRIVER 
        self.user_name = ""
        self.password = ""

        #llamamos a la funcion login
        self.login()

    #cargaar pagina y logear 
    def login(self):
        

        print("login")
        time.sleep(2)
        self.driver = webdriver.Chrome() 
        self.driver.get("https://www.icloud.com/photos/")
        self.driver.maximize_window()
        self.driver.refresh()
        time.sleep(5)
        

        try:
            wait = WebDriverWait(self.driver, 10)  
            wait.until(EC.presence_of_element_located((By.XPATH, '//*[contains(text(),"Iniciar sesión")]')))
            print("Element found, stopping refresh")
        except TimeoutException:
            print("Element NO found, REFRESH")
            self.driver.refresh()
            try:
                wait.until(EC.presence_of_element_located((By.XPATH, '//*[contains(text(),"Iniciar sesión")]')))
                print("Element found, stopping refresh")
            except TimeoutException:
                print("Element NO found,")

        time.sleep(2)
        self.driver.find_element(By.XPATH, '//*[contains(text(),"Iniciar sesión")]').click()
        print("login whit : ",self.user_name)
        time.sleep(5)
        # Cambiar al iframe por ID
        self.driver.switch_to.frame("aid-auth-widget-iFrame")
        login_user = self.driver.find_element(By.ID, 'account_name_text_field')
        login_user.send_keys(self.user_name)
        print('login_user',self.user_name)
        time.sleep(2)
        self.driver.find_element(By.XPATH, '//*[@id="sign-in"]').click()
        time.sleep(2)
        
        login_pass = self.driver.find_element(By.ID, 'password_text_field')
        login_pass.send_keys(self.password)
        print('login_pass, ***********')
        self.driver.find_element(By.XPATH, '//*[@id="sign-in"]').click()

        time.sleep(10)

        iframe = self.driver.find_element(By.XPATH, '//*[@id="root"]/ui-main-pane/div/div[3]/iframe')
        self.driver.switch_to.frame(iframe)
    
        while True:
            try:
                print('waiting form 10 seconds')
                WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[contains(text(),"Enlaces de iCloud")]')))
                
                print('Element found, stopping refresh')
                break  
            except:
                print('Element not found, refreshing page...')
                self.driver.refresh()
                time.sleep(5)

        time.sleep(2)

        print('Login DONE')


run = main()