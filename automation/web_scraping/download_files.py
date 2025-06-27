# main.py
from utils.logger import setup_logger
from utils.selectors import SELECTORS, get_element
from utils.clean_excel import clean_excel_files
from config import USER, PASSWORD, DOWNLOAD_PATH, PROJECT_FILE

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException

import os
import time
import pandas as pd
from datetime import datetime


class SREAutomation:
    def __init__(self):
        self.logger = setup_logger("sre_automation.log")
        self.driver = None
        self.projects = []
        self.current_date = datetime.now().strftime("%Y-%m-%d")

        # Clean folders
        clean_excel_files("files_output_task")
        clean_excel_files("Dowloawds_Task")

    def read_projects(self):
        try:
            df = pd.read_excel(PROJECT_FILE, dtype=str)
            self.projects = df['Project'].dropna().unique().tolist()
            self.logger.info(f"Loaded {len(self.projects)} projects.")
        except Exception as e:
            self.logger.error(f"Failed to load projects: {e}")

    def open_browser(self):
        options = Options()
        prefs = {
            "download.default_directory": DOWNLOAD_PATH,
            "safebrowsing.enabled": "false",
            "credentials_enable_service": False,
            "profile.password_manager_enabled": False
        }
        options.add_experimental_option("prefs", prefs)
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')

        try:
            self.driver = webdriver.Chrome(service=Service(), options=options)
            self.logger.info("Browser opened successfully.")
        except Exception as e:
            self.logger.error(f"Error starting browser: {e}")

    def login(self):
        self.driver.get("pagina")
        self.driver.maximize_window()

        try:
            WebDriverWait(self.driver, 60).until(
                EC.presence_of_element_located(SELECTORS['login_logo'])
            )

            get_element(self.driver, 'username').send_keys(USER)
            get_element(self.driver, 'password').send_keys(PASSWORD)
            get_element(self.driver, 'submit').click()

            WebDriverWait(self.driver, 120).until(
                EC.presence_of_element_located(SELECTORS['my_projects'])
            )
            self.logger.info("Login successful.")
        except Exception as e:
            self.logger.error(f"Login failed: {e}")

    def run(self):
        self.read_projects()
        self.open_browser()
        self.login()
        # Agrega tus funciones process_projects, changes_projects, etc. aquí...

        self.driver.quit()
        self.logger.info("Automation finished.")


if __name__ == "__main__":
    bot = SREAutomation()
    bot.run()
