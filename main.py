# https://www.instagram.com/leoo_esteves1/
# https://github.com/DotcomPeep

from selenium import webdriver
from selenium.webdriver.common.by import By
import logging
import time
import pandas as pd
import os
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from ips import get_ip_data

load_dotenv()

class PFSenseAutomation:

    PFSENSE_HOSTNAME = os.getenv('PFSENSE_HOSTNAME')
    PFSENSE_PASSWORD = os.getenv('PFSENSE_PASSWORD')

    def __init__(self, username, password):
        self.username = username
        self.password = password
        self.driver = None

    def initialize_browser(self):
        self.driver = webdriver.Chrome()
        self.driver.get('Enter your PFSense URL here')
        self.driver.maximize_window()
        logging.info("Opening Chrome and entering PFSense!")

    def login(self):
        assert "Erro de privacidade" in self.driver.title # change the message according to your browser language
        details_button = self.driver.find_element(By.ID, 'details-button')
        details_button.click()
        proceed_link = self.driver.find_element(By.ID, 'proceed-link')
        proceed_link.click()
        time.sleep(3)
        login_field = self.driver.find_element(By.NAME, 'usernamefld')
        login_field.clear()
        login_field.send_keys(self.username)
        logging.info("Entered username")
        password_field = self.driver.find_element(By.NAME, 'passwordfld')
        password_field.clear()
        password_field.send_keys(self.password)
        logging.info("Entered Password")
        login_button = self.driver.find_element(By.NAME, 'login')
        login_button.click()
        logging.info("User Logged!")
        time.sleep(3)

    def access_openvpn_menu(self):
        access_vpn = self.driver.find_element(By.XPATH, '//*[@id="pf-navbar"]/ul[1]/li[5]/a')
        access_vpn.click()
        logging.info("Accessing the OpenVPN page")
        time.sleep(1)
        openvpn = self.driver.find_element(By.XPATH, '//*[@id="pf-navbar"]/ul[1]/li[5]/ul/li[3]/a')
        openvpn.click()
        logging.info("OpenVPN page found")
        time.sleep(1)

    def get_user_ip_data(self):
        specified = self.driver.find_element(By.XPATH, '//*[@id="2"]/div/ul/li[3]/a')
        specified.click()
        logging.info("Found the list of Users and IP's. Checking...")
        time.sleep(3)

        users = self.driver.find_elements(By.XPATH, '//*[@id="2"]/div/div/div[2]/table/tbody/tr/td[2]')
        user_list = [user.text for user in users]
        logging.info("Getting list of Users...")

        ips = self.driver.find_elements(By.XPATH, '//*[@id="2"]/div/div/div[2]/table/tbody/tr/td[3]')
        ips_list = [ip.text for ip in ips]
        logging.info("Getting list of IP's...")

        data = [{"hostname": user, "ip_address": ip} for user, ip in zip(user_list, ips_list)]
        logging.info("List of created Hostnames and IPs")

        return data


class ExcelFileCreator:
    @staticmethod
    def create_excel_file(data, filename):
        df = pd.DataFrame(data)
        workbook = Workbook()
        sheet = workbook.active
        logging.info("Creating an Excel spreadsheet with Hostnames and IP's")

        for row in dataframe_to_rows(df, index=False, header=True):
            sheet.append(row)

        workbook.save(filename)
        logging.info("Excel file saved and finalized!")

def main():
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s %(name)s %(levelname)s %(message)s',
        filename='./logs.txt',
        filemode='w',
        encoding='utf-8'
    )

    pfsense_automation = PFSenseAutomation(PFSenseAutomation.PFSENSE_HOSTNAME, PFSenseAutomation.PFSENSE_PASSWORD)
    pfsense_automation.initialize_browser()
    pfsense_automation.login()
    pfsense_automation.access_openvpn_menu()
    data = pfsense_automation.get_user_ip_data()

    selenium_data = get_ip_data()
    ip_to_hostname = {d["ip_address"]: d["hostname"] for d in data}

    result = []
    for item in selenium_data:
        ip = item["ip_address"]
        if ip in ip_to_hostname:
            item["hostname"] = ip_to_hostname[ip]
        else:
            item["hostname"] = None
        result.append(item)

    ExcelFileCreator.create_excel_file(result, 'dados.xlsx')


if __name__ == "__main__":
    main()
