# https://www.instagram.com/leoo_esteves1/
# https://github.com/DotcomPeep

import openpyxl
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
        try:
            self.driver = webdriver.Chrome()
            self.driver.get('Enter your url here')
            self.driver.maximize_window()
            logging.info("Opening Chrome and entering PFSense!")
        except Exception as exc:
            logging.info(exc)

    def login(self):
        try:
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
        except Exception as exc:
            logging.info(exc)

    def access_openvpn_menu(self):
        try:
            access_vpn = self.driver.find_element(By.XPATH, '//*[@id="pf-navbar"]/ul[1]/li[5]/a')
            access_vpn.click()
            logging.info("Accessing the OpenVPN page")
            time.sleep(1)
        except Exception as exc:
            logging.info(exc)

        try:
            openvpn = self.driver.find_element(By.XPATH, '//*[@id="pf-navbar"]/ul[1]/li[5]/ul/li[3]/a')
            openvpn.click()
            logging.info("OpenVPN page found")
            time.sleep(1)
        except Exception as exc:
            logging.info(exc)

    def get_user_ip_data(self):
        try:
            specified = self.driver.find_element(By.XPATH, '//*[@id="2"]/div/ul/li[3]/a')
            specified.click()
            logging.info("Accessing the Client Specific Overrides page!")
            logging.info("Found the list of Common names and IP's. Checking...")
            time.sleep(3)
        except Exception as exc:
            logging.info(exc)

        try:
            users = self.driver.find_elements(By.XPATH, '//*[@id="2"]/div/div/div[2]/table/tbody/tr/td[2]')
            user_list = [user.text for user in users]
            logging.info("Getting list of Common names...")
        except Exception as exc:
            logging.info(exc)

        try:
            ips = self.driver.find_elements(By.XPATH, '//*[@id="2"]/div/div/div[2]/table/tbody/tr/td[3]')
            ips_list = [ip.text for ip in ips]
            logging.info("Getting list of IP's...")
        except Exception as exc:
            logging.info(exc)

        try:
            data = [{"hostname": user, "ip_address": ip} for user, ip in zip(user_list, ips_list)]
            logging.info("List of created Hostnames and IPs")
        except Exception as exc:
            logging.info(exc)

        return data

    def get_user_manager(self):
        username = []
        try:
            system = self.driver.find_element(By.XPATH, '//*[@id="pf-navbar"]/ul[1]/li[1]/a')
            system.click()
            logging.info("Accessing the System page")

            user_manager = self.driver.find_element(By.XPATH, '//*[@id="pf-navbar"]/ul[1]/li[1]/ul/li[11]')
            user_manager.click()
            logging.info("Accessing the user manager page")
            time.sleep(2)

            user_manager_list = self.driver.find_elements(By.XPATH, '//*[@id="2"]/div/form/div/div[2]/div/table/tbody/tr/td[2]')
            
            user_list = [username_list.text for username_list in user_manager_list]

            for users in user_list:
                obj = {}
                obj["users"] = users
                username.append(obj)
            logging.info("Getting users")

            # print(username)
        except Exception as exc:
            logging.error(exc)

        return username


class ExcelFileCreator:
    @staticmethod
    def create_excel_file(data, filename):
        try:
            df = pd.DataFrame(data)
            workbook = Workbook()
            sheet = workbook.active
            logging.info("Creating an Excel spreadsheet with Hostnames and IP's")

            for row in dataframe_to_rows(df, index=False, header=True):
                sheet.append(row)

            workbook.save(filename)
            logging.info("Excel file saved and finalized!")
        except Exception as exc:
            logging.error(exc)

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
    #data = pfsense_automation.get_user_ip_data()
    #pfsense_automation.get_user_manager()

    # Gets user lists from both sources
    ip_data = pfsense_automation.get_user_ip_data()
    user_manager_data = pfsense_automation.get_user_manager()

    # Extracts only usernames from both sources
    ip_usernames = [item["hostname"] for item in ip_data]
    user_manager_usernames = [item["users"] for item in user_manager_data]
    #print(user_manager_usernames)

    # Finds users who are not in both lists
    not_in_ip_data = [username for username in user_manager_usernames if username not in ip_usernames]
    not_in_user_manager = [username for username in ip_usernames if username not in user_manager_usernames]

    users_not_matching = not_in_ip_data + not_in_user_manager

    #print("Users who do not match in both sources:")
    #print(users_not_matching)

    selenium_data = get_ip_data()
    ip_to_hostname = {d["ip_address"]: d["hostname"] for d in ip_data}

    result = []
    for item in selenium_data:
        ip = item["ip_address"]
        if ip in ip_to_hostname:
            item["hostname"] = ip_to_hostname[ip]
        else:
            item["hostname"] = None
        result.append(item)

    ExcelFileCreator.create_excel_file(result, 'dados.xlsx')
    
    try:
        # Open Excel file to add missing hostnames
        excel_filename = 'dados.xlsx'
        workbook = openpyxl.load_workbook(excel_filename)
        sheet = workbook.active

        sheet['E1'] = "IPs livres"

        # Iterate over the result list to find IPs with missing hostnames
        none_hostnames = [item["ip_address"] for item in result if item["hostname"] is None]

        # Fill column E with IPs that have missing hostnames
        for i, ip in enumerate(none_hostnames):
            sheet[f'E{2 + i}'] = ip

        sheet['I1'] = "Usuários sem regras"

        # Fill column i with users who do not have a rule
        for i, username in enumerate(users_not_matching):
            sheet[f'I{2 + i}'] = username

        workbook.save(excel_filename)
    except Exception as exc:
            logging.error(exc)

if __name__ == "__main__":
    main()
