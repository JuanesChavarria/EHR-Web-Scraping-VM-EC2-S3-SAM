from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import datetime
from selenium.webdriver.support.ui import Select
import os
import shutil
from datetime import datetime, date
from dotenv import load_dotenv

# Load environment variables from the .env file
load_dotenv()
username = os.getenv("SITE_USERNAME")
password = os.getenv("SITE_PASSWORD")
login_url = os.getenv("SITE_LOGIN_URL")
appointments_report_url = os.getenv("SITE_APPOINTMENTS_REPORT_URL")
search_appointments_url = os.getenv("SITE_SEARCH_APPOINTMENTS_URL")

class UINavigator:
    def __init__(self):
        self.driver = webdriver.Chrome(ChromeDriverManager(driver_version='114.0.5735.90').install())

    def login(self):
        self.driver.get(login_url)
        username_box = self.driver.find_element(By.NAME,"username")
        username_box.send_keys(username)
        password_box = self.driver.find_element(By.NAME,"password")
        password_box.send_keys(password)
        wait = WebDriverWait(self.driver, 10)  # Create a WebDriverWait instance
        login_button = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/main/section/div/div/div/form/div[3]/button')))
        login_button.click()

    def get_monthly_appointment_report(self, month : int, year : int):
        self.driver.get(appointments_report_url)
        month_box = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_phFolderContent_cntrlMonthlyAppointments_ReportDate_Month"]')))
        month_box.clear()
        month_box.send_keys(month)
        year_box = self.driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_cntrlMonthlyAppointments_ReportDate_Year"]')
        year_box.clear()
        year_box.send_keys(year)
        group_by_button = self.driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_cntrlMonthlyAppointments_ddlGroupBy"]')
        group_by_dropdown = Select(group_by_button)
        group_by_dropdown.select_by_visible_text("Office")
        go_button = self.driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_cntrlMonthlyAppointments_btnGo"]')
        go_button.click()
        WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_phFolderContent_cntrlMonthlyAppointments_divReportFrame"]')))
        export_to_excel_button = self.driver.find_element(By.XPATH , '//*[@id="ctl00_phFolderContent_cntrlMonthlyAppointments_btnExportExcel"]')
        export_to_excel_button.click()
        time.sleep(2)

    def get_monthly_appointment_reports(self):
        chrome_options = webdriver.ChromeOptions()
        prefs = {'download.default_directory' : os.path.abspath("MonthlyAppointmentReports")}
        chrome_options.add_experimental_option('prefs', prefs)
        self.driver = webdriver.Chrome(options=chrome_options)
        self.login()
        today = date.today()
        try:
            shutil.rmtree("MonthlyAppointmentReports")
        except FileNotFoundError:
            pass
        os.mkdir("MonthlyAppointmentReports")
        month_years_to_download = []
        num_months_to_load = 12
        for i in range(num_months_to_load):
            month = ((today.month + i  - 1)) % 12 + 1
            year = today.year + ((today.month + i - 1) // 12)
            month_year = {'month' : month, 'year' : year}
            month_years_to_download.append(month_year)
        while len(month_years_to_download) > 0:
            month_year = month_years_to_download.pop(0)
            self.get_monthly_appointment_report(month_year.get('month'), month_year.get('year'))
            num_documents = len(os.listdir(os.path.abspath("MonthlyAppointmentReports")))
            if(num_documents != num_months_to_load - len(month_years_to_download)):
                month_years_to_download.append(month_year)
            
        for i in range(12):
            month = ((today.month + i  - 1)) % 12 + 1
            year = today.year + ((today.month + i - 1) // 12)
            self.get_monthly_appointment_report(month, year)

    def get_EHRAppointmentID(self, patient_id, appointment_date, appointment_time):
        try:
            self.driver.get(search_appointments_url)
            search_for_type_1_button = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_phFolderContent_apptSearch_ddlSearchFieldID"]'))) 
            search_for_type_1_dropdown = Select(search_for_type_1_button)
            search_for_type_1_dropdown.select_by_visible_text("Patient ID")
            search_box = self.driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_apptSearch_txtSearchText"]')
            search_box.send_keys(patient_id)
            search_option_button = self.driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_apptSearch_ddlSearchOption"]')
            search_option_dropdown = Select(search_option_button)
            search_option_dropdown.select_by_visible_text("Last 30 Days & Upcoming Appointments")
            search_button = self.driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_apptSearch_btnSearch"]')
            search_button.click()
            appointment_date = datetime.strptime(appointment_date, '%A, %m/%d/%Y')
            appointment_time = datetime.strptime(appointment_time, '%I:%M %p').time()
            appointment_datetime = datetime.combine(appointment_date.date(), appointment_time)

            rows = self.driver.find_elements(By.XPATH, '//*[@id="MainFolder"]/div[6]/div[2]/div/table/tbody/tr')
            for i, row in enumerate(rows[1:]):
                row_date = row.find_element(By.XPATH, './td[2]').text
                row_date = datetime.strptime(row_date,'%A\n%m/%d/%Y')
                row_time = row.find_element(By.XPATH, './td[3]').text
                row_time = datetime.strptime(row_time, '%I:%M %p').time()
                row_datetime = datetime.combine(row_date.date(), row_time)

                if(appointment_datetime == row_datetime):
                    appointment_id_XPATH = '//*[@id="AppointmentID' + str(i) + '"]'
                    EHR_AppointmentID = row.find_element(By.XPATH, appointment_id_XPATH).get_attribute("value")
                    return EHR_AppointmentID
        except:
            return ""

    def get_patient_reports(self):
        chrome_options = webdriver.ChromeOptions()
        prefs = {'download.default_directory' : os.path.abspath("PatientReports")}
        chrome_options.add_experimental_option('prefs', prefs)

        self.driver = webdriver.Chrome(ChromeDriverManager(driver_version='114.0.5735.90').install(), chrome_options=chrome_options)

        self.login()
        patient_charts_button = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, "patient-charts_tab")))
        patient_charts_button.click()
        reports_button = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="subMenu"]/ul/li[1]/div/div[1]')))
        reports_button.click()
        excel_export_button = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'patientList(CSVExcelExport)')))
        excel_export_button.click()
        status_button = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_phFolderContent_ucReport_lstStatus"]')))
        status_option_dropdown = Select(status_button)
        status_option_dropdown.select_by_visible_text("Active")
        alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        try:
            shutil.rmtree("PatientReports")
        except FileNotFoundError:
            pass
        os.mkdir("PatientReports")
        letters_to_download = []
        for letter in alphabet:
            letters_to_download.append(letter)
        while(len(letters_to_download) > 0):
            letter = letters_to_download.pop(0)
            from_last_name_button = self.driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_ucReport_lstFromLastName"]')
            from_last_name_dropdown = Select(from_last_name_button)
            from_last_name_dropdown.select_by_visible_text(letter)
            to_last_name_button = self.driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_ucReport_lstToLastName"]')
            to_last_name_dropdown = Select(to_last_name_button)
            to_last_name_dropdown.select_by_visible_text(letter)
            go_button = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_phFolderContent_ucReport_btnGo"]')))
            go_button.click()
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_phFolderContent_ucReport_divReportFrame"]')))
            export_to_excel_button = self.driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_ucReport_btnExportExcel"]')
            export_to_excel_button.click()
            time.sleep(2)
            num_documents = len(os.listdir(os.path.abspath("PatientReports")))
            if(num_documents != len(alphabet) - len(letters_to_download)):
                letters_to_download.append(letter)
