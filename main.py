from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from data import WebData
from excel_function import Excel_Functions
from locators import Locators



# Constants for column indices
USERNAME_COL = 6
PASSWORD_COL = 7
RESULT_COL = 8


def login(driver, username, password):
    try:
        wait = WebDriverWait(driver, 20)
        username_field = wait.until(EC.presence_of_element_located((By.NAME, Locators.username)))
        username_field.send_keys(username)

        password_field = wait.until(EC.presence_of_element_located((By.NAME, Locators.password)))
        password_field.send_keys(password)

        submit_button = wait.until(EC.presence_of_element_located((By.XPATH, Locators.submit_button)))
        submit_button.click()
        return driver.current_url
    except Exception as e:
        print("Login failed:", e)
        return None


def logout(driver):
    try:
        wait = WebDriverWait(driver, 20)
        dropdown_button = wait.until(EC.presence_of_element_located((By.XPATH, Locators.dropdown_button)))
        dropdown_button.click()

        logout_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Logout')]")))
        logout_button.click()
    except Exception as e:
        print("Logout failed:", e)


def main():
    data = WebData()
    Excel_Data = Excel_Functions(data.excel_file, data.sheet_number)
    driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()))

    with driver:
        driver.get(data.url)
        rows = Excel_Data.row_count()

        for row in range(2, rows + 1):
            username = Excel_Data.read_data(row, column_number=USERNAME_COL)
            password = Excel_Data.read_data(row, column_number=PASSWORD_COL)

            current_url = login(driver, username, password)

            if data.dashboard_url in current_url:
                Excel_Data.write_data(row, column_number=RESULT_COL, data="TEST PASSED")
                logout(driver)
            else:
                Excel_Data.write_data(row, column_number=RESULT_COL, data="TEST FAILED")
                driver.refresh()
    driver.quit()


if __name__ == "__main__":
    main()