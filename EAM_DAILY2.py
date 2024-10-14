import os
import datetime
import time
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from cryptography.fernet import Fernet
import config

def convert_date(timestamp):
    return datetime.datetime.utcfromtimestamp(timestamp).strftime('%d%b%Y')

def pause(enable):
    if enable:
        input("Press Enter to continue...")

def get_token():
    with open(config.TOKEN_FILE) as f:
        nid = f.readline().strip()
        npasswd = f.readline().strip()

    with open(config.KEY_FILE) as f:
        ref_key = f.read().strip()

    key_to_use = Fernet(ref_key.encode())
    uid = key_to_use.decrypt(nid.encode()).decode()
    token = key_to_use.decrypt(npasswd.encode()).decode()
    return uid, token

def wait_for_element(driver, locator, timeout=config.EXPLICIT_WAIT):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located(locator))

def click_element(driver, locator):
    element = wait_for_element(driver, locator)
    element.click()

def fill_form_field(driver, locator, value):
    field = wait_for_element(driver, locator)
    field.clear()
    field.send_keys(value)

def main(enable_pause):
    os.chdir(config.SCRIPT_DIR)

    yesterday = datetime.date.today() - datetime.timedelta(days=1)
    start_date = yesterday.strftime("%d-%b-%Y")
    end_date = datetime.datetime.today().strftime("%d-%b-%Y")

    print("*** Starting EAM report generation")
    driver = getattr(webdriver, config.WEBDRIVER_TYPE)()
    driver.get(config.EAM_URL)
    pause(enable_pause)

    uid, token = get_token()
    fill_form_field(driver, (By.NAME, 'userid'), uid)
    fill_form_field(driver, (By.NAME, 'password'), token)
    click_element(driver, (By.ID, 'button-1036-btnInnerEl'))

    print("*** Stage 0 Done - Login successful")
    pause(enable_pause)

    wait_for_element(driver, (By.ID, "module_header-1091-module-header-body"))
    click_element(driver, (By.ID, 'button-1044'))  # Menu Work
    click_element(driver, (By.ID, 'menuitem-1063'))  # Item#4
    click_element(driver, (By.ID, 'menuitem-1600-itemEl'))
    pause(enable_pause)

    fill_form_field(driver, (By.XPATH, '//*[@data-ref="inputEl"][@name="organization"]'), config.ORGANIZATION)
    fill_form_field(driver, (By.XPATH, '//*[@data-ref="inputEl"][@name="param6"]'), config.TYPE_FIELD)

    click_element(driver, (By.XPATH, '//*[@data-ref="inputEl"][@name="chkbox1"]'))  # Show parts
    click_element(driver, (By.XPATH, '//*[@data-ref="inputEl"][@name="chkbox2"]'))  # Only Show Completed Work Orders

    fill_form_field(driver, (By.XPATH, '//*[@data-ref="inputEl"][@name="genfromdate"]'), start_date)
    fill_form_field(driver, (By.XPATH, '//*[@data-ref="inputEl"][@name="genthrudate"]'), end_date)

    click_element(driver, (By.XPATH, '//*[@data-qtip="Print Preview"][@style="left: 367px; margin: 0px; top: 8px;"]'))

    print("*** Stage 1 Done - Selecting the Report (Daily Report)")
    pause(enable_pause)

    WebDriverWait(driver, config.EXPLICIT_WAIT).until(EC.number_of_windows_to_be(2))
    driver.switch_to.window(driver.window_handles[1])
    wait_for_element(driver, (By.ID, '_NS_runIn'), config.LONG_WAIT)

    print("*** Stage 2 Done - Preparing the Report")
    click_element(driver, (By.ID, '_NS_runIn'))  # PDF Button
    pause(enable_pause)
    click_element(driver, (By.ID, '_NS_viewInExcel'))  # XLS Button
    click_element(driver, (By.ID, '_NS_viewInspreadsheetML'))  # SXLS Button
    pause(enable_pause)

    time.sleep(5)
    print("*** Stage 3 Done - Report is ready")

    for handle in reversed(driver.window_handles[1:]):
        driver.switch_to.window(handle)
        driver.close()
    driver.switch_to.window(driver.window_handles[0])
    driver.close()

    print("*** Finish")

    src = os.path.join(config.DOWNLOAD_DIR, config.ORIGINAL_FILE_NAME)
    dst = os.path.join(config.DOWNLOAD_DIR, config.RENAMED_FILE_NAME)
    os.rename(src, dst)

    dfile = os.path.join(config.DESTINATION_DIR, config.RENAMED_FILE_NAME)
    if os.path.exists(dfile):
        os.remove(dfile)
        print(f"Deleted existing file: {dfile}")
    shutil.move(dst, config.DESTINATION_DIR)

if __name__ == "__main__":
    enable_pause = False
    main(enable_pause)