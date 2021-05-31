from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook

import time
executable_path ='./chromedriver'
driver = webdriver.Chrome(executable_path)

#Navigate to page
driver.get('https://cba.knec.ac.ke/')

username = driver.find_element_by_id("UserName")
password = driver.find_element_by_id("Password")

username.send_keys("20408005")
password.send_keys("shortman")

driver.find_element_by_xpath('//button[@type="submit"]').click()

try:
    element = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH,"/html/body/div[2]/div[3]/div[3]/a"))
    )
    element.click()

    element = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH,"/html/body/div[7]/div[2]/div/nav/div[1]/a/i"))
    )
    element.click()

    element = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH,"/html/body/div[7]/div[1]/div/div/div[3]/ul/li[4]/a/i"))
    )
    element.click()

    table = driver.find_elements_by_xpath("/html/body/div[7]/div[3]/div[3]/div[1]/table/tbody/tr")
    
    for row in table:
        butons = driver.find_elements_by_xpath("/html/body/div[7]/div[3]/div[3]/div[1]/table/tbody/tr/td/a")
        for buton in butons:
            buton.click()
            element = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH,"/html/body/div[7]/div[3]/div[3]/div/div/div[1]/label/select/option[4]"))
            )
            element.click()
            children = driver.find_elements_by_xpath("/html/body/div[7]/div[3]/div[3]/div/div/table/tbody/tr/td[2]")
            for child in children:
                print(child.text)
            driver.back()

    
except Exception as e:
    driver.quit()
    print(e)
    