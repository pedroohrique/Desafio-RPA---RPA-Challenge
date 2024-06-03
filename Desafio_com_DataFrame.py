import pandas as PD
from selenium import webdriver as Driver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pyautogui as tempo_carregamento


caminho_arquivo_df = (r"C:\Users\lixow\OneDrive\Documentos\SCRIPTS Python - Uteis\Automações - RPA\Automações\challenge.xlsx")
base_to_dataFrame = PD.read_excel(caminho_arquivo_df)

navegador = Driver.Chrome()
navegador.get("https://rpachallenge.com/")
navegador.find_element(By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button').click()
tempo_carregamento.sleep(0.3)

input_firstName = '[ng-reflect-name="labelFirstName"]'
input_lastName = '[ng-reflect-name="labelLastName"]'
input_companyName = '[ng-reflect-name="labelCompanyName"]'
input_role = '[ng-reflect-name="labelRole"]'
input_address = '[ng-reflect-name="labelAddress"]'
input_email = '[ng-reflect-name="labelEmail"]'
input_phoneNumber = '[ng-reflect-name="labelPhone"]'


first_name = base_to_dataFrame["First Name"]
last_name = base_to_dataFrame["Last Name "]
company_name = base_to_dataFrame["Company Name"]
role = base_to_dataFrame["Role in Company"]
address = base_to_dataFrame["Address"]
email = base_to_dataFrame["Email"]
phone_number = base_to_dataFrame["Phone Number"]

for indice in range(1, len(base_to_dataFrame) + 1):

    navegador.find_element(By.CSS_SELECTOR, input_firstName).send_keys(first_name[indice])
    tempo_carregamento.sleep(0.3)
    navegador.find_element(By.CSS_SELECTOR, input_lastName).send_keys(last_name[indice])
    tempo_carregamento.sleep(0.3)
    navegador.find_element(By.CSS_SELECTOR, input_companyName).send_keys(company_name[indice])
    tempo_carregamento.sleep(0.3)
    navegador.find_element(By.CSS_SELECTOR, input_role).send_keys(role[indice])
    tempo_carregamento.sleep(0.3)
    navegador.find_element(By.CSS_SELECTOR, input_address).send_keys(address[indice])
    tempo_carregamento.sleep(0.3)
    navegador.find_element(By.CSS_SELECTOR, input_email).send_keys(email[indice])
    tempo_carregamento.sleep(0.3)
    navegador.find_element(By.CSS_SELECTOR, input_phoneNumber).send_keys(str(phone_number[indice]))
    tempo_carregamento.sleep(0.3)
    navegador.find_element(By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input').click()
    tempo_carregamento.sleep(0.5)
tempo_carregamento.sleep(3)

