from selenium import webdriver as Driver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from openpyxl import load_workbook
import pyautogui as Automacoes


caminho_arquivo = (r"C:\Users\lixow\OneDrive\Documentos\SCRIPTS Python - Uteis\Automações - RPA\Automações\challenge.xlsx")
carrega_arquivo = load_workbook(filename=caminho_arquivo)
aba_selecionada = carrega_arquivo['Sheet1']

navegador = Driver.Chrome()
navegador.get('https://rpachallenge.com')
navegador.find_element(By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button').click()
Automacoes.sleep(1)

#Instanciando as variáveis que receberão o valor de cada label da tag "CSS_SELECTOR"
labelEmail_selector = '[ng-reflect-name="labelEmail"]'
labelname_selector = '[ng-reflect-name="labelFirstName"]'
labelCompany_selector = '[ng-reflect-name="labelCompanyName"]'
labelRole_selector = '[ng-reflect-name="labelRole"]'
labelAddress_selector = '[ng-reflect-name="labelAddress"]'
labelPhone_selector = '[ng-reflect-name="labelPhone"]'
labelLastName_selector = '[ng-reflect-name="labelLastName"]'

for linha in range(2, len(aba_selecionada['A'])):

    colunaA = 'A' + str(linha)
    colunaB = 'B' + str(linha)
    colunaC = 'C' + str(linha)
    colunaD = 'D' + str(linha)
    colunaE = 'E' + str(linha)
    colunaF = 'F' + str(linha)
    colunaG = 'G' + str(linha)

    name = aba_selecionada[colunaA].value
    last_name = aba_selecionada[colunaB].value
    company_name = aba_selecionada[colunaC].value
    role = aba_selecionada[colunaD].value
    address = aba_selecionada[colunaE].value
    email = aba_selecionada[colunaF].value
    phone_number = aba_selecionada[colunaG].value

    if name != None:
        navegador.find_element(By.CSS_SELECTOR, labelEmail_selector).send_keys(email)
        Automacoes.sleep(0.5)
        navegador.find_element(By.CSS_SELECTOR, labelname_selector).send_keys(name)
        Automacoes.sleep(0.5)
        navegador.find_element(By.CSS_SELECTOR, labelCompany_selector).send_keys(company_name)
        Automacoes.sleep(0.5)
        navegador.find_element(By.CSS_SELECTOR, labelRole_selector).send_keys(role)
        Automacoes.sleep(0.5)
        navegador.find_element(By.CSS_SELECTOR, labelAddress_selector).send_keys(address)
        Automacoes.sleep(0.5)
        navegador.find_element(By.CSS_SELECTOR, labelPhone_selector).send_keys(phone_number)
        Automacoes.sleep(0.5)
        navegador.find_element(By.CSS_SELECTOR, labelLastName_selector). send_keys(last_name)
        Automacoes.sleep(0.5)

        navegador.find_element(By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input').click()
        Automacoes.sleep(1.5)





