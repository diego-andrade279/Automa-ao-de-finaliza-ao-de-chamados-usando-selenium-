from datetime import datetime
from bs4 import BeautifulSoup
import pandas as pd
import ctypes
import os
import pyautogui
import keyboard
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    NoSuchElementException,
)
from time import sleep
import time

user_name = os.getenv("USERNAME")
dia_semana = 10
ctypes.windll.kernel32.SetThreadExecutionState(0x80000002)


def chamado_Offbording():
    ini = time.time()
    tabela = pd.read_excel(rf"C:\Users\{user_name}\Desktop\Bot_Sara\atualizacao_estoque.xlsx")
    global dia_semana

    option = Options()
    s = Service(ChromeDriverManager().install())
    option.add_argument("--start-maximized")
    option.add_argument("--force-device-scale-factor=0.6")
    option.add_argument(r"--profile-directory=Default")
    option.add_argument(
        r"--user-data-dir=C:\Users\{0}\AppData\Local\Google\Chrome\User Data".format(
            user_name
        )
    )

    for linha in tabela.index:
        print(f"N°{linha}:\n {tabela.loc[linha,['Nome','Número de série','MODELO','GARANTIA','Plug-ins - Histórico da movimentação - Nº de chamado de entrada']]}:\n")
        driver = webdriver.Chrome(service=s, options=option)
        driver.get(
            f"https://ifood.atlassian.net/issues/{tabela.loc[linha,'Plug-ins - Histórico da movimentação - Nº de chamado de entrada']}"
        )
        sleep(dia_semana)

        sleep(1)
        # Campo Responder cliente
        camp_Descricao = driver.find_element(
            By.XPATH,
            "//span[contains(@class, 'css-178ag6o') and text()='Responder ao cliente']",
        )
        camp_Descricao.click()
        sleep(5)

        camp_Descricao = driver.find_element(By.ID, "ak-editor-textarea")
        ActionChains(driver).scroll_to_element(camp_Descricao).perform()
        camp_Descricao.send_keys(
            "Equipamento recebido no estoque.\n",
            f" Modelo: {tabela.loc[linha,'MODELO']}\n",
            f" Patrimonio: {tabela.loc[linha,'PATRIMONIO']}\n ",
            f" Service Tag: {tabela.loc[linha,'Número de série']}\n",
            " Com carregador \n",
        )
        sleep(2)

        # Click no botao guardar
        botao_Aguardar = driver.find_element(
            By.CSS_SELECTOR, '[data-testid="comment-save-button"]'
        )
        ActionChains(driver).scroll_to_element(botao_Aguardar).perform()
        botao_Aguardar.click()
        sleep(1)

        # Validaçao se o chamado é devoluçao troca ou devouçao offbording
        validacao = tabela.loc[
            linha, "Plug-ins - Histórico da movimentação - Entrada"
        ].upper()
        print(validacao)
        if validacao == "OFFBOARDING":
            # Classificaçao De Filas 1
            botao_Classi_Fila = driver.find_element(
                By.CSS_SELECTOR,
                '[data-component-selector="jira-issue-field-cascading-select-read-view-container"]',
            )
            ActionChains(driver).scroll_to_element(botao_Classi_Fila).perform()
            sleep(1)
            botao_Classi_Fila.click()
            sleep(2)
            bt_classificacao = driver.find_element(
                By.CSS_SELECTOR, 'input[id^="react-select"]'
            )
            bt_classificacao = bt_classificacao.get_attribute("id")
            bt_classificacao = driver.find_element(By.ID, bt_classificacao)
            bt_classificacao.send_keys(str("Devolução"))
            sleep(1.5)
            bt_classificacao.send_keys(Keys.ENTER)

            # Campo 2 classificaçao de fila 2
            bt_classificacao = driver.find_element(
                By.CSS_SELECTOR, 'input[id^="react-select"]'
            )
            bt_classificacao = bt_classificacao.get_attribute("id")
            # quebrando string e reconstrindo para click em segundo campo
            valor = bt_classificacao[13:-6]
            print(valor)
            bt_classificacao = f"react-select-{int(valor) + 1}-input"
            print(bt_classificacao)
            bt_classificacao = driver.find_element(By.ID, bt_classificacao)
            bt_classificacao.send_keys(str("Devolução Offboarding"))
            sleep(1.5)
            bt_classificacao.send_keys(Keys.ENTER)

        elif validacao == "DEVOLUÇÃO":
            # Classificaçao De Filas 1
            botao_Classi_Fila = driver.find_element(By.CSS_SELECTOR,'[data-component-selector="jira-issue-field-cascading-select-read-view-container"]')
            ActionChains(driver).scroll_to_element(botao_Classi_Fila).perform()
            sleep(1)
            botao_Classi_Fila.click()
            sleep(2)
            bt_classificacao = driver.find_element(
                By.CSS_SELECTOR, 'input[id^="react-select"]'
            )
            bt_classificacao = bt_classificacao.get_attribute("id")
            bt_classificacao = driver.find_element(By.ID, bt_classificacao)
            bt_classificacao.send_keys(str("Devolução"))
            sleep(1.5)
            bt_classificacao.send_keys(Keys.ENTER)
            sleep(1)

            # Campo 2 classificaçao de fila 2
            bt_classificacao = driver.find_element(
                By.CSS_SELECTOR, 'input[id^="react-select"]'
            )
            bt_classificacao = bt_classificacao.get_attribute("id")
            # quebrando string e reconstrindo para click em segundo campo
            valor = bt_classificacao[13:-6]
            bt_classificacao = f"react-select-{int(valor) + 1}-input"
            bt_classificacao = driver.find_element(By.ID, bt_classificacao)
            bt_classificacao.send_keys(str("Devolução de Troca"))
            sleep(1.5)
            bt_classificacao.send_keys(Keys.ENTER)
            sleep(1)

        # Campo em Transito para em Andamento
        if validacao == "OFFBOARDING":
            sleep(2)
            botao_Em_Transito = driver.find_element(
                By.CSS_SELECTOR,
                '[data-testid="issue-field-status.ui.status-view.status-button.status-button"]',
            )
            botao_Em_Transito.click()
            sleep(5)

            # Campo Em Andamento
            try:
                botao_pen_retirada = driver.find_element(
                    By.CSS_SELECTOR, "div.css-363ehx-option"
                )
                id = botao_pen_retirada.get_attribute("id")
                id = id[:-1]
                id = f"{id}0"
                botao_pen_retirada = driver.find_element(By.ID, id)
                botao_pen_retirada.click()
                sleep(2)

            except NoSuchElementException:
                botao_pen_retirada = driver.find_element(
                    By.CSS_SELECTOR, "div.css-rz73v4-option"
                )
                id = botao_pen_retirada.get_attribute("id")
                id = id[:-1]
                id = f"{id}0"
                botao_pen_retirada = driver.find_element(By.ID, id)
                botao_pen_retirada.click()
                sleep(2)

            ##############################
            # Campo em Andamento finalizar chamado
            botao_Em_Andamento = driver.find_element(
                By.CSS_SELECTOR,
                '[data-testid="issue-field-status.ui.status-view.status-button.status-button"]',
            )
            botao_Em_Andamento.click()
            sleep(5)

            # Campo Em Andamento
            try:
                botao_pen_retirada = driver.find_element(
                    By.CSS_SELECTOR, "div.css-363ehx-option"
                )
                id = botao_pen_retirada.get_attribute("id")
                id = id[:-1]
                id = f"{id}9"
                botao_pen_retirada = driver.find_element(By.ID, id)
                botao_pen_retirada.click()
                sleep(2)

            except NoSuchElementException:
                botao_pen_retirada = driver.find_element(
                    By.CSS_SELECTOR, "div.css-rz73v4-option"
                )
                id = botao_pen_retirada.get_attribute("id")
                id = id[:-1]
                id = f"{id}9"
                botao_pen_retirada = driver.find_element(By.ID, id)
                botao_pen_retirada.click()
                sleep(2)

        elif validacao == "DEVOLUÇÃO":

            ##############################
            # Campo em Andamento finalizar chamado
            botao_Em_Andamento = driver.find_element(
                By.CSS_SELECTOR,
                '[data-testid="issue-field-status.ui.status-view.status-button.status-button"]',
            )
            botao_Em_Andamento.click()
            sleep(5)

            # Campo Em Andamento
            try:
                botao_pen_retirada = driver.find_element(
                    By.CSS_SELECTOR, "div.css-363ehx-option"
                )
                id = botao_pen_retirada.get_attribute("id")
                id = id[:-1]
                id = f"{id}9"
                botao_pen_retirada = driver.find_element(By.ID, id)
                botao_pen_retirada.click()
                sleep(2)

            except NoSuchElementException:
                botao_pen_retirada = driver.find_element(
                    By.CSS_SELECTOR, "div.css-rz73v4-option"
                )
                id = botao_pen_retirada.get_attribute("id")
                id = id[:-1]
                id = f"{id}9"
                botao_pen_retirada = driver.find_element(By.ID, id)
                botao_pen_retirada.click()
                sleep(2)

            sleep(5)

        # validaçao Defeito do equipamento
        camp_Defeito_equip = driver.find_element(
            By.CSS_SELECTOR, '[name="customfield_15907"]'
        )
        camp_Defeito_equip.click()
        sleep(5)

        estado_equip = (
            tabela.loc[
                linha,
                "Plug-ins - Histórico da movimentação - Entrada",
            ]
            .upper()
            .strip()
        )
        if estado_equip == "OFFBOARDING":
            camp_Defeito_equip = driver.find_element(
                By.CSS_SELECTOR, "#customfield_15907 > option:nth-child(4)"
            )
            camp_Defeito_equip.click()
            sleep(5)
        else:
            camp_Defeito_equip = driver.find_element(
                By.CSS_SELECTOR, "#customfield_15907 > option:nth-child(3)"
            )
            camp_Defeito_equip.click()
            sleep(5)

        # Campo troca
        estado_equip = (
            tabela.loc[linha, "Plug-ins - Histórico da movimentação - Entrada"]
            .upper()
            .strip()
        )
        print(estado_equip)
        if estado_equip == "OFFBOARDING":
            camp_troca = driver.find_element(
                By.CSS_SELECTOR, "#customfield_18105 > option:nth-child(3)"
            )
            camp_troca.click()
            sleep(5)
        elif estado_equip == "DEVOLUÇÃO":
            camp_troca = driver.find_element(
                By.CSS_SELECTOR, "#customfield_18105 > option:nth-child(2)"
            )
            camp_troca.click()
            sleep(5)

        # Patrimonio encerrando chamado

        pt = tabela.loc[linha, "PATRIMONIO"]
        input_patrimonio = driver.find_element(
            By.CSS_SELECTOR, '[name="customfield_14879"]'
        )
        input_patrimonio.send_keys(str(pt))
        sleep(1)

        # campo service tag 2
        st = tabela.loc[linha, "Número de série"]
        camp_st = driver.find_element(By.CSS_SELECTOR, '[id="customfield_15185"]')
        camp_st.click()
        camp_st.send_keys(str(st))
        sleep(2)

        # Botao Resolvido
        botao_resolvido = driver.find_element(
            By.CSS_SELECTOR, '[id="issue-workflow-transition-submit"]'
        )
        botao_resolvido.click()
        sleep(10)

        temp = []
        fim = time.time()
        temp = (fim - ini) / 60
        print(f"Tempo de gasto: {temp}")
        driver.quit()


if __name__ == "__main__":
    print("=============================================")
    print("= Pressione um numero para selecionar um dia=")
    print("= 1 - Segunda-Feira:                        =")
    print("= 2 - Terca-Feira:                          =")
    print("= 3 - Quata-Feira:                          =")
    print("= 4 - Quinta-Feira:                         =")
    print("= 5 - Sexta-Feira:                          =")
    print("=============================================")
    print(f"Pressione: ")
    opcao = keyboard.read_key()
    print(opcao)
    match opcao:
        case "1":
            dia_semana = 10
            chamado_Offbording()
        case "2":
            dia_semana = 20
            chamado_Offbording()
        case "3":
            dia_semana = 20
            chamado_Offbording()
        case "4":
            dia_semana = 20
            chamado_Offbording()
        case "5":
            dia_semana = 10
            chamado_Offbording()
        case _:
            print("Opçao invalida!!!")

ctypes.windll.kernel32.SetThreadExecutionState(0x00000002)
