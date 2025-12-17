from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from time import sleep
from dotenv import load_dotenv
import os


def setup():
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")
    # options.add_argument("--headless=new")

    return webdriver.Chrome(options=options)


def teardown(driver):
    driver.quit()


def collector():

    load_dotenv()

    n_access = os.getenv("n_access")
    user = os.getenv("user")
    passwd = os.getenv("passwd")
    
    auth_url = "https://goldenlog.gservice.com.br/"
    target_url = "https://goldenlog.gservice.com.br/checklist/filtro-checklist-cliente"
    
    driver = setup()

    try:
        driver.get(auth_url)
        sleep(3)
        driver.find_element(By.ID, "codigo-login").send_keys(n_access)
        driver.find_element(By.ID, "login").send_keys(user)
        driver.find_element(By.ID, "password").send_keys(passwd, Keys.RETURN)

        sleep(6)

        # After login, GET in page wheres the Data is located
        driver.get(target_url)
        sleep(1.5)

        now = datetime.now()

        minus_now = now - relativedelta(days=10)
        minus_now = minus_now.strftime("%d/%m/%Y")
        driver.find_element(By.ID, "date").send_keys(minus_now)
     
        plus_now = now + relativedelta(days=18)
        plus_now = plus_now.strftime("%d/%m/%Y")
        driver.find_element(By.ID, "dataFim").send_keys(plus_now)

        sleep(1.5)

        driver.find_element(
            By.XPATH, '//*[@id="root"]/div[1]/div/div[3]/form/div[6]/button[2]'
        ).click()

        sleep(5)

        complete_table = []

        next_btn = driver.find_element(
            By.XPATH,
            "/html/body/div/div[1]/div/div[3]/div[5]/div[2]/div/nav/ul/li[7]/button",
        )

        while next_btn.is_enabled():

            try:
                sleep(3)

                rows = driver.find_elements(
                    By.XPATH,
                    '//*[@id="root"]/div[1]/div/div[3]/div[3]/table/tbody/tr',
                )

                for row in rows:
                    columns = [
                        col.text.strip() for col in row.find_elements(By.TAG_NAME, "td")
                    ]

                    complete_table.append(columns)

                next_btn.click()

            except Exception as click_exception:
                print(f"Erro ao clicar no botão de próxima página: {click_exception}")
                driver.execute_script("arguments[0].click();", next_btn)

        print("=>   Tabela extraída com sucesso!\n=================================")

        return complete_table

    except Exception as e:
        print(f"ERRO durante setup!!!\n\n{e}")
        teardown(driver)

        # Reboot
        setup()