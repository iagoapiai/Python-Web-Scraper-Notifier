from selenium import webdriver
from pathlib import Path
import time
import openpyxl
import requests

diretorio_not = r'C:\Users\Iago Piai\Desktop\StatusBolt\Banco\BancoBolts.xlsx'
WEBHOOK_URL = 'WEBHOOK_DO_SEU_CANAL'

bolts_xpath = [
    ('//*[@id="root"]/div[1]/div[2]/div/div[1]/div[3]/div/div[2]/div/a[{}]/div/div[1]/div/span'.format(i),
     '//*[@id="root"]/div[1]/div[2]/div/div[1]/div[3]/div/div[2]/div/a[{}]/div/div[3]/div[2]/label'.format(i))
    for i in range(1, 7)
]

options = webdriver.ChromeOptions()
options.add_argument("--headless")
options.add_argument("--no-sandbox")
options.add_argument("--disable-infobars")
options.add_argument("--disable-logging")
options.add_argument("--disable-extensions")
options.add_argument("--disable-gpu")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-plugins")
options.add_argument("--window-size=1920,1080")
options.add_argument("--blink-settings=imagesEnabled=false")
options.add_argument("--disable-notifications")

driver = webdriver.Chrome(options=options)


def login():
    driver.get("https://retina.ibbx.tech")
    time.sleep(2)
    driver.find_element('xpath', '//*[@id="root"]/div[1]/div[2]/form/div[2]/input').send_keys('EMAIL_PESSOAL')
    driver.find_element('xpath', '//*[@id="root"]/div[1]/div[2]/form/div[3]/input').send_keys('SENHA_PESSOAL')
    driver.find_element('xpath', '//*[@id="root"]/div[1]/div[2]/form/button').click()
    time.sleep(2)


def join_company(facilities):
    driver.get(facilities)
    time.sleep(3)
    return driver.find_element('xpath', '//*[@id="root"]/div[1]/div[2]/div/div[1]/div[1]/div[1]/span/span').text


def send_message(webhook_url, content, thumbnail_url, color):
    data = {
        'embeds': [
            {
                'type': 'rich',
                'color': color,
                'fields': [
                    {'name': 'UNIDADE:', 'value': f'```{facilities_name}```'},
                    {'name': 'BOLT:', 'value': f'```{content}```\n[Acesse no Retina]({facilities})'}
                ],
                'thumbnail': {'url': thumbnail_url, 'height': 25, 'width': 25}
            }
        ]
    }
    requests.post(webhook_url, json=data)


def verify_bolts(qntbolt):
    bolts_data = []
    for i in range(int(qntbolt)):
        bolt_status = driver.find_element('xpath', bolts_xpath[i][0]).text
        bolt_name = driver.find_element('xpath', bolts_xpath[i][1]).text
        bolts_data.append((bolt_name, bolt_status))
    return bolts_data


def compare_bolts(sheet, bolts_data):
    for bolt_name, bolt_status in bolts_data:
        if bolt_name in sheet and bolt_status != sheet[bolt_name].value:
            sheet[bolt_name] = bolt_status
            color = 0x00ff00 if bolt_status == 'ONLINE' else 0xff0000
            send_message(WEBHOOK_URL, f"{bolt_name.upper()} ESTÁ {bolt_status}!", thumbnail_url, color)


def reset_variables():
    global facilities, facilities_name, thumbnail_url
    facilities = facilities_name = thumbnail_url = ''


login()

while True:
    workbook = openpyxl.load_workbook(diretorio_not)
    sheet = workbook.active
    for bolt_set in [('B2', 'B3', 'B4', 'B5'), ('B6', 'B7', 'B8', 'B9', 'B10', 'C10'),
                     ('B11', 'B12', 'B13'), ('B14', 'B15', 'B16', 'B17', 'B18'), ('B19',), ('B20',), ('E2', 'E3')]:
        facilities = 'FACILITIES_URL'
        facilities_name = 'EMPRESA CONFIDENCIAL'
        qntbolt = join_company(facilities)
        bolts_data = verify_bolts(qntbolt)
        compare_bolts(sheet, bolts_data)
        reset_variables()
        time.sleep(1)
    workbook.save(diretorio_not)
