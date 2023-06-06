from selenium import webdriver
from pathlib import Path
import time
import openpyxl
import requests

diretorio_not = r'C:\Users\Iago Piai\Desktop\StatusBolt\Banco\BancoBolts.xlsx'

WEBHOOK_URL = 'WEBHOOK_DO_SEU_CANAL'

facilitie_bolts = '//*[@id="root"]/div[1]/div[2]/div/div[1]/div[1]/div[1]/span/span'

bolt1_xpath = '//*[@id="root"]/div[1]/div[2]/div/div[1]/div[3]/div/div[2]/div/a[1]/div/div[1]/div/span'
bolt1_name_xpath = '//*[@id="root"]/div[1]/div[2]/div/div[1]/div[3]/div/div[2]/div/a[1]/div/div[3]/div[2]/label'

bolt2_xpath = '//*[@id="root"]/div[1]/div[2]/div/div[1]/div[3]/div/div[2]/div/a[2]/div/div[1]/div/span'
bolt2_name_xpath = '//*[@id="root"]/div[1]/div[2]/div/div[1]/div[3]/div/div[2]/div/a[2]/div/div[3]/div[2]/label'

bolt3_xpath = '//*[@id="root"]/div[1]/div[2]/div/div[1]/div[3]/div/div[2]/div/a[3]/div/div[1]/div/span'
bolt3_name_xpath = '//*[@id="root"]/div[1]/div[2]/div/div[1]/div[3]/div/div[2]/div/a[3]/div/div[3]/div[2]/label'

bolt4_xpath = '//*[@id="root"]/div[1]/div[2]/div/div[1]/div[3]/div/div[2]/div/a[4]/div/div[1]/div/span'
bolt4_name_xpath = '//*[@id="root"]/div[1]/div[2]/div/div[1]/div[3]/div/div[2]/div/a[4]/div/div[3]/div[2]/label'

bolt5_xpath = '//*[@id="root"]/div[1]/div[2]/div/div[1]/div[3]/div/div[2]/div/a[5]/div/div[1]/div/span'
bolt5_name_xpath = '//*[@id="root"]/div[1]/div[2]/div/div[1]/div[3]/div/div[2]/div/a[5]/div/div[3]/div[2]/label'

bolt6_xpath = '//*[@id="root"]/div[1]/div[2]/div/div[1]/div[3]/div/div[2]/div/a[6]/div/div[1]/div/span'
bolt6_name_xpath = '//*[@id="root"]/div[1]/div[2]/div/div[1]/div[3]/div/div[2]/div/a[6]/div/div[3]/div[2]/label'

# AS VARIAVEIS QUE POSSUIAM OS DADOS DAS EMPRESAS FORAM RETIRADOS POIS NÃO TENHO PERMISSÃO PARA MOSTRAR!

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
    driver.find_element('xpath', '//*[@id="root"]/div[1]/div[2]/form/div[3]/input').send_keys('SENHA PESSOAL')
    driver.find_element('xpath', '//*[@id="root"]/div[1]/div[2]/form/button').click()
    time.sleep(2)

def join_company():
    driver.get(facilities)
    time.sleep(3)
    qntbolt = driver.find_element('xpath', facilitie_bolts).text
    time.sleep(2)
    return qntbolt

def send_message(webhook_url, content, thumbnail_url):
    data = {
        'embeds': [
            {
                'type': 'rich',
                'color': 0x00ff00,
                'author': {
                    'name': '',
                },
                'fields': [
                    {
                        'name': f' UNIDADE:',
                        'value': f'```{facilities_name}```'
                    },
                    {
                        'name': f' BOLT:',
                        'value': f'```{content}```\n[Acesse no Retina]({facilities})'
                    }
                ],
                'thumbnail': {
                    'url': thumbnail_url,
                    'height': 25,  # Defina a altura da miniatura conforme necessário
                    'width': 25,  # Defina a largura da miniatura conforme necessário
                }
            }
        ]
    }



    response = requests.post(webhook_url, json=data)

    if response.status_code == 204:
        print('|--------------------------- Mensagem enviada com sucesso! ---------------------------|')

    else:
        print('|------------------------------ Erro ao enviar mensagem! ------------------------------|')




def send_message_offline(webhook_url, content, thumbnail_url):
    data = {
        'embeds': [
            {
                'type': 'rich',
                'color': 0xff0000,
                'author': {
                    'name': '',
                },
                'fields': [
                    {
                        'name': f' UNIDADE:',
                        'value': f'```{facilities_name}```'
                    },
                    {
                        'name': f' BOLT:',
                        'value': f'```{content}```\n[Acesse no Retina]({facilities})'
                    }
                ],
                'thumbnail': {
                    'url': thumbnail_url,
                    'height': 25,
                    'width': 25,
                }
            }
        ]
    }

    response = requests.post(webhook_url, json=data)
    if response.status_code == 204:
        print('|--------------------------- Mensagem enviada com sucesso! ---------------------------|')

    else:
        print('|------------------------------ Erro ao enviar mensagem! ------------------------------|')


def verify_bolts():
    if qntbolt =='1':
        bolt1_status = driver.find_element('xpath', bolt1_xpath).text
        bolt1_name = driver.find_element('xpath', bolt1_name_xpath).text
        return bolt1_name, bolt1_status
     
    elif qntbolt =='2':
        bolt1_status = driver.find_element('xpath', bolt1_xpath).text
        bolt1_name = driver.find_element('xpath', bolt1_name_xpath).text

        bolt2_status = driver.find_element('xpath', bolt2_xpath).text
        bolt2_name = driver.find_element('xpath', bolt2_name_xpath).text
        return bolt1_name, bolt1_status ,bolt2_name, bolt2_status

    elif qntbolt =='3':
        bolt1_status = driver.find_element('xpath', bolt1_xpath).text
        bolt1_name = driver.find_element('xpath', bolt1_name_xpath).text

        bolt2_status = driver.find_element('xpath', bolt2_xpath).text
        bolt2_name = driver.find_element('xpath', bolt2_name_xpath).text

        bolt3_status = driver.find_element('xpath', bolt3_xpath).text
        bolt3_name = driver.find_element('xpath', bolt3_name_xpath).text
        return bolt1_name, bolt1_status ,bolt2_name, bolt2_status, bolt3_name, bolt3_status

    elif qntbolt =='4':
        bolt1_status = driver.find_element('xpath', bolt1_xpath).text
        bolt1_name = driver.find_element('xpath', bolt1_name_xpath).text

        bolt2_status = driver.find_element('xpath', bolt2_xpath).text
        bolt2_name = driver.find_element('xpath', bolt2_name_xpath).text

        bolt3_status = driver.find_element('xpath', bolt3_xpath).text
        bolt3_name = driver.find_element('xpath', bolt3_name_xpath).text

        bolt4_status = driver.find_element('xpath', bolt4_xpath).text
        bolt4_name = driver.find_element('xpath', bolt4_name_xpath).text
        return bolt1_name, bolt1_status ,bolt2_name, bolt2_status, bolt3_name, bolt3_status, bolt4_name, bolt4_status

    elif qntbolt =='5':
        bolt1_status = driver.find_element('xpath', bolt1_xpath).text
        bolt1_name = driver.find_element('xpath', bolt1_name_xpath).text

        bolt2_status = driver.find_element('xpath', bolt2_xpath).text
        bolt2_name = driver.find_element('xpath', bolt2_name_xpath).text

        bolt3_status = driver.find_element('xpath', bolt3_xpath).text
        bolt3_name = driver.find_element('xpath', bolt3_name_xpath).text

        bolt4_status = driver.find_element('xpath', bolt4_xpath).text
        bolt4_name = driver.find_element('xpath', bolt4_name_xpath).text

        bolt5_status = driver.find_element('xpath', bolt5_xpath).text
        bolt5_name = driver.find_element('xpath', bolt5_name_xpath).text
        return bolt1_name, bolt1_status ,bolt2_name, bolt2_status, bolt3_name, bolt3_status, bolt4_name, bolt4_status, bolt5_name, bolt5_status
    
    elif qntbolt =='6':
        bolt1_status = driver.find_element('xpath', bolt1_xpath).text
        bolt1_name = driver.find_element('xpath', bolt1_name_xpath).text

        bolt2_status = driver.find_element('xpath', bolt2_xpath).text
        bolt2_name = driver.find_element('xpath', bolt2_name_xpath).text

        bolt3_status = driver.find_element('xpath', bolt3_xpath).text
        bolt3_name = driver.find_element('xpath', bolt3_name_xpath).text

        bolt4_status = driver.find_element('xpath', bolt4_xpath).text
        bolt4_name = driver.find_element('xpath', bolt4_name_xpath).text

        bolt5_status = driver.find_element('xpath', bolt5_xpath).text
        bolt5_name = driver.find_element('xpath', bolt5_name_xpath).text

        bolt6_status = driver.find_element('xpath', bolt6_xpath).text
        bolt6_name = driver.find_element('xpath', bolt6_name_xpath).text
        return bolt1_name, bolt1_status ,bolt2_name, bolt2_status, bolt3_name, bolt3_status, bolt4_name, bolt4_status, bolt5_name, bolt5_status, bolt6_name, bolt6_status

def compare_bolts():
    if bolt1 is not None and bolt1_status != sheet[bolt1].value:
        sheet[bolt1] = bolt1_status
        if bolt1_status == 'ONLINE':
            send_message(WEBHOOK_URL, f"{bolt1_name.upper()} ESTÁ {bolt1_status}!", thumbnail_url)
        else:
            send_message_offline(WEBHOOK_URL, f"{bolt1_name.upper()} ESTÁ {bolt1_status}!", thumbnail_url)

    if bolt2 is not None and bolt2_status != sheet[bolt2].value:
        sheet[bolt2] = bolt2_status
        if bolt2_status == 'ONLINE':
            send_message(WEBHOOK_URL, f"{bolt2_name.upper()} ESTÁ {bolt2_status}!", thumbnail_url)
        else:
            send_message_offline(WEBHOOK_URL, f"{bolt2_name.upper()} ESTÁ {bolt2_status}!", thumbnail_url)

    if bolt3 is not None and bolt3_status != sheet[bolt3].value:
        sheet[bolt3] = bolt3_status
        if bolt3_status == 'ONLINE':
            send_message(WEBHOOK_URL, f"{bolt3_name.upper()} ESTÁ {bolt3_status}!", thumbnail_url)
        else:
            send_message_offline(WEBHOOK_URL, f"{bolt3_name.upper()} ESTÁ {bolt3_status}!", thumbnail_url)

    if bolt4 is not None and bolt4_status != sheet[bolt4].value:
        sheet[bolt4] = bolt4_status
        if bolt4_status == 'ONLINE':
            send_message(WEBHOOK_URL, f"{bolt4_name.upper()} ESTÁ {bolt4_status}!", thumbnail_url)
        else:
            send_message_offline(WEBHOOK_URL, f"{bolt4_name.upper()} ESTÁ {bolt4_status}!", thumbnail_url)
        
    if bolt5 is not None and bolt5_status != sheet[bolt5].value:
        sheet[bolt5] = bolt5_status
        if bolt5_status == 'ONLINE':
            send_message(WEBHOOK_URL, f"{bolt5_name} ESTÁ {bolt5_status}!", thumbnail_url)
        else:
            send_message_offline(WEBHOOK_URL, f"{bolt5_name} ESTÁ {bolt5_status}!", thumbnail_url)

    if bolt6 is not None and bolt6_status != sheet[bolt6].value:
        sheet[bolt6] = bolt6_status
        if bolt6_status == 'ONLINE':
            send_message(WEBHOOK_URL, f"{bolt6_name} ESTÁ {bolt6_status}!", thumbnail_url)
        else:
            send_message_offline(WEBHOOK_URL, f"{bolt6_name} ESTÁ {bolt6_status}!", thumbnail_url)

    workbook.save(diretorio_not)

def reset_variables():
    global facilities, thumbnail_url, bolt5_name, bolt5_status, bolt6_status, bolt6_name, bolt6,bolt1, bolt2, bolt3, bolt4, bolt5, qntbolt, bolt1_name, bolt1_status ,bolt2_name, bolt2_status, bolt3_name, bolt3_status, bolt4_name, bolt4_status, facilities_name
    bolt1 = ''
    bolt2 = ''
    bolt3 = ''
    bolt4 = ''
    bolt5 = ''
    bolt6 = ''
    thumbnail_url = ''
    qntbolt = ''
    facilities_name = ''
    bolt1_name = ''
    bolt1_status = ''
    bolt2_name = ''
    bolt2_status = ''
    bolt3_name = ''
    bolt3_status = ''
    bolt4_name = ''
    bolt4_status = ''
    bolt5_name = ''
    bolt5_status = ''
    bolt6_name = ''
    bolt6_status = ''
    facilities = ''

login()

print('---------------------------------------------------------------------------------------')
print('------------------------------- Primeiro loop iníciado! -------------------------------')
print('---------------------------------------------------------------------------------------')

while True:
    workbook = openpyxl.load_workbook(diretorio_not)
    sheet = workbook.active
    facilities= variavel_da_empresa
    facilities_name = 'EMPRESA CONFIDENCIAL'
    qntbolt = join_company()
    bolt1 = 'B2'
    bolt2 = 'B3'
    bolt3 = 'B4'
    bolt4 = 'B5'
    bolt5 = None
    bolt6 = None
    verify_bolts()
    bolt1_name, bolt1_status ,bolt2_name, bolt2_status, bolt3_name, bolt3_status, bolt4_name, bolt4_status = verify_bolts()
    compare_bolts()
    reset_variables()
    time.sleep(1)
    
    workbook = openpyxl.load_workbook(diretorio_not)
    sheet = workbook.active
    facilities= variavel_da_empresa
    facilities_name = 'EMPRESA CONFIDENCIAL'
    qntbolt = join_company()
    bolt1 = 'B6'
    bolt2 = 'B7'
    bolt3 = 'B8'
    bolt4 = 'B9'
    bolt5 = 'B10'
    bolt6 = 'C10'
    verify_bolts()
    bolt1_name, bolt1_status ,bolt2_name, bolt2_status, bolt3_name, bolt3_status, bolt4_name, bolt4_status, bolt5_name, bolt5_status, bolt6_name, bolt6_status = verify_bolts()
    compare_bolts()
    reset_variables()
    time.sleep(1)

    workbook = openpyxl.load_workbook(diretorio_not)
    sheet = workbook.active
    facilities= variavel_da_empresa
    facilities_name = 'EMPRESA CONFIDENCIAL'
    qntbolt = join_company()
    bolt1 = 'B11'
    bolt2 = 'B12'
    bolt3 = 'B13'
    bolt4 = None
    bolt5 = None
    bolt6 = None
    verify_bolts()
    bolt1_name, bolt1_status ,bolt2_name, bolt2_status, bolt3_name, bolt3_status = verify_bolts()
    compare_bolts()
    reset_variables()
    time.sleep(1)

    workbook = openpyxl.load_workbook(diretorio_not)
    sheet = workbook.active
    facilities= variavel_da_empresa
    facilities_name = 'EMPRESA CONFIDENCIAL'
    qntbolt = join_company()
    bolt1 = 'B14'
    bolt2 = 'B15'
    bolt3 = 'B16'
    bolt4 = 'B17'
    bolt5 = 'B18'
    bolt6 = None
    verify_bolts()
    bolt1_name, bolt1_status ,bolt2_name, bolt2_status, bolt3_name, bolt3_status, bolt4_name, bolt4_status, bolt5_name, bolt5_status = verify_bolts()
    compare_bolts()
    reset_variables()
    time.sleep(1)

    workbook = openpyxl.load_workbook(diretorio_not)
    sheet = workbook.active
    facilities= variavel_da_empresa
    facilities_name = 'EMPRESA CONFIDENCIAL'
    qntbolt = join_company()
    bolt1 = 'B19'
    bolt2 = None
    bolt3 = None
    bolt4 = None
    bolt5 = None
    bolt6 = None
    verify_bolts()
    bolt1_name, bolt1_status = verify_bolts()
    compare_bolts()
    reset_variables()
    time.sleep(1)

    workbook = openpyxl.load_workbook(diretorio_not)
    sheet = workbook.active
    facilities= variavel_da_empresa
    facilities_name = 'EMPRESA CONFIDENCIAL'
    qntbolt = join_company()
    bolt1 = 'B20'
    bolt2 = None
    bolt3 = None
    bolt4 = None
    bolt5 = None
    bolt6 = None
    verify_bolts()
    bolt1_name, bolt1_status = verify_bolts()
    compare_bolts()
    reset_variables()
    time.sleep(1)

    workbook = openpyxl.load_workbook(diretorio_not)
    sheet = workbook.active
    facilities= variavel_da_empresa
    facilities_name = 'EMPRESA CONFIDENCIAL'
    qntbolt = join_company()
    bolt1 = 'E2'
    bolt2 = 'E3'
    bolt3 = None
    bolt4 = None
    bolt5 = None
    bolt6 = None
    verify_bolts()
    bolt1_name, bolt1_status ,bolt2_name, bolt2_status = verify_bolts()
    compare_bolts()
    reset_variables()
    time.sleep(1)

    workbook = openpyxl.load_workbook(diretorio_not)
    sheet = workbook.active
    facilities= variavel_da_empresa
    facilities_name = 'EMPRESA CONFIDENCIAL'
    qntbolt = join_company()
    bolt1 = 'E4'
    bolt2 = None
    bolt3 = None
    bolt4 = None
    bolt5 = None
    bolt6 = None
    verify_bolts()
    bolt1_name, bolt1_status = verify_bolts()
    compare_bolts()
    reset_variables()
    time.sleep(1)

    workbook = openpyxl.load_workbook(diretorio_not)
    sheet = workbook.active
    facilities = variavel_da_empresa
    facilities_name = 'EMPRESA CONFIDENCIAL'
    qntbolt = join_company()
    bolt1 = 'E7'
    bolt2 = 'E8'
    bolt3 = 'E9'
    bolt4 = 'E10'
    bolt5 = 'E11'
    bolt6 = 'E12'
    verify_bolts()
    bolt1_name, bolt1_status ,bolt2_name, bolt2_status, bolt3_name, bolt3_status, bolt4_name, bolt4_status, bolt5_name, bolt5_status, bolt6_name, bolt6_status = verify_bolts()
    compare_bolts()
    reset_variables()
    time.sleep(1)

    print('|-------------------------------------------------------------------------------------|')
    print('|--------------------------------- Loop Finalizado! ----------------------------------|')
    print('|-------------------------------------------------------------------------------------|')

    time.sleep(60)
 
    print('|-------------------------------------------------------------------------------------|')
    print('|---------------------------------- Iníciando loop! ----------------------------------|')
    print('|-------------------------------------------------------------------------------------|')



