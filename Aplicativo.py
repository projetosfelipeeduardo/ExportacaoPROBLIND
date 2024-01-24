import tkinter as tk
from tkinter import messagebox
from tkcalendar import DateEntry
from datetime import datetime
from time import sleep
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
import getpass

usuario = getpass.getuser()
def iniciar_driver():
    chrome_options = Options()
    arguments = ['--lang=pt-BR', '--window-size=1920,1080']
    for argument in arguments:
        chrome_options.add_argument(argument)

    prefs = {
        'download.prompt_for_download': False,
        'profile.default_content_setting_values.notifications': 2,
        'profile.default_content_setting_values.automatic_downloads': 1,
        'download.default_directory': f'C\\Users\\{usuario}\\Downloads',
    }

    chrome_options.add_experimental_option('prefs', prefs)

    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)
    return driver

def executar():
    data_inicial = entry_data_inicial.get()
    data_final = entry_data_final.get()

    try:
        data_inicial_obj = datetime.strptime(data_inicial, '%d/%m/%Y')
        data_final_obj = datetime.strptime(data_final, '%d/%m/%Y')

        data_inicial_formatada = data_inicial_obj.strftime('%d/%m/%Y')
        data_final_formatada = data_final_obj.strftime('%d/%m/%Y')

        navegador = iniciar_driver()
        navegador.get("https://carbonadmin.problind.com.br/#/login")

        campo_email = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@placeholder="Usuário"]')))
        campo_email.send_keys("felipe.pereira")

       
        campo_senha = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@placeholder="Senha"]')))
        campo_senha.send_keys("Carbon@10")

        btn_entrar = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[contains(text(),"ENTRAR")]')))
        btn_entrar.click()

        sleep(10)
        navegador.get("https://carbonadmin.problind.com.br/#/servicos/os/consulta")
        sleep(10)

        btn_calendario = navegador.find_element(By.XPATH, '//*[@class="btn grey reportrange"]')

        navegador.execute_script("arguments[0].scrollIntoView();", btn_calendario)
        btn_calendario.click()
        # btn_calendario= WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[3]/div[6]/div[2]/div/div[3]/div[1]/div[2]/form/div[1]/div[16]/div/div')))
        sleep(5)
        elemento_periodo = navegador.find_element(By.CSS_SELECTOR, 'body > div:nth-child(284) > div.ranges > ul > li:nth-child(7)')
        navegador.execute_script("arguments[0].scrollIntoView();", btn_calendario)
        elemento_periodo.click()
        sleep(1)
        elemento_dataIn = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//input[@class="input-mini active"]')))

        elemento_dataIn.send_keys(Keys.CONTROL + "a")
        elemento_dataIn.send_keys(Keys.BACKSPACE)
        elemento_dataIn.send_keys(data_inicial_formatada)

        sleep(1)
        elemento_dataFi= WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[6]/div[3]/div[1]/input')))
        elemento_dataFi.send_keys(Keys.CONTROL + "a")
        elemento_dataFi.send_keys(Keys.BACKSPACE)
        elemento_dataFi.send_keys(data_final_formatada)
        elemento_dataFi.send_keys(Keys.ENTER)
        sleep(10)

        btn_aplicar = navegador.find_element(By.XPATH, '/html/body/div[3]/div[6]/div[2]/div/div[3]/div[1]/div[2]/form/div[2]/button[1]')
        sleep(1)
        btn_aplicar.click()
        sleep(10)
        btn_download = navegador.find_element(By.XPATH, '/html/body/div[3]/div[6]/div[2]/div/div[3]/div[1]/div[2]/form/div[2]/button[4]')
        sleep(3)
        navegador.execute_script("arguments[0].scrollIntoView();", btn_calendario)
        btn_download.click()
        sleep(15)
        navegador.close()
        messagebox.showinfo("Aviso", "Arquivo Exportado com Sucesso")
    except ValueError:
        messagebox.showerror("Erro", "Formato de data inválido. Use DD/MM/AAAA.")

root = tk.Tk()
root.title("Seletor de Datas")
label_data_inicial = tk.Label(root, text="Data Inicial:")
label_data_inicial.pack(pady=5)
entry_data_inicial = DateEntry(root, date_pattern='dd/MM/yyyy', locale='pt_BR')
entry_data_inicial.pack(pady=5)

label_data_final = tk.Label(root, text="Data Final:")
label_data_final.pack(pady=5)
entry_data_final = DateEntry(root, date_pattern='dd/MM/yyyy', locale='pt_BR')
entry_data_final.pack(pady=5)

botao_executar = tk.Button(root, text="Executar", command=executar)
botao_executar.pack(pady=10)
root.mainloop()
