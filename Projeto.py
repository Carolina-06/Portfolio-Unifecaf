import tkinter as tk
from tkinter import messagebox
from datetime import datetime
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import os

ARQUIVO = "dados_clima.xlsx"
WEATHER_URL = "https://weather.com/weather/today/l/-23.55,-46.64?par=google"


if not os.path.exists(ARQUIVO):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Data/Hora", "Temperatura", "Status da umidade do ar"])
    wb.save(ARQUIVO)
    print(f"Arquivo {ARQUIVO} criado.")

def buscar_previsao():
    try:

        options = webdriver.ChromeOptions()
        options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument(
            "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        )

        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),
                                  options=options)
        wait = WebDriverWait(driver, 20)

        driver.get(WEATHER_URL)

        temp_el = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "span[data-testid='TemperatureValue']"))
        )
        temperatura = temp_el.text.strip()  

        temp_f = float(temperatura.replace("°",""))
        temp_c = round((temp_f - 32) * 5/9, 1)
        temperatura = f"{temp_c}°C"
        print(f"Temperatura capturada: {temperatura}")

        try:
            umidade_el = wait.until(
                EC.presence_of_element_located((By.XPATH, "//span[text()='Humidity']/following-sibling::span"))
            )
            umidade = umidade_el.text.strip()
        except:
            umidade_el = wait.until(
                EC.presence_of_element_located((By.XPATH, "(//span[contains(text(),'%')])[1]"))
            )
            umidade = umidade_el.text.strip()
        print(f"Umidade capturada: {umidade}")

        driver.quit()

        wb = openpyxl.load_workbook(ARQUIVO)
        ws = wb.active

        agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        ws.append([agora, temperatura, umidade])
        wb.save(ARQUIVO)
        print(f"Dados salvos em {ARQUIVO}")

        messagebox.showinfo(
            "Sucesso",
            f"Dados capturados e salvos em {ARQUIVO}:\n\n"
            f"Data/Hora: {agora}\n"
            f"Temperatura: {temperatura}\n"
            f"Umidade: {umidade}"
        )

    except Exception as e:
        print("Erro capturando dados:", e)
        messagebox.showerror("Erro", f"Falha ao capturar dados:\n{e}")

janela = tk.Tk()
janela.title("Captador de Temperatura e Umidade - São Paulo")
janela.geometry("450x200")

lbl = tk.Label(janela, text="Clique para captar a previsão agora (São Paulo)", font=("Arial", 12))
lbl.pack(pady=20)

btn = tk.Button(janela, text="Buscar previsão", font=("Arial", 12),
                bg="#2563eb", fg="white", command=buscar_previsao)
btn.pack(pady=10)

janela.mainloop()