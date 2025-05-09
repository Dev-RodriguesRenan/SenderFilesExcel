import os
import datetime
import time
from dotenv import load_dotenv
from pathlib import Path
import schedule
from vjwhats import WhatsApp
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

load_dotenv()


class WhatsApp_Handler:
    def __init__(self):
        self.profile_directory = "Developer"
        self.user_data_dir = os.path.expandvars(
            "%LOCALAPPDATA%\\Google\\Chrome\\User Data\\DevLittle"
        )
        self.profile_directory = "DevLittle"
        # Configurações do Chrome
        self.chrome_options = Options()
        self.chrome_options.add_argument(f"user-data-dir={self.user_data_dir}")
        self.chrome_options.add_argument(f"profile-directory={self.profile_directory}")
        # chrome_options.add_argument("--headless")  # Executa o Chrome em segundo plano
        self.service = Service()
        self.driver = None
        self.whatsapp = None

    def __enter__(self):
        # Inicializa o driver do Chrome
        for _ in range(3):
            try:
                self.driver = webdriver.Chrome(
                    service=self.service, options=self.chrome_options
                )
                self.whatsapp = WhatsApp(self.driver)
                return self
            except Exception as e:
                print(f"Erro ao inicializar o driver do Chrome: {e}")
                os.system("taskkill /F /IM chrome.exe")
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        # Fecha o driver do Chrome
        if self.driver:
            self.driver.quit()

    def enviar_mensagem(self, contact, message, attachment):
        message += "\n\n" + f'> {datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")}'

        self.whatsapp.find_by_username(contact)
        # Envia a mensagem
        self.whatsapp.send_file(attachment, which=2)
        self.whatsapp.send_message(message)


def main():
    message = """
Lembrete Importante!

Não deixem de acessar nossa ferramenta de análise de dados!
Lá, vocês conseguem acompanhar o desempenho diário, semanal, mensal e trimestral, tendo uma visão completa dos resultados alcançados.

Acompanhem seu progresso e mantenham-se informados sobre o seu desempenho!
Qualquer dúvida, estamos à disposição para ajudar!

Bom final de semana e ótimas análises!
"""
    attachment = Path("data/Links BI Vendedores.xlsx")
    contact = os.getenv("CONTATO")
    with WhatsApp_Handler() as whatsapp_handler:
        whatsapp_handler.enviar_mensagem(contact, message, attachment)
    print("Mensagem enviada com sucesso!")


if __name__ == "__main__":
    print(f'{time.strftime("%X")} - INFO - Iniciando o envio de mensagens...')
    # Schedule the job every Friday at 12:00 PM
    schedule.every().friday.at("12:00").do(main)
    while True:
        print(f'{time.strftime("%X")} - INFO - Aguardando o horário programado...')
        schedule.run_pending()
        time.sleep(1)
