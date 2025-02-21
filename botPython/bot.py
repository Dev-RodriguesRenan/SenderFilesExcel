import datetime
import os
import sys
import time
from urllib.parse import quote
import schedule
from dotenv import load_dotenv
from colorama import Fore, Style
from botcity.core import DesktopBot

load_dotenv()
data_path = os.path.join(os.getcwd(), "data")
os.makedirs(data_path, exist_ok=True)
"""python -m cookiecutter https://github.com/botcity-dev/bot-python-template/archive/v2.zip"""
ID_GROUP = os.getenv("ID_GROUP")
print(f"grupo á enviar conteudo: {ID_GROUP}")
PHONE_NUMBER = os.getenv("PHONE_NUMBER")
print(f"numero á enviar conteudo: {PHONE_NUMBER}")


class BiBot(DesktopBot):

    def __init__(self, debug):
        self.text = f"""
🔔 Lembrete Importante!

Não deixem de acessar nossa ferramenta de análise de dados! 📊
Lá, vocês conseguem acompanhar o desempenho diário, semanal, mensal e trimestral, tendo uma visão completa dos resultados alcançados.

🚀 Acompanhem seu progresso e mantenham-se informados sobre o seu desempenho!
Qualquer dúvida, estamos à disposição para ajudar!

Bom final de semana e ótimas análises! 💪😎

> {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}
"""
        self.debug = debug
        if not self.debug:
            self.url_whats = f"https://web.whatsapp.com/accept?code={ID_GROUP}"
        else:
            self.text = quote(self.text)
            self.url_whats = f"https://web.whatsapp.com/send/?phone=%2B{PHONE_NUMBER}&text={self.text}&type=phone_number&app_absent=0"

        print(self.url_whats)
        self.path_bi = data_path
        super().__init__()

    def start_browser(self):
        self.browse(self.url_whats)
        self.wait(30_000)

    def send_attachment(self):
        # Searching for element 'MORE '
        #  Procura o more pra precissionar enter e enviar o txt
        if not self.find("MORE", matching=0.87, waiting_time=10000):
            self.not_found("MORE")
        if not self.debug:
            self.paste(self.text)

        # Searching for element 'MORE '
        if not self.find("MORE", matching=0.87, waiting_time=15000):
            self.not_found("MORE")
        self.click()
        # Searching for element 'DOCUMENTS '
        if not self.find("DOCUMENTS", matching=0.87, waiting_time=15000):
            self.not_found("DOCUMENTS")
        self.click()
        # seleciona o arquivo
        self.select_attachment()
        # envia o arquivo
        self.subimit_attachment()

    def select_attachment(self):
        # Searching for element 'MENU_EXPLORER '
        if not self.find("MENU_EXPLORER", matching=0.87, waiting_time=15000):
            self.not_found("MENU_EXPLORER")

        self.control_key("l")
        self.paste(self.path_bi)
        self.enter()

        # Searching for element 'FILENAME '
        if not self.find("FILENAME", matching=0.87, waiting_time=10000):
            self.not_found("FILENAME")
            for _ in range(6):
                self.tab()
        else:
            self.click_relative(69, 13)
        self.paste("links BI Vendedores")
        # Searching for element 'OPEN_FILE '
        if not self.find("OPEN_FILE", matching=0.87, waiting_time=15000):
            self.not_found("OPEN_FILE")

        self.enter()

    def subimit_attachment(self):
        # Searching for element 'ATTACHMENT_XLSX '
        if not self.find("ATTACHMENT_XLSX", matching=0.87, waiting_time=15000):
            self.not_found("ATTACHMENT_XLSX")

        self.enter()

    def quit_browser(self, timeout=10):
        time.sleep(timeout)
        self.type_keys(["ctrl", "w"])

    def run(self):
        self.start_browser()
        self.send_attachment()
        self.quit_browser()

    def not_found(self, element):
        print(
            f'\n>>{Fore.CYAN}{time.strftime("%X")} -- {Fore.RED}O elemento {element} não foi encontrado{Style.RESET_ALL}\n'
        )


def main(debug=False):
    bi_bot = BiBot(debug=debug)
    bi_bot.run()


if __name__ == "__main__":
    if len(sys.argv) > 1:
        if sys.argv[1] == "debug":
            print(
                f">> {Fore.WHITE}{time.strftime('%X')} [INFO] {Fore.MAGENTA}Debug mode activate [INFO] {Style.RESET_ALL}",
                end="\r",
            )
            main(debug=True)
        elif sys.argv[1] == "debug-group":
            print(
                f">> {Fore.WHITE}{time.strftime('%X')} [INFO] {Fore.MAGENTA}Debug mode activate [INFO] {Style.RESET_ALL}",
                end="\r",
            )
            main(debug=False)
    else:
        schedule.every(1).day.at("12:03").do(main)
        while True:
            print(
                f">> {Fore.BLUE}{time.strftime('%X')} -- {Fore.GREEN}Aguardando o horario correto!!{Style.RESET_ALL}",
                end="\r",
            )
            schedule.run_pending()
            time.sleep(1)
