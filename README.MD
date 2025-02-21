
# Bot de Envio de BI

Este projeto é um bot automatizado que utiliza a biblioteca BotCity para enviar arquivos de BI via WhatsApp Web.

## Estrutura do Projeto

.
├── .gitignore
├── botPython/
│   ├── bot.py
│   ├── botPython.botproj
│   ├── build.bat
│   ├── build.ps1
│   ├── build.sh
│   ├── requirements.txt
│   └── resources/
│       ├── ATTACHMENT_XLSX.png
│       ├── DOCUMENTS.png
│       ├── FILENAME.png
│       ├── MENU_EXPLORER.png
│       ├── MORE.png
│       └── ...
├── data/
│   └── Links BI Vendedores.xlsx
└── README.MD

## Pré-requisitos

- Python 3.6 ou superior
- Virtualenv (opcional, mas recomendado)

## Instalação

1. Clone o repositório:

```sh
git clone <URL_DO_REPOSITORIO>
cd robo_de_envios_BI
```

2. Crie e ative um ambiente virtual (opcional, mas recomendado):

```sh
python -m venv venv
source venv/bin/activate  # No Windows: venv\Scripts\activate
```

3. Instale as dependências:

```sh
pip install -r botPython/requirements.txt
```

## Uso

Para rodar o bot em modo debug:

```sh
python botPython/bot.py debug
```

Para agendar o bot para rodar diariamente às 12:00:

```sh
python botPython/bot.py
```

## Estrutura do Código

- bot.py: Contém a implementação do bot.
- requirements.txt: Lista de dependências do projeto.
- build.sh, build.ps1, build.bat: Scripts para empacotar o projeto.
- resources: Contém os recursos de imagem utilizados pelo bot.

## Contribuição

1. Faça um fork do projeto.
2. Crie uma branch para sua feature (`git checkout -b feature/nova-feature`).
3. Commit suas mudanças (`git commit -am 'Adiciona nova feature'`).
4. Faça um push para a branch (`git push origin feature/nova-feature`).
5. Crie um novo Pull Request.

## Licença

Este projeto está licenciado sob a licença MIT. Veja o arquivo LICENSE para mais detalhes.
