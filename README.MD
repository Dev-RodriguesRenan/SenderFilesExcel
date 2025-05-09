# Bot de Envio de BI

Este projeto é um bot automatizado que utiliza as bibliotecas **vjwhats** e **Selenium** para enviar arquivos de BI via WhatsApp Web. Com ele, é possível agendar o envio de mensagens com anexos, facilitando o compartilhamento de relatórios e análises.

## Estrutura do Projeto

.
├── .gitignore
├── main.py
├── data/
│ └── Links BI Vendedores.xlsx
├── README.MD
└── .env

## Pré-requisitos

- Python 3.6 ou superior
- Virtualenv (opcional, mas recomendado)

## Instalação

1. Clone o repositório:

   ```sh
   git clone <URL_DO_REPOSITORIO>
   cd SenderBI
   ```

2. Crie e ative um ambiente virtual (opcional, mas recomendado):

   ```sh
   python -m venv venv
   # No Windows:
   venv\Scripts\activate
   ```

3. Instale as dependências:
   ```sh
   pip install -r requirements.txt
   ```

## Uso

Para rodar o bot em modo de teste:

```sh
python main.py
```

O script agenda o envio de mensagens para ocorrer automaticamente toda sexta-feira às 12:00. Para ajustar o agendamento, modifique a configuração no próprio arquivo `main.py`.

## Estrutura do Código

- **main.py**: Contém a implementação do bot, que gerencia a inicialização do WebDriver, o envio de mensagens e o agendamento das tarefas.
- **requirements.txt**: Lista de dependências do projeto.
- **data/**: Contém o arquivo Excel com os dados para envio.
- **.env**: Arquivo de configuração com variáveis de ambiente, como o contato para envio.

## Contribuição

1. Faça um fork do projeto.
2. Crie uma branch para sua feature:
   ```sh
   git checkout -b feature/nova-feature
   ```
3. Commit suas mudanças:
   ```sh
   git commit -am 'Adiciona nova feature'
   ```
4. Faça um push para a branch:
   ```sh
   git push origin feature/nova-feature
   ```
5. Crie um novo Pull Request.

## Licença

Este projeto está licenciado sob a licença MIT. Veja o arquivo LICENSE para mais detalhes.
