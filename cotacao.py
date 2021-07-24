import requests  # Pega informações de API's
from datetime import datetime  # Data e hora atual
from openpyxl import Workbook  # Cria arquivo Excel
from openpyxl.styles import Alignment, Font  # Estilos para células
import pandas as pd  # Nesse caso, estou usando p/ transformar em html

# Busca as informações do site
requisicao = requests.get(
    "https://economia.awesomeapi.com.br/last/USD-BRL,EUR-BRL,BTC-BRL")

# Converte para json
dados = requisicao.json()

# Pegando os dados: dólar, euro e bitcoin
cotacao_dolar = dados["USDBRL"]["bid"]
cotacao_euro = dados["EURBRL"]["bid"]
cotacao_bitcoin = dados["BTCBRL"]["bid"]

# Criando um arquivo cotacao.xlsx
wb = Workbook()
tabela = wb.active

# Título da tabela
tabela.title = "Cotacao Moedas"

# Nomeando as Células
tabela["A1"] = "Moedas"
tabela["B1"] = "Cotação"
tabela["C1"] = "Última Atualização"

# Nome das moedas
tabela["A2"] = "Dólar"
tabela["A3"] = "Euro"
tabela["A4"] = "Bitcoin"

# Valor das cotações
tabela["B2"] = float(cotacao_dolar)
tabela["B3"] = float(cotacao_euro)
tabela["B4"] = float(cotacao_bitcoin)

# Formatação da data atual
tabela["C2"] = datetime.now().strftime("Às %Hh:%Mm:%Ss - %d/%m/%Y")
tabela["C3"] = datetime.now().strftime("Às %Hh:%Mm:%Ss - %d/%m/%Y")
tabela["C4"] = datetime.now().strftime("Às %Hh:%Mm:%Ss - %d/%m/%Y")

# Alinha todas as células preenchidas
celulas = ["a", "b", "c"]
for letra in celulas:
    for i in range(1, 5):
        if i == 1:
            # Deixa a primeira linha em negrito
            tabela[letra+str(i)].font = Font(bold=True)
        tabela[letra+str(i)].alignment = Alignment("center", "center")

# Salvando a planilha
wb.save("cotacao.xlsx")

# Transformando em html, caso queira enviar para e-mails.
tabela = pd.read_excel("cotacao.xlsx")
tabela.to_html("cotacao.html")
print(tabela)
