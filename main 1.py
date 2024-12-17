import pandas as pd
import time
import smtplib
from datetime import datetime  # Importando para capturar data e hora
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


# Lê os dados do arquivo Excel
df = pd.read_csv("Esp8266_Receiver (1).csv", encoding='utf-8')

# Define as colunas que representam os valores das esteiras
esteira1 = df["esteira1"]
esteira2 = df["esteira2"]
esteira3 = df["esteira3"]


def relatorio(esteira, estado, valor):
    """Cria um relatório e salva no arquivo Excel, incluindo a data e hora."""
    # Captura a data e hora atuais
    current_time = datetime.now()
    date = current_time.strftime("%Y-%m-%d")  # Formato da data: ano-mês-dia
    time_of_day = current_time.strftime("%H:%M:%S")  # Formato da hora: hora:minuto:segundo

    # Cria o DataFrame com os dados a serem adicionados, incluindo Date e Time
    novo_dado = pd.DataFrame({
        "esteira": [esteira],
        "valor": [valor],
        "estado": [estado],
        "Date": [date],
        "Time": [time_of_day]  # Hora do processamento
    })

    # Salva os dados no arquivo de relatório Excel
    arquivo_relatorio = "Relatorio.xlsx"

    try:
        with pd.ExcelWriter(arquivo_relatorio, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            startrow = writer.sheets['Sheet1'].max_row  # Determina a próxima linha disponível
            novo_dado.to_excel(writer, index=False, header=False, sheet_name='Sheet1', startrow=startrow)
        print(
            f"Relatório atualizado com dados: Esteira={esteira}, Valor={valor}, Estado={estado}, Date={date}, Time={time_of_day}")
    except FileNotFoundError:
        # Cria um novo arquivo se o relatório ainda não existir
        novo_dado.to_excel(arquivo_relatorio, index=False, sheet_name='Sheet1')
        print(
            f"Novo relatório criado com dados: Esteira={esteira}, Valor={valor}, Estado={estado}, Date={date}, Time={time_of_day}")

def enviar_email(teste : str):
    try:
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login("nicolecamachorose873@gmail.com", "scht dbuf lvog aueu")

        # Criando a mensagem de e-mail
        msg = MIMEMultipart()
        msg['From'] = "nicolecamachorose873@gmail.com"
        msg['To'] = "nicolecamachorose873@gmail.com"
        msg['Subject'] = "Alerta de Estoque"
        
        # Definindo o conteúdo do e-mail com a codificação utf-8
        body = MIMEText(teste, 'plain', 'utf-8')
        msg.attach(body)

        # Envia o e-mail
        server.sendmail(msg['From'], msg['To'], msg.as_string())
        server.quit()
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")

def checar_valor(esteira, valor):
    """Verifica o estado da esteira com base no valor e gera um relatório."""
    if valor == 1:
        estado = "Estoque baixo"
        enviar_email(estado)
    elif valor == 2:
        estado = "Estoque médio"
        enviar_email(estado)
    elif valor == 3:
        estado = "Estoque cheio"
        enviar_email(estado)
    else:
        estado = "Valor inválido"
        enviar_email(estado)

    print(f"{esteira}: {valor} - {estado}")
    if estado != "Valor inválido":
        relatorio(esteira, estado, valor)


def ler_linhas(esteira1, esteira2, esteira3):
    """Itera pelas linhas das esteiras e verifica os valores."""
    for valor1, valor2, valor3 in zip(esteira1, esteira2, esteira3):
        print("Processando Esteira1...")
        checar_valor("Esteira1", valor1)
        time.sleep(1)

        print("Processando Esteira2...")
        checar_valor("Esteira2", valor2)
        time.sleep(1)

        print("Processando Esteira3...")
        checar_valor("Esteira3", valor3)
        time.sleep(1)


#função para o envio dos emails




# # Configuração do envio de e-mails (se necessário)

# server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
# server.login("nicolecamachorose873@gmail.com", "scht dbuf lvog aueu")
# server.sendmail(
#     "nicolecamachorose873@gmail.com",
#     "nicolecamachorose873@gmail.com",
#     "")
# server.quit()


# Inicia o processamento das esteiras
ler_linhas(esteira1, esteira2, esteira3)
