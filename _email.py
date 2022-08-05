from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
from datetime import datetime

data = datetime.today().strftime('%d%m%Y')
hora = datetime.now().strftime('%H')

# Estabelecendo conexão SMTP
def connect_smtp(servidor, usuario, senha):
    m = smtplib.SMTP(servidor, 587)
    m.starttls()
    m.login(usuario, senha)
    return m

# Parametrizando envio
def enviar_email(msg):
    servidor = 'smtp.suaempresa.com.br'
    usuario = 'seuemail.com.br'
    senha = 'sua senha'
    remetente = usuario
    destinatario = 'destinatario@.com.br'

    assunto = f"Automação dia:-{data}"
    mensagem = msg
    e = connect_smtp(servidor, usuario, senha)

    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = assunto

    msgAlernative = MIMEMultipart('alternative')
    msg.attach(msgAlernative)

    msgtext = MIMEText(mensagem, 'html')
    msgAlernative.attach(msgtext)

    e.sendmail(remetente, destinatario, msg.as_string())
    e.quit()