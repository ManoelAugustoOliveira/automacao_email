# ================================================== Projeto Automação Envio de Email ================================#
# imports
import smtplib # Simple Mail Tranfer Protocol
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# ======================================================== ENVIANDO O EMAIL ==========================================#
# Outlook
# smtp.office365.com
# Porta: 587

#1 - STARTAR O SERVIDOR SMTP
host = 'smtp.office365.com'
port = '587'
login = 'email(colocar email outlook)'
senha = 'colocar senha do email outlook'

server = smtplib.SMTP(host, port)
server.ehlo()
server.starttls()
server.login(login, senha)

#===================================================== CONSTRUIR EMAIL TIPO MIME ======================================#

corpo = f"""<p>Olá!</p>
            <p>Abaixo segue relatório com as informações de vendas do mês.</p>
            <b>Total:</b>
            <table style = 'width:35%; border: 1px solid black'>
                <tr>
                    <th style ='border:1px solid black'>Indicador</th>
                    <th style ='border:1px solid black'>Valor</th>
                </tr>
                <tr style = 'border: 1px solid black'>
                    <td style = 'border: 1px solid black'>FATURAMENTO TOTAL</td> 
                    <td style = 'border: 1px solid black'>colocar</td> 
                <tr style = 'border: 1px solid black'>
                    <td style = 'border: 1px solid black'>TOTAL DE VENDAS (NF-e)</td>
                    <td style = 'border: 1px solid black'>Colocar Valor Aqui</td>
                <tr style = 'border: 1px solid black'>
                    <td style = 'border: 1px solid black'>TICKET MÉDIO POR VENDA</td>
                    <td style = 'border: 1px solid black'>Colocar Valor Aqui</td>
                <tr style = 'border: 1px solid black'>
                    <td style = 'border: 1px solid black'>CMV</td>
                    <td style = 'border: 1px solid black'>Colocar valor aqui</td>
                <tr style = 'border: 1px solid black'>
                    <td style = 'border: 1px solid black'>ICMS/ISSQN</td>
                    <td style = 'border: 1px solid black'>Colocar valor aqui</td>
                <tr style = 'border: 1px solid black'>
                    <td style = 'border: 1px solid black'>PIS</td>
                    <td style = 'border: 1px solid black'>Colocar Valor Aqui</td>
                <tr style = 'border: 1px solid black'>
                    <td style = 'border: 1px solid black'>COFINS</td>
                    <td style = 'border: 1px solid black'>Colocar Valor Aqui</td>
                <tr style = 'border: 1px solid black'>
                    <td style = 'border: 1px solid black'>RESULTADO BRUTO DO DIA</td>
                    <td style = 'border: 1px solid black'>Colocar Valor Aqui</td>
            </table>
            <p> Atenciosamente,<p/>"""


email_msg = MIMEMultipart()
email_msg['From'] = login
email_msg['To'] =  login
email_msg['Subject'] = f"Relatório de faturamento"
email_msg.attach(MIMEText(corpo,'html'))

# Abrir o arquivo em modo leitura e binary
cam_arq = "C:/Users/atend/OneDrive/Área de Trabalho/email/datasets/statusinvest-busca-avancada.xlsx"
attchement = open(cam_arq, 'rb')

# base64
att = MIMEBase('application', 'octet-stream')
att.set_payload(attchement.read())
encoders.encode_base64(att)

# cabeçalho  no tipo anexo
att.add_header("Content-Disposition", f'attchement; filename = statusinvest-busca-avancada.xlsx')
attchement.close()
email_msg.attach(att)


#3- ENVIAR O EMAIL TIPO MIME NO SERVIDOR
server.sendmail(email_msg['From'], email_msg['To'], email_msg.as_string())
server.quit()

print('Email enviado')
