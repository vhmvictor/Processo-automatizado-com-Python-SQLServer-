#--------------------------------------------------------CONEXÃO COM ORACLE-------------------------------------------------------------

import cx_Oracle
import smtplib
import csv

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

#CONECTANDO NO BANCO
con = cx_Oracle.connect("usuario/senha@ip-servidor/instancia")
cursor = con.cursor()

#GERANDO O ARQUIVO CSV 
csv_file = open("ARQUIVO.csv", "w")

#ESTRUTURANDO O ARQUIVO E EXECUTANDO SQL
writer = csv.writer(csv_file, delimiter=';', lineterminator="\n", quoting=csv.QUOTE_NONNUMERIC)
r = cursor.execute('select * from TABELA')

#NOME DA TABELA
col_names = [row[0] for row in cursor.description]
writer.writerow(col_names)

#GRAVANDO INFORMAÇÕES DA TABELA
for row in cursor:
    writer.writerow(row)
   
cursor.close()
con.close()
csv_file.close()

print('Exportação Concluída !')

#-----------------------------------------------------CONEXAO SMTP - ENVIO DE E-MAIL---------------------------------------------------

email_user = 'exemplo-email@hotmail.com'
email_password = ''
email_send = 'exemplo-email@hotmail.com'

cc = 'exemplo-email-copia@hotmail.com'

subject = 'Teste Envio de E-mail - PYTHON'

msg = MIMEMultipart()
msg['From'] = email_user
msg['To'] = email_send 
msg['Cc'] = cc
msg['Subject'] = subject

print ('Enviando E-mail para: ' + email_send + '...')

body = 'Teste E-mail'
msg.attach(MIMEText(body,'plain'))

filename ='operacao.csv'
attachment  = open(filename,'rb')

part = MIMEBase('application','octet-stream')
part.set_payload((attachment).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition',"attachment; filename = " + filename)

msg.attach(part)
text = msg.as_string()
#UTILIZANDO SERVIDOR SEGURO
server = smtplib.SMTP("rsmtp-mail.outlook.com",587)
#server = smtplib.SMTP("smtp-mail.outlook.com",587)
#server.starttls()
#server.login(email_user,email_password)

server.sendmail(email_user,email_send,text)

print ('E-mail enviado!')

server.quit()