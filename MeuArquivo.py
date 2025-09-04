import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# --- Ler planilha ---
tabela_vendas = pd.read_excel('Vendas.xlsx')

# --- Estabelecer variáveis ---
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})

# --- Montar corpo do e-mail ---
corpo_email = f"""
<p>Olá,</p>
<p>Segue o relatório de vendas por loja:</p>

<p><b>Faturamento:</b><br>{faturamento.to_html(formatters = {'Valor Final': 'R${:,.2f}'.format})}</p>
<p><b>Quantidade Vendida:</b><br>{quantidade.to_html()}</p>
<p><b>Ticket Médio:</b><br>{ticket_medio.to_html(formatters = {'Ticket Médio': 'R${:,.2f}'.format})}</p>

<p>Atenciosamente,<br>Eduarda</p>
"""

# --- Configurações do e-mail ---
remetente = "dudaavelinobonfim@gmail.com"
senha = "migr desq vuyl wevn"
destinatario = "dudaavelinobonfim@gmail.com"

# --- Criar a mensagem ---
msg = MIMEMultipart("alternative")
msg["From"] = remetente
msg["To"] = destinatario
msg["Subject"] = "Relatório de Vendas por Loja"
msg.attach(MIMEText(corpo_email, "html"))

# --- Enviar e-mail via SMTP do Gmail ---
try:
    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()  # conexão segura
        server.login(remetente, senha)
        server.sendmail(remetente, destinatario, msg.as_string())
    print("✅ E-mail enviado com sucesso!")
except Exception as e:
    print("❌ Erro ao enviar e-mail:", e)
















# <p>Prezados,</p>

#<p>Segue o Relatório de Vendas por cada Loja.</p>

#<p>Faturamento:</p>
#{faturamento.to_html(formatters = {'Valor Final': 'R${:,.2f}'.format})}

#<p>Quantidade Vendida:</p>
#{quantidade.to_html()}

#<p>Ticket Médio dos Produtos em cada Loja:</p>
#{ticket_medio.to_html(formatters = {'Ticket Médio': 'R${:,.2f}'.format})}

#<p>Qualquer dúvida estou à disposição.</p>

#<p>Atenciosamente,</p>
#<p>Eduarda</p>
#'''



