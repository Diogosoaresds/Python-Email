import os
import openpyxl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def filtrar_planilha(caminho_arquivo, mes_ano_inicio, mes_ano_fim):
    linhas_filtradas = []
    
    try:
        workbook = openpyxl.load_workbook(caminho_arquivo)
        sheet = workbook.active
        
        for row in sheet.iter_rows(min_row=2, values_only=True):
            data_row = row[4]  # Data na coluna 5 (index 4)
            if data_row is not None:
                mes_ano_data = data_row.strftime('%m-%Y')
                if mes_ano_inicio <= mes_ano_data <= mes_ano_fim:
                    linhas_filtradas.append(row)

        return linhas_filtradas
    except FileNotFoundError:
        print("Arquivo não encontrado!")
    except Exception as e:
        print("Ocorreu um erro:", e)

def enviar_email(destinatario):
    try:
        smtp_server = 'smtp.gmail.com'
        port = 587
        sender_email = '@gmail.com'
        password = 'senha_app'
        assunto = "Teste"
        mensagem = "Olá,\n\nEsta é uma mensagem de teste."

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = destinatario
        msg['Subject'] = assunto
        msg.attach(MIMEText(mensagem, 'plain'))

        server = smtplib.SMTP(smtp_server, port)
        server.starttls()
        server.login(sender_email, password)

        server.send_message(msg)
        
        server.quit()
        
        print(f"E-mail enviado para {destinatario} com sucesso!")
    except Exception as e:
        print(f"Erro ao enviar e-mail para {destinatario}: {e}")

if __name__ == "__main__":
    caminho_arquivo = "caminho do arquivo .xlsx"
    data_inicio = input("Digite a data de início (formato MM-AAAA): ")
    data_fim = input("Digite a data de fim (formato MM-AAAA): ")

    linhas_filtradas = filtrar_planilha(caminho_arquivo, data_inicio, data_fim)

    if linhas_filtradas:
        for linha in linhas_filtradas:
            destinatario = linha[7]  # Destinatário na coluna 8
            enviar_email(destinatario)
    else:
        print("Nenhuma linha encontrada no período especificado.")
        
