import pandas as pd
import cx_Oracle
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv
import os

# Carregar variáveis do .env
load_dotenv()

# Configurações de Oracle
ORACLE_USER = os.getenv('ORACLE_USER')
ORACLE_PASSWORD = os.getenv('ORACLE_PASSWORD')
ORACLE_DSN = os.getenv('ORACLE_DSN')

# Caminho do Excel
EXCEL_PATH = os.getenv('EXCEL_PATH')

# Email settings
EMAIL_ADDRESS = os.getenv('EMAIL_ADDRESS')
EMAIL_ADDRESS_TO_SENT = os.getenv('EMAIL_ADDRESS_TO_SENT').split(',')
SMTP_SERVER = os.getenv('SMTP_SERVER')
SMTP_PORT = int(os.getenv('SMTP_PORT'))
SMTP_USERNAME = os.getenv('SMTP_USERNAME')
SMTP_PASSWORD = os.getenv('SMTP_PASSWORD')

def obter_dados_oracle():
    conn = cx_Oracle.connect(ORACLE_USER, ORACLE_PASSWORD, ORACLE_DSN)
    cursor = conn.cursor()
    query = """
        WITH paises_extraidos AS (
            SELECT 
                CASE 
                    WHEN SUBSTR(expe_deev_id, 7, 2) <> 'PT' THEN SUBSTR(expe_deev_id, 7, 2)
                    ELSE SUBSTR(expe_deev_id, 1, 2)
                END AS pais
            FROM IPS_AQ.TIPSREGFALTAINTIPSCI
            WHERE data_apuramento < TRUNC(SYSDATE)
        )
        SELECT pais, COUNT(*) AS total
        FROM paises_extraidos
        GROUP BY pais
        ORDER BY total DESC
    """
    cursor.execute(query)
    resultados = cursor.fetchall()
    cursor.close()
    conn.close()
    return pd.DataFrame(resultados, columns=["Pais", "Total"])

def carregar_excel():
    df = pd.read_excel(EXCEL_PATH, header=None, names=["Pais", "Responsavel"])
    df["Pais"] = df["Pais"].str.strip()  # Limpar espaços
    return df

def cruzar_dados(df_oracle, df_responsaveis):
    df_merged = pd.merge(df_oracle, df_responsaveis, on="Pais", how="left", validate="many_to_one")
    return df_merged

def gerar_corpo_email(df_final):
    # Substituir valores NaN na coluna "Responsavel" por "Não atribuído"
    df_final["Responsavel"] = df_final["Responsavel"].fillna("Não atribuído")

    # Ordenar o resumo por total de pedidos (descendente) e nome do responsável (ascendente)
    resumo = df_final.groupby("Responsavel")["Total"].sum().reset_index()
    resumo = resumo.sort_values(by=["Total", "Responsavel"], ascending=[False, True])

    # Ordenar o detalhe por total (descendente) e nome do responsável (ascendente)
    df_final = df_final.sort_values(by=["Total", "Responsavel", "Pais"], ascending=[False, True, True])

    # Gerar o corpo do email
    corpo = "<p>Olá,</p><p>Segue o resumo das expedições a tratar:</p><ul>"
    for _, row in resumo.iterrows():
        corpo += f"<li><b>{row['Responsavel']}:</b> {row['Total']} pedidos</li>"
    corpo += "</ul><p>Detalhe por país:</p><table border='1' cellpadding='5' cellspacing='0'>"
    corpo += "<tr><th>País</th><th>Total</th><th>Responsável</th></tr>"
    for _, row in df_final.iterrows():
        corpo += f"<tr><td>{row['Pais']}</td><td>{row['Total']}</td><td>{row['Responsavel']}</td></tr>"
    corpo += "</table><p>Cumprimentos,<br>Nuno Franco</p>"
    return corpo

def enviar_email(corpo_html):
    message = MIMEMultipart()
    message["From"] = f"NUNO FRANCO <{EMAIL_ADDRESS}>"
    message["To"] = ", ".join(EMAIL_ADDRESS_TO_SENT)
    message["Subject"] = "Resumo das Expedições a tratar por País e Responsável"
    
    message.attach(MIMEText(corpo_html, "html"))
    
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(SMTP_USERNAME, SMTP_PASSWORD)
        server.sendmail(EMAIL_ADDRESS, EMAIL_ADDRESS_TO_SENT, message.as_string())
    print("Email enviado com sucesso.")

def main():
    print("Obtendo dados da BD Oracle...")
    df_oracle = obter_dados_oracle()

    print("Carregando Excel de responsáveis...")
    df_responsaveis = carregar_excel()

    print("Cruzando dados...")
    df_final = cruzar_dados(df_oracle, df_responsaveis)

    print("Gerando email...")
    corpo_email = gerar_corpo_email(df_final)

    print("Enviando email...")
    enviar_email(corpo_email)

if __name__ == "__main__":
    main()
