import logging;import win32com.client;import os;import winreg;import time;import datetime
#Definindo a pasta documento

def get_documents_folder():
    try:
        # Abre a chave do registro do Windows que contém o caminho da pasta Documentos
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                             r'Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders')

        # Lê o valor da chave que contém o caminho da pasta Documentos
        value, _ = winreg.QueryValueEx(key, 'Personal')
        # Retorna o caminho para a pasta Documentos
        return value
    except WindowsError:
        # Caso ocorra um erro ao tentar ler a chave do registro, retorna o caminho padrão da pasta Documentos
        return os.path.expanduser("~\\Documentos")



vld_dire = get_documents_folder()[0]
if vld_dire == "C":
    documents_folder = get_documents_folder()
else:
    documents_folder = os.path.dirname(os.path.realpath(__file__))


#Criando nomenclatura para o Log
from datetime import datetime

# Obtendo a data atual
current_date = datetime.now()
dt_formatada = current_date.strftime("\%Y%m%d") # Definindo data com base na data atual
hr_exec = current_date.strftime("%d/%m/%Y-%H:%M:%S") # Definindo nomenclatura para Log's
hr_log = current_date.strftime("%Y%m%d%H%M%S") # Definindo nomenclatura para arquivo em anexo para consulta

logging.basicConfig(filename=f'{documents_folder}{dt_formatada}.log',
                    level=logging.DEBUG,
                    filemode='w',
                    format='%(message)s')  # Definindo padrão do Log e destinatario

# Adicionando uma mensagem ao log

logging.debug(f"Data registrada - Hora Registrada - Atividade - Status - Tars")
logging.debug(f"{hr_exec} - Inicio BkpPy - 0 - Sistema")

time.sleep(1)

#Conectar com o OUTLOOOK

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
logging.debug(f"{hr_exec} - Conexão com o Outlook realizada com sucesso - 2 - Sistema")
inbox = outlook.GetDefaultFolder(6)  # 6 é o código para a pasta de entrada
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
Pasta_despache="C:\\Windows\\Temp\\"
erro="Error"
today = datetime.now().date()
for message in messages:
    if "Duplidata Backup Log" in message.Subject and message.ReceivedTime.date() == today:
        tars = message.Subject[-5:]
        logging.debug(f"{hr_exec} - Email do tars assunto '{tars}' encontrado com sucesso - 3 - {tars}")

        if message.Attachments.Count == 0:
            logging.debug(f"{hr_exec} - E-mail do tars {tars} sem anexo - 6 - {tars}")
        else:
                for attachment in message.Attachments:
                    attachment_name = attachment.FileName
                    primeiras_quatro_letras = attachment_name[:4]

                    if attachment.Size > 0:


                        if primeiras_quatro_letras == "diff":  # Se encontrar um documento Diff
                            logging.debug(f"{hr_exec} - Localizado anexo Diff em anexo - D - {tars}")
                            if attachment.FileName.endswith(".txt"):
                                file_path = f"{Pasta_despache}{hr_log}.txt"
                                try:
                                    attachment.SaveAsFile(file_path)
                                    with open(file_path, "r") as file:
                                        content = file.read()
                                        logging.debug(
                                            f"{hr_exec} - realizando a leitura do anexo {primeiras_quatro_letras} - 1 - Sistema")
                                        if erro in content:
                                            logging.debug(
                                                f"{hr_exec} - Erro encontrado no log, um ou mais documentos não foram copiados - K - {tars}")
                                        else:
                                            logging.debug(f"{hr_exec} - Log do hotel está OK - L - {tars}")
                                except ValueError as e:
                                    logging.debug(f"{hr_exec} - Erro ao copiar ao conteudo - X - {tars}")
                                    logging.debug(f"{hr_exec} - Tentativa de leitura com metodo diferente - B - {tars}")
                                    try:
                                        with open(file_path, "r", encoding="utf-8") as file:
                                            content = file.read()
                                            if erro in content:
                                                logging.debug(
                                                    f"{hr_exec} - Erro encontrado no log, um ou mais documentos não foram copiados - K - {tars}")
                                            else:
                                                logging.debug(f"{hr_exec} - Log do hotel está OK - L - {tars}")
                                    except Exception as e:
                                        logging.debug(f"{hr_exec} - Tentativa com o novo metodo falhado - 8 - {tars}")
                            else:
                                logging.debug(
                                    f"{hr_exec} - Documento {primeiras_quatro_letras} está com conteudo vazio - 9 - {tars}")

                        if primeiras_quatro_letras == "full": # Se encontrar um documento Full
                            logging.debug(f"{hr_exec} - Localizado anexo Full em anexo - F - {tars}")
                            if attachment.FileName.endswith(".txt"):
                                file_path = f"{Pasta_despache}{hr_log}.txt"
                                try:
                                    attachment.SaveAsFile(file_path)
                                    with open(file_path, "r") as file:
                                        content = file.read()
                                        logging.debug(
                                            f"{hr_exec} - realizando a leitura do anexo {primeiras_quatro_letras} - 1 - {tars}")
                                        if erro in content:
                                            logging.debug(
                                                f"{hr_exec} - Erro encontrado no log, um ou mais documentos não foram copiados - M - {tars}")
                                        else:
                                            logging.debug(f"{hr_exec} - Log do hotel está OK - N - {tars}")
                                except ValueError as e:
                                    logging.debug(f"{hr_exec} - Erro ao copiar ao conteudo - Z - {tars}")
                                    logging.debug(f"{hr_exec} - Tentativa de leitura com metodo diferente - A - {tars}")
                                    try:
                                        with open(file_path, "r", encoding="utf-8") as file:
                                            content = file.read()
                                            if erro in content:
                                                logging.debug(
                                                    f"{hr_exec} - Erro encontrado no log, um ou mais documentos não foram copiados - K - {tars}")
                                            else:
                                                logging.debug(f"{hr_exec} - Log do hotel está OK - L - {tars}")
                                    except Exception as e:
                                        logging.debug(f"{hr_exec} - Tentativa com o novo metodo falhado - 8 - {tars}")
                            if primeiras_quatro_letras != "diff" and primeiras_quatro_letras!="full":
                                print(message.Subject)
                    else:
                        logging.debug(f"{hr_exec} - Documento {primeiras_quatro_letras} está com conteudo vazio - 9 - {tars}")



        # fazer o e-mail lido mover para alguma pasta dentro do outlook
print(f"{documents_folder}{dt_formatada}")
time.sleep(30)