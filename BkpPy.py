import logging;import win32com.client;import os;import winreg;import time
#Definindo a pasta documento

email=["TESTE 01","TESTE 2","TESTE 3","Americanaas"]


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



documents_folder = get_documents_folder()
print(documents_folder)

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

logging.debug(f"Data registrada - Hora Registrada - Atividade - Status ")
logging.debug(f"{hr_exec} - Inicio BkpPy - 0")

time.sleep(1)

#Conectar com o OUTLOOOK

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
logging.debug(f"{hr_exec} - Conexão com o Outlook realizada com sucesso - 2")
inbox = outlook.GetDefaultFolder(6)  # 6 é o código para a pasta de entrada
messages = inbox.Items
Pasta_despache="C:\\Windows\\Temp\\"
erro="Error"

for subject in email:
    message = messages.Find(f"[Subject] = '{subject}'")
    if message:
        logging.debug(f"{hr_exec} - Email com assunto '{subject}' encontrado com sucesso - 3")

        if message.Attachments.Count == 0:
            logging.debug(f"{hr_exec} - E-mail sem anexo 6")
        else:
                for attachment in message.Attachments:
                    attachment_name = attachment.FileName
                    primeiras_quatro_letras = attachment_name[:4]

                    if attachment.Size > 0:


                        if primeiras_quatro_letras == "diff":  # Se encontrar um documento Diff
                            logging.debug(f"{hr_exec} - Localizado anexo Diff em anexo - D")
                            if attachment.FileName.endswith(".txt"):
                                file_path = f"{Pasta_despache}{hr_log}.txt"
                                try:
                                    attachment.SaveAsFile(file_path)
                                    with open(file_path, "r") as file:
                                        content = file.read()
                                        logging.debug(
                                            f"{hr_exec} - realizando a leitura do anexo {primeiras_quatro_letras} - 1")
                                        if erro in content:
                                            logging.debug(
                                                f"{hr_exec} - Erro encontrado no log, um ou mais documentos não foram copiados - K")
                                        else:
                                            logging.debug(f"{hr_exec} - Log do hotel está OK - L")
                                except ValueError as e:
                                    logging.debug(f"{hr_exec} - Erro ao copiar ao conteudo - X")
                                    logging.debug(f"{hr_exec} - Tentativa de leitura com metodo diferente - B")
                                    try:
                                        with open(file_path, "r", encoding="utf-8") as file:
                                            content = file.read()
                                            if erro in content:
                                                logging.debug(
                                                    f"{hr_exec} - Erro encontrado no log, um ou mais documentos não foram copiados - K")
                                            else:
                                                logging.debug(f"{hr_exec} - Log do hotel está OK - L")
                                    except Exception as e:
                                        logging.debug(f"{hr_exec} - Tentativa com o novo metodo falhado - 8")
                            else:
                                logging.debug(
                                    f"{hr_exec} - Documento {primeiras_quatro_letras} está com conteudo vazio - 9")

                        if primeiras_quatro_letras == "full": # Se encontrar um documento Full
                            logging.debug(f"{hr_exec} - Localizado anexo Full em anexo - F")
                            if attachment.FileName.endswith(".txt"):
                                file_path = f"{Pasta_despache}{hr_log}.txt"
                                try:
                                    attachment.SaveAsFile(file_path)
                                    with open(file_path, "r") as file:
                                        content = file.read()
                                        logging.debug(
                                            f"{hr_exec} - realizando a leitura do anexo {primeiras_quatro_letras} - 1")
                                        if erro in content:
                                            logging.debug(
                                                f"{hr_exec} - Erro encontrado no log, um ou mais documentos não foram copiados - M")
                                        else:
                                            logging.debug(f"{hr_exec} - Log do hotel está OK - N")
                                except ValueError as e:
                                    logging.debug(f"{hr_exec} - Erro ao copiar ao conteudo - Z")
                                    logging.debug(f"{hr_exec} - Tentativa de leitura com metodo diferente - A")
                                    try:
                                        with open(file_path, "r", encoding="utf-8") as file:
                                            content = file.read()
                                            if erro in content:
                                                logging.debug(
                                                    f"{hr_exec} - Erro encontrado no log, um ou mais documentos não foram copiados - K")
                                            else:
                                                logging.debug(f"{hr_exec} - Log do hotel está OK - L")
                                    except Exception as e:
                                        logging.debug(f"{hr_exec} - Tentativa com o novo metodo falhado - 8")
                    else:
                        logging.debug(f"{hr_exec} - Documento {primeiras_quatro_letras} está com conteudo vazio - 9")

    else:
        logging.debug(f"{hr_exec} - Email com assunto '{subject}' não encontrado - 4")

        # fazer o e-mail lido mover para alguma pasta dentro do outlook
