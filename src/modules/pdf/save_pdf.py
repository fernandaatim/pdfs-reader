import pythoncom
import win32com.client
import os
import asyncio
from datetime import datetime

def connect_outlook(account_name=None):
    try:
        pythoncom.CoInitialize()  
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        if account_name:
            account = get_account(outlook, account_name)
            if account:
                return account
            else:
                raise Exception(f"Conta '{account_name}' não encontrada.")
        else:
            return outlook
    except Exception as e:
        print(f"Erro ao conectar ao Outlook: {e}")
        raise

def get_account(outlook, account_name):
    try:
        for account in outlook.Folders:
            if account.Name.lower() == account_name.lower():
                return account
        print(f"Conta '{account_name}' não encontrada.")
        return None
    except Exception as e:
        print(f"Erro ao acessar contas: {e}")
        raise

def get_folder(outlook, folder_name):
    try:
        for account in outlook.Folders:
            folder = search_folder(account.Folders, folder_name)
            if folder:
                return folder
        print(f"Pasta '{folder_name}' não encontrada em nenhuma conta do Outlook.")
        return None
    except Exception as e:
        print(f"Erro ao acessar a pasta: {e}")
        raise

def search_folder(folders, folder_name):
    for folder in folders:
        if folder.Name.lower() == folder_name.lower():
            return folder
        subfolder = search_folder(folder.Folders, folder_name)
        if subfolder:
            return subfolder
    return None

def get_documents_folder():
    return os.path.join(os.environ['USERPROFILE'], 'Documents') if os.name == 'nt' else os.path.join(os.environ['HOME'], 'Documents')

async def save_pdfs_from_folder():
    await asyncio.to_thread(_save_pdfs_from_folder)

def create_save_folder(email_date=None):
    if email_date is None:
        email_date = datetime.now()
    month_year = email_date.strftime("%B-%Y")
    save_folder = os.path.join(get_documents_folder(), f"pdfs-{month_year}")
    save_folder_network = ''
    save_folder_network = os.path.join(save_folder_network, f"pdfs-{month_year}")
    os.makedirs(save_folder, exist_ok=True)
    os.makedirs(save_folder_network, exist_ok=True)
    return save_folder, save_folder_network

def _save_pdfs_from_folder():
    outlook = connect_outlook()
    folder = get_folder(outlook, "pdfs_medicoes")
    if not folder:
        raise Exception("Pasta 'pdfs_medicoes' não encontrada.")
    try:
        messages = folder.Items
        messages.Sort("[ReceivedTime]", True)
        for message in messages:
            try:
                received_time = getattr(message, "ReceivedTime", None)
                if received_time is None:
                    received_time = datetime.now()
                save_folder, save_folder_network = create_save_folder(received_time)
                saved_files_local = set(os.listdir(save_folder))
                saved_files_network = set(os.listdir(save_folder_network))
                attachments = message.Attachments
                for attachment in attachments:
                    if "Medições" in attachment.FileName and attachment.FileName.endswith(".pdf"):
                        file_path_local = os.path.join(save_folder, attachment.FileName)
                        file_path_network = os.path.join(save_folder_network, attachment.FileName)
                        if attachment.FileName in saved_files_local and attachment.FileName in saved_files_network:
                            print(f"PDF já existe em ambos os locais, ignorando: {attachment.FileName}")
                        else:
                            if attachment.FileName not in saved_files_local:
                                attachment.SaveAsFile(file_path_local)
                                print(f"Novo PDF salvo localmente: {attachment.FileName}")
                                saved_files_local.add(attachment.FileName)
                            if attachment.FileName not in saved_files_network:
                                attachment.SaveAsFile(file_path_network)
                                print(f"Novo PDF salvo na rede: {attachment.FileName}")
                                saved_files_network.add(attachment.FileName)
            except Exception as e:
                print(f"Erro ao processar e-mail: {e}")
    except Exception as e:
        print(f"Erro ao acessar mensagens: {e}")
