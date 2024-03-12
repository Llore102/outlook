

import win32com.client
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import datetime

def download_attachment(attachment, download_folder):
    attachment.SaveAsFile(os.path.join(download_folder, attachment.FileName))

def attach_file(msg, file_path):
    attachment = MIMEBase('application', 'octet-stream')
    with open(file_path, 'rb') as file:
        attachment.set_payload(file.read())
    encoders.encode_base64(attachment)
    attachment.add_header('Content-Disposition', f'attachment; filename={os.path.basename(file_path)}')
    msg.attach(attachment)

def create_eml(item, eml_file_path, download_folder):
    msg = MIMEMultipart()
    msg['From'] = item.SenderEmailAddress
    msg['To'] = item.To
    msg['Subject'] = item.Subject
    msg['Date'] = item.ReceivedTime.strftime('%a, %d %b %Y %H:%M:%S %z')

    msg.attach(MIMEText(item.Body, 'plain'))

    for attachment in item.Attachments:
        download_attachment(attachment, download_folder)

    for attachment in item.Attachments:
        attach_file(msg, os.path.join(download_folder, attachment.FileName))

    with open(eml_file_path, 'w', encoding='utf-8') as eml_file:
        eml_file.write(msg.as_string())

    # print(f"Correo guardado como EML: {eml_file_path}")

def capture_emails(account, subject_prefix_list, system_date_format='%Y-%m-%d'):
    outlook_app = win32com.client.Dispatch("Outlook.Application")
    outlook_namespace = outlook_app.GetNamespace("MAPI")


    inbox = outlook_namespace.Folders(account).Folders("Bandeja de entrada") ##! Inbox

    current_date = datetime.datetime.now().strftime(system_date_format)

    for item in inbox.Items:
        subject = item.Subject.lower()
        received_date = item.ReceivedTime.strftime(system_date_format)

        # Verificar si el asunto comienza con alguno de los prefijos y la fecha coincide con la del sistema
        if any(subject.startswith(prefix.lower()) for prefix in subject_prefix_list) and received_date == current_date:
            received_time = item.ReceivedTime.strftime('%Y%m%d_%H%M%S')
            folder_name = f"{subject}_{received_time}"
            folder_path = os.path.join('C:/Users/llore/BackEnd/outlook', folder_name)

            counter = 1
            while os.path.exists(folder_path):
                folder_name = f"{subject}_{received_time}_{counter}"
                folder_path = os.path.join('C:/Users/llore/BackEnd/outlook', folder_name)
                counter += 1

            os.makedirs(folder_path)

            eml_file_path = os.path.join(folder_path, f"{subject}_{received_time}.eml")
            download_folder = os.path.join(folder_path, 'DescargasAdjuntos')
            os.makedirs(download_folder)

            create_eml(item, eml_file_path, download_folder)

            print(f"Correo electrónico con asunto que comienza por '{subject}' y recibido en la fecha '{current_date}' capturado y contenido guardado en '{folder_path}'.")

if __name__ == "__main__":
    account_to_capture = "lloreda102@outlook.es"  
    subject_prefix_list_to_capture = ["Santander", "Robot"]
    capture_emails(account_to_capture, subject_prefix_list_to_capture)












