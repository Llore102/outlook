import win32com.client
import os
import datetime

def save_attachments(item, attachments_folder):
    for attachment in item.Attachments:
        attachment.SaveAsFile(os.path.join(attachments_folder, attachment.FileName))
        print(f"Archivo adjunto guardado: {attachment.FileName}")

def capture_email(subject):
    outlook_app = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook_app.GetDefaultFolder(6)  # 6 corresponds to the Inbox folder

    for item in inbox.Items:
        if item.Subject == subject:
            
            current_datetime = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")# Crear un nombre de carpeta basado en la fecha y hora actual
            folder_name = f"{subject}_{current_datetime}"

            
            folder_path = os.path.join('C:/Users/llore/BackEnd/outlook', folder_name)# Crear la ruta completa de la carpeta

            
            os.makedirs(folder_path)# Crear la carpeta para guardar el contenido del correo

           
            email_file_path = os.path.join(folder_path, f"{subject}_{current_datetime}.txt") # Guardar el contenido del correo en un archivo
            with open(email_file_path, 'w', encoding='utf-8') as file:
                file.write(f"Asunto: {item.Subject}\n")
                file.write(f"De: {item.SenderName}\n")
                file.write(f"Fecha: {item.ReceivedTime}\n\n")
                file.write(item.Body)

            print(f"Correo electrónico con asunto '{subject}' capturado y contenido guardado en '{email_file_path}'.")

            
            attachments_folder = os.path.join(folder_path, 'ArchivosAdjuntos')# Guardar archivos adjuntos
            os.makedirs(attachments_folder)
            save_attachments(item, attachments_folder)

            break

if __name__ == "__main__":
    subject_to_capture = "Informe 2024-02-23"
    capture_email(subject_to_capture)

