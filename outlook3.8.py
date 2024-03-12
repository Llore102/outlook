import os
import win32com.client
import datetime

prefijos_a_buscar = ['Santander', 'Robot', 'OtroPrefijo']
carpeta_local = 'C:/Users/llore/BackEnd/outlook'

if not os.path.exists(carpeta_local):
    os.makedirs(carpeta_local)

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.Folders['ejemplo@outlook.es'].Folders['Inbox']
messages = inbox.Items

today = datetime.date.today()
todos_correos_guardados = False


day_folder = os.path.join(carpeta_local, today.strftime('%Y%m%d'))
if not os.path.exists(day_folder):
    os.makedirs(day_folder)

for message in messages:
    if message.Class == 43 and message.Senton.date() == today:
        for prefijo in prefijos_a_buscar:
            if prefijo in message.Subject:
                print(f"Subject: {message.Subject}, Received: {message.Senton}")

                eml_filename = os.path.join(day_folder, f"{message.Subject}_{message.Senton.strftime('%Y%m%d%H%M%S')}.eml")
                message.SaveAs(eml_filename)

                print(f"Correo guardado en: {eml_filename}")

                # Guardar archivos adjuntos
                for attachment in message.Attachments:
                    attach_filename = os.path.join(day_folder, attachment.FileName)
                    attachment.SaveAsFile(attach_filename)
                    print(f"Adjunto guardado en: {attach_filename}")

                todos_correos_guardados = True

if todos_correos_guardados:
    print("Todos los correos han sido guardados.")
else:
    print("No se encontraron correos para guardar.")
