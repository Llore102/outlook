# Automatic Outlook Email Downloader

<div align="center" style="width: 200px;">
  <img alt="GIF" src="https://media2.giphy.com/media/v1.Y2lkPTc5MGI3NjExYWhsODZ0Y2tpcW1zNHFjemJ6eTFvcWxmaGNtYm4wYXk3OHAxeTQzdSZlcD12MV9pbnRlcm5hbF9naWZfYnlfaWQmY3Q9Zw/tQIrvVQXJCttlXaTAD/giphy.gif" width="50%"/>
</div>

----------------

## Overview

Este script de Python te permite capturar y descargar correos electrónicos y sus archivos adjuntos desde una cuenta de correo electrónico de Outlook, basándose en prefijos específicos en el asunto. Utiliza la biblioteca `win32com.client` para interactuar con la aplicación Outlook.

Se han proporcionado dos versiones de scripts para interactuar con las versiones antigua y actual de Outlook.

### Tecnologías Utilizadas
![Python](https://www.vectorlogo.zone/logos/python/python-ar21.svg) ![Microsoft](https://www.vectorlogo.zone/logos/microsoft/microsoft-ar21.svg)


## Requirements

1. Instala las bibliotecas requeridas:

   ```bash
   pip install -r requirements.txt


## Usage

1. Configurar parametros del script

    * Proporcionar el correo electronico del cual se extraeran los archivos account_to_capture = "ejemplo@outlook.es" --> "Script outlook.py"
    * Proporcionar el correo electronico del cual se extraeran los archivos inbox = outlook.Folders['ejemplo@outlook.es'].Folders['Inbox'] --> "Script outlook3.8.py"
    * La bandeja de entrada se debe configurar segun el idioma en el que se encuentre (Ingles "Inbox") | (Español "Bandeja de entrada")
    * Proporcionar una lista de prefijos, porla cual se buscaran los mensajes a descargar ['Reporte', 'Quejas', 'Informacion']
    


2. Ejecutar Script para la version requerida:
 
   * outlook.py ultima version de outlok
   * outlook3.8.py Version antigua de outlook

Los correos y adjuntos se guardaran en una carpeta creada con la fecha. 


