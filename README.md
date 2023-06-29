# Script de Envío de Recordatorios por correo y descargar archivos adjuntos

Este repositorio contiene un script en Python que se encarga de enviar recordatorios y recuperar correos y adjuntos utilizando las bibliotecas win32com.client, pandas, os, openpyxl y datetime.

## Funcionalidad
El script consta de dos funciones principales:

- Enviar recordatorios
La función enviar_recordatorios() lee los datos de un archivo Excel llamado Datos.xlsx y envía recordatorios por correo electrónico a los responsables. Para cada registro en el archivo, se extraen el nombre del responsable, el correo electrónico, el mensaje y la fecha límite. Luego, se calcula la cantidad de días restantes hasta la fecha límite y se envía un correo electrónico utilizando Outlook con la información correspondiente.

- Recuperar correos y adjuntos
La función trae_correos_y_adjuntos() recupera los últimos tres correos electrónicos de la bandeja de entrada de Outlook y guarda la información relevante en un archivo Excel llamado correos.xlsx. Para cada correo, se extrae el remitente, la fecha y el asunto. Además, se verifica si hay adjuntos en formato PDF, DOC o DOCX, y se guarda la lista de adjuntos junto con enlaces para acceder a ellos.

## Requisitos
Antes de ejecutar el script, asegúrate de tener los siguientes requisitos:

- Python 3.x instalado en tu sistema.
- Las bibliotecas win32com, pandas, openpyxl instaladas. Puedes instalarlas mediante el siguiente comando:
  `pip install pywin32 pandas openpyxl`

## Instrucciones de uso
Sigue los pasos a continuación para utilizar este script:

- Clona este repositorio o descarga el archivo del código fuente.

- Asegúrate de tener el archivo Datos.xlsx con los datos de los responsables en el mismo directorio que el script.

- Ejecuta el script en Python. Los recordatorios serán enviados por correo electrónico y la información de los correos y adjuntos será guardada en el archivo correos.xlsx.

Nota: Es posible que se te solicite iniciar sesión en Outlook la primera vez que ejecutes el script para permitir el acceso a la bandeja de entrada y enviar correos electrónicos.

## Consideraciones:
- La recuperación de correos y archivos adjuntos se limita a los últimos tres correos de la bandeja de entrada. Puedes ajustar el número de correos recuperados modificando el valor de 3 en el bucle for dentro de la función trae_correos_y_adjuntos().

- El script guarda los archivos adjuntos en la misma ubicación que el script. Puedes modificar la ubicación de guardado cambiando el valor de os.getcwd() en la línea correspondiente.
