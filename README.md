# Reminder Email Sending and Email Retrieval with Attachments Script

This repository contains a Python script that handles sending reminders and retrieving emails with their attachments using the win32com.client, pandas, os, openpyxl, and datetime libraries.

## Functionality
The script consists of two main functions:

- Sending reminders
The `enviar_recordatorios()` function reads data from an Excel file called `Datos.xlsx` and sends reminder emails to the responsible individuals. For each record in the file, it extracts the responsible person's name, email, message, and deadline date. It then calculates the number of days remaining until the deadline and sends an email using Outlook with the corresponding information.

- Retrieving emails and attachments
The `trae_correos_y_adjuntos()` function retrieves the latest three emails from the Outlook inbox and saves the relevant information in an Excel file called `correos.xlsx`. For each email, it extracts the sender, date, and subject. Additionally, it checks for attachments in PDF, DOC, or DOCX format and saves the list of attachments along with links to access them.

## Requirements
Before running the script, make sure you have the following requirements:

- Python 3.x installed on your system.
- The `win32com`, `pandas`, and `openpyxl` libraries installed. You can install them using the following command:
   `pip install pywin32 pandas openpyxl`

## Usage Instructions
Follow the steps below to use this script:

- Clone this repository or download the source code file.

- Make sure you have the `Datos.xlsx` file with responsible person data in the same directory as the script.

- Run the script in Python. Reminders will be sent via email, and email and attachment information will be saved in the `correos.xlsx` file.

**Note**: You may be prompted to log in to Outlook the first time you run the script to allow access to the inbox and send emails.

## Considerations:
- Email and attachment retrieval is limited to the latest three emails in the inbox. You can adjust the number of retrieved emails by modifying the value of 3 in the `for` loop inside the `trae_correos_y_adjuntos()` function.

- The script saves the attachments in the same location as the script. You can modify the save location by changing the value of `os.getcwd()` in the corresponding line.
