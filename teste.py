from PySimpleGUI import PySimpleGUI as sg
import imaplib
import email
from email.header import decode_header
import webbrowser
import os, time, keyboard
from time import sleep
import email, smtplib, ssl, random, pdfkit
from email import encoders
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.message import Message
from email.mime.base import MIMEBase
import pdfkit
from bs4 import BeautifulSoup


#Layout
sg.theme('DarkGreen1')
layout = [
    [sg.Text('Email de consulta'),sg.Input(key='email')],
    [sg.Text('Senha'),sg.Input(key='codigo')],
    [sg.Button('Iniciar')]
]
#Window
window = sg.Window('Digite seus dados', layout)
event, values = window.read()
#Events
while True:
    eventos, valores = window.read()
    if eventos == sg.WINDOW_CLOSED:
        quit()
    if eventos == 'Iniciar':
        conta = values['email']
        senha = values['codigo']
        sleep(1)
        window.close()
        break

pasta = 'Pedidos'
lista_teste = ['vendas@caneparts.com.br','comercial@caneparts.com.br']
#folders for consulting old emails
folder_path = os.getcwd()
folder_path2 = folder_path+"\html_emails"
##############################################################################################

# Looks for unread email, and if it finds some, look inside its body and see if it's composed by html, if yes converts it into seperate PDF files
mail = imaplib.IMAP4_SSL("outlook.office365.com")
mail.login(conta, senha)

# Select the inbox folder
mail.select('Pedidos')

# Search for unread emails
status, email_ids = mail.search(None, '(UNSEEN)')
email_ids = email_ids[0].split()
if not email_ids:
    print('Nenhum novo pedido')
# Create a folder to store the HTML files
os.makedirs("html_emails", exist_ok=True)
teste = False
# Loop through each email
for email_id in email_ids:
    # Fetch the email message
    status, msg = mail.fetch(email_id, "(RFC822)")
    msg = msg[0][1]

    # Parse the email message
    email_message = email.message_from_bytes(msg)

    # Check if the email has an HTML body
    if email_message.is_multipart():
        for part in email_message.walk():
            content_type = part.get_content_type()
            if content_type == "text/html":
                html_body = part.get_payload(decode=True)
    
                with open(f"html_emails/email_{email_id}.html", "wb") as file:
                    file.write(html_body)
    else:
        content_type = email_message.get_content_type()
        if content_type == "text/html":
            html_body = email_message.get_payload(decode=True)
            with open(f"html_emails/email_{email_id}.html", "wb") as file:
                file.write(html_body)
                teste = True
if teste == True:
# Specify the path to the folder containing the HTML files
    html_path = "html_emails"

# Loop through each file in the folder
    for filename in os.listdir(html_path):
    # Check if the file is an HTML file
        if filename.endswith(".html"):

            # Load HTML file
            with open(os.path.join(html_path, filename), "r") as f:
                html = f.read()

            # Parse HTML with BeautifulSoup
            soup = BeautifulSoup(html, "html.parser")

            # Find all text within the body tag
            body_text = soup.body.get_text()

            # Check if "python" is in the list of words
            if "Keite" in body_text:
                lista_teste = ["comercial@caneparts.com.br"]
            elif "keite" in body_text:
                lista_teste = ["comercial@caneparts.com.br"]
            elif "Alexandre" in body_text:
                lista_teste = ["vendas@caneparts.com.br"]
            elif "alexandre" in body_text:
                lista_teste = ["vendas@caneparts.com.br"]
            else:
                break


    # Create the PDF file path
            pdf_file_path = os.path.splitext(filename)[0] + ".pdf"

        # Convert the HTML file to PDF
            pdfkit.from_file(os.path.join(html_path, filename), pdf_file_path)
            pdfs = []

            # get the current time
            current_time = time.time()

            # find all pdf files in the folder that were modified 1 minute ago
            pdf_files = [f for f in os.listdir(".") if f.endswith(".pdf") and current_time - os.path.getmtime(f) <= 60]
    	    
            # put each pdf file into a separate variable
            for i, pdf in enumerate(pdf_files):
                globals()[f"pdf_{i+1}"] = pdf

            # check if any pdf files were found
            if pdf_files:
                print("The following pdf files were found:")
                for i, pdf in enumerate(pdf_files):
                    print(f"{pdf}")
                    pdfs = pdf_files

    ####### Gets all the new pdf files and store them inside the body of an email and send it to a random receiver

            #determins which person will receive the email
            recebedor = str(random.choice(lista_teste))

            #credentials
            port = '587'
            server = 'smtp.outlook.com'
            login = conta
            senha1 = senha
            subject = "Novo pedido de orçamento | Caneparts"
            body = ""
            sender = conta
            receiver = recebedor
            msg = MIMEMultipart()
            msg["Subject"] = subject
            msg["From"] = sender
            msg["To"] = receiver

            #creates a body for the email
            msg.attach(MIMEText(body, "plain"))
            # Transform the list into string variables for each item in the list
            for i, item in enumerate(pdfs):
                globals()[f"item_{i+1}"] = item

                # Open PDF file in binary mode
                with open(item, "rb") as attachment:
                    # Add file as application/octet-stream
                    # Email client can usually download this automatically as attachment
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())

                # Encode file in ASCII characters to send by email    
                encoders.encode_base64(part)

                # Add header as key/value pair to attachment part
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename= {item}",
                )

                # Add attachment to message and convert message to string
                msg.attach(part)

                #sends final email to receiver
                server_run = smtplib.SMTP('smtp.outlook.com')
                server_run.connect('smtp.outlook.com', 587)
                server_run.starttls()
                server_run.login(login, senha1)
                server_run.send_message(msg)

                sleep(3) 

                # Checks for "old" orders and deletes all their file
                    # Get the current time
                current_time = time.time()
                    # Loop through each file in the folder
                for filename in os.listdir(folder_path):
                            # Check if the file is a PDF file
                        if filename.endswith(".pdf"):
                            # Get the file's modification time
                            file_path = os.path.join(folder_path, filename)
                            modification_time = os.path.getmtime(file_path)

                            # Check if the file is older than 2 seconds
                            if current_time - modification_time > 2:
                                # Delete the file
                                os.remove(file_path)

                print('Sent')
                print("____________________________________________")
    print('\n','Todos os pedidos foram enviados com sucesso') 

################################################################################################
current_time = time.time()
#Loop through each file in the folder, erasing "old" emails
for filename in os.listdir(folder_path2):
    # Check if the file is a PDF file
    if filename.endswith(".html"):
    # Get the file's modification time
        file_path = os.path.join(folder_path2, filename)
        modification_time = os.path.getmtime(file_path)

    # Check if the file is older than 2 seconds
    if current_time - modification_time > 1:
    # Delete the file
        os.remove(file_path)
#####################################################    
    sleep(5)

mail.close()
mail.logout()