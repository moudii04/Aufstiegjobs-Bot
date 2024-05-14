import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import csv
from motiv import generate_cover_letter
from conv import convert_docx_to_pdf
from variables import MY_NAME, MY_EMAIL, SECRET, FILENAME, STUDIENGANG

counter = 1

gmail_server = "smtp.gmail.com"
gmail_port = 587

resume_files = ['C:\\Users\\moudi\\Desktop\\Naim\\Lebenslauf.pdf',
                'C:\\Users\\moudi\\Desktop\\Naim\\Zeugnisse.pdf',
                'Anschreiben.pdf']


# Starting connection
my_server = smtplib.SMTP(gmail_server, gmail_port)
my_server.ehlo()
my_server.starttls()

# Login with your email and password
my_server.login(MY_EMAIL, SECRET)

with open(FILENAME, encoding="utf-8") as csv_file:
    jobs = csv.reader(csv_file)
    next(jobs)  # Skip header row

    for name, address, email in jobs:
        generate_cover_letter(name, address)
        convert_docx_to_pdf(os.getcwd(), os.getcwd())

        # Create a new message object for each email
        message = MIMEMultipart("alternative")
        message["Subject"] = f"Bewerbung um einen Duales Studium Platz ({STUDIENGANG})"

        text_content = """Sehr geehrter {name}, 
        
im Anhang finden Sie meine Bewerbungsunterlagen für die ausgeschriebene Stelle als dualer Student ({STUDIENGANG}). Ich freue mich über die Möglichkeit, Teil Ihres Teams zu werden und meine Motivation sowie Eignung für die Ausbildung näher zu erläutern.

Ich habe bereits mit Ihrer Partnerschule (DHFPG) gesprochen und ich erfülle all die Voraussetzungen.

Für Fragen oder die Bereitstellung weiterer Unterlagen stehe ich Ihnen gerne zur Verfügung.

Mit freundlichen Grüßen,

{MY_NAME}""".format(MY_NAME=MY_NAME, name=name, STUDIENGANG=STUDIENGANG)

        text_content = text_content.encode('utf-8')

        message.attach(MIMEText(text_content, 'plain', 'utf-8'))

        # Attach resume file
        for resume_file in resume_files:
            with open(resume_file, 'rb') as f:
                file = MIMEApplication(
                    f.read(),
                    name=os.path.basename(resume_file)
                )
                file['Content-Disposition'] = f'attachment; filename="{os.path.basename(resume_file)}"'
                message.attach(file)

        my_server.sendmail(
            from_addr=MY_EMAIL,
            to_addrs=email,
            msg=message.as_string()
        )
        print(f"{counter}/210")
        counter += 1

my_server.quit()
