# Aufstiegjobs Bot
## Description
The Aufstiegjobs Bot is designed to facilitate sending emails to applicants from the Aufstiegjobs website. Below are the key files and their purposes:

- **variables.py**: Contains the information that needs to be customized.
- **motiv.py**: Generates the motivation letter. You can modify the cover letter text here.
- **conv.py**: Converts the cover letter to PDF (from .docx to .pdf).
- **main.py**: The main executable file.

## How to Use
Customize variables: Open and edit the "variables.py" file to include the necessary information. For obtaining secret keys, follow these steps: Obtaining API Key for Gmail API.

Change FILENAME: In **"variables.py"**, modify the FILENAME variable to either "Herr.csv" or "Frau.csv" depending on the recipient's gender.

### Note: By default, in "main.py" and "motiv.py", the salutation is set to "Sehr geehrter ..." . When changing to "Frau.csv", ensure to change it to "Sehr geehrte".

Execute main.py: Run the "main.py" file to initiate the email sending process.

