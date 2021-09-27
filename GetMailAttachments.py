from imaplib import IMAP4_SSL
import email

attachmentsPath = 'Attachments\\'
# Outlook Credentials
login = "your email here"
pwd = "your password here"

# Connect to outlook
objConn = IMAP4_SSL('outlook.office365.com')
objConn.login(login, pwd)

# Select Inbox folder (you can pass any folder)
objConn.select(mailbox='Inbox', readonly=True)
# Get mailsId in selected folder
response, mailsId = objConn.search(None, 'All')

# Loop through all mails in selected folder
for mailId in mailsId[0].split():
    # Get a mail by id
    response, data = objConn.fetch(mailId, '(RFC822)')
    # Decode mail content
    emailContent = email.message_from_string(data[0][1].decode('utf-8'))
    # Check for a pattern at the subject
    if '[Anexo]' in emailContent.get('Subject'):
        # Loop throught mail content parts
        for part in emailContent.walk():
            # Leaves loop if content maintype is multipart
            if part.get_content_maintype() == 'multipart':
                continue
            # Leaves loop if Content-Disposition is None
            if part.get('Content-Disposition') is None:
                continue
            # Get attachment file name
            fileName = part.get_filename()
            print(fileName)

            # Save file
            with open(attachmentsPath + fileName, 'wb') as file:
                file.write(part.get_payload(decode=True))
