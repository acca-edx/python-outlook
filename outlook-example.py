### Add code libraries we are going to use with import statements.
# win32com is the package used to call Outlook.
###

import win32com.client

### Create an outlook Email.
# win32com.client.Dispatch = Tell windows to open outlook.
# CreateItem(0) = Create an email for sending.
###

o = win32com.client.Dispatch("Outlook.Application")
msg = o.CreateItem(0)

### Format required email information.
# Now we format the email with all the standard information.
# to = List of addresses we are sending to.
# Subject = The Subject Line of the email.
# Body = The body of the email in plain text.
###

msg.to = "email@example.com"
msg.Subject = "This is a test"
msg.Body = "Hi, this is your email"

### Adding an attachment.
# Define the attachment variable.
#   The r before the "filename" means plain text.
# Attachments.Add = Add the file from your local computer.
###

msg.Attachments.Add(r'C:\Users\name\Downloads\data_report.csv')

### Send the message
# Now that the message is ready, it is time to send it.
###

msg.Send()
