import win32com.client

ol=win32com.client.Dispatch("outlook.application")

olmailitem=0x0 #size of the new email

newmail=ol.CreateItem(olmailitem)

newmail.Subject= 'Testing Mail'
newmail.To='insert-email-here'
# newmail.CC='insert-email-here'
newmail.Body= 'Hello, this is a test email to showcase how to send emails from Python and Outlook.'
# attach='C:\\Users\\admin\\Desktop\\Python\\Sample.xlsx'
# newmail.Attachments.Add(attach)
# To display the mail before sending it
# newmail.Display() 
newmail.Send()