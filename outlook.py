opt1:
import win32com.client
import os
get_path = os.getcwd()
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
message2 = messages.GetLast()
subject = message2.Subject
body = message2.body
sender = message2.Sender
attachments = message2.Attachments
for m in messages:
    if m.Subject == "Test Mail":
        for x in message2.Attachments:
            x.SaveASFile(os.path.join(get_path,x.FileName))
            print "successfully downloaded attachments"
            
            
#opt 2:
from exchangelib import DELEGATE, Account, Credentials

credentials = Credentials(
    username='MYWINDOMAIN\\myusername',  # Or myusername@example.com for O365
    password='topsecret'
)
account = Account(
    primary_smtp_address='john@example.com', 
    credentials=credentials, 
    autodiscover=True, 
    access_type=DELEGATE
)
# Print first 100 inbox messages in reverse order
for item in account.inbox.all().order_by('-datetime_received')[:100]:
    print(item.subject, item.body, item.attachments)
