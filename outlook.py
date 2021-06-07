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