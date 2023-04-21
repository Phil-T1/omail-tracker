import win32com.client
import os
from datetime import datetime, timedelta

# Outlook app
outlook = win32com.client.Dispatch('outlook.application')

# Inbox
inbox = outlook.GetNamespace('MAPI').GetDefaultFolder(6)

# Filter (restriction) parameters
start_dt = datetime.now() - timedelta(days=1)
end_dt = datetime.now()

restriction = (f'''
[ReceivedTime] >= '{start_dt.strftime("%d/%m/%Y %H:%M %p")}'
''')

# Get filtered messages
messages = inbox.Items.Restrict(restriction)
for message in list(messages):
    print(message.subject)
print(len(messages))
# sender
# to
# cc
# bcc
# bodyformat
# body
# htmlbody
# subject
# display()
# receivedtime
# senton