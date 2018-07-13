import win32com.client
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd

class OutlookLib:
        
    def __init__(self, settings={}):
        self.settings = settings
        
    def get_messages(self, user, folder="Inbox", match_field="all", match="all"):      
        outlook = win32com.client.Dispatch("Outlook.Application")
        myfolder = outlook.GetNamespace("MAPI").Folders[user] 
        inbox = myfolder.Folders[folder] # Inbox
        if match_field == "all" and match =="all":
            return inbox.Items
        else:
            messages = []
            for msg in inbox.Items:
                try:
                    if match_field == "Sender":
                        if msg.SenderName.find(match) >= 0:
                            messages.append(msg)
                    elif match_field == "Subject":
                        if msg.Subject.find(match) >= 0:
                            messages.append(msg)
                    elif match_field == "Body":
                        if msg.Body.find(match) >= 0:
                            messages.append(msg)
                    #print msg.To
                    # msg.Attachments
                    # a = item.Attachments.Item(i)
                    # a.FileName
                except:
                    pass
            return messages
        
    def get_body(self, msg):
        return msg.Body
    
    def get_subject(self, msg):
        return msg.Subject
    
    def get_sender(self, msg):
        return msg.SenderName
    
    def get_recipient(self, msg):
        return msg.To
    
    def get_attachments(self, msg):
        return msg.Attachments

#main
outlook = OutlookLib()

messages = outlook.get_messages('rj.oosthuizen@acctech.biz')
wordsPerMessage = []
messageCount = []
counter = 1
senders = []
for msg in messages:
    print(msg.Subject)
    print(msg.Body)
    print(len(msg.body))
    wordsPerMessage.append(len(msg.body))
    messageCount.append(counter)
    counter += 1
    senders.append(msg.SenderName)
# newList = pd.DataFrame(messages)
# print(newList)

print(len(messages))
print(wordsPerMessage)
print(senders)
shortSenders = []
for sent in senders:
    space = sent.find(" ")
    #print(space)
    if(space != -1):
      shortSenders.append(sent[:space + 1])
    else:
      shortSenders.append(sent[:7])
print(shortSenders)

#plt.plot(messageCount, wordsPerMessage)
plt.plot(shortSenders, wordsPerMessage)
plt.title('Words per email in my inbox')
#plt.xlabel(senders, fontsize=8)
plt.show()