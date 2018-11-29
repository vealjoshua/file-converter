import win32com.client
import win32com
import os
import sys
import matplotlib.pyplot as plt
import matplotlib.image as mpimg

f = open("testfile.docx","w+")

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
account = win32com.client.Dispatch("Outlook.Application").Session.Accounts[0]

def emailleri_al(folder):
    messages = folder.Items
    a=len(messages)
    if a>0:
        for message2 in messages:
             try:
                sender = message2.SenderEmailAddress
                if sender != "":
                    print(sender, file=f)
                    print(message2.Subject, file=f)
                    print(message2.Body, file=f)
                    print(message2.Attachments[0])
             except:
                print("Error")
                print(account.DeliveryStore.DisplayName)
                pass

             try:
                message2.Save
                message2.Close(0)
             except:
                 pass

global inbox
inbox = outlook.Folders(account.DeliveryStore.DisplayName)
print("****Account Name**********************************",file=f)
print(account.DisplayName,file=f)
print(account.DisplayName)

print("***************************************************",file=f)
folder = inbox.Folders['TechShorts']
print("****Folder Name**********************************", file=f)
print(folder, file=f)
print("*************************************************", file=f)
emailleri_al(folder)

print("Finished Succesfully")